import logging
import os
from itertools import product
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

MISE_DEFAUT = 5000
SEUIL_BENEFICE = 1.015
historique = {}

class AlboraEngine:
    @staticmethod
    def detect_arbitrage(cotes):
        s = sum(1 / c for c in cotes)
        if s < 1:
            profit = round((1 / s - 1) * 100, 2)
            stakes = {f"cote_{i+1}": round((1 / c / s) * 100, 2) for i, c in enumerate(cotes)}
            return {"arbitrage": True, "profit": profit, "stakes": stakes, "sum": round(s, 4)}
        return {"arbitrage": False, "sum": round(s, 4)}

    @staticmethod
    def combo_matchs(matchs, mise):
        labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        resultats = []
        seuil = mise * SEUIL_BENEFICE
        for combo in product(*[range(len(m)) for m in matchs]):
            cotes_sel = [matchs[i][combo[i]] for i in range(len(matchs))]
            noms = [f"{labels[i]}{combo[i]+1}" for i in range(len(matchs))]
            cote_totale = round(eval("*".join(str(c) for c in cotes_sel)), 4)
            gain = round(mise * cote_totale, 0)
            benefice_net = round(gain - mise, 0)
            benefice_pct = round((cote_totale - 1) * 100, 2)
            if gain >= seuil:
                resultats.append({
                    "noms": "+".join(noms),
                    "cotes": cotes_sel,
                    "cote_totale": cote_totale,
                    "gain": int(gain),
                    "benefice_net": int(benefice_net),
                    "benefice_pct": benefice_pct,
                })
        resultats.sort(key=lambda x: x["cote_totale"])
        return resultats


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "ALBORAA BOT - Arbitrage et Combinaisons\n"
        "Commandes disponibles:\n\n"
        "/arbitrage 1.21 3.02 3.10\n"
        "/combo 1.21 3.02 | 1.22 3.01\n"
        "/simuler 1.85 10000\n"
        "/bilan\n"
        "/effacer\n"
        "/export\n\n"
        "Devise: Franc CFA\n"
        "Mise defaut: 5000 FCFA\n"
        "Seuil benefice: +1.50% minimum"
    )
    await update.message.reply_text(msg)


async def arbitrage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) < 2 or len(args) > 3:
        await update.message.reply_text("Format: /arbitrage 1.21 3.02 3.10")
        return
    try:
        cotes = list(map(float, args))
        result = AlboraEngine.detect_arbitrage(cotes)
        if result["arbitrage"]:
            stakes_str = "\n".join([f"Cote {k[-1]} ({cotes[int(k[-1])-1]}) -> {v}%" for k, v in result["stakes"].items()])
            msg = f"ARBITRAGE DETECTE\nProfit: {result['profit']}%\nSomme: {result['sum']}\n\nRepartition:\n{stakes_str}"
            _save_historique(update.effective_user.id, "arbitrage", cotes, result)
        else:
            msg = f"Pas d'arbitrage. Somme: {result['sum']} (doit etre < 1.00)"
        await update.message.reply_text(msg)
    except ValueError:
        await update.message.reply_text("Cotes invalides. Exemple: /arbitrage 1.21 3.02 3.10")


async def combo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text("Format: /combo 1.21 3.02 | 1.22 3.01")
        return
    try:
        ligne = " ".join(context.args)
        tokens = ligne.split()
        mise = MISE_DEFAUT
        try:
            derniere = float(tokens[-1])
            if "|" not in tokens[-1] and derniere > 100:
                mise = derniere
                ligne = " ".join(tokens[:-1])
        except ValueError:
            pass
        matchs_raw = ligne.split("|")
        matchs = []
        for m in matchs_raw:
            cotes = [float(x) for x in m.strip().split() if x.strip()]
            if cotes:
                matchs.append(cotes)
        if len(matchs) < 2:
            await update.message.reply_text("Minimum 2 matchs separes par |")
            return
        if len(matchs) > 10:
            await update.message.reply_text("Maximum 10 matchs.")
            return
        resultats = AlboraEngine.combo_matchs(matchs, mise)
        nb_total = 1
        for m in matchs:
            nb_total *= len(m)
        if not resultats:
            msg = f"Aucune combinaison rentable.\n{nb_total} combinaisons analysees.\nMise: {int(mise)} FCFA"
        else:
            lignes = []
            for i, r in enumerate(resultats, 1):
                lignes.append(f"{i}. {r['noms']} = {r['cote_totale']} -> {r['gain']} FCFA (+{r['benefice_pct']}%)")
            msg = f"ALBORAA - {len(matchs)} matchs\nMise: {int(mise)} FCFA\n{len(resultats)} rentables / {nb_total} total\n\n" + "\n".join(lignes)
        _save_historique(update.effective_user.id, "combo", matchs, resultats)
        await update.message.reply_text(msg)
    except Exception:
        await update.message.reply_text("Erreur de format. Exemple: /combo 1.21 3.02 | 1.22 3.01")


async def simuler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) != 2:
        await update.message.reply_text("Format: /simuler 1.85 10000")
        return
    try:
        cote = float(args[0])
        mise = float(args[1])
        gain = round(mise * cote, 0)
        benefice = round(gain - mise, 0)
        pct = round((cote - 1) * 100, 2)
        msg = f"SIMULATION\nCote: {cote}\nMise: {int(mise)} FCFA\nGain: {int(gain)} FCFA\nBenefice: +{int(benefice)} FCFA (+{pct}%)"
        _save_historique(update.effective_user.id, "simuler", [cote, mise], {"gain": gain, "benefice": benefice})
        await update.message.reply_text(msg)
    except ValueError:
        await update.message.reply_text("Erreur. Exemple: /simuler 1.85 10000")


async def bilan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if uid not in historique or not historique[uid]:
        await update.message.reply_text("Aucun historique pour cette session.")
        return
    lignes = [f"BILAN SESSION - {len(historique[uid])} operation(s)"]
    for i, h in enumerate(historique[uid], 1):
        lignes.append(f"{i}. [{h['type'].upper()}] {h['resume']} - {h['heure']}")
    await update.message.reply_text("\n".join(lignes))


async def effacer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    historique[update.effective_user.id] = []
    await update.message.reply_text("Historique efface.")


async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if uid not in historique or not historique[uid]:
        await update.message.reply_text("Aucune donnee a exporter.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alboraa Session"
    headers = ["#", "Type", "Resume", "Heure"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1a1a2e")
        cell.alignment = Alignment(horizontal="center")
    for i, h in enumerate(historique[uid], 2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=h["type"].upper())
        ws.cell(row=i, column=3, value=h["resume"])
        ws.cell(row=i, column=4, value=h["heure"])
    ws.column_dimensions["C"].width = 60
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    await update.message.reply_document(
        document=buf,
        filename=f"alboraa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        caption="Export session Alboraa"
    )


def _save_historique(uid, type_op, data, result):
    if uid not in historique:
        historique[uid] = []
    heure = datetime.now().strftime("%H:%M:%S")
    if type_op == "arbitrage":
        arb = result.get("arbitrage", False)
        resume = f"Cotes {data} -> {'ARB +' + str(result.get('profit','')) + '%' if arb else 'Pas arbitrage'}"
    elif type_op == "combo":
        resume = f"{len(data)} matchs -> {len(result)} combo(s) rentable(s)"
    elif type_op == "simuler":
        resume = f"Cote {data[0]} / Mise {int(data[1])} FCFA -> Gain {int(result['gain'])} FCFA"
    else:
        resume = str(data)
    historique[uid].append({"type": type_op, "resume": resume, "heure": heure})


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("arbitrage", arbitrage))
    app.add_handler(CommandHandler("combo", combo))
    app.add_handler(CommandHandler("simuler", simuler))
    app.add_handler(CommandHandler("bilan", bilan))
    app.add_handler(CommandHandler("effacer", effacer))
    app.add_handler(CommandHandler("export", export))
    print("Bot Alboraa demarre")
    app.run_polling()


if __name__ == "__main__":
    main()

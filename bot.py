import logging
import os
import re
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

MISE_DEFAUT = 50000
TAXE = 0.10
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
    def combo_matchs(matchs, noms_matchs, mise_totale):
        labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        resultats = []

        nb_total = 1
        for m in matchs:
            nb_total *= len(m)

        mise_par_combo = round(mise_totale / nb_total, 0)

        # Labels issus/domicile/extérieur
        issue_labels = ["1", "N", "2"]

        for combo in product(*[range(len(m)) for m in matchs]):
            cotes_sel = [matchs[i][combo[i]] for i in range(len(matchs))]
            codes = [f"{labels[i]}{combo[i]+1}" for i in range(len(matchs))]

            # Noms lisibles par match
            noms_combo = []
            for i, idx in enumerate(combo):
                match_nom = noms_matchs[i] if i < len(noms_matchs) else f"Match {i+1}"
                equipes = match_nom.split(" vs ")
                if len(equipes) == 2:
                    if idx == 0:
                        noms_combo.append(equipes[0].strip())
                    elif idx == 1:
                        noms_combo.append("Nul")
                    else:
                        noms_combo.append(equipes[1].strip())
                else:
                    noms_combo.append(f"{match_nom}({issue_labels[idx] if idx < 3 else idx+1})")

            cote_totale = 1.0
            for c in cotes_sel:
                cote_totale *= c
            cote_totale = round(cote_totale, 4)

            gain_brut = round(mise_par_combo * cote_totale, 0)
            taxe = round(gain_brut * TAXE, 0)
            gain_net = round(gain_brut - taxe, 0)
            benefice_net = round(gain_net - mise_totale, 0)
            benefice_pct = round((gain_net / mise_totale - 1) * 100, 2)

            # Probabilité implicite
            prob = round((1 / cote_totale) * 100, 2)

            resultats.append({
                "codes": "+".join(codes),
                "noms_combo": " + ".join(noms_combo),
                "cotes": cotes_sel,
                "cote_totale": cote_totale,
                "mise_par_combo": int(mise_par_combo),
                "gain_brut": int(gain_brut),
                "taxe": int(taxe),
                "gain_net": int(gain_net),
                "benefice_net": int(benefice_net),
                "benefice_pct": benefice_pct,
                "prob": prob,
                "rentable": gain_net > mise_totale,
            })

        resultats.sort(key=lambda x: x["cote_totale"])
        return resultats, nb_total, int(mise_par_combo)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "ALBORAA BOT - Arbitrage et Combinaisons\n"
        "Commandes disponibles:\n\n"
        '/combo "Barça vs Real" 1.30 2.21 3.12 | "Mali vs Sénégal" 1.45 3.10 4.20\n'
        "/arbitrage 1.21 3.02 3.10\n"
        "/simuler 1.85 10000\n"
        "/bilan\n"
        "/effacer\n"
        "/export\n\n"
        "Devise: Franc CFA\n"
        f"Mise defaut: {MISE_DEFAUT:,} FCFA\n"
        f"Taxe Mali: {int(TAXE*100)}% (Fonds Securite)\n"
        "Seuil benefice: gain net > mise totale"
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
            stakes_str = "\n".join([
                f"Cote {k[-1]} ({cotes[int(k[-1])-1]}) -> {v}%"
                for k, v in result["stakes"].items()
            ])
            msg = (
                f"ARBITRAGE DETECTE\n"
                f"Profit: {result['profit']}%\n"
                f"Somme: {result['sum']}\n\n"
                f"Repartition:\n{stakes_str}"
            )
            _save_historique(update.effective_user.id, "arbitrage", cotes, result)
        else:
            msg = f"Pas d'arbitrage. Somme: {result['sum']} (doit etre < 1.00)"
        await update.message.reply_text(msg)
    except ValueError:
        await update.message.reply_text("Cotes invalides. Exemple: /arbitrage 1.21 3.02 3.10")


async def combo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text(
            'Format: /combo "Barça vs Real" 1.30 2.21 3.12 | "Mali vs Sénégal" 1.45 3.10 4.20'
        )
        return
    try:
        ligne = " ".join(context.args)

        # Mise personnalisée en fin de ligne
        mise_totale = MISE_DEFAUT
        tokens = ligne.split()
        try:
            derniere = float(tokens[-1])
            if derniere > 100 and '"' not in tokens[-1] and "|" not in tokens[-1]:
                mise_totale = derniere
                ligne = " ".join(tokens[:-1])
        except ValueError:
            pass

        # Séparer les matchs par |
        segments = ligne.split("|")
        matchs = []
        noms_matchs = []

        for seg in segments:
            seg = seg.strip()
            # Extraire le nom entre guillemets
            nom_match = re.search(r'"([^"]+)"', seg)
            if nom_match:
                nom = nom_match.group(1)
                seg_sans_nom = seg.replace(nom_match.group(0), "").strip()
            else:
                nom = f"Match {len(matchs)+1}"
                seg_sans_nom = seg

            cotes = [float(x) for x in seg_sans_nom.split() if x.strip()]
            if cotes:
                matchs.append(cotes)
                noms_matchs.append(nom)

        if len(matchs) < 2:
            await update.message.reply_text("Minimum 2 matchs separes par |")
            return
        if len(matchs) > 10:
            await update.message.reply_text("Maximum 10 matchs.")
            return

        resultats, nb_total, mise_par_combo = AlboraEngine.combo_matchs(matchs, noms_matchs, mise_totale)

        rentables = [r for r in resultats if r["rentable"]]
        gain_min = resultats[0]["gain_net"]
        gain_max = resultats[-1]["gain_net"]

        # Meilleure combo = la plus probable (cote la plus basse)
        meilleure = resultats[0]

        lignes = []
        for i, r in enumerate(resultats, 1):
            signe = "+" if r["benefice_pct"] >= 0 else ""
            rentable_tag = "✅" if r["rentable"] else "❌"
            lignes.append(
                f"{rentable_tag} {i}. {r['noms_combo']}\n"
                f"   Mise: {r['mise_par_combo']:,} FCFA | Cote: {r['cote_totale']}\n"
                f"   Gain brut: {r['gain_brut']:,} FCFA | Taxe: -{r['taxe']:,} FCFA\n"
                f"   Gain net: {r['gain_net']:,} FCFA ({signe}{r['benefice_pct']}%) | Prob: {r['prob']}%"
            )

        en_tete = (
            f"ALBORAA - {len(matchs)} matchs\n"
            f"Mise totale: {int(mise_totale):,} FCFA\n"
            f"Mise par combo: {mise_par_combo:,} FCFA\n"
            f"Combos rentables: {len(rentables)}/{nb_total}\n"
            f"{'='*30}\n\n"
        )

        meilleure_section = (
            f"\n{'='*30}\n"
            f"⭐ MEILLEURE COMBO\n"
            f"{meilleure['noms_combo']}\n"
            f"Prob: {meilleure['prob']}%\n"
            f"Mise: {meilleure['mise_par_combo']:,} FCFA\n"
            f"Gain brut: {meilleure['gain_brut']:,} FCFA\n"
            f"Taxe (10%): -{meilleure['taxe']:,} FCFA\n"
            f"Gain net: {meilleure['gain_net']:,} FCFA\n"
        )
        avertissement = (
            f"\n{'='*30}\n"
            f"⚠️ AVERTISSEMENT\n"
            f"Mise totale engagee: {int(mise_totale):,} FCFA\n"
            f"Taxe Mali 10% appliquee.\n"
            f"Une seule combo gagne.\n"
            f"Gain net min: {gain_min:,} FCFA\n"
            f"Gain net max: {gain_max:,} FCFA\n"
            f"👉 Choisissez UNE seule combinaison."
        )
        msg = en_tete + "\n".join(lignes) + meilleure_section + avertissement
        _save_historique(update.effective_user.id, "combo", matchs, resultats)

        # Telegram limite à 4096 caractères — découper si nécessaire
        if len(msg) > 4096:
            await update.message.reply_text(en_tete + "\n".join(lignes[:15]))
            await update.message.reply_text("\n".join(lignes[15:]) + meilleure_section + avertissement)
        else:
            await update.message.reply_text(msg)

    except Exception as e:
        await update.message.reply_text(
            'Erreur de format.\nExemple: /combo "Barça vs Real" 1.30 2.21 3.12 | "Mali vs Sénégal" 1.45 3.10 4.20'
        )
async def simuler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) != 2:
        await update.message.reply_text("Format: /simuler 1.85 10000")
        return
    try:
        cote = float(args[0])
        mise = float(args[1])
        gain_brut = round(mise * cote, 0)
        taxe = round(gain_brut * TAXE, 0)
        gain_net = round(gain_brut - taxe, 0)
        benefice = round(gain_net - mise, 0)
        pct = round((gain_net / mise - 1) * 100, 2)
        msg = (
            f"SIMULATION\n"
            f"Cote: {cote}\n"
            f"Mise: {int(mise):,} FCFA\n"
            f"Gain brut: {int(gain_brut):,} FCFA\n"
            f"Taxe (10%): -{int(taxe):,} FCFA\n"
            f"Gain net: {int(gain_net):,} FCFA\n"
            f"Benefice: {'+' if benefice >= 0 else ''}{int(benefice):,} FCFA ({'+' if pct >= 0 else ''}{pct}%)"
        )
        _save_historique(update.effective_user.id, "simuler", [cote, mise], {"gain": gain_net, "benefice": benefice})
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
        rentables = [r for r in result if r.get("rentable")]
        resume = f"{len(data)} matchs -> {len(rentables)} combo(s) rentable(s)"
    elif type_op == "simuler":
        resume = f"Cote {data[0]} / Mise {int(data[1]):,} FCFA -> Gain net {int(result['gain']):,} FCFA"
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
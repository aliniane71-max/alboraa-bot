import logging
from itertools import product
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
TELEGRAM_TOKEN = "8935700557:AAFF00ot8CoQ18gQ0XRZT1D4o0v7krQFw"
Logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
MISE_DEFAUT = 5000
SEUIL_BENEFICE = 1.015  # 1.50%
# Historique session par utilisateur
historique = {}
# ─────────────────────────────────────────────
# MOTEUR ARBITRAGE
# ─────────────────────────────────────────────
class AlboraEngine:

    @staticmethod
    def detect_arbitrage(cotes: list[float]) -> dict:
        """Détecte un arbitrage sur 2 ou 3 cotes d'un même match."""
        s = sum(1 / c for c in cotes)
        if s < 1:
            profit = round((1 / s - 1) * 100, 2)
            stakes = {f"cote_{i+1}": round((1 / c / s) * 100, 2) for i, c in enumerate(cotes)}
            return {"arbitrage": True, "profit": profit, "stakes": stakes, "sum": round(s, 4)}
        return {"arbitrage": False, "sum": round(s, 4)}

    @staticmethod
    def combo_matchs(matchs: list[list[float]], mise: float) -> list[dict]:
        """
        Génère toutes les combinaisons entre matchs (1 cote par match).
        Filtre : gain >= mise * 1.015
        Trie du plus petit au plus grand.
        """
        labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        resultats = []
        seuil = mise * SEUIL_BENEFICE

        # Génère toutes les combinaisons (1 cote par match)
        for combo in product(*[range(len(m)) for m in matchs]):
            cotes_selectionnees = [matchs[i][combo[i]] for i in range(len(matchs))]
            noms = [f"{labels[i]}{combo[i]+1}" for i in range(len(matchs))]

            cote_totale = 1
            for c in cotes_selectionnees:
                cote_totale *= c
            cote_totale = round(cote_totale, 4)

            gain = round(mise * cote_totale, 0)
            benefice_net = round(gain - mise, 0)
            benefice_pct = round((cote_totale - 1) * 100, 2)

            if gain >= seuil:
                resultats.append({
                    "noms": "+".join(noms),
                    "cotes": cotes_selectionnees,
                    "cote_totale": cote_totale,
                    "gain": int(gain),
                    "benefice_net": int(benefice_net),
                    "benefice_pct": benefice_pct,
                })

        resultats.sort(key=lambda x: x["cote_totale"])
        return resultats


# ─────────────────────────────────────────────
# COMMANDES
# ─────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "⚡ *ALBORAA BOT — Arbitrage & Combinaisons*\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "📌 *Commandes disponibles :*\n\n"
        "🔹 `/arbitrage 1.21 3.02 3.10`\n"
        "   → Détecte un arbitrage sur un match\n\n"
        "🔹 `/combo 1.21 3.02 3.10 | 1.22 3.01 3.11`\n"
        "   → Combinaisons multi-matchs (2 à 10)\n"
        "   → Ajoutez la mise : `/combo ... 10000`\n\n"
        "🔹 `/simuler 1.85 10000`\n"
        "   → Simule un gain pour une cote et mise\n\n"
        "🔹 `/bilan` → Historique de la session\n"
        "🔹 `/effacer` → Efface l'historique\n"
        "🔹 `/export` → Exporte l'historique Excel\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "💰 Devise : *Franc CFA (FCFA)*\n"
        "📊 Mise défaut : *5 000 FCFA*\n"
        "📈 Seuil bénéfice : *+1.50% minimum*"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


async def arbitrage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) < 2 or len(args) > 3:
        await update.message.reply_text(
            "⚠ Format : `/arbitrage 1.21 3.02 3.10`\n"
            "Ou 2 cotes : `/arbitrage 1.45 2.10`",
            parse_mode="Markdown"
        )
        return
    try:
        cotes = list(map(float, args))
        result = AlboraEngine.detect_arbitrage(cotes)

        if result["arbitrage"]:
            stakes_str = "\n".join(
                [f"  • Cote {k[-1]} ({cotes[int(k[-1])-1]}) → {v}%" for k, v in result["stakes"].items()]
            )
            msg = (
                f"✅ *ARBITRAGE DÉTECTÉ !*\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📈 Profit garanti : *{result['profit']}%*\n"
                f"📊 Somme inverses : {result['sum']}\n\n"
                f"💰 *Répartition des mises (sur 100%) :*\n"
                f"{stakes_str}"
            )
            _save_historique(update.effective_user.id, "arbitrage", cotes, result)
        else:
            msg = (
                f"❌ *Pas d'arbitrage*\n"
                f"Somme des inverses : {result['sum']} (doit être < 1.00)"
            )

        await update.message.reply_text(msg, parse_mode="Markdown")

    except ValueError:
        await update.message.reply_text("⚠ Cotes invalides. Exemple : `/arbitrage 1.21 3.02 3.10`", parse_mode="Markdown")


async def combo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Format : /combo 1.21 3.02 3.10 | 1.22 3.01 3.11 | 1.45 2.80 3.05 [mise]
    Séparateur : | entre les matchs
    Dernier argument numérique sans | = mise en FCFA
    """
    if not context.args:
        await update.message.reply_text(
            "⚠ Format :\n`/combo 1.21 3.02 3.10 | 1.22 3.01 3.11`\n\n"
            "Avec mise : `/combo 1.21 3.02 3.10 | 1.22 3.01 3.11 10000`\n\n"
            "Jusqu'à 10 matchs séparés par `|`",
            parse_mode="Markdown"
        )
        return

    try:
        # Reconstruire la ligne complète
        ligne = " ".join(context.args)

        # Détecter si la mise est à la fin (dernier token sans |)
        tokens = ligne.split()
        mise = MISE_DEFAUT
        try:
            derniere = float(tokens[-1])
            # Si dernier token est un nombre et n'est pas précédé d'un |
            if "|" not in tokens[-1] and derniere > 100:
                mise = derniere
                ligne = " ".join(tokens[:-1])
        except ValueError:
            pass

        # Séparer les matchs par |
        matchs_raw = ligne.split("|")
        matchs = []
        for m in matchs_raw:
            cotes = [float(x) for x in m.strip().split() if x.strip()]
            if cotes:
                matchs.append(cotes)

        if len(matchs) < 2:
            await update.message.reply_text("⚠ Minimum 2 matchs séparés par `|`", parse_mode="Markdown")
            return
        if len(matchs) > 10:
            await update.message.reply_text("⚠ Maximum 10 matchs.", parse_mode="Markdown")
            return

        # Calcul
        resultats = AlboraEngine.combo_matchs(matchs, mise)

        # Construire le message
        nb_combos_total = 1
        for m in matchs:
            nb_combos_total *= len(m)

        nb_filtrees = nb_combos_total - len(resultats)
        seuil_fcfa = int(mise * SEUIL_BENEFICE)

        if not resultats:
            msg = (
                f"❌ *Aucune combinaison rentable*\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📊 {nb_combos_total} combinaisons analysées\n"
                f"💰 Mise : {int(mise):,} FCFA\n"
                f"📉 Seuil minimum : {seuil_fcfa:,} FCFA (+1.50%)\n\n"
                f"Toutes les combinaisons sont sous le seuil de rentabilité."
            )
        else:
            lignes = []
            for i, r in enumerate(resultats, 1):
                cotes_str = " × ".join(str(c) for c in r["cotes"])
                lignes.append(
                    f"{i}. *{r['noms']}* = {cotes_str} = *{r['cote_totale']}*\n"
                    f"   💵 {r['gain']:,} FCFA  (+{r['benefice_pct']}% | +{r['benefice_net']:,} FCFA)"
                )

            bloc = "\n\n".join(lignes)
            msg = (
                f"⚽ *ALBORAA — Combinaisons {len(matchs)} matchs*\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"💰 Mise : *{int(mise):,} FCFA*  |  Seuil : +1.50%\n"
                f"📊 {len(resultats)} rentables / {nb_combos_total} total  |  {nb_filtrees} éliminées\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n\n"
                f"{bloc}\n\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"🏆 *Meilleure :* {resultats[-1]['noms']} → {resultats[-1]['gain']:,} FCFA"
            )

        _save_historique(update.effective_user.id, "combo", matchs, resultats)
        await update.message.reply_text(msg, parse_mode="Markdown")

    except Exception as e:
        await update.message.reply_text(
            f"⚠ Erreur de format.\nExemple : `/combo 1.21 3.02 3.10 | 1.22 3.01 3.11`",
            parse_mode="Markdown"
        )


async def simuler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    args = context.args
    if len(args) != 2:
        await update.message.reply_text("Format : `/simuler 1.85 10000`", parse_mode="Markdown")
        return
    try:
        cote = float(args[0])
        mise = float(args[1])
        gain = round(mise * cote, 0)
        benefice = round(gain - mise, 0)
        pct = round((cote - 1) * 100, 2)

        msg = (
            f"📊 *SIMULATION*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"🎯 Cote : *{cote}*\n"
            f"💰 Mise : *{int(mise):,} FCFA*\n"
            f"━━━━━━━━━━━━━━━\n"
            f"💵 Gain total : *{int(gain):,} FCFA*\n"
            f"📈 Bénéfice net : *+{int(benefice):,} FCFA* (+{pct}%)"
        )
        _save_historique(update.effective_user.id, "simuler", [cote, mise], {"gain": gain, "benefice": benefice})
        await update.message.reply_text(msg, parse_mode="Markdown")
    except ValueError:
        await update.message.reply_text("⚠ Erreur. Exemple : `/simuler 1.85 10000`", parse_mode="Markdown")


async def bilan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if uid not in historique or not historique[uid]:
        await update.message.reply_text("📭 Aucun historique pour cette session.")
        return

    lignes = [f"📋 *BILAN SESSION — {len(historique[uid])} opération(s)*\n━━━━━━━━━━━━━━━"]
    for i, h in enumerate(historique[uid], 1):
        lignes.append(f"{i}. [{h['type'].upper()}] {h['resume']} — {h['heure']}")

    await update.message.reply_text("\n".join(lignes), parse_mode="Markdown")


async def effacer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    historique[uid] = []
    await update.message.reply_text("🗑 Historique effacé.")


async def export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if uid not in historique or not historique[uid]:
        await update.message.reply_text("📭 Aucune donnée à exporter.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alboraa Session"

    # En-têtes
    headers = ["#", "Type", "Résumé", "Heure"]
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
        filename=f"alboraa_session_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        caption="📊 Export session Alboraa"
    )


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def _save_historique(uid: int, type_op: str, data, result):
    if uid not in historique:
        historique[uid] = []

    heure = datetime.now().strftime("%H:%M:%S")

    if type_op == "arbitrage":
        arb = result.get("arbitrage", False)
        resume = f"Cotes {data} → {'ARB +' + str(result.get('profit','')) + '%' if arb else 'Pas d arbitrage'}"
    elif type_op == "combo":
        nb = len(result)
        resume = f"{len(data)} matchs → {nb} combo(s) rentable(s)"
    elif type_op == "simuler":
        resume = f"Cote {data[0]} / Mise {int(data[1]):,} FCFA → Gain {int(result['gain']):,} FCFA"
    else:
        resume = str(data)

    historique[uid].append({"type": type_op, "resume": resume, "heure": heure})


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("arbitrage", arbitrage))
    app.add_handler(CommandHandler("combo", combo))
    app.add_handler(CommandHandler("simuler", simuler))
    app.add_handler(CommandHandler("bilan", bilan))
    app.add_handler(CommandHandler("effacer", effacer))
    app.add_handler(CommandHandler("export", export))

    print("🤖 Bot Alboraa démarré — En attente de messages Telegram...")
    app.run_polling()


if __name__ == "__main__":
    main()

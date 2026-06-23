#!/usr/bin/env python3
"""
ALBORAA BOT – Arbitrage Sportif
Version VPS – 23/06/2026
Déploiement : systemd service sur Linux (Ubuntu/Debian)
"""

import logging
import os
from io import BytesIO
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
)

# ============================================================
# CONFIGURATION – Remplacez par votre token
# ============================================================
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "VOTRE_TOKEN_ICI")

# (Optionnel) Limitez l'accès à vos propres Telegram user IDs
# Laissez vide [] pour autoriser tout le monde
ALLOWED_USERS: list[int] = []

# ============================================================
# LOGGING
# ============================================================
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler("alboraa.log"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# ============================================================
# STOCKAGE EN MÉMOIRE (sessions actives uniquement)
# Pour une persistance entre redémarrages, utilisez SQLite.
# ============================================================
historique: list[dict] = []


# ============================================================
# MOTEUR D'ARBITRAGE
# ============================================================
class ArbitrageEngine:
    """Calcule les opportunités d'arbitrage sur 2 ou 3 issues."""

    @staticmethod
    def calculer(cotes: list[float]) -> dict:
        if len(cotes) not in (2, 3):
            raise ValueError("Fournissez 2 ou 3 cotes.")

        s = sum(1 / c for c in cotes)
        profit_pct = round((1 / s - 1) * 100, 3)
        mises_pct = [round((1 / c / s) * 100, 2) for c in cotes]

        labels = (
            ["1 (Home)", "2 (Away)"]
            if len(cotes) == 2
            else ["1 (Home)", "X (Draw)", "2 (Away)"]
        )

        return {
            "arbitrage": s < 1,
            "s": round(s, 4),
            "profit_pct": profit_pct,
            "mises_pct": dict(zip(labels, mises_pct)),
        }

    @staticmethod
    def simuler_mises(result: dict, mise_totale: float) -> dict:
        """Calcule les mises réelles en € pour une mise totale donnée."""
        return {
            label: round(pct / 100 * mise_totale, 2)
            for label, pct in result["mises_pct"].items()
        }


# ============================================================
# GARDE D'ACCÈS
# ============================================================
def acces_autorise(user_id: int) -> bool:
    return not ALLOWED_USERS or user_id in ALLOWED_USERS


# ============================================================
# HANDLERS
# ============================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not acces_autorise(update.effective_user.id):
        return

    texte = (
        "⚡ *Alboraa – Bot d'Arbitrage Sportif*\n\n"
        "📌 *Commandes disponibles :*\n\n"
        "`/arbitrage <c1> <c2> [c3]`\n"
        "→ Analyse 2 ou 3 cotes\n\n"
        "`/simuler <c1> <c2> [c3] <mise_totale>`\n"
        "→ Calcule les mises réelles en €\n\n"
        "`/bilan`\n"
        "→ Exporte l'historique en Excel\n\n"
        "`/effacer`\n"
        "→ Vide l'historique de cette session\n\n"
        "*Exemple :*\n"
        "`/arbitrage 1.21 3.02 3.21`\n"
        "`/simuler 1.21 3.02 3.21 1000`"
    )
    await update.message.reply_text(texte, parse_mode="Markdown")


async def arbitrage(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not acces_autorise(update.effective_user.id):
        return

    args = context.args
    if len(args) not in (2, 3):
        await update.message.reply_text(
            "❗ Format : `/arbitrage cote1 cote2` ou `/arbitrage cote1 cote2 cote3`",
            parse_mode="Markdown",
        )
        return

    try:
        cotes = list(map(float, args))
        result = ArbitrageEngine.calculer(cotes)
    except ValueError as e:
        await update.message.reply_text(f"❗ Erreur : {e}")
        return
    except Exception:
        await update.message.reply_text("❗ Utilisez des nombres valides (ex: 1.10 3.02 3.21)")
        return

    # Enregistrement dans l'historique
    historique.append(
        {
            "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "cotes": cotes,
            "s": result["s"],
            "arbitrage": result["arbitrage"],
            "profit_pct": result["profit_pct"],
        }
    )

    if result["arbitrage"]:
        lignes_mises = "\n".join(
            f"  • {label} : {pct} %" for label, pct in result["mises_pct"].items()
        )
        msg = (
            f"🔔 *ARBITRAGE DÉTECTÉ !*\n\n"
            f"📈 Profit garanti : `{result['profit_pct']} %`\n"
            f"📊 Somme S : `{result['s']}`\n\n"
            f"💰 *Répartition des mises :*\n{lignes_mises}\n\n"
            f"_Utilisez `/simuler` pour calculer les montants réels._"
        )
    else:
        msg = (
            f"❌ *Pas d'arbitrage*\n\n"
            f"📊 Somme S : `{result['s']}` (doit être < 1.000 pour arbitrer)\n"
            f"📉 Marge bookmaker : `{abs(result['profit_pct']):.2f} %`"
        )

    await update.message.reply_text(msg, parse_mode="Markdown")


async def simuler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not acces_autorise(update.effective_user.id):
        return

    args = context.args
    if len(args) not in (3, 4):
        await update.message.reply_text(
            "❗ Format : `/simuler c1 c2 mise` ou `/simuler c1 c2 c3 mise`",
            parse_mode="Markdown",
        )
        return

    try:
        *cotes_str, mise_str = args
        cotes = list(map(float, cotes_str))
        mise_totale = float(mise_str)
        result = ArbitrageEngine.calculer(cotes)
    except Exception:
        await update.message.reply_text("❗ Vérifiez vos valeurs (nombres décimaux).")
        return

    if not result["arbitrage"]:
        await update.message.reply_text(
            f"❌ Pas d'arbitrage possible (S = {result['s']}). Simulation annulée."
        )
        return

    mises_reelles = ArbitrageEngine.simuler_mises(result, mise_totale)
    gain_net = round(mise_totale * result["profit_pct"] / 100, 2)

    lignes = "\n".join(
        f"  • {label} : `{montant} €`" for label, montant in mises_reelles.items()
    )
    msg = (
        f"🧮 *Simulation – Mise totale : {mise_totale} €*\n\n"
        f"📈 Profit : `{result['profit_pct']} %` → `+{gain_net} €`\n\n"
        f"💰 *Mises à placer :*\n{lignes}"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


async def bilan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not acces_autorise(update.effective_user.id):
        return

    if not historique:
        await update.message.reply_text("📭 Aucun calcul enregistré dans cette session.")
        return

    # Création du fichier Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bilan Alboraa"

    # En-têtes
    entetes = ["Date", "Cotes", "Somme S", "Arbitrage ?", "Profit (%)"]
    for col, titre in enumerate(entetes, 1):
        cell = ws.cell(row=1, column=col, value=titre)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center")

    # Données
    for i, h in enumerate(historique, 2):
        cotes_str = " | ".join(str(c) for c in h["cotes"])
        ws.cell(row=i, column=1, value=h["date"])
        ws.cell(row=i, column=2, value=cotes_str)
        ws.cell(row=i, column=3, value=h["s"])
        ws.cell(row=i, column=4, value="✅ OUI" if h["arbitrage"] else "❌ NON")
        profit_cell = ws.cell(row=i, column=5, value=h["profit_pct"])
        if h["arbitrage"]:
            profit_cell.font = Font(color="006400", bold=True)

    # Ajustement largeurs
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    await update.message.reply_document(
        document=buffer,
        filename=f"bilan_alboraa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        caption=f"📊 *Bilan Alboraa* – {len(historique)} calcul(s) exporté(s).",
        parse_mode="Markdown",
    )


async def effacer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not acces_autorise(update.effective_user.id):
        return

    n = len(historique)
    historique.clear()
    await update.message.reply_text(f"🗑️ Historique effacé ({n} entrée(s) supprimée(s)).")


# ============================================================
# POINT D'ENTRÉE
# ============================================================
def main():
    if TELEGRAM_TOKEN == "VOTRE_TOKEN_ICI":
        logger.error("❌ Veuillez définir TELEGRAM_TOKEN dans les variables d'environnement.")
        return

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("arbitrage", arbitrage))
    app.add_handler(CommandHandler("simuler", simuler))
    app.add_handler(CommandHandler("bilan", bilan))
    app.add_handler(CommandHandler("effacer", effacer))

    logger.info("🤖 Bot Alboraa démarré (mode polling).")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

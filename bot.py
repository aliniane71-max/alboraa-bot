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
    def detect_arbitrage(cotes: list[float]) -> dict:
        s = sum(1 / c for c in cotes)
        if s < 1:
            profit = round((1 / s - 1) * 100, 2)
            stakes = {f"cote_{i+1}": round((1 / c / s) * 100, 2) for i, c in enumerate(cotes)}
            return {"arbitrage": True, "profit": profit, "stakes": stakes, "sum": round(s, 4)}
        return {"arbitrage": False, "sum": round(s, 4)}

    @staticmethod
    def combo_matchs(matchs: list[list[float]], mise: float) -> list[dict]:
        labels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        resultats = []
        seuil = mise * SEUIL_BENEFICE
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


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = (
        "⚡ *ALBORAA BOT — Arbitrage & Combinaisons*\n"
        "━━━━━━━━━━━━━━━━━━━━━━\n"
        "📌 *Commandes disponibles :*\n\n"
        "🔹 `/arbitrage 1.


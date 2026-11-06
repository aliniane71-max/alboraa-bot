"""Alboraa - bot Telegram qui surveille les marches de FOOT sur Polymarket.

V1 : lecture seule. Le bot ne mise jamais. Il suit des marches et te previent
quand le prix bouge de plus de X points (ex: 62% -> 70% = +8 points).

Variables Railway :
  TELEGRAM_TOKEN     (obligatoire)
  ALLOWED_USER_IDS   ids Telegram autorises, separes par des virgules (conseille)
  DB_PATH            chemin de la base SQLite (defaut: /data/alboraa.db si /data existe)
  SCAN_INTERVAL      secondes entre deux scans (defaut 300)
  SEUIL_DEFAUT       seuil d'alerte en points (defaut 5)
"""
import asyncio
import json
import logging
import os
import sqlite3
import threading
import time
from functools import wraps

import requests
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes

TOKEN = os.getenv("TELEGRAM_TOKEN")
ALLOWED_USER_IDS = set(
    int(x.strip()) for x in os.getenv("ALLOWED_USER_IDS", "").split(",") if x.strip().isdigit()
)
GAMMA = "https://gamma-api.polymarket.com"
CLOB = "https://clob.polymarket.com"
GEOBLOCK = "https://polymarket.com/api/geoblock"
INTERVALLE_SCAN = int(os.getenv("SCAN_INTERVAL", "300"))
SEUIL_DEFAUT = float(os.getenv("SEUIL_DEFAUT", "5"))
DB_PATH = os.getenv("DB_PATH") or ("/data/alboraa.db" if os.path.isdir("/data") else "alboraa.db")
MAX_SUIVIS = 30

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
# httpx logue l'URL Telegram complete (avec le token) : on le fait taire.
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)

_db_lock = threading.Lock()


# ─────────────────────────────────────────────
# Base de donnees (SQLite)
# ─────────────────────────────────────────────
def _db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute(
        "CREATE TABLE IF NOT EXISTS suivis ("
        "uid INTEGER, market_id TEXT, question TEXT, token_id TEXT, ref_prix REAL, "
        "PRIMARY KEY (uid, market_id))"
    )
    conn.execute("CREATE TABLE IF NOT EXISTS reglages (uid INTEGER PRIMARY KEY, seuil REAL)")
    return conn


def db_ajouter(uid, market_id, question, token_id, prix):
    with _db_lock:
        conn = _db()
        with conn:
            conn.execute(
                "INSERT OR REPLACE INTO suivis (uid, market_id, question, token_id, ref_prix) VALUES (?,?,?,?,?)",
                (uid, str(market_id), question, token_id, prix),
            )
        conn.close()


def db_supprimer(uid, market_id):
    with _db_lock:
        conn = _db()
        with conn:
            conn.execute("DELETE FROM suivis WHERE uid=? AND market_id=?", (uid, str(market_id)))
        conn.close()


def db_suivis(uid=None):
    with _db_lock:
        conn = _db()
        if uid is None:
            rows = conn.execute("SELECT * FROM suivis").fetchall()
        else:
            rows = conn.execute("SELECT * FROM suivis WHERE uid=? ORDER BY rowid", (uid,)).fetchall()
        conn.close()
    return [dict(r) for r in rows]


def db_maj_ref(uid, market_id, prix):
    with _db_lock:
        conn = _db()
        with conn:
            conn.execute("UPDATE suivis SET ref_prix=? WHERE uid=? AND market_id=?", (prix, uid, str(market_id)))
        conn.close()


def db_seuil(uid):
    with _db_lock:
        conn = _db()
        row = conn.execute("SELECT seuil FROM reglages WHERE uid=?", (uid,)).fetchone()
        conn.close()
    return float(row["seuil"]) if row else SEUIL_DEFAUT


def db_set_seuil(uid, seuil):
    with _db_lock:
        conn = _db()
        with conn:
            conn.execute("INSERT OR REPLACE INTO reglages (uid, seuil) VALUES (?,?)", (uid, seuil))
        conn.close()


# ─────────────────────────────────────────────
# Polymarket (lecture seule, sans cle)
# ─────────────────────────────────────────────
def http_get(url, params=None):
    try:
        r = requests.get(url, params=params, timeout=15)
        if r.status_code == 200:
            return r.json()
        logging.warning(f"HTTP {r.status_code} sur {url.split('?')[0]}")
    except Exception as e:
        logging.error(f"Erreur reseau sur {url.split('?')[0]}: {e}")
    return None


def _liste(valeur):
    """Gamma renvoie souvent des listes sous forme de texte JSON : '["0.6","0.4"]'."""
    if isinstance(valeur, str):
        try:
            valeur = json.loads(valeur)
        except Exception:
            return []
    return valeur if isinstance(valeur, list) else []


def _vers_marche(m):
    """Transforme un marche Gamma en dict simple, ou None s'il est inutilisable."""
    if m.get("closed") or m.get("active") is False:
        return None
    tokens = _liste(m.get("clobTokenIds"))
    prix = _liste(m.get("outcomePrices"))
    if not tokens or not prix:
        return None
    try:
        p = float(prix[0])
    except (TypeError, ValueError):
        return None
    try:
        volume = float(m.get("volume24hr") or m.get("volume") or 0)
    except (TypeError, ValueError):
        volume = 0.0
    return {
        "id": str(m.get("id")),
        "question": m.get("question") or m.get("slug") or "?",
        "token": str(tokens[0]),  # index 0 = le "Oui"
        "prix": p,
        "volume": volume,
    }


def _evenements_foot(mot):
    if mot:
        data = http_get(f"{GAMMA}/public-search", {"q": mot})
        events = (data or {}).get("events", []) if isinstance(data, dict) else []
        # filtre souple : si les evenements ont des tags, on garde ceux de football
        filtres = []
        for ev in events:
            slugs = {str(t.get("slug", "")).lower() for t in (ev.get("tags") or []) if isinstance(t, dict)}
            if not slugs or slugs & {"soccer", "football"}:
                filtres.append(ev)
        return filtres
    tag = http_get(f"{GAMMA}/tags/slug/soccer")
    tag_id = tag.get("id") if isinstance(tag, dict) else None
    if tag_id is None:
        return []
    data = http_get(f"{GAMMA}/events", {"tag_id": tag_id, "active": "true", "closed": "false", "limit": 50})
    return data if isinstance(data, list) else []


def trouver_marches_foot(mot=None, limite=10):
    marches = []
    for ev in _evenements_foot(mot):
        for m in ev.get("markets") or []:
            mm = _vers_marche(m)
            if mm:
                marches.append(mm)
    marches.sort(key=lambda x: x["volume"], reverse=True)
    return marches[:limite]


def prix_actuel(token_id):
    """Prix du 'Oui' (0 a 1). On essaie le milieu du carnet, puis le prix d'achat."""
    d = http_get(f"{CLOB}/midpoint", {"token_id": token_id})
    if isinstance(d, dict):
        for k in ("mid", "mid_price"):
            if k in d:
                try:
                    return float(d[k])
                except (TypeError, ValueError):
                    pass
    d = http_get(f"{CLOB}/price", {"token_id": token_id, "side": "buy"})
    if isinstance(d, dict) and "price" in d:
        try:
            return float(d["price"])
        except (TypeError, ValueError):
            pass
    return None


# ─────────────────────────────────────────────
# Scanner : cherche les mouvements de prix
# ─────────────────────────────────────────────
def pct(x):
    return f"{x * 100:.0f}%"


def scanner_une_fois():
    """Retourne la liste des (uid, message) a envoyer."""
    messages = []
    cache = {}
    for s in db_suivis():
        token = s["token_id"]
        if token not in cache:
            cache[token] = prix_actuel(token)
        p = cache[token]
        if p is None:
            continue
        ecart = (p - s["ref_prix"]) * 100
        if abs(ecart) >= db_seuil(s["uid"]):
            signe = "+" if ecart >= 0 else ""
            messages.append((
                s["uid"],
                f"MOUVEMENT DE PRIX\n{s['question']}\n"
                f"{pct(s['ref_prix'])} -> {pct(p)} ({signe}{ecart:.1f} points)",
            ))
            db_maj_ref(s["uid"], s["market_id"], p)
    return messages


def scanner_loop(bot_token):
    from telegram import Bot
    bot = Bot(token=bot_token)
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    logging.info(f"Scanner demarre (toutes les {INTERVALLE_SCAN}s).")
    while True:
        try:
            for uid, msg in scanner_une_fois():
                try:
                    loop.run_until_complete(bot.send_message(chat_id=uid, text=msg))
                except Exception as e:
                    logging.error(f"Envoi alerte: {e}")
        except Exception as e:
            logging.error(f"Erreur scanner: {e}")
        time.sleep(INTERVALLE_SCAN)


# ─────────────────────────────────────────────
# Commandes Telegram
# ─────────────────────────────────────────────
def protege(fonction):
    @wraps(fonction)
    async def enveloppe(update: Update, context: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if ALLOWED_USER_IDS and uid not in ALLOWED_USER_IDS:
            await update.message.reply_text("Acces non autorise.")
            return
        return await fonction(update, context)
    return enveloppe


@protege
async def start(update, context):
    await update.message.reply_text(
        "ALBORAA - Foot sur Polymarket (lecture seule)\n\n"
        "/foot - marches foot les plus actifs\n"
        "/foot mot - chercher (ex: /foot Mali)\n"
        "/suivre 3 - suivre le marche numero 3 de la liste\n"
        "/liste - voir tes marches suivis et leur prix\n"
        "/oublier 2 - arreter de suivre le numero 2\n"
        "/alerte 5 - alerte si le prix bouge de 5 points ou plus\n"
        "/diagnostic - tester la connexion a Polymarket\n\n"
        "Le bot ne mise jamais. Il te previent, c'est toi qui decides."
    )


@protege
async def foot(update, context):
    mot = " ".join(context.args).strip() or None
    await update.message.reply_text("Recherche en cours...")
    marches = await asyncio.to_thread(trouver_marches_foot, mot)
    if not marches:
        await update.message.reply_text(
            "Rien trouve. Essaie /foot avec un nom d'equipe, ou lance /diagnostic."
        )
        return
    context.user_data["liste"] = marches
    lignes = [f"{i}. {m['question']} - {pct(m['prix'])}" for i, m in enumerate(marches, 1)]
    await update.message.reply_text("\n".join(lignes) + "\n\nPour suivre : /suivre 1")


@protege
async def suivre(update, context):
    uid = update.effective_user.id
    liste = context.user_data.get("liste")
    if not liste:
        await update.message.reply_text("Lance d'abord /foot pour avoir une liste.")
        return
    try:
        n = int(context.args[0])
        if n < 1:
            raise IndexError
        m = liste[n - 1]
    except (IndexError, ValueError):
        await update.message.reply_text(f"Format: /suivre 1 (un numero entre 1 et {len(liste)})")
        return
    if len(db_suivis(uid)) >= MAX_SUIVIS:
        await update.message.reply_text(f"Maximum {MAX_SUIVIS} marches suivis. Utilise /oublier.")
        return
    db_ajouter(uid, m["id"], m["question"], m["token"], m["prix"])
    await update.message.reply_text(
        f"Suivi : {m['question']}\nPrix de depart : {pct(m['prix'])}\n"
        f"Alerte a partir de {db_seuil(uid):g} points."
    )


@protege
async def liste(update, context):
    uid = update.effective_user.id
    suivis = db_suivis(uid)
    if not suivis:
        await update.message.reply_text("Aucun marche suivi. Essaie /foot.")
        return
    lignes = []
    for i, s in enumerate(suivis, 1):
        p = await asyncio.to_thread(prix_actuel, s["token_id"])
        actuel = pct(p) if p is not None else "?"
        lignes.append(f"{i}. {s['question']}\n   depart {pct(s['ref_prix'])} -> maintenant {actuel}")
    await update.message.reply_text("\n".join(lignes))


@protege
async def oublier(update, context):
    uid = update.effective_user.id
    suivis = db_suivis(uid)
    try:
        n = int(context.args[0])
        if n < 1:
            raise IndexError
        s = suivis[n - 1]
    except (IndexError, ValueError):
        await update.message.reply_text("Format: /oublier 1 (numero dans /liste)")
        return
    db_supprimer(uid, s["market_id"])
    await update.message.reply_text(f"Plus suivi : {s['question']}")


@protege
async def alerte(update, context):
    uid = update.effective_user.id
    try:
        seuil = float(context.args[0].replace(",", "."))
        if not 0.5 <= seuil <= 50:
            raise ValueError
    except (IndexError, ValueError):
        await update.message.reply_text(
            f"Format: /alerte 5 (entre 0.5 et 50 points). Actuel : {db_seuil(uid):g}"
        )
        return
    db_set_seuil(uid, seuil)
    await update.message.reply_text(f"Alerte reglee a {seuil:g} points.")


def faire_diagnostic():
    lignes = []
    geo = http_get(GEOBLOCK)
    if isinstance(geo, dict):
        etat = "BLOQUE pour miser" if geo.get("blocked") else "autorise"
        lignes.append(f"Pays vu par Polymarket : {geo.get('country', '?')} - {etat}")
    else:
        lignes.append("Test du pays : echec (le bot lit quand meme les prix)")
    tag = http_get(f"{GAMMA}/tags/slug/soccer")
    lignes.append("Gamma (marches) : " + ("OK" if isinstance(tag, dict) and tag.get("id") else "ECHEC"))
    marches = trouver_marches_foot(limite=5)
    lignes.append(f"Marches foot trouves : {len(marches)}")
    if marches:
        p = prix_actuel(marches[0]["token"])
        lignes.append("CLOB (prix) : " + ("OK" if p is not None else "ECHEC"))
    return "\n".join(lignes)


@protege
async def diagnostic(update, context):
    await update.message.reply_text("Test en cours...")
    texte = await asyncio.to_thread(faire_diagnostic)
    await update.message.reply_text(texte)


def main():
    if not TOKEN:
        raise SystemExit("TELEGRAM_TOKEN manquant.")
    if not ALLOWED_USER_IDS:
        logging.warning("ALLOWED_USER_IDS vide : tout le monde peut utiliser le bot.")
    _db().close()  # cree les tables
    threading.Thread(target=scanner_loop, args=(TOKEN,), daemon=True).start()
    app = Application.builder().token(TOKEN).build()
    for nom, fn in [
        ("start", start), ("foot", foot), ("suivre", suivre), ("liste", liste),
        ("oublier", oublier), ("alerte", alerte), ("diagnostic", diagnostic),
    ]:
        app.add_handler(CommandHandler(nom, fn))
    print("Bot Alboraa Foot-Polymarket demarre - V1 lecture seule")
    app.run_polling()


if __name__ == "__main__":
    main()

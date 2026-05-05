#!/usr/bin/env python3
"""
Agente SWAPinn — bot de Telegram para gestión de apartamentos.
Lee RESERVAS_2026.xlsx y LIMPIEZAS_PROXIMAS.xlsx en tiempo real.
"""

import datetime
import json
import logging
import shutil
import subprocess
import time
from pathlib import Path
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from telegram import Update
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from telegram.constants import ParseMode

# ── Paths ─────────────────────────────────────────────────────────────────────

SCRIPT_DIR    = Path(__file__).parent
CONFIG_FILE   = SCRIPT_DIR / "config.json"
BASE_DIR      = Path("/Users/albava/Documents/Claude/Projects/AIRBNB herramientas")
RESERVAS_XLS  = BASE_DIR / "RESERVAS_2026.xlsx"
LIMPIEZAS_XLS = BASE_DIR / "LIMPIEZAS_2026.xlsx"

SHEETS = ["Sweet 2026", "Sunset 2026", "Center 2026", "Open Sky 2026", "Duplex 2026"]
APT_EMOJI = {
    "Sweet":    "🔵",
    "Sunset":   "🔴",
    "Center":   "🟠",
    "Open Sky": "🟢",
    "Duplex":   "🟡",
}
PLATFORM_EMOJI = {"Airbnb": "🏠", "Booking": "📅", "Holidu": "🌐", "Directa": "✉️"}

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


# ── Config ────────────────────────────────────────────────────────────────────

def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


# ── Drive cache ───────────────────────────────────────────────────────────────

CACHE_FILE         = SCRIPT_DIR / "data_cache.json"
TOKEN_FILE         = SCRIPT_DIR / "token.json"
DRIVE_CACHE_FOLDER = "15stGrk0IgfNAGupejzvpjVir0a-XqOT6"  # carpeta BUSSINESS
DRIVE_SCOPES       = ["https://www.googleapis.com/auth/drive.readonly"]

_cache_data: dict | None = None
_cache_ts: float = 0.0
CACHE_TTL = 300  # segundos

def _get_drive():
    creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), DRIVE_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        TOKEN_FILE.write_text(creds.to_json())
    return build("drive", "v3", credentials=creds)

# ── Cache reader ──────────────────────────────────────────────────────────────

def parse_date(v):
    if not v:
        return None
    if hasattr(v, "date"):
        return v.date()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(str(v), fmt).date()
        except ValueError:
            pass
    return None

def _load_cache():
    global _cache_data, _cache_ts
    now = time.time()
    if _cache_data is not None and (now - _cache_ts) < CACHE_TTL:
        return _cache_data, None
    try:
        svc     = _get_drive()
        results = svc.files().list(
            q=f"name='data_cache.json' and '{DRIVE_CACHE_FOLDER}' in parents and trashed=false",
            orderBy="modifiedTime desc",
            fields="files(id)",
            pageSize=1,
        ).execute()
        files = results.get("files", [])
        if files:
            content     = svc.files().get_media(fileId=files[0]["id"]).execute()
            _cache_data = json.loads(content)
            _cache_ts   = now
            return _cache_data, None
    except Exception as e:
        logger.warning(f"Drive cache error: {e}")
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text(encoding="utf-8")), None
        except Exception:
            pass
    return None, "Sin caché disponible. La rutina remota aún no ha generado datos."

def load_reservas():
    cache, err = _load_cache()
    if err:
        return [], err
    reservas = []
    for r in cache.get("reservas", []):
        reservas.append({
            "apt":      r["apt"],
            "nombre":   r["nombre"],
            "entrada":  parse_date(r["entrada"]),
            "salida":   parse_date(r["salida"]),
            "noches":   r.get("noches", 0),
            "total":    r.get("total", 0),
            "platform": r.get("platform", "—"),
            "adultos":  r.get("adultos", 0),
            "ninos":    r.get("ninos", 0),
            "tasas":    r.get("tasas", 0),
            "extras":   r.get("extras", ""),
        })
    return reservas, None

def load_limpiezas():
    cache, err = _load_cache()
    if err:
        return [], err
    rows = []
    for l in cache.get("limpiezas", []):
        rows.append({
            "apt":     l["apt"],
            "fecha":   parse_date(l["fecha"]),
            "montaje": l.get("montaje", ""),
            "dia":     l.get("dia", ""),
            "hecha":   l.get("hecha", False),
        })
    return rows, None

def cache_updated():
    cache, _ = _load_cache()
    return cache.get("updated") if cache else None


# ── Formatters ────────────────────────────────────────────────────────────────

def fmt_date(d):
    return d.strftime("%d/%m") if d else "—"

def apt_icon(apt):
    for k, v in APT_EMOJI.items():
        if k.lower() in apt.lower():
            return v
    return "🏡"

def plat_icon(p):
    return PLATFORM_EMOJI.get(p, "")


# ── Command handlers ──────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user    = update.effective_user
    chat_id = update.effective_chat.id

    # Auto-guardar el chat_id la primera vez
    cfg = load_config()
    if not cfg.get("chat_id"):
        cfg["chat_id"] = chat_id
        save_config(cfg)

    text = (
        f"👋 ¡Hola {user.first_name}! Soy el *Agente SWAPinn*.\n\n"
        f"🆔 Tu chat ID es: `{chat_id}`\n\n"
        "📋 *Comandos disponibles:*\n\n"
        "🗓 /hoy — entradas y salidas de hoy\n"
        "📆 /semana — próximos 7 días\n"
        "🧹 /limpiezas — próximas limpiezas\n"
        "💶 /ingresos — ingresos del mes actual\n"
        "📊 /resumen — resumen semanal completo\n"
        "ℹ️ /ayuda — esta ayuda\n"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await cmd_start(update, context)


async def cmd_hoy(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.date.today()
    reservas, err = load_reservas()
    if err:
        await update.message.reply_text(f"❌ Error leyendo reservas: {err}")
        return

    entradas  = [r for r in reservas if r["entrada"] == today]
    salidas   = [r for r in reservas if r["salida"]  == today]
    en_curso  = [r for r in reservas if r["entrada"] < today and (r["salida"] or today) > today]

    lines = [f"*📅 Hoy {today.strftime('%A %d/%m/%Y')}*\n"]

    if entradas:
        lines.append("✅ *Check-ins hoy:*")
        for r in sorted(entradas, key=lambda x: x["apt"]):
            icon = apt_icon(r["apt"])
            pax  = f"{r['adultos']}ad" + (f"+{r['ninos']}ni" if r["ninos"] else "")
            lines.append(f"  {icon} *{r['apt']}* — {r['nombre']} ({pax}, {r['noches']}n) {plat_icon(r['platform'])}")
    else:
        lines.append("✅ Sin check-ins hoy")

    lines.append("")

    if salidas:
        lines.append("🚪 *Check-outs hoy:*")
        for r in sorted(salidas, key=lambda x: x["apt"]):
            icon = apt_icon(r["apt"])
            lines.append(f"  {icon} *{r['apt']}* — {r['nombre']}")
    else:
        lines.append("🚪 Sin check-outs hoy")

    lines.append("")
    lines.append(f"🏡 *Apartamentos ocupados ahora:* {len(en_curso)}/5")

    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.MARKDOWN)


async def cmd_semana(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = datetime.date.today()
    fin   = today + datetime.timedelta(days=7)
    reservas, err = load_reservas()
    if err:
        await update.message.reply_text(f"❌ Error: {err}")
        return

    lines = [f"*📆 Próximos 7 días ({fmt_date(today)} – {fmt_date(fin)})*\n"]

    # Ocupación por apartamento
    lines.append("*Ocupación:*")
    for sheet in SHEETS:
        apt = sheet.replace(" 2026", "")
        icon = apt_icon(apt)
        apt_reservas = [r for r in reservas if r["apt"] == apt]
        # Días ocupados en la ventana
        occupied_days = 0
        for r in apt_reservas:
            if not r["entrada"] or not r["salida"]:
                continue
            overlap_start = max(r["entrada"], today)
            overlap_end   = min(r["salida"],  fin)
            if overlap_end > overlap_start:
                occupied_days += (overlap_end - overlap_start).days
        pct = int(occupied_days / 7 * 100)
        bar = "█" * (occupied_days) + "░" * (7 - occupied_days)
        lines.append(f"  {icon} {apt:<9} {bar} {pct}%")

    lines.append("")

    # Movimientos por día
    movimientos = {}
    for r in reservas:
        if r["entrada"] and today <= r["entrada"] <= fin:
            movimientos.setdefault(r["entrada"], []).append(("IN", r))
        if r["salida"] and today <= r["salida"] <= fin:
            movimientos.setdefault(r["salida"], []).append(("OUT", r))

    if movimientos:
        lines.append("*Movimientos:*")
        for dia in sorted(movimientos.keys()):
            lines.append(f"\n  📅 *{dia.strftime('%a %d/%m')}*")
            for tipo, r in movimientos[dia]:
                arrow = "→" if tipo == "IN" else "←"
                icon  = apt_icon(r["apt"])
                lines.append(f"    {arrow} {icon} {r['apt']} — {r['nombre']}")
    else:
        lines.append("Sin movimientos en los próximos 7 días.")

    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.MARKDOWN)


async def cmd_limpiezas(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today     = datetime.date.today()
    limpiezas, err = load_limpiezas()
    if err:
        await update.message.reply_text(f"❌ Error: {err}")
        return

    # Ventana: lo que queda de esta semana + semana completa siguiente
    fin_esta   = today + datetime.timedelta(days=(6 - today.weekday()))  # domingo de esta semana
    ini_prox   = fin_esta + datetime.timedelta(days=1)                    # lunes siguiente
    fin_prox   = ini_prox + datetime.timedelta(days=6)                   # domingo siguiente

    proximas = sorted(
        [l for l in limpiezas if not l["hecha"] and today <= l["fecha"] <= fin_prox],
        key=lambda x: x["fecha"],
    )

    if not proximas:
        await update.message.reply_text(
            f"✅ Sin limpiezas hasta el {fmt_date(fin_prox)}.", parse_mode=ParseMode.MARKDOWN
        )
        return

    lines = [f"*🧹 Limpiezas hasta el {fmt_date(fin_prox)}:*\n"]
    prev_fecha = None
    for l in proximas:
        if l["fecha"] != prev_fecha:
            dias_hasta = (l["fecha"] - today).days
            if   dias_hasta == 0: etiq = "🔴 HOY"
            elif dias_hasta == 1: etiq = "🟡 Mañana"
            elif dias_hasta <= 3: etiq = f"🟠 en {dias_hasta}d"
            else:                 etiq = f"⬜ en {dias_hasta}d"
            lines.append(f"\n*{l['dia']} {fmt_date(l['fecha'])}* {etiq}")
            prev_fecha = l["fecha"]
        icon    = apt_icon(l["apt"])
        montaje = f" · {l['montaje']}" if l["montaje"] else ""
        lines.append(f"  {icon} {l['apt']}{montaje}")

    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.MARKDOWN)


async def cmd_ingresos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today    = datetime.date.today()
    reservas, err = load_reservas()
    if err:
        await update.message.reply_text(f"❌ Error: {err}")
        return

    # Mes actual: todas las reservas que empiecen o estén en curso este mes
    mes_ini  = today.replace(day=1)
    mes_fin  = (mes_ini + datetime.timedelta(days=32)).replace(day=1)

    lines = [f"*💶 Ingresos — {today.strftime('%B %Y')}*\n"]
    gran_total = 0

    for sheet in SHEETS:
        apt = sheet.replace(" 2026", "")
        apt_r = [r for r in reservas if r["apt"] == apt and r["entrada"] and mes_ini <= r["entrada"] < mes_fin]
        if not apt_r:
            continue
        icon   = apt_icon(apt)
        total  = sum(r["total"] for r in apt_r)
        tasas  = sum(r["tasas"] for r in apt_r)
        netos  = total - tasas
        noches = sum(r["noches"] for r in apt_r)
        gran_total += netos
        lines.append(f"{icon} *{apt}* — {len(apt_r)} reservas, {noches} noches")
        lines.append(f"     Bruto: {total:,.0f}€  Tasas: {tasas:,.0f}€  *Neto: {netos:,.0f}€*")
        for r in sorted(apt_r, key=lambda x: x["entrada"]):
            pf = plat_icon(r["platform"])
            lines.append(f"     · {fmt_date(r['entrada'])}→{fmt_date(r['salida'])} {r['nombre']} {r['total']:,.0f}€ {pf}")
        lines.append("")

    if gran_total == 0 and not any(True for s in SHEETS for r in reservas if r["apt"] == s.replace(" 2026","") and r["entrada"] and mes_ini <= r["entrada"] < mes_fin):
        lines.append("Sin reservas registradas este mes.")
    else:
        lines.append(f"*TOTAL NETO: {gran_total:,.0f}€*")

    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.MARKDOWN)


async def cmd_resumen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Resumen semanal completo — el mismo que se envía automáticamente los lunes."""
    await update.message.reply_text("⏳ Generando resumen semanal...", parse_mode=ParseMode.MARKDOWN)
    text = build_resumen_semanal()
    # Telegram tiene límite de 4096 chars por mensaje
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)


def build_resumen_semanal():
    today    = datetime.date.today()
    lunes    = today - datetime.timedelta(days=today.weekday())
    domingo  = lunes + datetime.timedelta(days=6)
    prox_lun = lunes + datetime.timedelta(days=7)
    prox_dom = prox_lun + datetime.timedelta(days=6)

    reservas,  err1 = load_reservas()
    limpiezas, err2 = load_limpiezas()

    lines = [
        f"*📊 RESUMEN SEMANAL SWAPINN*",
        f"_Semana del {fmt_date(lunes)} al {fmt_date(domingo)}_\n",
    ]

    if err1:
        lines.append(f"❌ Error reservas: {err1}")
        return "\n".join(lines)

    # ── Ocupación esta semana ──────────────────────────────────────────────────
    lines.append("*🏡 Ocupación esta semana:*")
    for sheet in SHEETS:
        apt  = sheet.replace(" 2026", "")
        icon = apt_icon(apt)
        apt_r = [r for r in reservas if r["apt"] == apt]
        dias_ocu = 0
        for r in apt_r:
            if not r["entrada"] or not r["salida"]:
                continue
            overlap_s = max(r["entrada"], lunes)
            overlap_e = min(r["salida"],  domingo + datetime.timedelta(days=1))
            if overlap_e > overlap_s:
                dias_ocu += (overlap_e - overlap_s).days
        lines.append(f"  {icon} {apt:<9} {'█'*dias_ocu}{'░'*(7-dias_ocu)} {dias_ocu}/7 noches")
    lines.append("")

    # ── Movimientos semana actual ──────────────────────────────────────────────
    entradas_s = sorted([r for r in reservas if r["entrada"] and lunes <= r["entrada"] <= domingo], key=lambda x: x["entrada"])
    salidas_s  = sorted([r for r in reservas if r["salida"]  and lunes <= r["salida"]  <= domingo], key=lambda x: x["salida"])

    if entradas_s:
        lines.append(f"*✅ Check-ins esta semana ({len(entradas_s)}):*")
        for r in entradas_s:
            icon = apt_icon(r["apt"])
            pax  = f"{r['adultos']}ad" + (f"+{r['ninos']}ni" if r["ninos"] else "")
            lines.append(f"  {icon} {fmt_date(r['entrada'])} *{r['apt']}* — {r['nombre']} ({pax}, {r['noches']}n, {r['total']:,.0f}€) {plat_icon(r['platform'])}")
        lines.append("")

    if salidas_s:
        lines.append(f"*🚪 Check-outs esta semana ({len(salidas_s)}):*")
        for r in salidas_s:
            icon = apt_icon(r["apt"])
            lines.append(f"  {icon} {fmt_date(r['salida'])} *{r['apt']}* — {r['nombre']}")
        lines.append("")

    # ── Próxima semana ─────────────────────────────────────────────────────────
    entradas_p = sorted([r for r in reservas if r["entrada"] and prox_lun <= r["entrada"] <= prox_dom], key=lambda x: x["entrada"])
    lines.append(f"*📆 Próxima semana ({fmt_date(prox_lun)}–{fmt_date(prox_dom)}):*")
    if entradas_p:
        for r in entradas_p:
            icon = apt_icon(r["apt"])
            pax  = f"{r['adultos']}ad" + (f"+{r['ninos']}ni" if r["ninos"] else "")
            lines.append(f"  {icon} {fmt_date(r['entrada'])} *{r['apt']}* — {r['nombre']} ({pax}, {r['noches']}n)")
    else:
        lines.append("  Sin entradas registradas.")
    lines.append("")

    # ── Limpiezas próximas ─────────────────────────────────────────────────────
    if not err2:
        prox_limp = sorted([l for l in limpiezas if not l["hecha"] and l["fecha"] >= today][:8], key=lambda x: x["fecha"])
        if prox_limp:
            lines.append("*🧹 Próximas limpiezas:*")
            for l in prox_limp:
                icon    = apt_icon(l["apt"])
                dias    = (l["fecha"] - today).days
                urgente = " 🔴" if dias == 0 else (" 🟡" if dias == 1 else "")
                montaje = f" · {l['montaje']}" if l["montaje"] else ""
                lines.append(f"  {icon} {l['dia']} {fmt_date(l['fecha'])}{urgente} — {l['apt']}{montaje}")
            lines.append("")

    # ── Ingresos del mes ──────────────────────────────────────────────────────
    mes_ini = today.replace(day=1)
    mes_fin = (mes_ini + datetime.timedelta(days=32)).replace(day=1)
    apt_mes = [r for r in reservas if r["entrada"] and mes_ini <= r["entrada"] < mes_fin]
    if apt_mes:
        total_mes = sum(r["total"] - r["tasas"] for r in apt_mes)
        nreservas = len(apt_mes)
        nnoches   = sum(r["noches"] for r in apt_mes)
        lines.append(f"*💶 Ingresos {today.strftime('%B')}:* {nreservas} reservas · {nnoches} noches · *{total_mes:,.0f}€ neto*")
    lines.append("")
    lines.append(f"_Generado: {today.strftime('%d/%m/%Y')}_")

    return "\n".join(lines)


# ── Lenguaje natural ─────────────────────────────────────────────────────────

def find_claude_bin():
    path = shutil.which("claude")
    if path:
        return path
    base = Path.home() / "Library/Application Support/Claude/claude-code"
    if base.exists():
        for v in sorted(base.iterdir(), key=lambda p: p.name, reverse=True):
            candidate = v / "claude.app/Contents/MacOS/claude"
            if candidate.exists():
                return str(candidate)
    return None

CLAUDE_BIN = find_claude_bin()

def build_data_context(reservas, limpiezas):
    today = datetime.date.today()
    fin   = today + datetime.timedelta(days=30)
    lines = [f"Fecha actual: {today.strftime('%A %d/%m/%Y')}"]

    # Reservas próximas 30 días
    proximas = sorted(
        [r for r in reservas if r["entrada"] and r["entrada"] <= fin],
        key=lambda x: x["entrada"],
    )
    lines.append("\nRESERVAS (en curso y próximas 30 días):")
    for r in proximas:
        estado = "EN CURSO" if r["entrada"] <= today <= (r["salida"] or today) else ""
        pax = f"{r['adultos']}ad" + (f"+{r['ninos']}ni" if r["ninos"] else "")
        extras = f" extras:{r['extras']}" if r["extras"] else ""
        lines.append(
            f"  {r['apt']} | {r['entrada'].strftime('%d/%m')}→{r['salida'].strftime('%d/%m') if r['salida'] else '?'} "
            f"| {r['nombre']} | {pax} | {r['noches']}n | {r['total']:.0f}€ | {r['platform']}{extras} {estado}"
        )

    # Limpiezas próximas 2 semanas
    fin_limp  = today + datetime.timedelta(days=14)
    prox_limp = sorted([l for l in limpiezas if not l["hecha"] and today <= l["fecha"] <= fin_limp], key=lambda x: x["fecha"])
    if prox_limp:
        lines.append("\nLIMPIEZAS PRÓXIMAS (2 semanas):")
        for l in prox_limp:
            montaje = f" · {l['montaje']}" if l["montaje"] else ""
            lines.append(f"  {l['apt']} | {l['fecha'].strftime('%d/%m')} {l['dia']}{montaje}")

    # Ingresos mes actual
    mes_ini = today.replace(day=1)
    mes_fin = (mes_ini + datetime.timedelta(days=32)).replace(day=1)
    mes_r   = [r for r in reservas if r["entrada"] and mes_ini <= r["entrada"] < mes_fin]
    if mes_r:
        total_mes = sum(r["total"] - r["tasas"] for r in mes_r)
        lines.append(f"\nINGRESOS {today.strftime('%B').upper()}: {len(mes_r)} reservas · {sum(r['noches'] for r in mes_r)} noches · {total_mes:.0f}€ neto")
        for sheet in SHEETS:
            apt = sheet.replace(" 2026", "")
            apt_r = [r for r in mes_r if r["apt"] == apt]
            if apt_r:
                lines.append(f"  {apt}: {sum(r['total']-r['tasas'] for r in apt_r):.0f}€ neto ({len(apt_r)} reservas)")

    return "\n".join(lines)


def ask_claude(user_text, data_context):
    if not CLAUDE_BIN:
        logger.error("CLAUDE_BIN no encontrado")
        return "❌ No puedo procesar la pregunta ahora mismo (claude CLI no disponible)."
    logger.info(f"ask_claude: '{user_text[:60]}'")
    prompt = (
        "Eres el Agente SWAPinn, asistente de gestión de apartamentos turísticos de Alba. "
        "Los 5 apartamentos son: Sweet, Sunset, Center, Open Sky, Duplex.\n\n"
        f"DATOS ACTUALES:\n{data_context}\n\n"
        "Responde en español, de forma concisa y directa. "
        "Usa emojis. Sin Markdown especial — texto plano con saltos de línea.\n\n"
        f"Pregunta de Alba: {user_text}"
    )
    try:
        result = subprocess.run(
            [CLAUDE_BIN, "-p", prompt, "--dangerously-skip-permissions"],
            capture_output=True, text=True, timeout=90,
        )
        logger.info(f"ask_claude respuesta: {len(result.stdout)} chars, rc={result.returncode}")
        return result.stdout.strip() or "No he podido generar una respuesta."
    except subprocess.TimeoutExpired:
        logger.warning("ask_claude: timeout")
        return "⏳ Tardé demasiado en responder. Inténtalo de nuevo."
    except Exception as e:
        logger.exception("ask_claude error")
        return f"❌ Error: {e}"


async def cmd_mensaje_libre(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.reply_chat_action("typing")
        reservas,  _ = load_reservas()
        limpiezas, _ = load_limpiezas()
        data_context = build_data_context(reservas, limpiezas)

        import asyncio
        loop = asyncio.get_event_loop()
        respuesta = await loop.run_in_executor(
            None, ask_claude, update.message.text, data_context
        )

        for chunk in [respuesta[i:i+4000] for i in range(0, len(respuesta), 4000)]:
            try:
                await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
            except Exception:
                # Si falla el Markdown (caracteres especiales), enviamos sin formato
                await update.message.reply_text(chunk)
    except Exception as e:
        logger.exception("Error en cmd_mensaje_libre")
        try:
            await update.message.reply_text(f"❌ Error inesperado: {e}")
        except Exception:
            pass


async def cmd_unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "No reconozco ese comando. Escribe /ayuda para ver los disponibles."
    )


# ── Envío automático (para cron) ──────────────────────────────────────────────

async def enviar_resumen_automatico(app):
    """Llamado por el job semanal — envía resumen al chat_id guardado."""
    cfg = load_config()
    chat_id = cfg.get("chat_id")
    if not chat_id:
        logger.warning("No hay chat_id guardado, no se puede enviar resumen automático.")
        return
    text = build_resumen_semanal()
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        await app.bot.send_message(chat_id=chat_id, text=chunk, parse_mode=ParseMode.MARKDOWN)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    cfg = load_config()
    token = cfg.get("telegram_token")
    if not token:
        raise ValueError("Falta telegram_token en config.json")

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start",     cmd_start))
    app.add_handler(CommandHandler("ayuda",     cmd_ayuda))
    app.add_handler(CommandHandler("hoy",       cmd_hoy))
    app.add_handler(CommandHandler("semana",    cmd_semana))
    app.add_handler(CommandHandler("limpiezas", cmd_limpiezas))
    app.add_handler(CommandHandler("ingresos",  cmd_ingresos))
    app.add_handler(CommandHandler("resumen",   cmd_resumen))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, cmd_mensaje_libre))
    app.add_handler(MessageHandler(filters.COMMAND, cmd_unknown))

    # Resumen automático los lunes a las 8:00
    job_queue = app.job_queue
    now = datetime.datetime.now()
    days_until_monday = (7 - now.weekday()) % 7
    first_monday = now.replace(hour=8, minute=0, second=0, microsecond=0) + datetime.timedelta(days=days_until_monday if days_until_monday > 0 else 7)
    job_queue.run_repeating(
        lambda ctx: ctx.application.create_task(enviar_resumen_automatico(ctx.application)),
        interval=datetime.timedelta(weeks=1),
        first=first_monday,
        name="resumen_semanal",
    )

    logger.info("🤖 Agente SWAPinn arrancado. Ctrl+C para parar.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

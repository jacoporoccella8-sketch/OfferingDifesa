"""
Defence Market Daily Scraper - NTT DATA Italia
==============================================

Market analysis automatizzata sul segmento Difesa italiano.

Lo script interroga l'API Anthropic (con web search integrato) per raccogliere,
per ogni player monitorato: fatturato, stima spesa IT, servizi IT acquistati,
piano strategico digitale, storico gare ANAC, trend futuri e note utili per
NTT DATA. Genera un Excel con tre fogli (Dashboard, Storico Gare, Dettaglio
Player), calcola il delta rispetto al giorno precedente e invia una mail HTML
riassuntiva con l'Excel in allegato.

VINCOLO TIER 1 (budget ~5 USD): l'account ha rate limit stringenti. Per questo
lo script NON interroga tutti gli 11 player in un singolo run, ma adotta uno
SCAGLIONAMENTO A ROTAZIONE: processa solo PLAYERS_PER_RUN player per esecuzione
(default 3), tenendo traccia dell'indice in data/rotation_state.json. In 4
giorni copre tutti gli 11 player, poi ricomincia aggiornando i dati. I player
non processati in un dato giorno mantengono l'ultimo dato valido salvato in
data/history.json, cosi ogni report contiene comunque il quadro completo.
"""

import os
import sys
import json
import time
import smtplib
import logging
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

from dotenv import load_dotenv
from anthropic import Anthropic

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configurazione generale
# ---------------------------------------------------------------------------

load_dotenv()

# Stringa modello VALIDA E ATTUALE per la Messages API.
# NON usare claude-sonnet-4-20250514 (deprecato, errore 404).
MODEL = "claude-sonnet-4-6"

# Scaglionamento a rotazione: quanti player processare per esecuzione.
# Tier 1 budget limitato => 3 player per run. Alzare questo valore se si passa
# a un tier API superiore (es. 11 per processare tutti i player in un solo run).
PLAYERS_PER_RUN = 3

# Pausa tra un player e l'altro (secondi).
PAUSE_BETWEEN_PLAYERS = 10

# Retry con backoff esponenziale sul 429.
MAX_RETRIES = 3
BACKOFF_SECONDS = [20, 40, 80]

# Numero massimo di ricerche web per player (compromesso qualita/quota).
WEB_SEARCH_MAX_USES = 4

# Percorsi file di stato e output.
DATA_DIR = "data"
OUTPUT_DIR = "output"
HISTORY_PATH = os.path.join(DATA_DIR, "history.json")
ROTATION_STATE_PATH = os.path.join(DATA_DIR, "rotation_state.json")

# Lista player monitorati (ordine fisso, non modificare l'ordine).
PLAYERS = [
    {"name": "Ministero della Difesa", "market": "Pubblico", "complexity": "Molto alta"},
    {"name": "Forze Armate (Esercito/Marina/AM)", "market": "Pubblico", "complexity": "Alta"},
    {"name": "Difesa Servizi", "market": "Pubblico", "complexity": "Media"},
    {"name": "Segretariato Generale Difesa/DNA", "market": "Pubblico", "complexity": "Alta"},
    {"name": "Leonardo", "market": "Privato", "complexity": "Alta"},
    {"name": "Fincantieri", "market": "Privato", "complexity": "Alta"},
    {"name": "MBDA Italia", "market": "Privato", "complexity": "Alta"},
    {"name": "Elettronica Group", "market": "Privato", "complexity": "Alta"},
    {"name": "Thales Alenia Space Italia", "market": "Privato", "complexity": "Alta"},
    {"name": "Avio Aero", "market": "Privato", "complexity": "Alta"},
    {"name": "Leonardo DRS", "market": "Privato", "complexity": "Alta"},
]


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("defence_scraper")


# ---------------------------------------------------------------------------
# Persistenza stato (history + rotazione)
# ---------------------------------------------------------------------------

def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (json.JSONDecodeError, OSError) as exc:
        logger.error("Impossibile leggere %s: %s. Uso il default.", path, exc)
        return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def load_history():
    return load_json(HISTORY_PATH, {})


def save_history(history):
    save_json(HISTORY_PATH, history)


def load_rotation_index():
    state = load_json(ROTATION_STATE_PATH, {"next_index": 0})
    idx = state.get("next_index", 0)
    if not isinstance(idx, int) or idx < 0 or idx >= len(PLAYERS):
        idx = 0
    return idx


def save_rotation_index(next_index):
    save_json(ROTATION_STATE_PATH, {
        "next_index": next_index,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    })


def select_players_for_run(start_index):
    """Restituisce gli indici dei player da processare in questo run e il
    prossimo indice di partenza, gestendo il wrap-around circolare."""
    selected = []
    idx = start_index
    for _ in range(min(PLAYERS_PER_RUN, len(PLAYERS))):
        selected.append(idx)
        idx = (idx + 1) % len(PLAYERS)
    return selected, idx


# ---------------------------------------------------------------------------
# Interrogazione API Anthropic con web search e loop multi-turno
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = (
    "Sei un analista di mercato senior specializzato nel segmento Difesa "
    "italiano. Rispondi sempre in italiano. Usa la ricerca web per trovare "
    "dati aggiornati e verificabili, citando le fonti. Quando un dato non e "
    "reperibile, dichiaralo esplicitamente come 'N/D' senza inventare numeri."
)

# Schema JSON atteso dal modello, descritto in linguaggio naturale nel prompt.
PLAYER_PROMPT_TEMPLATE = (
    "Effettua una market analysis sul seguente player del segmento Difesa "
    "italiano per conto di NTT DATA Italia, che vuole aggredire questo "
    "mercato con un cluster Difesa unificato.\n\n"
    "PLAYER: {name}\n"
    "MERCATO: {market}\n"
    "COMPLESSITA PROCUREMENT: {complexity}\n\n"
    "Cerca attivamente sul web e raccogli i seguenti dati:\n"
    "1. Fatturato piu recente disponibile (con anno).\n"
    "2. Stima della spesa IT annua, con la fonte o la metodologia di stima.\n"
    "3. Servizi IT acquistati o di interesse (cloud, cybersecurity, system "
    "integration, software, ecc.).\n"
    "4. Piano strategico digitale o di trasformazione IT, se esiste.\n"
    "5. Storico delle gare/contratti rilevanti (preferibilmente da banca dati "
    "ANAC) con CIG, importo e aggiudicatario quando disponibili.\n"
    "6. Trend futuri attesi (investimenti, programmi, digitalizzazione).\n"
    "7. Note utili per NTT DATA (opportunita, punti di ingresso, rischi).\n\n"
    "Rispondi ESCLUSIVAMENTE con un oggetto JSON valido, senza testo prima o "
    "dopo, senza backtick e senza markdown. Struttura esatta:\n"
    "{{\n"
    '  "fatturato": "stringa",\n'
    '  "spesa_it_stimata": "stringa",\n'
    '  "fonte_stima": "stringa",\n'
    '  "servizi_it_acquistati": "stringa",\n'
    '  "piano_strategico_it": "stringa",\n'
    '  "trend_futuro": "stringa",\n'
    '  "note_ntt_data": "stringa",\n'
    '  "gare": [\n'
    '    {{"cig": "stringa", "oggetto": "stringa", "importo": "stringa", '
    '"aggiudicatario": "stringa", "anno": "stringa"}}\n'
    "  ]\n"
    "}}\n"
    "Se un campo non e reperibile usa il valore 'N/D'. Se non trovi gare usa "
    "una lista vuota []."
)


def extract_text_from_content(content_blocks):
    """Concatena tutti i blocchi di testo da una risposta dell'API."""
    parts = []
    for block in content_blocks:
        if getattr(block, "type", None) == "text":
            parts.append(block.text)
    return "\n".join(parts).strip()


def parse_player_json(raw_text):
    """Estrae l'oggetto JSON dalla risposta del modello in modo robusto."""
    text = raw_text.strip()
    # Rimuove eventuali fence markdown residui.
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    # Isola il primo oggetto JSON ben formato.
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Nessun oggetto JSON trovato nella risposta")
    candidate = text[start:end + 1]
    return json.loads(candidate)


def call_api_for_player(client, player):
    """Interroga l'API per un singolo player gestendo il loop multi-turno
    (tool_use -> tool_result -> end_turn) del web search server-side.

    Con il tool web_search server-side di Anthropic l'esecuzione della ricerca
    avviene lato server: il loop tipico termina con stop_reason 'end_turn' in
    una sola chiamata, ma gestiamo comunque iterazioni multiple per robustezza.
    Solleva l'eccezione originale in caso di errore, cosi il chiamante puo
    classificarla (429, timeout, ecc.)."""

    prompt = PLAYER_PROMPT_TEMPLATE.format(
        name=player["name"],
        market=player["market"],
        complexity=player["complexity"],
    )

    messages = [{"role": "user", "content": prompt}]
    tools = [{
        "type": "web_search_20250305",
        "name": "web_search",
        "max_uses": WEB_SEARCH_MAX_USES,
    }]

    max_turns = 8
    for turn in range(max_turns):
        response = client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            tools=tools,
            messages=messages,
        )

        if response.stop_reason == "end_turn" or response.stop_reason == "stop_sequence":
            return extract_text_from_content(response.content)

        if response.stop_reason == "tool_use":
            # Il web search e server-side: i risultati arrivano gia nel content.
            # Accodiamo la risposta dell'assistant e proseguiamo il loop, cosi
            # il modello puo continuare a ricercare o produrre la risposta finale.
            messages.append({"role": "assistant", "content": response.content})
            # Per il web search server-side non dobbiamo costruire tool_result
            # manualmente: il server li inserisce. Lasciamo che il modello
            # prosegua con un turno di continuazione neutro.
            messages.append({
                "role": "user",
                "content": "Prosegui e fornisci la risposta finale in JSON.",
            })
            continue

        if response.stop_reason == "max_tokens":
            # Risposta troncata: proviamo comunque a usare il testo disponibile.
            return extract_text_from_content(response.content)

        # Stop reason inatteso: restituiamo quel che abbiamo.
        return extract_text_from_content(response.content)

    # Esaurito il numero massimo di turni.
    return extract_text_from_content(response.content)


def classify_exception(exc):
    """Classifica un'eccezione API per il logging dell'errore vero."""
    status = getattr(exc, "status_code", None)
    name = exc.__class__.__name__
    if status == 429 or "rate_limit" in str(exc).lower() or "RateLimit" in name:
        return "429 rate_limit_error"
    if "timeout" in str(exc).lower() or "Timeout" in name:
        return "timeout"
    if status == 404 or "not_found" in str(exc).lower():
        return "404 model/endpoint not found"
    if status is not None:
        return "HTTP %s" % status
    return name


def is_rate_limit(exc):
    status = getattr(exc, "status_code", None)
    name = exc.__class__.__name__
    return status == 429 or "rate_limit" in str(exc).lower() or "RateLimit" in name


def fetch_player_data(client, player):
    """Recupera i dati di un player con retry+backoff sul 429.

    Restituisce una tupla (data_dict_or_None, status_string).
    status_string descrive esplicitamente l'esito per il logging."""

    attempt = 0
    while True:
        try:
            raw = call_api_for_player(client, player)
            try:
                data = parse_player_json(raw)
            except (ValueError, json.JSONDecodeError) as parse_exc:
                logger.error(
                    "Player '%s': JSON non parsato (%s). Risposta troncata: %.200s",
                    player["name"], parse_exc, raw,
                )
                return None, "JSON non parsato"
            logger.info("Player '%s': SUCCESSO", player["name"])
            return data, "successo"

        except Exception as exc:  # noqa: BLE001 - vogliamo classificare tutto
            kind = classify_exception(exc)
            if is_rate_limit(exc) and attempt < MAX_RETRIES:
                wait = BACKOFF_SECONDS[min(attempt, len(BACKOFF_SECONDS) - 1)]
                logger.error(
                    "Player '%s': %s (tentativo %d/%d). Attendo %ds e riprovo.",
                    player["name"], kind, attempt + 1, MAX_RETRIES, wait,
                )
                time.sleep(wait)
                attempt += 1
                continue
            # Errore non recuperabile o retry esauriti.
            logger.error(
                "Player '%s': ERRORE definitivo dopo %d tentativi: %s",
                player["name"], attempt, kind,
            )
            return None, kind


# ---------------------------------------------------------------------------
# Normalizzazione record e calcolo delta
# ---------------------------------------------------------------------------

def build_record(player, data, status, timestamp):
    """Costruisce un record normalizzato per lo storico."""
    gare = data.get("gare", []) if isinstance(data, dict) else []
    if not isinstance(gare, list):
        gare = []
    return {
        "name": player["name"],
        "market": player["market"],
        "complexity": player["complexity"],
        "fatturato": (data or {}).get("fatturato", "N/D"),
        "spesa_it_stimata": (data or {}).get("spesa_it_stimata", "N/D"),
        "fonte_stima": (data or {}).get("fonte_stima", "N/D"),
        "servizi_it_acquistati": (data or {}).get("servizi_it_acquistati", "N/D"),
        "piano_strategico_it": (data or {}).get("piano_strategico_it", "N/D"),
        "trend_futuro": (data or {}).get("trend_futuro", "N/D"),
        "note_ntt_data": (data or {}).get("note_ntt_data", "N/D"),
        "gare": gare,
        "n_gare": len(gare),
        "status": status,
        "updated_at": timestamp,
    }


def compute_delta(old_record, new_record):
    """Calcola un testo che descrive le variazioni rispetto al giorno prima."""
    if old_record is None:
        return "Primo rilevamento"
    changes = []
    for field, label in [
        ("spesa_it_stimata", "Spesa IT"),
        ("fatturato", "Fatturato"),
        ("trend_futuro", "Trend"),
    ]:
        old_val = str(old_record.get(field, "N/D"))
        new_val = str(new_record.get(field, "N/D"))
        if old_val != new_val:
            changes.append(label)
    old_gare = old_record.get("n_gare", 0)
    new_gare = new_record.get("n_gare", 0)
    if new_gare != old_gare:
        diff = new_gare - old_gare
        sign = "+" if diff > 0 else ""
        changes.append("Gare (%s%d)" % (sign, diff))
    if not changes:
        return "Nessuna variazione"
    return "Variazioni: " + ", ".join(changes)


# ---------------------------------------------------------------------------
# Generazione Excel (Dashboard, Storico Gare, Dettaglio Player)
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(color="1F3864", bold=True, size=14)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")


def style_header_row(ws, row_idx, n_cols):
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER


def autosize_columns(ws, widths):
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def build_excel(records_by_name, deltas_by_name, run_date):
    wb = openpyxl.Workbook()

    # --- Foglio 1: Dashboard ---
    ws = wb.active
    ws.title = "Dashboard"
    ws.cell(row=1, column=1, value="Market Analysis Difesa Italia - NTT DATA").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Aggiornato al %s" % run_date).font = Font(italic=True, color="595959")

    headers = [
        "Player", "Mercato", "Complessita Procurement", "Fatturato",
        "Spesa IT Stimata", "Fonte Stima", "Servizi IT Acquistati",
        "Piano Strategico IT", "N. Gare Trovate", "Trend Futuro",
        "Delta vs Ieri", "Note per NTT DATA",
    ]
    header_row = 4
    for col, head in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=head)
    style_header_row(ws, header_row, len(headers))

    row = header_row + 1
    for player in PLAYERS:
        rec = records_by_name.get(player["name"])
        if rec is None:
            rec = build_record(player, {}, "nessun dato storico", run_date)
        values = [
            rec["name"], rec["market"], rec["complexity"], rec["fatturato"],
            rec["spesa_it_stimata"], rec["fonte_stima"],
            rec["servizi_it_acquistati"], rec["piano_strategico_it"],
            rec["n_gare"], rec["trend_futuro"],
            deltas_by_name.get(player["name"], "N/D"), rec["note_ntt_data"],
        ]
        for col, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = WRAP
            cell.border = BORDER
        row += 1

    autosize_columns(ws, [26, 12, 14, 22, 22, 24, 30, 30, 12, 30, 22, 36])
    ws.freeze_panes = "A%d" % (header_row + 1)

    # --- Foglio 2: Storico Gare ---
    ws2 = wb.create_sheet("Storico Gare")
    ws2.cell(row=1, column=1, value="Storico Gare ANAC e contratti rilevanti").font = TITLE_FONT
    gare_headers = ["Player", "CIG", "Oggetto", "Importo", "Aggiudicatario", "Anno"]
    for col, head in enumerate(gare_headers, start=1):
        ws2.cell(row=3, column=col, value=head)
    style_header_row(ws2, 3, len(gare_headers))

    grow = 4
    any_gara = False
    for player in PLAYERS:
        rec = records_by_name.get(player["name"])
        if not rec:
            continue
        for gara in rec.get("gare", []):
            any_gara = True
            values = [
                rec["name"],
                gara.get("cig", "N/D"),
                gara.get("oggetto", "N/D"),
                gara.get("importo", "N/D"),
                gara.get("aggiudicatario", "N/D"),
                gara.get("anno", "N/D"),
            ]
            for col, val in enumerate(values, start=1):
                cell = ws2.cell(row=grow, column=col, value=val)
                cell.alignment = WRAP
                cell.border = BORDER
            grow += 1
    if not any_gara:
        ws2.cell(row=4, column=1, value="Nessuna gara trovata nei dati disponibili.")
    autosize_columns(ws2, [26, 20, 44, 18, 30, 8])
    ws2.freeze_panes = "A4"

    # --- Foglio 3: Dettaglio Player ---
    ws3 = wb.create_sheet("Dettaglio Player")
    drow = 1
    field_labels = [
        ("market", "Mercato"),
        ("complexity", "Complessita Procurement"),
        ("fatturato", "Fatturato"),
        ("spesa_it_stimata", "Spesa IT Stimata"),
        ("fonte_stima", "Fonte Stima"),
        ("servizi_it_acquistati", "Servizi IT Acquistati"),
        ("piano_strategico_it", "Piano Strategico IT"),
        ("trend_futuro", "Trend Futuro"),
        ("note_ntt_data", "Note per NTT DATA"),
        ("n_gare", "N. Gare Trovate"),
        ("status", "Esito ultimo rilevamento"),
        ("updated_at", "Ultimo aggiornamento"),
    ]
    for player in PLAYERS:
        rec = records_by_name.get(player["name"])
        if rec is None:
            rec = build_record(player, {}, "nessun dato storico", run_date)
        title_cell = ws3.cell(row=drow, column=1, value=rec["name"])
        title_cell.font = TITLE_FONT
        ws3.merge_cells(start_row=drow, start_column=1, end_row=drow, end_column=2)
        drow += 1
        for key, label in field_labels:
            lc = ws3.cell(row=drow, column=1, value=label)
            lc.font = Font(bold=True)
            lc.alignment = Alignment(vertical="top")
            lc.border = BORDER
            vc = ws3.cell(row=drow, column=2, value=rec.get(key, "N/D"))
            vc.alignment = WRAP
            vc.border = BORDER
            drow += 1
        drow += 1  # riga vuota tra player
    autosize_columns(ws3, [28, 70])

    filename = "market_analysis_difesa_%s.xlsx" % run_date
    path = os.path.join(OUTPUT_DIR, filename)
    wb.save(path)
    logger.info("Excel generato: %s", path)
    return path


# ---------------------------------------------------------------------------
# Invio mail HTML
# ---------------------------------------------------------------------------

def build_email_html(records_by_name, deltas_by_name, run_date, kpis):
    rows_html = []
    for player in PLAYERS:
        rec = records_by_name.get(player["name"])
        if rec is None:
            rec = build_record(player, {}, "nessun dato storico", run_date)
        rows_html.append(
            "<tr>"
            "<td style='padding:8px;border:1px solid #ddd;'>%s</td>"
            "<td style='padding:8px;border:1px solid #ddd;'>%s</td>"
            "<td style='padding:8px;border:1px solid #ddd;'>%s</td>"
            "<td style='padding:8px;border:1px solid #ddd;text-align:center;'>%s</td>"
            "<td style='padding:8px;border:1px solid #ddd;'>%s</td>"
            "</tr>" % (
                rec["name"], rec["market"], rec["spesa_it_stimata"],
                rec["n_gare"], deltas_by_name.get(player["name"], "N/D"),
            )
        )

    kpi_box = (
        "<div style='display:inline-block;width:23%%;min-width:140px;margin:6px;"
        "padding:16px;background:#1F3864;color:#fff;border-radius:8px;"
        "text-align:center;vertical-align:top;'>"
        "<div style='font-size:28px;font-weight:bold;'>%s</div>"
        "<div style='font-size:12px;opacity:0.9;'>%s</div></div>"
    )
    kpi_html = "".join(
        kpi_box % (value, label) for label, value in kpis
    )

    return (
        "<html><body style='font-family:Arial,sans-serif;color:#222;'>"
        "<h2 style='color:#1F3864;'>Market Analysis Difesa Italia - NTT DATA</h2>"
        "<p>Report automatico del <strong>%s</strong>. In allegato l'Excel completo "
        "(Dashboard, Storico Gare, Dettaglio Player).</p>"
        "<div style='text-align:center;margin:16px 0;'>%s</div>"
        "<h3 style='color:#1F3864;'>Quadro player</h3>"
        "<table style='border-collapse:collapse;width:100%%;font-size:13px;'>"
        "<thead><tr style='background:#1F3864;color:#fff;'>"
        "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Player</th>"
        "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Mercato</th>"
        "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Spesa IT Stimata</th>"
        "<th style='padding:8px;border:1px solid #ddd;'>N. Gare</th>"
        "<th style='padding:8px;border:1px solid #ddd;text-align:left;'>Delta vs Ieri</th>"
        "</tr></thead><tbody>%s</tbody></table>"
        "<p style='font-size:11px;color:#888;margin-top:20px;'>Scaglionamento a "
        "rotazione attivo: %d player processati per run, ciclo di 4 giorni per "
        "coprire tutti gli 11. I player non aggiornati oggi riportano l'ultimo "
        "dato valido.</p>"
        "</body></html>" % (
            run_date, kpi_html, "".join(rows_html), PLAYERS_PER_RUN,
        )
    )


def send_email(html_body, attachment_path, run_date):
    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ["SMTP_USER"]
    password = os.environ["SMTP_PASS"]
    recipient = os.environ.get("RECIPIENT_EMAIL", "jacopo.roccella@nttdata.com")

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = recipient
    msg["Subject"] = "Market Analysis Difesa Italia - %s" % run_date
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with open(attachment_path, "rb") as fh:
        part = MIMEApplication(fh.read(), _subtype="xlsx")
    part.add_header(
        "Content-Disposition", "attachment",
        filename=os.path.basename(attachment_path),
    )
    msg.attach(part)

    logger.info("Invio mail a %s via %s:%d", recipient, host, port)
    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(user, password)
        server.sendmail(user, [recipient], msg.as_string())
    logger.info("Mail inviata correttamente.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ensure_dirs()
    run_date = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    timestamp = datetime.now(timezone.utc).isoformat()

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logger.error("ANTHROPIC_API_KEY mancante. Impossibile proseguire.")
        sys.exit(1)

    client = Anthropic(api_key=api_key)

    history = load_history()
    start_index = load_rotation_index()
    selected_indices, next_index = select_players_for_run(start_index)

    selected_names = [PLAYERS[i]["name"] for i in selected_indices]
    logger.info(
        "Run del %s. Scaglionamento a rotazione: processo %d player: %s",
        run_date, len(selected_indices), ", ".join(selected_names),
    )

    fresh_success = 0
    errors = 0

    for pos, idx in enumerate(selected_indices):
        player = PLAYERS[idx]
        logger.info("Interrogo player %d/%d: %s",
                    pos + 1, len(selected_indices), player["name"])

        data, status = fetch_player_data(client, player)

        if data is not None:
            new_record = build_record(player, data, status, timestamp)
            old_record = history.get(player["name"])
            new_record["delta"] = compute_delta(old_record, new_record)
            history[player["name"]] = new_record
            fresh_success += 1
        else:
            errors += 1
            logger.error("Player '%s' non aggiornato in questo run (%s). "
                         "Mantengo l'ultimo dato storico se presente.",
                         player["name"], status)

        if pos < len(selected_indices) - 1:
            logger.info("Pausa di %d secondi prima del prossimo player.",
                        PAUSE_BETWEEN_PLAYERS)
            time.sleep(PAUSE_BETWEEN_PLAYERS)

    # Aggiorna lo stato della rotazione SOLO dopo aver tentato il run.
    save_rotation_index(next_index)
    save_history(history)

    # Costruisce le strutture per Excel e mail usando lo storico completo
    # (player processati oggi + ultimo dato valido degli altri).
    records_by_name = {}
    deltas_by_name = {}
    players_with_it_estimate = 0
    total_gare = 0
    total_changes = 0

    for player in PLAYERS:
        rec = history.get(player["name"])
        if rec is None:
            continue
        records_by_name[player["name"]] = rec
        delta = rec.get("delta")
        if delta is None:
            delta = "Dato storico (non aggiornato oggi)"
        deltas_by_name[player["name"]] = delta
        if rec.get("spesa_it_stimata", "N/D") not in ("N/D", "", None):
            players_with_it_estimate += 1
        total_gare += rec.get("n_gare", 0)
        if delta.startswith("Variazioni"):
            total_changes += 1

    # GUARDIA SULL'INVIO MAIL: se nessun player ha dati validi (ne freschi ne
    # storici), non inviare nulla ed esci in errore.
    if not records_by_name:
        logger.error(
            "ZERO player con dati validi (freschi: %d, errori: %d, storico vuoto). "
            "Non invio la mail. Esco con codice 1 per segnalare l'anomalia.",
            fresh_success, errors,
        )
        sys.exit(1)

    if fresh_success == 0:
        logger.error(
            "Nessun player aggiornato con successo in questo run (errori: %d). "
            "Procedo comunque con i dati storici disponibili per %d player.",
            errors, len(records_by_name),
        )

    kpis = [
        ("Player monitorati", len(records_by_name)),
        ("Con stima spesa IT", players_with_it_estimate),
        ("Gare trovate", total_gare),
        ("Variazioni vs ieri", total_changes),
    ]

    excel_path = build_excel(records_by_name, deltas_by_name, run_date)
    html_body = build_email_html(records_by_name, deltas_by_name, run_date, kpis)

    try:
        send_email(html_body, excel_path, run_date)
    except Exception as exc:  # noqa: BLE001
        logger.error("Invio mail fallito: %s", exc)
        sys.exit(1)

    logger.info(
        "Run completato. Player freschi: %d, errori: %d, totale in report: %d.",
        fresh_success, errors, len(records_by_name),
    )


if __name__ == "__main__":
    main()

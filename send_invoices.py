import win32com.client
from datetime import datetime
import time
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import re
from collections import defaultdict, OrderedDict


# --- Helpers ---
def wait_ready(session, timeout=60.0, poll=0.1):
    """Wait until SAP session is not busy or timeout."""
    t0 = time.time()
    while True:
        try:
            if not session.Busy:
                return
        except Exception:
            pass
        if time.time() - t0 > timeout:
            raise TimeoutError("SAP session stayed busy for too long.")
        time.sleep(poll)

def press_with_tooltip(session, btn_id: str, expected_tooltip_substring: str):
    """Verify tooltip contains expected text, then press."""
    btn = session.findById(btn_id)
    tip = getattr(btn, "Tooltip", "") or getattr(btn, "toolTip", "")
    if expected_tooltip_substring not in tip:
        raise RuntimeError(f"Unexpected tooltip for {btn_id}. Got: '{tip}'  Expected to contain: '{expected_tooltip_substring}'")
    btn.press()
    
def norm_header(s: str) -> str:
        # Normalize headers like "Opret. d." -> "Opret. d." (keep dots), but trim/space-normalize
        return " ".join((s or "").strip().split())

def send_invoice(orchestrator_connection: OrchestratorConnection):
    # --- SAP session ---
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    # --- Get robot username from Orchestrator (you already have this available) ---
    RobotCredential = orchestrator_connection.get_credential("OpusBruger")
    RobotUsername = RobotCredential.username

    # --- Go to ZVF04 ---
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZVF04"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)  # Enter
    wait_ready(session)

    # --- Fill fields ---
    today = datetime.today().strftime("%d.%m.%Y")  # dd.MM.yyyy
    date_field = session.findById("wnd[0]/usr/ctxtP_FKDAT")
    date_field.text = today
    date_field.caretPosition = len(today)

    user_field = session.findById("wnd[0]/usr/txtS_ERNAM-LOW")
    user_field.text = RobotUsername
    user_field.caretPosition = len(RobotUsername)

    # Press the "Execute/Check" type button (btn[8]) on the app toolbar
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    wait_ready(session)

    # --- Verify and press "Marker alle (F5)" -> btn[5] ---
    press_with_tooltip(session, "wnd[0]/tbar[1]/btn[5]", "Marker alle   (F5)")
    wait_ready(session)

    # --- Verify and press "Gem (Ctrl+S)" -> btn[11] ---
    press_with_tooltip(session, "wnd[0]/tbar[0]/btn[11]", "Gem   (Ctrl+S)")
    wait_ready(session)

    container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")

    LBL_RE = re.compile(r".*/lbl\[(\d+),(\d+)\]$")


    cells = []          # (col:int, row:int, text:str, id:str)
    non_table_labels = []

    for child in container.Children:
        if "lbl" in child.Id:
            m = LBL_RE.match(child.Id)
            if not m:
                # Label under usr but not in the lbl[i,j] grid -> not part of table
                non_table_labels.append((child.Id, getattr(child, "Text", "").strip()))
                continue
            col, row = int(m.group(1)), int(m.group(2))
            try:
                text = (child.Text or "").strip()
            except Exception:
                text = ""
            cells.append((col, row, text, child.Id))

    # Make sure we only have table labels 
    if non_table_labels:
        raise RuntimeError(
            "Der findes labels udenfor tabellen (ikke i formatet lbl[col,row]): "
            + ", ".join(f"{i}='{t}'" for i,t in non_table_labels[:10])
            + (" ..." if len(non_table_labels) > 10 else "")
        )

    # Build header (row==1) and data rows (row>=3, odd)
    headers = {col: norm_header(text) for col, row, text, _ in cells if row == 1}
    if not headers:
        raise RuntimeError("Ingen tabel-headers (row=1) fundet.")

    data_cells = [(col, row, text) for col, row, text, _ in cells if row >= 3 and row % 2 == 1]

    if not data_cells:
        # Still OK if there are zero data rows; we’ll just validate Fejl header exists and is empty-by-definition
        pass

    # Sanity checks: rows only 1 or odd >=3 (else it's probably not the table)
    unexpected = [(c, r, t) for c, r, t, _ in cells if not (r == 1 or (r >= 3 and r % 2 == 1))]
    if unexpected:
        raise RuntimeError(
            "Uventede label-rækker (ikke row=1 eller en ulige række >=3): "
            + ", ".join(f"[{c},{r}]='{t}'" for c, r, t in unexpected[:10])
            + (" ..." if len(unexpected) > 10 else "")
        )

    # Order columns and build a column name list
    sorted_cols = sorted(headers.keys())
    column_names = [headers[c] for c in sorted_cols]

    # Find Fejl column (case-insensitive, punctuation tolerant)
    def is_fejl(h):
        return h.lower().strip(".:") == "fejl"

    fejl_col = None
    for c in sorted_cols:
        if is_fejl(headers[c]):
            fejl_col = c
            break

    if fejl_col is None:
        raise RuntimeError("Kolonnen 'Fejl' blev ikke fundet i header-rækken.")

    # Group data by row index and map to {header: value}


    rows_by_index = defaultdict(dict)
    for col, row, text in data_cells:
        rows_by_index[row][col] = text

    # Convert to list of dicts in row order
    table_rows = []
    for row_idx in sorted(rows_by_index.keys()):
        rowmap = rows_by_index[row_idx]
        record = OrderedDict()
        for c in sorted_cols:
            record[headers[c]] = rowmap.get(c, "")
        table_rows.append(record)

    # Validate: every Fejl cell must be empty string
    bad_fejl = []
    for i, rec in enumerate(table_rows, start=1):
        val = (rec.get(headers[fejl_col]) or "").strip()
        if val != "":
            bad_fejl.append((i, val))

    if bad_fejl:
        # Build a helpful error message with first few offending rows
        preview = ", ".join(f"række {i}: '{v}'" for i, v in bad_fejl[:10])
        raise RuntimeError(
            f"Fejl-kolonnen skal være tom i alle rækker, men fandt værdier: {preview}"
            + (" ..." if len(bad_fejl) > 10 else "")
        )

    print("Tabel verificeret (kun grid-labels, korrekt header/data-rækker).")
    print(f"Kolonner: {column_names}")
    print(f"Antal rækker: {len(table_rows)}")
    # Access example: print each row
    for idx, rec in enumerate(table_rows,  start=1):
        print(f"Row {idx}: {dict(rec)}")
    session.findById("wnd[0]/tbar[0]/btn[12]").press()
    session.findById("wnd[0]/tbar[0]/btn[12]").press()


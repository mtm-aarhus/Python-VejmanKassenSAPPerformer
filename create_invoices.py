import win32com.client
import time
import re
import csv
import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

def is_cvr(cvr: str) -> bool:
    """
    Validates a Danish CVR number using modulus-11.
    
    Args:
        cvr (str): The CVR number as a string (must be 8 digits).
    
    Returns:
        bool: True if valid, False otherwise.
    """
    # Check if the CVR consists of exactly 8 digits
    if not cvr.isdigit() or len(cvr) != 8:
        return False

    weights = [2, 7, 6, 5, 4, 3, 2, 1]
    total = sum(int(digit) * weight for digit, weight in zip(cvr, weights))
    
    return total % 11 == 0

def generate_row(debitornummer):
    return [
        str(debitornummer),    # Col 1: Debitor number
        "0020",                # Col 2
        "SE",                  # Col 3: Identifier type (CVR = SE)
        "0020",                # Col 4
        "20",                  # Col 5
        "20",                  # Col 6
        "", "", "", "", "",    # Col 7–11
        "DK",                  # Col 12
        "",                    # Col 13: EAN
        "91401000",            # Col 14
        "",                    # Col 15
        "1",                   # Col 16
        "Z003",                # Col 17
        "1",                   # Col 18
        "DA",                  # Col 19
        "DKK",                 # Col 20
        "MWST",                # Col 21
        ""                     # Col 22
    ]

def generate_csv(debitors, output_filename):
    with open(output_filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        #writer.writerow(header)
        for debitornummer in debitors:
            row = generate_row(debitornummer)
            writer.writerow(row)
    return os.path.abspath(output_filename)
        


def wait_for_element(session, element_id, timeout=10):
    """Wait dynamically for an element to be available."""
    for _ in range(timeout * 10):
        try:
            el = session.findById(element_id)
            return el
        except:
            time.sleep(0.1)
    raise TimeoutError(f"Element {element_id} not found after {timeout} seconds")

def check_label_text(session, element_id, expected_text):
    try:
        label = session.findById(element_id)
        current_text = label.Text.strip()
        print(f"Current Text: '{current_text}'")
        if expected_text.strip().lower() in current_text.lower():
            print("Match found!")
            return True
        else:
            print("❌ Text does not match expected.")
            return False
    except Exception as e:
        print(f"❌ Could not find label {element_id}: {str(e)}")
        return False

def run_zfi_fakturagrundlag(filepath, orchestrator_connection: OrchestratorConnection):
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)


    # Navigate to transaction
    tx_input = wait_for_element(session, "wnd[0]/tbar[0]/okcd")
    tx_input.text = "ZFI_FAKTURAGRUNDLAG"
    session.findById("wnd[0]").sendVKey(0)  # Confirm navigation

    # Set file path
    path_field = wait_for_element(session, "wnd[0]/usr/ctxtP_PATH")
    path_field.text = filepath
    path_field.caretPosition = len(filepath)
    session.findById("wnd[0]").sendVKey(0)  # Confirm path

    # Set test mode (radio button)
    test_radio = wait_for_element(session, "wnd[0]/usr/radP_TEST")
    test_radio.select()  # more semantic than .setFocus + VKey

    # Execute (F8)
    execute_button = wait_for_element(session, "wnd[0]/tbar[1]/btn[8]")
    execute_button.press()

    container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")
    texts = []
    for child in container.Children:
        if "lbl" in child.Id:  # filter only labels
            try:
                text = child.Text.strip()
                texts.append(text)
            except:
                continue 
    combined = " | ".join(texts)
    print(f"All label texts combined:\n{combined}")

    print(f"Current Text: '{combined}'")
    if "Input filen er fejlfri - klar til opdatering.".strip().lower() in combined.lower():
        print("Fejlfri indlæsning")
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        opret_radio = wait_for_element(session, "wnd[0]/usr/radP_OPDAT")
        opret_radio.select()  # more semantic than .setFocus + VKey
        execute_button = wait_for_element(session, "wnd[0]/tbar[1]/btn[8]")
        execute_button.press()
        container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")

        # Collect all label texts in order
        texts = []
        for child in container.Children:
            if "lbl" in child.Id:  # SAP GUI labels usually have 'lbl' in their Id
                try:
                    texts.append((child.Id, (child.Text or "").strip()))
                except Exception:
                    continue

        # Convenience: just the text values
        labels = [t for _, t in texts]
        orchestrator_connection.log_info("All label texts combined:\n" + " | ".join(labels))

        # Find the split point
        try:
            split_idx = labels.index("Række Fejltekst")
        except ValueError:
            orchestrator_connection.log_error("Kunne ikke finde 'Række Fejltekst' i labels; kan ikke validere.")
            raise RuntimeError("Kunne ikke finde 'Række Fejltekst' i labels; kan ikke validere.")

        after = labels[split_idx + 1 :]

        # Pattern: "KMD Standardordre <digits> gemt"
        pat = re.compile(r"^KMD\s+Standardordre\s+(\d+)\s+gemt$", re.IGNORECASE)

        standardordre_ids = []
        bad_entries = []

        for text in after:
            if text == "":  # empty is allowed, skip it
                continue
            m = pat.match(text)
            if m:
                standardordre_ids.append(m.group(1))
            else:
                bad_entries.append(text)

        # If any non-empty entry didn't match, that's an error
        if bad_entries:
            raise RuntimeError(
                "Uventet tekst efter 'Række Fejltekst' (skal være 'KMD Standardordre <xyz> gemt' eller tom). "
                f"Fandt i stedet: {bad_entries}"
            )

        # At this point, everything non-empty was valid and we've captured all xyz values
        print(f"Valideret. Fangede {len(standardordre_ids)} Standardordre-id(s): {standardordre_ids}")
        
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/tbar[0]/btn[12]").press()

        print("DONE")
        return True, standardordre_ids

    items = [p.strip() for p in combined.split('|') if p.strip()]

    # Explicitly remove header line if present
    if items and "fejlliste vedr. indlæsning" in items[0].lower():
        items = items[1:]
    
    if items and "række fejltekst" in items[1].lower():
        items = items[2:]

    # Group into [row, (optional blank), message]
    rows = []
    i = 0
    while i + 1 < len(items):
        row_num = items[i]
        message = items[i + 1]
        rows.append((row_num, message))
        i += 2

    if i < len(items):
        leftover = items[i]
        orchestrator_connection.log_error(f"Uventet uparret fejltekst i slutningen:\n{leftover}")
        raise ValueError(f"Uventet uparret fejltekst i slutningen:\n{leftover}")
    
    # Accepted error formats
    valid_patterns = [
        r"^Ordregiver\s+(\d{10,})\s+er ikke aktiv i Salgsområde\s+\d+(?:\s\d+)*\.?$",
        r"^Fakturamodtager\s+(\d{10,})\s+er ikke aktiv i Salgsområde\s+\d+(?:\s\d+)*\.?$"
    ]

    extracted_ids = set()
    invalid_rows = []

    for row_number, message in rows:
        matched = False
        for pattern in valid_patterns:
            match = re.match(pattern, message)
            if match:
                raw_id = match.group(1)
                clean_id = raw_id[2:] if raw_id.startswith("00") else raw_id
                extracted_ids.add(clean_id)
                matched = True
                break
        if not matched:
            invalid_rows.append(f"Række {row_number}: {message}")

    #Exit current screen
    session.findById("wnd[0]/tbar[0]/btn[12]").press()
    session.findById("wnd[0]/tbar[0]/btn[12]").press()

    if invalid_rows:
        orchestrator_connection.log_error("Uventede fejlmeddelelser:\n" + "\n".join(invalid_rows))
        raise ValueError("Uventede fejlmeddelelser:\n" + "\n".join(invalid_rows))
    
    if not extracted_ids:
        orchestrator_connection.log_error("Ingen gyldige CVR-numre blev fundet.")
        raise ValueError("Ingen gyldige CVR-numre blev fundet.")
    
    return False, list(extracted_ids)
 
    
def create_debitors(file_path, orchestrator_connection: OrchestratorConnection):
    # Start SAP GUI scripting engine
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)



    # Enter transaction code
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFIE_OPRETDEB"
    session.findById("wnd[0]").sendVKey(0)

    # Set checkbox P_TEST to True
    checkbox = session.findById("wnd[0]/usr/chkP_TEST")
    if not checkbox.selected:
        checkbox.selected = True

    # Set file path
    session.findById("wnd[0]/usr/ctxtP_FILNAM").text = file_path

    # Press F8 (Execute)
    session.findById("wnd[0]").sendVKey(8)
    
    
    # Get the usr container
    usr_container = session.findById("wnd[0]/usr")

    # Loop through all children and combine text from GuiLabel elements
    combined_text = ""
    for child in usr_container.Children:
        if child.Type == "GuiLabel":
            text = child.Text.strip()
            if text:  # skip empty labels
                combined_text += text + "\n"
    print(combined_text)
    
    if "alt er ok" in combined_text.lower() and not "ikke korrekt" in combined_text.lower():
        print("Debitorfil klar til indlæsning")
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        checkbox = session.findById("wnd[0]/usr/chkP_TEST")
        if checkbox.selected:
            checkbox.selected = False
        session.findById("wnd[0]").sendVKey(8)
        container = session.findById("/app/con[0]/ses[0]/wnd[0]/usr")

        # Grab all lbl texts in order
        labels = []
        for child in container.Children:
            if "lbl" in child.Id:
                try:
                    text = (child.Text or "").strip()
                except Exception:
                    text = ""
                labels.append(text)

        print("Labels found:", labels)

        # Locate the marker line ("      1")
        try:
            marker_index = labels.index("1")  # "      1" will be stripped to "1"
        except ValueError:
            orchestrator_connection.log_error("Marker '      1' not found in labels.")
            raise RuntimeError("Marker '      1' not found in labels.")

        # Lines after the marker
        after_lines = labels[marker_index + 1 :]
        
        if not after_lines:
            orchestrator_connection.log_error("Ingen linjer fundet efter overskrift, debitoroprettelse er muligvis fejlet.")
            raise RuntimeError("Ingen linjer fundet efter overskrift, debitoroprettelse er muligvis fejlet.")

        # Required phrase
        required_phrase = "Følgende debitorer er operttet korrekt"

        # Validate all lines after marker contain the phrase
        bad_lines = [line for line in after_lines if required_phrase not in line]

        if bad_lines:
            orchestrator_connection.log_error(
                f"Nogle linjer efter står ikke som oprettet korrekt, da de mangler teksten '{required_phrase}':\n" +
                "\n".join(repr(l) for l in bad_lines)
            )
            raise RuntimeError(
                f"Nogle linjer efter står ikke som oprettet korrekt, da de mangler teksten '{required_phrase}':\n" +
                "\n".join(repr(l) for l in bad_lines)
            )

        print("Alle linjer indeholder den krævede tekst.")
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        
    else:
        raise Exception("Fejl i debitoroprettelse, stopper kørsel.")
    
    
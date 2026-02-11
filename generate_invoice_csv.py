# Minimal dispatcher: read DB rows and write one CSV per matching row
# - No Selenium/Vejman
# - No emails
# - No DB updates

import os
import csv
import locale
from datetime import datetime, timedelta
import re
import math
import time
import pyodbc
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

# ---------- Helpers ----------

def format_decimal(value, decimals=None):
    """
    DK formatting. If decimals is None:
      - ints => no decimals
      - floats => 2 decimals unless it's an integer-like value
    If decimals given => force that many.
    """
    if isinstance(value, int):
        if decimals is None:
            return str(locale.format_string("%d", value, grouping=False))
        else:
            return str(locale.format_string(f"%.{decimals}f", value, grouping=False))
    elif isinstance(value, float):
        if value.is_integer() and decimals is None:
            return str(locale.format_string("%d", int(value), grouping=False))
        else:
            if decimals is None:
                return str(locale.format_string("%.2f", value, grouping=False))
            else:
                return str(locale.format_string(f"%.{decimals}f", value, grouping=False))
    else:
        return "" if value is None else str(value)
    
    

        
def generate_invoice_csv(orchestrator_connection: OrchestratorConnection, conn: pyodbc.Connection, cursor: pyodbc.Cursor, row: pyodbc.Row):
    locale.setlocale(locale.LC_NUMERIC, 'da_DK')


    tilladelsestype = row.TilladelsesType
    # Fetch the matching fakturatekster row

    conn.commit()

    cursor.execute("""
        SELECT TOP (1) *
        FROM [dbo].[VejmanFakturaTekster]
        WHERE Fakturalinje = ?
    """, (tilladelsestype,))
    fakturarow = cursor.fetchone()
    Fakturalinje = fakturarow.Fakturalinje
    fordringstype = fakturarow.Fordringstype
    psp_element = fakturarow.PSPElement
    materiale_nr_opus = fakturarow.MaterialeNrOpus
    formatted_material_number = f'{int(materiale_nr_opus):018}'
    top_text = fakturarow.Toptekst
    forklaring = fakturarow.Forklaring
    SAP_NOTE = (
        "Bemærk: Forfaldsdatoen angiver periodens start og er ikke betalingsfristen. Betalingsfristen fremgår øverst på fakturaen."
    )

    # Assign variables directly using column names
    ID = row.ID
    VejmanID = row.VejmanID
    FørsteSted = row.FørsteSted
    Tilladelsesnr = row.Tilladelsesnr
    Ansøger = row.Ansøger
    CvrNr = row.CvrNr
    Enhedspris = row.Enhedspris
    Meter = row.Meter
    Startdato = (datetime.strptime(row.Startdato, '%Y-%m-%d')if row.Startdato else None)
    Slutdato = (datetime.strptime(row.Slutdato, '%Y-%m-%d') if row.Slutdato else None)
    AntalDage = row.AntalDage
    TotalPris = row.TotalPris
    kunde_ref_id = row.ATT

    # Ensure specific columns have the correct types
    Enhedspris = float(Enhedspris) if Enhedspris is not None else None
    Meter = float(Meter) if Meter is not None else None
    TotalPris = float(TotalPris) if TotalPris is not None else None
    AntalDage = int(AntalDage) if AntalDage is not None else None
    
    
    def format_decimal(value, decimals=None):
        if isinstance(value, int):
            if decimals is None:
                # Format the integer without decimal places
                return str(locale.format_string("%d", value, grouping=False))
            else:
                # Force formatting with the specified number of decimals
                return str(locale.format_string(f"%.{decimals}f", value, grouping=False))
        elif isinstance(value, float):
            # Check if the value is a whole number and decimals is None (e.g., 19.0 should be formatted as 19)
            if value.is_integer() and decimals is None:
                return str(locale.format_string("%d", int(value), grouping=False))
            else:
                # Format the float with the specified number of decimal places
                if decimals is None:
                    return str(locale.format_string("%.2f", value, grouping=False))
                else:
                    # Force formatting with the specified number of decimals
                    return str(locale.format_string(f"%.{decimals}f", value, grouping=False))
        else:
            # If it's not a number, return the original value
            return str(value)

    
    formatted_cvr_number = f'{int(CvrNr):010}'
    
    
    today = datetime.now().strftime('%d-%m-%Y')
    future_date = (datetime.now() + timedelta(days=30)).strftime('%d-%m-%Y')        
    short_start_date = Startdato.strftime('%d-%m-%Y')
    short_end_date = Slutdato.strftime('%d-%m-%Y')
    
    # Format numbers inside the f-string expressions
    opus_price = format_decimal(round(Meter*Enhedspris,2),2)
    unit_price = format_decimal(Enhedspris)
    length = format_decimal(Meter)
    days_period_formatted = format_decimal(AntalDage,3)
    total_calculated_price = format_decimal(TotalPris)
        # Use eval to evaluate them as f-strings
    top_text_evaluated = eval(top_text)
    forklaring_evaluated = eval(forklaring)

    
    # Prepare rows for writing
    row_H = [
        'H', formatted_cvr_number, '', today, today, '0020', '20', '20', 'ZRA', Tilladelsesnr, '', 
        '', '', '', kunde_ref_id, 
        top_text_evaluated,
        '', '', '', '', '', '', '', short_start_date, short_start_date, short_end_date, '', '', short_start_date, short_end_date, '', fordringstype, '', '', short_start_date, future_date
    ]
    
    row_L = [
        'L', formatted_material_number, Fakturalinje, days_period_formatted, opus_price, 'NEJ', psp_element, '', '', '', 
        '', SAP_NOTE, forklaring_evaluated,
        '', '', '', '', '', '', '', '', '','', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S-%f")[:-3]  # milliseconds
    csvname = f"{timestamp}_Fakturaer_{ID}.csv"

    full_path = os.path.abspath(csvname)  # get absolute path in current working dir
    
    # Write to the CSV
    with open(full_path, mode='a', newline='', encoding='windows-1252') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(row_H)
        writer.writerow(row_L)
    return full_path


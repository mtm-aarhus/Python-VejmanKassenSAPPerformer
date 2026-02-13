"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import pyodbc
from datetime import datetime
import json
import os

from create_invoices import run_zfi_fakturagrundlag, generate_csv, create_debitors
from generate_invoice_csv import generate_invoice_csv
from send_invoices import send_invoice
from update_vejman import update_case



# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    vejmantoken = orchestrator_connection.get_credential("VejmanToken").password

    sql_server = orchestrator_connection.get_constant("SqlServer").value
    conn_string = "DRIVER={SQL Server};"+f"SERVER={sql_server};DATABASE=VejmanKassen;Trusted_Connection=yes;"
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()
    
    faktura = json.loads(queue_element.data)
    sql_id = faktura.get("ID")
    vejmanid = faktura.get("VejmanID")
    orchestrator_connection.log_info(f"Running for SQL row with ID: {sql_id} - {vejmanid}")

    cursor.execute("""
        SELECT *
        FROM [VejmanKassen].[dbo].[VejmanFakturering]
        WHERE ID = ? AND FakturaStatus = 'TilFakturering'
    """, sql_id)
    row = cursor.fetchone()
    fakturafil = generate_invoice_csv(orchestrator_connection, conn, cursor, row)

    ordernumber = None

    success, debitorsororder = run_zfi_fakturagrundlag(fakturafil)
    # Output file name based on date
    if not success:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # remove last 3 digits → milliseconds
        filename = f"{sql_id}_Debitorer_CSV_{timestamp}.csv"
        debitor_csv = generate_csv(debitorsororder, filename)
        create_debitors(debitor_csv)
        #os.remove(debitor_csv)
        success, debitorsororder = run_zfi_fakturagrundlag(fakturafil)
    if success:
        if len(debitorsororder) == 1:
            ordernumber = debitorsororder[0]  # Extract the only item
        else:
            raise RuntimeError("Flere ordrenumre fundet, der burde kun være et.")
    else:
        raise RuntimeError("Fejlede indlæsning efter debitoroprettelse")
            
    
    send_invoice(orchestrator_connection)
    cursor.execute("""
        UPDATE [VejmanKassen].[dbo].[VejmanFakturering]
        SET FakturaStatus = 'Faktureret',
            FakturaDato        = CAST(GETDATE() AS date),
            Ordrenummer        = ?
        WHERE ID = ?
    """, ordernumber, sql_id)
    conn.commit()
    os.remove(fakturafil)
    if not vejmanid == "Henstilling":
        update_case(vejmanid, vejmantoken)
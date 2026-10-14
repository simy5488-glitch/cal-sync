import os, requests, re, time
from io import BytesIO
from O365 import Account
from datetime import datetime, timedelta
from openpyxl import load_workbook

def main():
    credentials = (os.environ['CLIENT_ID'], os.environ['CLIENT_SECRET'])
    account = Account(credentials, tenant_id=os.environ['TENANT_ID'], auth_flow_type='credentials')
    if not account.authenticate(scopes=['https://graph.microsoft.com/.default']): 
        print("Errore autenticazione Microsoft")
        return

    # Preso dalle variabili d'ambiente (Secret)
    calendar_email = os.environ['CALENDAR_EMAIL']
    schedule = account.schedule(resource=calendar_email)
    
    # --- 1. SELEZIONE CALENDARIO ---
    calendars = schedule.list_calendars()
    target_calendar = next((c for c in calendars if c.name == 'Università'), None)
    if not target_calendar:
        target_calendar = schedule.get_default_calendar()

    # --- 2. PULIZIA TOTALE (ANNO 2026) ---
    print("Pulizia in corso...")
    s_iso, e_iso = "2026-01-01T00:00:00Z", "2026-12-31T23:59:59Z"
    while True:
        try:
            eventi = list(target_calendar.get_events(limit=999, start_recurring=s_iso, end_recurring=e_iso, include_recurring=True))
            if not eventi: break
            for ev in eventi: ev.delete()
            time.sleep(2)
        except Exception as e:
            print(f"Fine pulizia: {e}")
            break

    # --- 3. DOWNLOAD E LETTURA EXCEL ---
    print("Download e inserimento nuovi dati...")
    # Preso dalle variabili d'ambiente (Secret)
    drive_file_id = os.environ['DRIVE_FILE_ID']
    url = f"https://drive.google.com/uc?export=download&id={drive_file_id}"
    
    res = requests.get(url)
    if res.status_code != 200:
        print(f"Errore download: Status {res.status_code}")
        return

    wb = load_workbook(BytesIO(res.content), data_only=True)
    sheet = wb.active 
    rows = list(sheet.rows)

    base = None

    for r_idx, row in enumerate(rows):
        # 1. Controlla se in questa riga c'è la dicitura della SETTIMANA
        row_values_text = " ".join([str(cell.value or "") for cell in row])
        if "SETTIMANA" in row_values_text.upper():
            match = re.search(r'(\d{2}-\d{2})', row_values_text)
            if match:
                base = datetime.strptime(match.group(1) + "-2026", "%d-%m-%Y")
                print(f"\n--- Inizio {match.group(0)} (Data Lunedì: {base.strftime('%d/%m/%Y')}) ---")
            continue

        if not base:
            continue

        # 2. Controlla se la Colonna A contiene una fascia oraria
        col_a = str(row[0].value or "").replace(" ", "").replace("–", "-")
        time_match = re.match(r'^(\d{1,2})-(\d{1,2})$', col_a)
        
        if time_match:
            start_hour_ita = int(time_match.group(1))
            duration = int(time_match.group(2)) - start_hour_ita

            for col_off in range(1, 6):
                if col_off >= len(row):
                    continue

                target_cell = row[col_off]
                mat = str(target_cell.value or "").strip()

                if not mat or mat.lower() == 'none' or len(mat) < 3:
                    continue

                if target_cell.font and target_cell.font.strikethrough:
                    print(f"Salto '{mat}' ({col_a}) perché barrata.")
                    continue

                dt = base + timedelta(days=col_off - 1)
                utc_hour = start_hour_ita - 2 

                ev = target_calendar.new_event()
                ev.subject = mat
                ev.start = dt.replace(hour=utc_hour, minute=0)
                ev.end = ev.start + timedelta(hours=duration)
                ev.location = "Aula Università"
                ev.is_reminder_on = False
                ev.save()
                
                print(f"Inserito: {mat} | {dt.strftime('%d/%m')} ore {start_hour_ita}:00-{start_hour_ita+duration}:00")

    print("\nSincronizzazione completata con successo!")

if __name__ == "__main__":
    main()

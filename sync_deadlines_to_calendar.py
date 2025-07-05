import os
import datetime
import gspread
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

# Настройки

SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

SPREADSHEET_ID = '1lwA3DqC5NxrHsd5jVXzEck0YSvielIUT4e1YepJtRDo' 
SHEET_NAME = 'Дедлайны'    

# Авторизация в Google APIs

def get_calendar_service():
    print("Авторизация в Google Calendar API...")
    creds = None
    if os.path.exists('token.json'):
        print("Токен найден, загрузка...")
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Обновление устаревшего токена...")
            creds.refresh(Request())
        else:
            print("Запуск OAuth flow для авторизации...")
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            print("Токен успешно сохранён!")
    service = build('calendar', 'v3', credentials=creds)
    print("Авторизация в Google Calendar завершена.")
    return service

# Чтение данных из Google Таблицы

def read_sheet_data():
    print("Подключение к Google Sheets...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    gc_creds = ServiceAccountCredentials.from_json_keyfile_name("service-account.json", scope) 
    client = gspread.authorize(gc_creds)

    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_records()
    print(f"Получено {len(data)} записей из таблицы '{SHEET_NAME}'.")
    return sheet, pd.DataFrame(data)

# Создание события в календаре

def create_calendar_event(service, sheet, df_row):
    try:
        deadline_date = datetime.datetime.strptime(df_row['Дата дедлайна'], "%d.%m.%Y").date()
    except ValueError:
        print(f"❌ Ошибка формата даты для заявки {df_row['Заявка']}: {df_row['Дата дедлайна']}")
        return

    event_summary = f"{df_row['Заявка']} — {df_row['Действие']}"
    event_description = (
        f"Клиент: {df_row['Клиент']}\n"
        f"Статус: {df_row['Статус']}\n"
        f"Ответственный: {df_row['Ответственный']}"
    )

    event = {
        'summary': event_summary,
        'description': event_description,
        'start': {
            'date': deadline_date.strftime('%Y-%m-%d'),
            'timeZone': 'Europe/Moscow',
        },
        'end': {
            'date': deadline_date.strftime('%Y-%m-%d'),
            'timeZone': 'Europe/Moscow',
        },
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 60 * 24 * 3},  # за 3 дня
                {'method': 'popup', 'minutes': 60}           # за 1 час
            ],
        },
    }

    print(f"✅ Создаю событие: {event_summary} | Дата: {deadline_date}")
    created_event = service.events().insert(calendarId='primary', body=event).execute()
    event_id = created_event['id']
    print(f"✅ Событие создано: {event_summary} | Event ID: {event_id}")

    # Обновляем ячейку с Event ID в таблице
    cell = sheet.find(df_row['Заявка'])
    sheet.update_cell(cell.row, cell.col + 6, event_id)  # Event ID в 7-м столбце
    print(f"📌 Event ID записан в таблицу для строки {cell.row}")

# Основной процесс

def main():
    calendar_service = get_calendar_service()
    try:
        sheet, df = read_sheet_data()
    except Exception as e:
        print(f"❌ Ошибка при подключении к таблице: {e}")
        return

    for index, row in df.iterrows():
        event_id = row.get('Event ID')
        if pd.isna(event_id) or event_id == '':
            try:
                create_calendar_event(calendar_service, sheet, row)
            except Exception as e:
                print(f"❌ Ошибка при создании события: {e}")
        else:
            print(f"ℹ️ Пропущено: событие уже существует для {row['Заявка']}")

    print("🎉 Синхронизация завершена.")

if __name__ == '__main__':
    main()
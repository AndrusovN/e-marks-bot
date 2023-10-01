import telebot
import json
import os
from dotenv import load_dotenv
from google.oauth2 import service_account
import pandas as pd
import gspread

load_dotenv()
# Создаем экземпляр бота
bot = telebot.TeleBot(os.getenv('BOT_API_TOKEN'))
google_json = {}
with open("starlit-link-340011-78e591ae7d80.json", 'r') as f:
    google_json = json.loads(f.read())
credentials = service_account.Credentials.from_service_account_info(google_json)
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds_with_scope = credentials.with_scopes(scope)
client = gspread.authorize(creds_with_scope)


def get_sheet():
    spreadsheet = client.open_by_url(
        'https://docs.google.com/spreadsheets/d/16pjAVfmtai-qZEwXMWHzykUEH2WX1ZcMaX5RxUnDP_4/edit?usp=sharing')
    worksheet = spreadsheet.get_worksheet(0)
    records_data = worksheet.get_values()
    records_df = pd.DataFrame.from_dict(records_data)
    return records_df


def get_user(sheet, user):
    row = -1
    for id, cell in enumerate(sheet[1]):
        if cell == '@' + user.username or cell == user.first_name:
            row = id
            break
    if row == -1:
        return -1, None
    return row, sheet[2][row], sheet[3][row]


@bot.message_handler(commands=['get_marks'])
def get_marks(m):

    sheet = get_sheet()
    row, name, level = get_user(sheet, m.from_user)
    if row == -1:
        bot.send_message(m.from_user.id, "Извините, вы не найдены в таблице учеников")
        return
    result_message = ""
    for column_id in sheet:
        column = sheet[column_id]
        if column[1] == "Оценка" and column[row] != "":
            result_message += f"{column[0]} (вес {column[2]}): {column[row]}\n"
    result_message += "На этом пока всё!"
    bot.send_message(m.from_user.id, result_message)


@bot.message_handler(commands=['start'])
def start(m):
    sheet = get_sheet()
    _, name, level = get_user(sheet, m.from_user)
    bot.send_message(m.from_user.id, f"""Привет! Я бот для оценок. 
По команде /get_marks я могу написать твои текущие оценки.
Пожалуйста, проверь, что я правильно определил, кто ты: {name}, уровень: {level}
Если это не ты или у тебя неправильно указан уровень, то напиши об этой ошибке @n_andrusov как можно скорее!""")


bot.polling()

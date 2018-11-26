import telebot
import constants
from openpyxl import load_workbook

bot = telebot.TeleBot(constants.token)

try:
    wb = load_workbook('оборудование.xlsx')
    users_sheet = wb['users']
except FileNotFoundError as error:
    msg = 'Не могу найти файл "оборудование.xlsx", а без него работать не могу'
    print(msg)
    raise SystemExit

for cell in users_sheet['B']:
    if cell.row != 1:
        msg = "Спасибо, ненадолго прервёмся, выключаю бот"
        bot.send_message(cell.value, msg)

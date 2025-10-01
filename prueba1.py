# import telebot


# print("Hola")
# def EnvioFoto(cedula):
#     #TELEGRAM_BOT_TOKEN = '7241959128:AAEVyjfp1HPF1Ytpwe2gLNoSYQU3ZllgVx0'
#     #TELEGRAM_CHAT_ID = '-1002227166533'
#     # GRUPO ROBOT CENSEL
#     TELEGRAM_BOT_TOKEN = '8058959859:AAFIiLBcd4JI2DSrnEo2sKzFFCRikFctv2k'
#     TELEGRAM_CHAT_ID = '7411433556'
#     bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)
#     # foto = open(f"{folder_path}{cedula}.png", "rb")
#     img = cedula
#     foto = open(f"{img}", "rb")
#     bot.send_photo(chat_id=TELEGRAM_CHAT_ID, photo=foto, caption=f"<b>{cedula}!!</b>", parse_mode="html")
#     # logging.info("Envió imagen al Telegram")
#     print("Envió imagen al Telegram")
#     # log_cedulas("Envió imagen al Telegram")

# EnvioFoto("C:/ContratosSuperflex/fotosPDV/11405158.png")
# import requests

# url = "https://www.emsa-esp.com.co:441/factura/consulta_factura.php?cuenta=274416732"

# payload = {}
# headers = {}

# # response = requests.request(url, timeout=10)
# response = requests.request("GET", url, headers=headers, data=payload)
# # status_code = response.status_code

# # try:
# #     # response = requests.request(url, timeout=10)
# #     response = requests.request("GET", url, headers=headers, data=payload)
# #     # status_code = response.status_code
# # except requests.exceptions.RequestException:
# #     response = None

# print(response.text)
# Guardar en variable de Rocketbot

import requests, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.emsa-esp.com.co:441/factura/consulta_factura.php?cuenta=274416732"
r = requests.get(url, timeout=15, verify=False)  # ?? inseguro
print(r.status_code)
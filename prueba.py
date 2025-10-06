import requests, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = 'https://www.emsa-esp.com.co:441/factura/consulta_factura.php?cuenta=14420010'

try:
  response = requests.get(url, timeout=30, verify=False)
  status_code = response.status_code
  txt_response = response.text
except requests.exceptions.RequestException:
    status_code = None

if txt_response == "No se genero la factura. Intente nuevamente o verifique los datos ingresados":
   status_code = 300
print(response)
print(txt_response)
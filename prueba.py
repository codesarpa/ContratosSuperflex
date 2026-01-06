import requests, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# facturta correcta
# url = 'https://www.emsa-esp.com.co:441/factura/consulta_factura.php?cuenta=399274086'

# factura erronea
url = 'https://www.emsa-esp.com.co:441/factura/consulta_factura.php?cuenta=451245565'

try:
  response = requests.get(url, timeout=30, verify=False)
  status_code = response.status_code
  content_type = response.headers.get("Content-Type")
  txt_response = response.text
except requests.exceptions.RequestException:
    status_code = None

if content_type != 'application/pdf':
   print("factura buena")
# print(response)
print(response.headers.get("Content-Type"))




import requests, urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = GetVar('url_descarga')

try:
  response = requests.get(url, timeout=30, verify=False)
  status_code = response.status_code
except requests.exceptions.RequestException:
    status_code = None
    
SetVar("status_code", status_code)

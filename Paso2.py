import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains



# driver = webdriver.Firefox()
#service = Service(executable_path='ruta/a/geckodriver')
#driver = webdriver.Firefox(service=service, options=options)
def IrPaginaContratos(driver):

    #driver.get("http://10.1.1.22:8181/BusinessNET-WEB/XHTML/azar/adminventa/contratopersonas.xhtml")
    driver.get("http://10.167.32.73:8130/BusinessNET-WEB/XHTML/azar/adminventa/contratopersonas.xhtml")
    print("Hola")
    # Esperar hasta que un elemento clave esté presente en la página
    elemento_clave = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "tbwContratoPersonas:frmPersonas:txtPersonaDocumentoPrincipal"))
    )
    print(elemento_clave)
    print("La página se ha cargado completamente.")
    time.sleep(3)


def Login(driver):
    try:        
        # Abrir la página de inicio de sesión
        #driver.get("http://10.1.1.22:8181/BusinessNET-WEB/XHTML/general/login.xhtml")
        # PRODUCCION driver.get("http://10.167.32.73:8130/BusinessNET-WEB/XHTML/general/login.xhtml")
        driver.get("http://10.1.1.22:8181/BusinessNET-WEB/XHTML/general/login.xhtml") #PRUEBAS
        time.sleep(3)
        driver.maximize_window()
        # Esperar hasta que el campo de usuario esté disponible
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "idFormLogin:user"))
        )

        # Ingresar el usuario
        username_field = driver.find_element(By.ID, "idFormLogin:user")
        username_field.send_keys("CP1121900795")

        # Ingresar la contraseña
        password_field = driver.find_element(By.ID, "idFormLogin:password")
        password_field.send_keys("795CP")
        print(password_field)
        # Hacer clic en el botón de ingresar
        login_button = driver.find_element(By.ID, "idFormLogin:ingresar")
        login_button.click()

        time.sleep(10)
        # Redireccionar a la página deseada después del inicio de sesión
        IrPaginaContratos()
        time.sleep(5)
        # Aquí puedes continuar interactuando con la página después de la redirección

    except Exception as e:
        print("Error al iniciar sesión:", e)
        #driver.quit()
        
#Login()

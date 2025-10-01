from httpcore import TimeoutException
from Paso2 import *
import Paso1
from Paso1 import *
import os
import telebot


folder_path = './fotosPDV/'

def ConsultarDocumento(cedula, driver):
    # Esperar a que el campo esté presente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.ID, "tbwContratoPersonas:frmPersonas:txtPersonaDocumentoPrincipal"))
    )

    # Encontrar el campo de texto y enviar el número de documento
    campo_documento = driver.find_element(
        By.ID, "tbwContratoPersonas:frmPersonas:txtPersonaDocumentoPrincipal")
    campo_documento.send_keys(cedula)
    campo_documento.send_keys(Keys.ENTER)

    # Esperar a que el campo de nombre tenga un valor distinto de vacío
    WebDriverWait(driver, 10).until(
        lambda d: d.find_element(
            By.ID, "tbwContratoPersonas:frmPersonas:txtPersonaNombre").get_attribute('value') != ""
    )
    # Hacer clic en el enlace "Personal ventas"
    enlace_personal_ventas = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Personal ventas"))
    )
    enlace_personal_ventas.click()
    time.sleep(5)


def GrupoPlanes(driver):
    boton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.ID, "tbwContratoPersonas:frmPersonasVenta:txtCodgrupoplanes"))
    )
    boton.click()
    time.sleep(2)

    input_box = driver.find_element(By.ID, "tbwContratoPersonas:frmPersonasVenta:txtCodgrupoplanes")

    input_box.clear()  # Limpia la caja de texto (opcional)
    input_box.send_keys("618")
    time.sleep(1)
    input_box.send_keys(Keys.RETURN)

    time.sleep(3)


def ValidarMensaje(driver):
    try:
        # Busca el elemento con la clase específica
        elemento = driver.find_element_by_css_selector(
            '.ui-growl-item-container.ui-state-highlight.ui-corner-all')

        # Verifica si el elemento es visible
        if elemento.is_displayed():
            print("El elemento está visible en pantalla.")
        else:
            print("El elemento existe pero no es visible en pantalla.")
    except NoSuchElementException:  # type: ignore
        print("El elemento no existe en el DOM.")


def vaciarCarpeta():

    # Iterar sobre todos los archivos en la carpeta
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # Eliminar archivos o enlaces simbólicos
            elif os.path.isdir(file_path):
                os.rmdir(file_path)  # Eliminar directorios vacíos
        except Exception as e:
            print(f'Error al eliminar {file_path}. Razón: {e}')


def CapturaPantalla(cedula, driver):
    # Captura la pantalla completa
    screenshot_path = f"{folder_path}{cedula}.png"
    driver.save_screenshot(screenshot_path)


def EnvioFoto(cedula):

    bot = telebot.TeleBot("7241959128:AAEVyjfp1HPF1Ytpwe2gLNoSYQU3ZllgVx0")
    foto = open(f"{folder_path}{cedula}.png", "rb")
    bot.send_photo(7411433556, foto, f"<b>{cedula}!!</b>", parse_mode="html")


def GuardarContrato(cedula, driver):
    # click en el botón guardar
    boton_guardar = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "#tbwContratoPersonas\\:frmPersonasVenta\\:btnGuardarContratosventa")))
    # boton_guardar.click()
    action = ActionChains(driver)
    action.double_click(boton_guardar).perform()
    time.sleep(3)
    loading_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "j_idt20"))
    )

    # Bucle para esperar mientras el elemento tiene la clase que indica que está visible y cargando
    while True:
        # Obtiene la clase del elemento
        element_class = loading_element.get_attribute("class")

        # Verifica si el elemento sigue visible y cargando
        if "ui-overlay-visible" in element_class:
            time.sleep(2)  # Espera 2 segundos antes de volver a verificar
        else:
            break

    time.sleep(10)
    print("salir ciclo")
    #EXTRAS
    CapturaPantalla(f"{cedula}")
    time.sleep(5)
    EnvioFoto(f"{cedula}")
    time.sleep(1)
    vaciarCarpeta()
    time.sleep(5)
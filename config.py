import logging
from datetime import datetime

# LOGS
logging.basicConfig(
    filename='Logs.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Logger especifico para cedulas
cedulas_logger = logging.getLogger('cedulas')
cedulas_handler = logging.FileHandler('cedulas.log')
cedulas_handler.setLevel(logging.INFO)
cedulas_formatter = logging.Formatter('%(asctime)s - %(message)s')
cedulas_handler.setFormatter(cedulas_formatter)
cedulas_logger.addHandler(cedulas_handler)

def log_cedulas(log):
    cedulas_logger.info(log)

# Ejemplo de uso
log_cedulas("Registro de cedula: 123456789")

hora_inicio = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
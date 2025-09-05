import warnings, time
from funciones import *

# Rutas base
from dotenv import load_dotenv
load_dotenv()  # carga .env desde el cwd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def main():
    inicio = time.time()
    carpetas = "Summary Report", "Terminal Report"
    for carpeta in carpetas:
        descompresor(carpeta)
    
    exportarReportes("Terminal Report", formato="csv")
    exportarReportes("Summary Report", formato="xlsx")
        
    tiempoEjecucion(inicio)
    
if __name__ == "__main__":
    main()
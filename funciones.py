import os, zipfile, shutil, time
import pandas as pd
from pathlib import Path


def tiempoEjecucion(inicio):
    final = time.time()                                                                                                    # Debug
    tiempo_total = final - inicio                                                                                          # Debug
    minutos, segundos = divmod(tiempo_total, 60)                                                                           # Debug
    print(f"Tiempo total de ejecución: {int(minutos)} minutos y {round(segundos, 2)} segundos")                            # Debug 
        
def eliminarArchivos(ruta: str, nombre: str = ''):
    archivos = os.listdir(ruta)
    for archivo in archivos:                                                                            
        ruta_completa = os.path.join(ruta, archivo)
        try:
            if os.path.isfile(ruta_completa):
                if nombre == '':
                    os.remove(ruta_completa)                                                            # Elimina cada archivo en la carpeta
                elif archivo == nombre:
                    os.remove(ruta_completa)                                                            # Elimina un archivo especifico 
        except Exception as e:
            print(f"No se pudo eliminar {ruta_completa}: {e}")
                

def descompresor(carpeta: str):
    carpeta_origen = Path(f"documentos/comprimidos/{carpeta}")    
    carpeta_destino = Path(f"documentos/descomprimidos/{carpeta}")
    carpeta_destino.mkdir(parents=True, exist_ok=True)

    eliminarArchivos(carpeta_destino)
    
    for archivo_comprimido in carpeta_origen.glob("*.zip"):                                
        nombre_sin_extension = archivo_comprimido.stem
        carpeta_temp = carpeta_destino / f"tmp_{nombre_sin_extension}"
        carpeta_temp.mkdir(exist_ok=True)

        try:
            with zipfile.ZipFile(archivo_comprimido, 'r') as zip_ref:
                zip_ref.extractall(carpeta_temp)
            #print("Descompresión completa del archivo {nombre_sin_extension}")                                # Debug

            archivos_encontrados = list(carpeta_temp.rglob("*.*"))                                  # Busca archivos dentro de todos los subniveles del directorio

            if not archivos_encontrados:
                print(f"⚠️ No se encontraron archivos dentro de {archivo_comprimido.name}")
            else:
                for archivo in archivos_encontrados:
                    if archivo.is_file():
                        destino_final = carpeta_destino / archivo.name

                        if destino_final.exists():                                                 # Si el nombre ya existe, agrega un sufijo
                            destino_final = carpeta_destino / f"{archivo.stem}_{nombre_sin_extension}{archivo.suffix}"

                        shutil.copy(archivo, destino_final)

        except zipfile.BadZipFile:
            print(f"❌ Error: {archivo_comprimido.name} no es un archivo ZIP válido.\n")

        shutil.rmtree(carpeta_temp)
        

def exportarReportes(carpeta: str, formato: str = "xlsx"):
    formato = formato.lower().strip()
    maximo_filas = int(os.getenv('CANTIDAD_MAXIMA_FILAS'))
    if formato == "xlsx":
        return unificarExcelXlsx(carpeta, maximo_filas=maximo_filas)
    elif formato == "csv":
        return unificarExcelsCsv(carpeta)
    else:
        raise ValueError("Formato no soportado. Use 'xlsx' o 'csv'.")


def leerHojas(archivo: Path):                                                                       # Elige engine dependiendo de la extension de los archivos
    engine = "openpyxl" if archivo.suffix.lower() == ".xlsx" else "xlrd"
    hojas = pd.read_excel(archivo, sheet_name=None, engine=engine)
    return hojas


def unificarExcelXlsx(carpeta: str, maximo_filas: int):             # Varios archivos xlsx (según maximo de filas)
    carpeta_origen = Path(f"documentos/descomprimidos/{carpeta}")
    carpeta_salida = Path("documentos/finales")
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    base_nombre = carpeta.strip().replace('/', '_')                                                     # Conserva espacios

    archivos_excel = list(carpeta_origen.glob("*.xlsx")) + list(carpeta_origen.glob("*.xls"))
    if not archivos_excel:
        print(f"⚠️ No se encontraron Excel en {carpeta_origen}")
        return

    hojas_acumuladas: dict[str, list[pd.DataFrame]] = {}                                        # Acumula por hoja
    for archivo in archivos_excel:
        print(f"📂 Procesando: {archivo.name}")                                                 # Debug
        try:
            hojas = leerHojas(archivo)
            for hoja_nombre, df in hojas.items():
                df["__archivo_origen__"] = archivo.name  # opcional
                hojas_acumuladas.setdefault(hoja_nombre, []).append(df)
        except Exception as e:
            print(f"❌ Error al procesar {archivo.name}: {e}")

    if not hojas_acumuladas:
        print("⚠️ No hubo datos legibles para unificar.")
        return

    fragmento_por_hoja: dict[str, list[pd.DataFrame]] = {}                                               # Particiona por límite de filas
    maxima_cantidad_fragmentos = 0

    for hoja_nombre, lista_df in hojas_acumuladas.items():
        df_completo = pd.concat(lista_df, ignore_index=True)
        total = len(df_completo)
        if total == 0:
            continue

        cant_chunks = (total + maximo_filas - 1) // maximo_filas
        maxima_cantidad_fragmentos = max(maxima_cantidad_fragmentos, cant_chunks)

        fragmentos = []
        for i in range(cant_chunks):
            ini = i * maximo_filas
            fin = min((i + 1) * maximo_filas, total)
            fragmentos.append(df_completo.iloc[ini:fin])
        fragmento_por_hoja[hoja_nombre] = fragmentos
        print(f"🧩 Hoja '{hoja_nombre}': {total} filas → {cant_chunks} chunk(s)")

    if maxima_cantidad_fragmentos == 0:
        print("⚠️ No hubo datos para escribir.")
        return

    if maxima_cantidad_fragmentos == 1:                                                             # Un solo archivo (sin sufijo)
        archivo_salida_unico = carpeta_salida / f"{base_nombre}.xlsx"
        print(f"📝 Generando archivo único: {archivo_salida_unico}")
        with pd.ExcelWriter(archivo_salida_unico, engine="xlsxwriter") as writer:
            for hoja_nombre, fragmentos in fragmento_por_hoja.items():
                df_fragmento = fragmentos[0]
                if not df_fragmento.empty:
                    sheet_name = hoja_nombre[:31]  # límite Excel
                    df_fragmento.to_excel(writer, sheet_name=sheet_name, index=False)
        print("✅ Proceso finalizado.")
        return

    print(f"📝 Se generarán {maxima_cantidad_fragmentos} archivo(s) .xlsx")                         # Multiples archivos con sufijo _1, _2, ...
    for idx in range(maxima_cantidad_fragmentos):
        # Verificar si hay al menos una hoja con datos para este volumen
        hay_datos = any(idx < len(chs) and not chs[idx].empty for chs in fragmento_por_hoja.values())
        if not hay_datos:
            continue

        archivo_salida = carpeta_salida / f"{base_nombre}_{idx+1}.xlsx"
        print(f"  ➕ Creando: {archivo_salida}")

        with pd.ExcelWriter(archivo_salida, engine="xlsxwriter") as writer:
            for hoja_nombre, fragmentos in fragmento_por_hoja.items():
                if idx < len(fragmentos):
                    df_fragmento = fragmentos[idx]
                    if not df_fragmento.empty:
                        sheet_name = hoja_nombre[:31]
                        df_fragmento.to_excel(writer, sheet_name=sheet_name, index=False)

    print("✅ Proceso finalizado.")


def unificarExcelsCsv(carpeta: str):                                                                    # Unico archivo csv
    carpeta_origen = Path(f"documentos/descomprimidos/{carpeta}")
    carpeta_salida = Path("documentos/finales")
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    base_nombre = carpeta.strip().replace('/', '_')  # mantener espacios
    ruta_csv = carpeta_salida / f"{base_nombre}.csv"

    archivos_excel = list(carpeta_origen.glob("*.xlsx")) + list(carpeta_origen.glob("*.xls"))
    if not archivos_excel:
        print(f"⚠️ No se encontraron Excel en {carpeta_origen}")
        return

    columnas_ordenadas: list[str] = []
    def acumular_columnas(columnas: list[str]):
        for columna in columnas:
            if columna not in columnas_ordenadas and columna not in ("__archivo_origen__", "__hoja_origen__"):
                columnas_ordenadas.append(columna)

    for archivo in archivos_excel:
        try:
            engine = "openpyxl" if archivo.suffix.lower() == ".xlsx" else "xlrd"
            with pd.ExcelFile(archivo, engine=engine) as xls:
                for hoja_nombre in xls.sheet_names:
                    df_encabezado = xls.parse(hoja_nombre, nrows=0)
                    acumular_columnas(list(df_encabezado.columns))
        except Exception as e:
            print(f"❌ Error leyendo encabezados de {archivo.name}: {e}")

    columnas_ordenadas += ["__archivo_origen__", "__hoja_origen__"]

    if ruta_csv.exists():
        ruta_csv.unlink()

    wrote_header = False
    for archivo in archivos_excel:
        print(f"📂 Procesando: {archivo.name}")
        try:
            engine = "openpyxl" if archivo.suffix.lower() == ".xlsx" else "xlrd"
            with pd.ExcelFile(archivo, engine=engine) as xls:
                for hoja_nombre in xls.sheet_names:
                    df = xls.parse(hoja_nombre)
                    if df.empty:
                        continue

                    df["__archivo_origen__"] = archivo.name
                    df["__hoja_origen__"] = hoja_nombre

                    df = df.reindex(columns=columnas_ordenadas)

                    df.to_csv(
                        ruta_csv,
                        mode="a",
                        index=False,
                        header=(not wrote_header),
                        encoding="utf-8",
                        lineterminator="\n",
                        sep=";",
                    )
                    wrote_header = True
        except Exception as e:
            print(f"❌ Error al procesar {archivo.name}: {e}")

    print(f"✅ CSV único generado en: {ruta_csv}")
import win32com.client
import pandas as pd
import os
import time
import sys
import pythoncom # <--- NECESARIO PARA EL MENÚ

class SapBotPallet:
    def __init__(self):
        self.RUTA_SAP_EXPORT = r"C:\SAP_TEMP"
        self.TIPOS_ALMACEN = ["PBK", "PFW", "RCK", "CGO"]
        if not os.path.exists(self.RUTA_SAP_EXPORT):
            try: os.makedirs(self.RUTA_SAP_EXPORT)
            except: pass

    def get_excel_app(self):
        try:
            return win32com.client.GetActiveObject("Excel.Application")
        except:
            raise Exception("No encontré ningún Excel abierto. Por favor abre el archivo destino primero.")

    def leer_datos_sap_robusto(self, ruta_archivo):
        try:
            with open(ruta_archivo, 'r', encoding='latin1') as f:
                lineas = f.readlines()
            if not lineas: return pd.DataFrame()

            fila_header = -1
            sep = '\t'
            for i, linea in enumerate(lineas):
                if "MATERIAL" in linea.upper() and "LOTE" in linea.upper():
                    fila_header = i
                    if "|" in linea: sep = "|"
                    break
            
            if fila_header == -1: fila_header = 5

            df = pd.read_csv(ruta_archivo, sep=sep, encoding='latin1', skiprows=fila_header, on_bad_lines='skip', dtype=str)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.columns = df.columns.str.replace('|', '').str.strip()
            df = df.loc[:, df.columns != '']
            if not df.empty:
                col_ref = df.columns[0]
                df = df[~df[col_ref].str.contains('^----', na=False)]
                df = df[~df[col_ref].str.contains('^____', na=False)]
                for col in df.columns:
                    df[col] = df[col].astype(str).str.replace('|', '').str.strip()
            return df
        except Exception as e:
            print(f"Error leyendo archivo: {e}")
            return pd.DataFrame()

    def exportar_lx02(self, nombre_archivo):
        # 1. INICIALIZAR COM PARA HILOS (Vital para el Menú)
        pythoncom.CoInitialize()

        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
        except:
            raise Exception("SAP no está abierto o no hay sesión.")

        ruta_completa = os.path.join(self.RUTA_SAP_EXPORT, nombre_archivo)
        if os.path.exists(ruta_completa):
            try: os.remove(ruta_completa)
            except: pass

        # Entrar a LX02
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02"
        session.findById("wnd[0]").sendVKey(0)
        
        # 2. INTENTO DE FILTROS MANUALES (BLINDADO)
        # Si falla el ID del botón, saltamos este paso y confiamos en la variante
        try:
            print("   Configurando filtros...")
            # Intentar abrir selección múltiple
            session.findById("wnd[0]/usr/btn%_S1_LGTYP_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press() # Borrar anteriores
            
            base_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I"
            for i, tipo in enumerate(self.TIPOS_ALMACEN):
                try: session.findById(f"{base_id}[1,{i}]").text = tipo
                except: pass
            
            session.findById("wnd[1]/tbar[0]/btn[8]").press() # Aceptar filtros
        except:
            print("   ⚠️ No se pudo filtrar manualmente (Botón no encontrado). Se usará variante o descarga completa.")

        # 3. APLICAR VARIANTE (Si existe, sobrescribe lo manual, lo cual es bueno)
        try:
            session.findById("wnd[0]/usr/ctxtP_VARI").text = "/JESTAY"
        except: pass
        
        # EJECUTAR
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # 4. EXPORTAR
        print("   Exportando...")
        try:
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
            time.sleep(1)
            try:
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass
            
            time.sleep(0.5)
            try:
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.RUTA_SAP_EXPORT
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo
            except:
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ruta_completa

            session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except Exception as e:
            # Si falla la exportación, puede ser que no haya datos
            if "No existen" in session.findById("wnd[0]/sbar").Text:
                return None # Retornamos None para indicar vacío
            raise Exception(f"Fallo exportando LX02: {e}")

        # Esperar archivo
        for i in range(15):
            if os.path.exists(ruta_completa) and os.path.getsize(ruta_completa) > 0:
                return ruta_completa
            time.sleep(1)
        
        return None # Timeout

    def run(self):
        print("--- INICIANDO PALLET ALTURA ---")
        try:
            xl = self.get_excel_app()
            wb = xl.ActiveWorkbook
            if not wb: raise Exception("No hay libro activo en Excel. Abre el archivo destino.")
            
            try:
                nombre_archivo = str(wb.Sheets(1).Range("Z1").Value).strip()
            except: nombre_archivo = "None"
            
            if not nombre_archivo or nombre_archivo == "None": nombre_archivo = "lx02altura.txt"
            if not nombre_archivo.lower().endswith(".txt"): nombre_archivo += ".txt"

            ruta_export = self.exportar_lx02(nombre_archivo)
            
            if not ruta_export:
                print("⚠️ No se generó reporte (Posiblemente sin datos en SAP).")
                return

            df = self.leer_datos_sap_robusto(ruta_export)
            
            if df.empty: 
                print("⚠️ El reporte de SAP está vacío.")
                return

            # Pegar en Excel
            try:
                ws = wb.Sheets("BBD")
            except:
                ws = wb.Sheets.Add()
                ws.Name = "BBD"
            
            ws.Cells.Clear()
            ws.Cells.NumberFormat = "@" # Texto para conservar ceros
            
            columnas = list(df.columns)
            datos = df.values.tolist()
            
            # Pegado rápido
            ws.Range(ws.Cells(1, 1), ws.Cells(1, len(columnas))).Value = columnas
            if datos:
                ws.Range(ws.Cells(2, 1), ws.Cells(2 + len(datos) - 1, len(columnas))).Value = datos
            
            ws.Columns.AutoFit()
            print("✅ PROCESO PALLET TERMINADO CORRECTAMENTE.")
            
        except Exception as e:
            print(f"❌ Error Pallet: {e}")
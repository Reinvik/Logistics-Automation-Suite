import win32com.client
import pandas as pd
import os
import time
import sys
import pythoncom

class SapBotPallet:
    def __init__(self):
        self.RUTA_SAP_EXPORT = r"C:\SAP_TEMP"
        self.TIPOS_ALMACEN = ["PBK", "PFW", "RCK", "CGO"]
        if not os.path.exists(self.RUTA_SAP_EXPORT):
            try: os.makedirs(self.RUTA_SAP_EXPORT)
            except: pass

    def abrir_excel(self, ruta_archivo):
        """Abre el archivo Excel seleccionado."""
        try:
            # Intentamos conectar a Excel o abrir una instancia nueva
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True # Hacemos visible para que veas que trabaja
            
            # Abrimos el libro
            wb = xl.Workbooks.Open(ruta_archivo)
            return xl, wb
        except Exception as e:
            raise Exception(f"No se pudo abrir el Excel: {e}")

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

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nLX02"
        session.findById("wnd[0]").sendVKey(0)
        
        # Filtros
        try:
            print("   Configurando filtros...")
            session.findById("wnd[0]/usr/btn%_S1_LGTYP_%_APP_%-VALU_PUSH").press()
            session.findById("wnd[1]/tbar[0]/btn[16]").press() 
            base_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I"
            for i, tipo in enumerate(self.TIPOS_ALMACEN):
                try: session.findById(f"{base_id}[1,{i}]").text = tipo
                except: pass
            session.findById("wnd[1]/tbar[0]/btn[8]").press()
        except:
            print("   ⚠️ Botón de filtros no encontrado. Usando variante o descarga total.")

        try:
            session.findById("wnd[0]/usr/ctxtP_VARI").text = "/JESTAY"
        except: pass
        
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Exportar
        print("   Exportando reporte...")
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
            if "No existen" in session.findById("wnd[0]/sbar").Text:
                return None
            raise Exception(f"Fallo exportando: {e}")

        for i in range(15):
            if os.path.exists(ruta_completa) and os.path.getsize(ruta_completa) > 0:
                return ruta_completa
            time.sleep(1)
        return None

    def run(self, ruta_excel):
        print("--- INICIANDO PALLET ALTURA ---")
        
        # 1. ABRIR EL EXCEL SELECCIONADO
        print(f"Abriendo Excel: {os.path.basename(ruta_excel)}")
        xl, wb = self.abrir_excel(ruta_excel)
        
        # 2. OBTENER NOMBRE DE REPORTE (De celda Z1 o default)
        try:
            nombre_archivo = str(wb.Sheets(1).Range("Z1").Value).strip()
        except: nombre_archivo = "None"
        
        if not nombre_archivo or nombre_archivo == "None": nombre_archivo = "lx02altura.txt"
        if not nombre_archivo.lower().endswith(".txt"): nombre_archivo += ".txt"

        # 3. EXPORTAR DESDE SAP
        ruta_export = self.exportar_lx02(nombre_archivo)
        
        if not ruta_export:
            print("⚠️ No se generó reporte (Posiblemente sin datos).")
            return

        # 4. LEER Y PEGAR
        df = self.leer_datos_sap_robusto(ruta_export)
        
        if df.empty: 
            print("⚠️ Reporte vacío.")
            return

        print("Pegando datos en pestaña 'BBD'...")
        try:
            ws = wb.Sheets("BBD")
        except:
            ws = wb.Sheets.Add()
            ws.Name = "BBD"
        
        ws.Cells.Clear()
        ws.Cells.NumberFormat = "@" 
        
        columnas = list(df.columns)
        datos = df.values.tolist()
        
        ws.Range(ws.Cells(1, 1), ws.Cells(1, len(columnas))).Value = columnas
        if datos:
            ws.Range(ws.Cells(2, 1), ws.Cells(2 + len(datos) - 1, len(columnas))).Value = datos
        
        ws.Columns.AutoFit()
        
        # Guardar (Opcional, si quieres que guarde solo, descomenta la linea)
        # wb.Save()
        
        print("✅ PROCESO PALLET TERMINADO.")
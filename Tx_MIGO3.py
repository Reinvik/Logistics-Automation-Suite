import win32com.client
import sys
import time
import pandas as pd
import os
import math
import shutil
import re # Para buscar el numero de documento
from datetime import datetime

class SapMigoBotTurbo:
    def __init__(self):
        self.session = None
        self.table = None
        self.cols = {} 
        self.connect_to_sap()

    def connect_to_sap(self):
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
            connection = application.Children(0)
            self.session = connection.Children(0)
            print("--- Conectado a SAP exitosamente ---")
        except Exception as e:
            print("Error conectando a SAP. Asegúrate de tener SAP abierto.")
            sys.exit(1)

    def start_transaction(self):
        print("--- Iniciando Transacción MIGO ---")
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(2.0)
            try: self.session.findById("wnd[0]").maximize()
            except: pass
            try:
                self.session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell/shellcont[0]/shell").pressButton("CLOSE")
            except: pass
        except Exception as e:
            print(f"Advertencia al abrir MIGO: {e}")

    def clean_data(self, df):
        df = df.astype(str)
        def remove_trailing_zero(val):
            val = val.strip()
            if val.endswith(".0"):
                return val[:-2]
            return val
        for col in df.columns:
            df[col] = df[col].apply(remove_trailing_zero)
            df[col] = df[col].replace({'nan': '', 'None': ''})
        return df

    def read_excel_dynamic(self, excel_path):
        filename = os.path.basename(excel_path)
        try:
            excel_app = win32com.client.GetActiveObject("Excel.Application")
            wb_found = None
            for wb in excel_app.Workbooks:
                if wb.Name == filename:
                    wb_found = wb
                    break
            if wb_found:
                print("   [INFO] Leyendo desde Excel ABIERTO (En vivo)...")
                ws = wb_found.Worksheets(1)
                used_range = ws.UsedRange.Value
                if used_range:
                    headers = list(used_range[0])
                    data = used_range[1:]
                    df = pd.DataFrame(data, columns=headers)
                    df = df.loc[:, ~df.columns.duplicated()]
                    return self.clean_data(df)
        except Exception:
            pass

        print("   [INFO] Leyendo desde archivo en disco...")
        temp_file = excel_path + ".temp_bot.xlsx"
        try:
            shutil.copyfile(excel_path, temp_file)
            df = pd.read_excel(temp_file, dtype=str)
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.fillna("")
            return self.clean_data(df)
        except Exception as e:
            print(f"[ERROR] No se pudo leer el Excel: {e}")
            sys.exit(1)
        finally:
            if os.path.exists(temp_file):
                try: os.remove(temp_file)
                except: pass

    def find_migo_table(self):
        possible_paths = [
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM",
            "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0008/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"
        ]
        start_time = time.time()
        while time.time() - start_time < 5:
            for path in possible_paths:
                try:
                    table = self.session.findById(path)
                    _ = table.RowCount 
                    return table
                except: continue
            time.sleep(0.2)
        print("Error Crítico: No se encontró la tabla.")
        sys.exit(1)

    def map_columns(self):
        print("--- Mapeando columnas según variante ---")
        search_keys = {
            "MAT": ["MATNR", "MAKTX"], "QTY": ["ERFMG"], "UNIT": ["ERFME"],
            "PLANT_O": ["WERKS", "NAME1"], "LOC_O": ["LGORT", "LGOBE"], "BATCH_O": ["CHARG"],
            "PLANT_D": ["UMWRK", "UMNAME1"], "LOC_D": ["UMLGO", "UMLGOBE"], "BATCH_D": ["UMCHA"]
        }
        self.cols = {k: -1 for k in search_keys}
        try:
            for i in range(self.table.Columns.Count):
                name = self.table.Columns.Item(i).Name
                for key, possibilities in search_keys.items():
                    if self.cols[key] == -1:
                        for p in possibilities:
                            if p in name:
                                self.cols[key] = i
                                break
        except: pass
        if self.cols["BATCH_D"] != -1:
            print(f"   -> Columna Lote Destino detectada en índice: {self.cols['BATCH_D']}")
        else:
            print("   [AVISO] Lote Destino (UMCHA) no detectado.")

    def write_header(self, text):
        if pd.isna(text) or str(text).strip() == "": return
        try:
            self.session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").Text = str(text)
        except: pass

    def set_val_robust(self, col_idx, row_vis, val):
        if hasattr(val, 'iloc'): val = val.iloc[0]
        if pd.isna(val) or str(val).strip() == "": return
        if col_idx == -1: return

        for _ in range(2):
            try:
                if col_idx >= 10:
                    try:
                        if self.table.FirstVisibleColumn < (col_idx - 5):
                            self.table.FirstVisibleColumn = col_idx - 2
                    except: pass
                else:
                    try:
                        if self.table.FirstVisibleColumn > 0:
                            self.table.FirstVisibleColumn = 0
                    except: pass
                
                cell = self.table.GetCell(row_vis, col_idx)
                try: cell.SetFocus()
                except: pass

                if cell.Changeable:
                    cell.Text = str(val)
                    return 
            except:
                try: self.table = self.find_migo_table()
                except: pass
        print(f"   [!] Fallo escritura: Col {col_idx} Row {row_vis} Val {val}")

    def finalizar_y_mb51(self):
        print("\n--- Guardando Documento ---")
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(2.0)
            
            # Capturar mensaje
            mensaje = self.session.findById("wnd[0]/sbar").Text
            print(f"Mensaje SAP: {mensaje}")
            
            numeros = re.findall(r'\d+', mensaje)
            doc_material = ""
            for num in numeros:
                if len(num) >= 9:
                    doc_material = num
                    break
            
            if doc_material:
                print(f"Detectado Documento: {doc_material}. Yendo a MB51...")
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.5)
                try:
                    self.session.findById("wnd[0]/usr/ctxtMBLNR-LOW").text = doc_material
                    self.session.findById("wnd[0]/usr/txtMJAHR-LOW").text = str(datetime.now().year)
                    self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                    print("--- MB51 Ejecutada ---")
                except:
                    print("Error llenando campos de MB51 (IDs distintos).")
            else:
                print("No se detectó número de documento para ir a MB51.")
        except Exception as e:
            print(f"Error al finalizar: {e}")

    def run(self, excel_path):
        self.start_transaction()
        df = self.read_excel_dynamic(excel_path)
        if df is None or df.empty: return

        self.table = self.find_migo_table()
        self.map_columns()
        df.columns = df.columns.str.strip()
        
        if 'Texto_Cabecera' in df.columns:
             header_val = df.iloc[0]['Texto_Cabecera']
             if hasattr(header_val, 'iloc'): header_val = header_val.iloc[0]
             self.write_header(header_val)

        BLOCK_SIZE = 10
        total_rows = len(df)
        num_blocks = math.ceil(total_rows / BLOCK_SIZE)
        
        print(f"Iniciando carga: {total_rows} registros en {num_blocks} bloques.")
        current_sap_scroll = 0 

        for block_idx in range(num_blocks):
            start_idx = block_idx * BLOCK_SIZE
            end_idx = min((block_idx + 1) * BLOCK_SIZE, total_rows)
            print(f"\n--- Bloque {block_idx + 1} ---")
            block_df = df.iloc[start_idx:end_idx]
            
            self.table = self.find_migo_table()
            try: self.table.VerticalScrollbar.Position = current_sap_scroll
            except: 
                self.table = self.find_migo_table()
                self.table.VerticalScrollbar.Position = current_sap_scroll
            time.sleep(0.3)
            self.table = self.find_migo_table()

            # PASO 1
            print("   -> Llenando Origen...")
            for i in range(len(block_df)):
                row_data = block_df.iloc[i]
                visual_row = i 
                self.set_val_robust(self.cols["MAT"], visual_row, row_data.get('Material'))
                self.set_val_robust(self.cols["QTY"], visual_row, row_data.get('Cantidad'))
                self.set_val_robust(self.cols["UNIT"], visual_row, row_data.get('Unidad'))
                self.set_val_robust(self.cols["LOC_O"], visual_row, row_data.get('Alm_Orig'))
                if 'Centro_Orig' in row_data:
                    self.set_val_robust(self.cols["PLANT_O"], visual_row, row_data.get('Centro_Orig'))

            print("   -> Validando Origen...")
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.2)
            self.table = self.find_migo_table()
            try: 
                if self.table.VerticalScrollbar.Position != current_sap_scroll:
                    self.table.VerticalScrollbar.Position = current_sap_scroll
            except: pass

            # PASO 2
            print("   -> Llenando Lote/Destinos...")
            has_dest = False
            for i in range(len(block_df)):
                row_data = block_df.iloc[i]
                visual_row = i
                self.set_val_robust(self.cols["BATCH_O"], visual_row, row_data.get('Lote_Orig'))
                if 'Centro_Dest' in row_data and str(row_data.get('Centro_Dest')).strip() != "":
                    has_dest = True
                    self.set_val_robust(self.cols["PLANT_D"], visual_row, row_data.get('Centro_Dest'))
                    self.set_val_robust(self.cols["LOC_D"], visual_row, row_data.get('Alm_Dest'))
            
            if has_dest:
                print("   -> Validando Destinos...")
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.2)
                self.table = self.find_migo_table()
                try: 
                    if self.table.VerticalScrollbar.Position != current_sap_scroll:
                        self.table.VerticalScrollbar.Position = current_sap_scroll
                except: pass

                print("   -> Llenando Lote Destino...")
                try: self.table.FirstVisibleColumn = 12
                except: pass
                for i in range(len(block_df)):
                    row_data = block_df.iloc[i]
                    if 'Lote_Dest' in row_data:
                        visual_row = i
                        self.set_val_robust(self.cols["BATCH_D"], visual_row, row_data.get('Lote_Dest'))
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(1.0)
            else:
                 self.session.findById("wnd[0]").sendVKey(0)
                 time.sleep(1.0)

            current_sap_scroll += len(block_df)
            if current_sap_scroll >= 20:
                print("--- Límite de 20 líneas alcanzado ---")
                break

        print("--- Carga finalizada ---")
        # LLAMAR A LA FUNCIÓN DE GUARDADO Y MB51
        self.finalizar_y_mb51()

if __name__ == "__main__":
    carpeta = r"c:\Users\ariel.mella\OneDrive - CIAL Alimentos\Archivos de Operación  Outbound CD - 16.-Inventario Critico"
    nombre_archivo = "carga_migo.xlsx" 
    archivo_excel = os.path.join(carpeta, nombre_archivo)

    if os.path.exists(archivo_excel):
        bot = SapMigoBotTurbo()
        bot.run(archivo_excel)
    else:
        print(f"No encuentro el archivo: {archivo_excel}")
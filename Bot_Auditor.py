import win32com.client
import pandas as pd
import datetime
import sys
import time
import os
import pythoncom  # <--- IMPORTANTE: Necesario para trabajar con hilos

class SapBotAuditor:
    def __init__(self):
        self.RUTA_BASE = r"C:\SAP_TEMP"
        if not os.path.exists(self.RUTA_BASE):
            try: os.makedirs(self.RUTA_BASE)
            except: self.RUTA_BASE = os.path.join(os.getenv('TEMP'), 'SAP_BOT')
            
    def conectar_sap(self):
        try:
            # ESTA LÍNEA ES LA MAGIA PARA QUE FUNCIONE DESDE EL MENÚ
            pythoncom.CoInitialize() 
            
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not SapGuiAuto:
                raise Exception("No se encontró el objeto SAPGUI.")
                
            application = SapGuiAuto.GetScriptingEngine
            if not application:
                raise Exception("Scripting Engine no disponible.")
                
            connection = application.Children(0)
            if not connection:
                raise Exception("No hay conexión activa.")
                
            session = connection.Children(0)
            if not session:
                raise Exception("No hay sesión activa.")
                
            return session
        except Exception as e:
            # Ahora mostramos el error real técnico, no un mensaje genérico
            raise Exception(f"Fallo de conexión técnica: {e}")

    def guardar_como_macro(self, session, nombre_archivo):
        ruta_completa = os.path.join(self.RUTA_BASE, nombre_archivo)
        print(f"   💾 Exportando {nombre_archivo}...")
        
        if os.path.exists(ruta_completa):
            try: os.remove(ruta_completa)
            except: pass

        try:
            session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select()
            time.sleep(0.5)
            try:
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass 
            
            time.sleep(0.5)

            try:
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.RUTA_BASE
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo
            except:
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ruta_completa

            session.findById("wnd[1]/tbar[0]/btn[11]").press() 
        except Exception as e:
            print(f"      ⚠️ Error exportando: {e}")
            return False

        for i in range(15):
            if os.path.exists(ruta_completa) and os.path.getsize(ruta_completa) > 0:
                return True
            time.sleep(1)
        return False

    def leer_archivo_sap(self, ruta):
        if not os.path.exists(ruta): return pd.DataFrame()
        try:
            with open(ruta, 'r', encoding='latin1') as f:
                lineas = f.readlines()
            
            fila_header = 0
            for i, linea in enumerate(lineas):
                if "Material" in linea or "Mat." in linea or "Centro" in linea:
                    fila_header = i
                    break
            
            df = pd.read_csv(ruta, sep="\t", encoding='latin1', skiprows=fila_header, on_bad_lines='skip', dtype=str)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.dropna(how='all', inplace=True)
            return df
        except: 
            return pd.DataFrame()

    def run(self, almacen_argumento=None):
        print("--- INICIANDO AUDITOR DE STOCK ---")
        
        if not almacen_argumento:
            print("❌ Error: No se recibió ningún almacén para auditar.")
            return
            
        ALMACEN = almacen_argumento
        CENTRO = "SGSJ"
        MOVIMIENTOS = ["301", "653", "311", "101", "102"] 
        DIAS_HISTORIA = 90
        
        FILE_MB52 = f"dump_mb52_{ALMACEN}.txt"
        FILE_MB51 = f"dump_mb51_{ALMACEN}.txt"
        RUTA_MB52 = os.path.join(self.RUTA_BASE, FILE_MB52)
        RUTA_MB51 = os.path.join(self.RUTA_BASE, FILE_MB51)

        try:
            session = self.conectar_sap()
        except Exception as e:
            print(f"❌ ERROR CRÍTICO CONECTANDO SAP: {e}")
            return

        # 1. MB52
        print(f"📥 1. Descargando MB52 ({ALMACEN})...")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB52"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
            session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = ""
            session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = CENTRO
            session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ALMACEN
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            if "No existen" in session.findById("wnd[0]/sbar").Text:
                print(f"⚠️ El almacén {ALMACEN} está vacío.")
                return

            if not self.guardar_como_macro(session, FILE_MB52):
                print("Error exportando MB52.")
                return
            df_mb52 = self.leer_archivo_sap(RUTA_MB52)

            # Filtrar stock
            try:
                col_stock = next((c for c in df_mb52.columns if "Libr" in c or "Utiliz" in c), None)
                if col_stock:
                    df_mb52[col_stock] = df_mb52[col_stock].astype(str).str.replace(".", "").str.replace(",", ".")
                    df_mb52[col_stock] = pd.to_numeric(df_mb52[col_stock], errors='coerce')
                    df_mb52 = df_mb52[df_mb52[col_stock] > 0]
            except: pass

            if df_mb52.empty:
                print("Stock es 0 o error leyendo MB52.")
                return

            # 2. MB51
            print(f"🕵️ 2. Descargando MB51 (Últimos {DIAS_HISTORIA} días)...")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB51"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
            session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = ""
            try: session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "" 
            except: pass
            session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = CENTRO
            session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = ALMACEN

            try:
                session.findById("wnd[0]/usr/btn%_BWART_%_APP_%-VALU_PUSH").press()
                session.findById("wnd[1]/tbar[0]/btn[16]").press()
                base_id = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I"
                for idx, mov in enumerate(MOVIMIENTOS):
                    try: session.findById(f"{base_id}[1,{idx}]").text = mov
                    except: pass
                session.findById("wnd[1]/tbar[0]/btn[8]").press()
            except: pass

            hoy = datetime.date.today()
            inicio = hoy - datetime.timedelta(days=DIAS_HISTORIA)
            session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = inicio.strftime("%d.%m.%Y")
            session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = hoy.strftime("%d.%m.%Y")
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            df_mb51 = pd.DataFrame()
            msg = session.findById("wnd[0]/sbar").Text
            if "No existen" not in msg and "No se han seleccionado" not in msg:
                if self.guardar_como_macro(session, FILE_MB51):
                    df_mb51 = self.leer_archivo_sap(RUTA_MB51)

            # 3. CRUCE
            print("🧠 3. Cruzando Datos...")
            def buscar_col(df, keywords):
                for col in df.columns:
                    for kw in keywords:
                        if kw.upper() in col.upper(): return col
                return None

            def normalizar(serie):
                return serie.astype(str).str.split('.').str[0].str.strip().str.lstrip("0")

            col_mat_52 = buscar_col(df_mb52, ["Material", "Mat."])
            col_lot_52 = buscar_col(df_mb52, ["Lote", "Lot"])
            
            if not col_mat_52 and len(df_mb52.columns) > 0: col_mat_52 = df_mb52.columns[0]
            if not col_lot_52 and len(df_mb52.columns) > 4: col_lot_52 = df_mb52.columns[4]

            if col_mat_52:
                df_mb52["KEY_MAT"] = normalizar(df_mb52[col_mat_52])
                df_mb52["KEY_LOT"] = df_mb52[col_lot_52].astype(str).str.strip() if col_lot_52 else ""

                if not df_mb51.empty:
                    col_mat_51 = buscar_col(df_mb51, ["Material", "Mat."])
                    col_lot_51 = buscar_col(df_mb51, ["Lote", "Lot"])
                    col_fech_51 = buscar_col(df_mb51, ["Fe.contab", "Fecha de con", "Contabiliz", "Fecha"])

                    if not col_mat_51 and len(df_mb51.columns) > 0: col_mat_51 = df_mb51.columns[0]
                    
                    if not col_fech_51:
                        for c in df_mb51.columns:
                            if ("FE" in c.upper() or "CONTA" in c.upper()) and "DOC" not in c.upper():
                                col_fech_51 = c
                                break
                    
                    if col_mat_51 and col_fech_51:
                        df_mb51["KEY_MAT"] = normalizar(df_mb51[col_mat_51])
                        df_mb51["KEY_LOT"] = df_mb51[col_lot_51].astype(str).str.strip() if col_lot_51 else ""
                        
                        df_mb51[col_fech_51] = df_mb51[col_fech_51].astype(str).str.strip()
                        df_mb51["DT"] = pd.to_datetime(df_mb51[col_fech_51], format="%d.%m.%Y", errors='coerce')
                        if df_mb51["DT"].isna().all():
                             df_mb51["DT"] = pd.to_datetime(df_mb51[col_fech_51], dayfirst=True, errors='coerce')

                        df_mb51 = df_mb51.sort_values("DT", ascending=False).drop_duplicates(["KEY_MAT", "KEY_LOT"])
                        df_final = pd.merge(df_mb52, df_mb51[["KEY_MAT", "KEY_LOT", "DT"]], on=["KEY_MAT", "KEY_LOT"], how="left")
                    else:
                        df_final = df_mb52.copy()
                        df_final["DT"] = pd.NaT
                else:
                    df_final = df_mb52.copy()
                    df_final["DT"] = pd.NaT

                now = pd.Timestamp.now()
                df_final["DIAS_SIN_MOV"] = (now - df_final["DT"]).dt.days
                
                def tag(d):
                    if pd.isna(d): return f"💀 CRÍTICO (> {DIAS_HISTORIA} días)"
                    if d <= 2: return "🟢 FRESCO (0-2 días)"
                    if d <= 7: return "🟡 PENDIENTE (3-7 días)"
                    return f"🔴 LENTO ({int(d)} días)"

                df_final["ESTADO"] = df_final["DIAS_SIN_MOV"].apply(tag)
                
                for c in ["KEY_MAT", "KEY_LOT"]:
                    if c in df_final.columns: del df_final[c]
                
                df_final = df_final.sort_values("DIAS_SIN_MOV", ascending=False, na_position='first')

                ruta_escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
                timestamp = datetime.datetime.now().strftime("%H%M%S")
                ruta_reporte = os.path.join(ruta_escritorio, f"Auditoria_{ALMACEN}_{timestamp}.xlsx")
                
                print("💾 Generando Excel...")
                with pd.ExcelWriter(ruta_reporte) as writer:
                    df_final.to_excel(writer, sheet_name='Reporte_Zombies', index=False)
                    df_mb52.to_excel(writer, sheet_name='RAW_MB52', index=False)
                    if not df_mb51.empty:
                        df_mb51.to_excel(writer, sheet_name='RAW_MB51', index=False)
                
                print(f"🚀 LISTO: {ruta_reporte}")
                os.startfile(ruta_reporte)
            else:
                print("❌ Error de columnas.")

        except Exception as e:
            print(f"❌ Error General: {e}")
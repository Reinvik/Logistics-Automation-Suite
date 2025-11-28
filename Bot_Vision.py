import google.generativeai as genai
import pandas as pd
import os
import glob
import re
from PIL import Image
import json
from datetime import datetime, timedelta

class BotVisionPizarra:
    def run(self):
        print("--- INICIANDO BOT DE VISIÓN IA ---")
        
        API_KEY = "AIzaSyDLu2CjsEZPPM_31c9-Da_zyn0v7SM_5Hc" 
        ANIO_ACTUAL = datetime.now().year
        
        MOTIVOS_DICT = {
            1: "1.- Problemas con SAP", 2: "2.- Atraso en Picking / falta producto",
            3: "3.- Error detectado control", 4: "4.- Atraso en Planificación",
            5: "5.- Falta dotación despacho", 6: "6.- Retraso de llegada de camión",
            7: "7.- Falla de equipos moviles", 8: "8.- Otros"
        }

        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]

        try:
            genai.configure(api_key=API_KEY)
            model = genai.GenerativeModel('gemini-2.5-flash', safety_settings=safety_settings)

            ruta_base = os.path.join(os.path.expanduser("~"), r"OneDrive - CIAL Alimentos\Archivos de Operación  Outbound CD - 16.-Inventario Critico")
            archivo_master = os.path.join(ruta_base, "Base_Datos_Pizarra_Master.xlsx")
            patron = os.path.join(ruta_base, "pizarra_semana_*.jpg")

            lista = glob.glob(patron)
            if not lista: 
                raise Exception("No hay fotos de pizarra (pizarra_semana_X.jpg) en la carpeta.")
            
            archivo_foto = max(lista, key=os.path.getctime)
            print(f"📸 Analizando imagen: {os.path.basename(archivo_foto)}")

            match = re.search(r"semana_(\d+)", os.path.basename(archivo_foto), re.IGNORECASE)
            semana_archivo = int(match.group(1)) if match else 0

            prompt = """
            Actúa como Auditor Logístico. Extrae datos de la imagen.
            JSON ESPERADO: 
            {
                "Semana_Detectada": 47, 
                "Datos": [
                    {
                        "Turno": "...", "Dia": "...", "Zonal": "...", 
                        "Hora_Plan": "HH:MM", "Hora_Real": "HH:MM", 
                        "Motivo_Leido": 6, "Sigla_T": "...", "Sigla_C": "..."
                    }
                ]
            }
            """
            
            img = Image.open(archivo_foto)
            response = model.generate_content([prompt, img], generation_config={"response_mime_type": "application/json"})
            
            if not response.text: raise Exception("IA devolvió respuesta vacía.")
            
            data_raw = json.loads(response.text.strip())
            df = pd.DataFrame(data_raw["Datos"])

            semana_final = semana_archivo if semana_archivo > 0 else data_raw.get("Semana_Detectada", 0)
            if semana_final == 0: semana_final = datetime.now().isocalendar()[1]

            lunes_iso = datetime.strptime(f'{ANIO_ACTUAL}-W{semana_final}-1', "%G-W%V-%u")

            def limpiar_hora(h):
                if pd.isna(h) or str(h).strip() == "": return None
                h = str(h).replace(".", ":").strip()
                try: return datetime.strptime(h, "%H:%M").time()
                except: return None

            def validar_retraso(row):
                plan = limpiar_hora(row.get('Hora_Plan'))
                real = limpiar_hora(row.get('Hora_Real'))
                motivo_ia = row.get('Motivo_Leido')
                
                if plan is None or real is None: return motivo_ia, "Falta Hora"

                dt_plan = datetime.combine(datetime.today(), plan)
                dt_real = datetime.combine(datetime.today(), real)
                
                if dt_real < dt_plan and dt_plan.hour > 20 and dt_real.hour < 6:
                    dt_real += timedelta(days=1)
                
                delta_minutos = (dt_real - dt_plan).total_seconds() / 60
                
                if delta_minutos <= 0: return None, "A Tiempo" 
                else:
                    if pd.isna(motivo_ia) or str(motivo_ia) == "": return None, "⚠ ATRASO SIN MOTIVO"
                    else: return motivo_ia, f"Atraso de {int(delta_minutos)} min"

            if not df.empty:
                validacion = df.apply(validar_retraso, axis=1, result_type='expand')
                df['Motivo_Final'] = validacion[0]
                df['Estado_Auditoria'] = validacion[1]
                
                def get_motivo_texto(num):
                    if pd.isna(num): return "" 
                    try: return MOTIVOS_DICT.get(int(num), str(num))
                    except: return str(num)

                df['Descripcion_Retraso'] = df['Motivo_Final'].apply(get_motivo_texto)

                mapa_dias = {"Lunes": 0, "Martes": 1, "Miércoles": 2, "Jueves": 3, "Viernes": 4, "Sábado": 5, "Domingo": 6}
                
                def get_fecha(dia):
                    d = mapa_dias.get(str(dia).capitalize().strip(), 0)
                    return (lunes_iso + timedelta(days=d)).strftime("%d.%m.%Y")
                
                df['Fecha'] = df['Dia'].apply(get_fecha)
                df['Semana'] = semana_final
                df.fillna("", inplace=True)
                
                cols = ["Fecha", "Semana", "Turno", "Dia", "Zonal", "Hora_Plan", "Hora_Real", "Estado_Auditoria", "Descripcion_Retraso", "Sigla_T", "Sigla_C"]
                df_final = df[[c for c in cols if c in df.columns]]

                if os.path.exists(archivo_master):
                    try:
                        df_master = pd.read_excel(archivo_master)
                        df_total = pd.concat([df_master, df_final], ignore_index=True)
                        df_total.drop_duplicates(subset=['Semana', 'Turno', 'Dia', 'Zonal'], keep='last', inplace=True)
                    except: df_total = df_final
                else: df_total = df_final
                    
                df_total.to_excel(archivo_master, index=False)
                print(f"✅ AUDITORÍA FINALIZADA. Archivo: {os.path.basename(archivo_master)}")
            else:
                print("⚠️ La IA no detectó datos.")

        except Exception as e:
            print(f"❌ Error Técnico: {e}")
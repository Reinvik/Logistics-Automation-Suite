import customtkinter as ctk
import threading
import sys
import os
from tkinter import filedialog

# --- IMPORTAR TODOS LOS BOTS ---
print("Cargando bots...")
try:
    from Tx_MIGO3 import SapMigoBotTurbo 
    print("✅ MIGO cargado.")
except: SapMigoBotTurbo = None

try:
    from Bot_Pallet import SapBotPallet
    print("✅ Pallet cargado.")
except: SapBotPallet = None

try:
    from Bot_Transporte import SapBotTransporte
    print("✅ Transporte cargado.")
except: SapBotTransporte = None

try:
    from Bot_Vision import BotVisionPizarra
    print("✅ Visión cargado.")
except: BotVisionPizarra = None

try:
    from Bot_Auditor import SapBotAuditor
    print("✅ Auditor cargado.")
except: SapBotAuditor = None

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Panel de Control SAP - CIAL")
        self.geometry("900x600")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.bot_actual_class = None

        # --- MENU LATERAL ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=4, sticky="nsew")
        
        ctk.CTkLabel(self.sidebar, text="🤖 BOTS SAP", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        self.crear_boton("Carga MIGO", self.panel_migo)
        self.crear_boton("Pallet Altura (LX02)", self.panel_pallet)
        self.crear_boton("Transporte (VT11)", self.panel_transporte)
        self.crear_boton("Visión Pizarra (IA)", self.panel_vision)
        self.crear_boton("Auditor Zombies", self.panel_auditor)

        # --- PANEL CENTRAL ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        
        self.setup_ui_generica()

    def crear_boton(self, texto, comando):
        btn = ctk.CTkButton(self.sidebar, text=texto, command=comando, height=40)
        btn.pack(pady=5, padx=20)

    def setup_ui_generica(self):
        self.lbl_title = ctk.CTkLabel(self.main_frame, text="Selecciona un Bot", font=ctk.CTkFont(size=24))
        self.lbl_title.pack(pady=20)
        
        self.lbl_info = ctk.CTkLabel(self.main_frame, text="", justify="left")
        self.lbl_info.pack(pady=10)
        
        self.entry_file = ctk.CTkEntry(self.main_frame, placeholder_text="Ruta archivo...", width=400)
        self.btn_select = ctk.CTkButton(self.main_frame, text="Buscar Archivo", command=self.sel_archivo, fg_color="gray")
        
        self.btn_run = ctk.CTkButton(self.main_frame, text="EJECUTAR", command=self.run_migo_thread, height=50, fg_color="#2CC985", state="disabled")
        self.btn_run.pack(pady=20)
        
        self.log_box = ctk.CTkTextbox(self.main_frame, width=600, height=250)
        self.log_box.pack(fill="both", expand=True)

    def reset_ui(self, titulo, info, necesita_archivo=False):
        self.lbl_title.configure(text=titulo)
        self.lbl_info.configure(text=info)
        self.log_box.delete("0.0", "end")
        
        if necesita_archivo:
            self.entry_file.pack(pady=5)
            self.btn_select.pack(pady=5)
        else:
            self.entry_file.pack_forget()
            self.btn_select.pack_forget()
        
        self.btn_run.configure(state="normal")

    # --- PANELES ---
    def panel_migo(self):
        self.bot_actual_class = SapMigoBotTurbo
        self.reset_ui("Carga MIGO", "Requiere Excel 'carga_migo.xlsx'.", True)

    def panel_pallet(self):
        self.bot_actual_class = SapBotPallet
        # CAMBIO 1: Ahora necesita archivo = True
        self.reset_ui("Pallet Altura", "Selecciona el Excel donde pegar los datos.", False)

    def panel_transporte(self):
        self.bot_actual_class = SapBotTransporte
        self.reset_ui("Reporte Transporte", "Extrae VT11 de ayer y envía correo.", False)

    def panel_vision(self):
        self.bot_actual_class = BotVisionPizarra
        self.reset_ui("Visión IA", "Analiza la foto más reciente en la carpeta OneDrive.", False)

    def panel_auditor(self):
        self.bot_actual_class = SapBotAuditor
        self.reset_ui("Auditor Stock", "Al ejecutar, te pedirá el Almacén.", False)

    # --- LOGICA ---
    def sel_archivo(self):
        f = filedialog.askopenfilename()
        if f: 
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, f)

    def log(self, msg):
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")

    def run_migo_thread(self):
        ruta = None
        arg_extra = None

        # CAMBIO 2: Agregamos SapBotPallet a la lista de bots que piden archivo
        if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet]:
            ruta = self.entry_file.get()
            if not ruta:
                self.log("[ERROR] Selecciona un archivo Excel primero.")
                return

        elif self.bot_actual_class == SapBotAuditor:
            almacenes = ["SGVT", "CDNW", "SGTR", "TAVI", "SGSD", "AVAS", "SGBC", "SGVE", "SGEN"]
            dialog = ctk.CTkInputDialog(text=f"Almacenes: {', '.join(almacenes)}\n\nEscribe el Almacén:", title="Auditor MM")
            arg_extra = dialog.get_input()
            if not arg_extra:
                self.log("Cancelado.")
                return
            arg_extra = arg_extra.upper().strip()

        self.btn_run.configure(state="disabled", text="Ejecutando...")
        self.log(f"--- Iniciando {self.bot_actual_class.__name__}... ---")
        
        threading.Thread(target=self.execute, args=(ruta, arg_extra)).start()

    def execute(self, ruta, arg_extra):
        try:
            if not self.bot_actual_class: raise Exception("Bot no encontrado.")
            bot = self.bot_actual_class()
            
            # CAMBIO 3: Lógica de ejecución
            if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet]:
                bot.run(ruta) # Ambos reciben la ruta ahora
            
            elif self.bot_actual_class == SapBotAuditor:
                bot.run(arg_extra)
            
            else:
                bot.run()
                
            self.log("✅ PROCESO FINALIZADO CON ÉXITO")
        except Exception as e:
            self.log(f"❌ ERROR: {e}")
        finally:
            self.btn_run.configure(state="normal", text="EJECUTAR")

if __name__ == "__main__":
    app = App()
    app.mainloop()
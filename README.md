
### `main.py`

```python
"""Panel de control para bots SAP.

Este módulo contiene la interfaz gráfica principal que permite al usuario seleccionar
y ejecutar distintos bots de SAP mediante un menú lateral.  Utiliza CustomTkinter
para crear una aplicación moderna y responsiva.

Las clases de bots se importan desde el paquete local `bots`.
"""

import customtkinter as ctk
import threading
from tkinter import filedialog

# Importar bots desde el paquete `bots`
from bots.Tx_MIGO3 import SapMigoBotTurbo
from bots.Bot_Pallet import SapBotPallet
from bots.Bot_Transporte import SapBotTransporte
from bots.Bot_Vision import BotVisionPizarra
from bots.Bot_Auditor import SapBotAuditor

# Configuración de apariencia de CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    """Clase principal de la aplicación GUI.

    Contiene la barra lateral con los distintos bots disponibles y un panel
    central donde se muestran los controles para cada bot.  Utiliza hilos
    (`threading.Thread`) para ejecutar los bots en segundo plano sin bloquear
    la interfaz de usuario.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("Panel de Control SAP - CIAL")
        self.geometry("900x600")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.bot_actual_class = None  # Clase del bot actualmente seleccionado

        # --- MENÚ LATERAL ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=4, sticky="nsew")

        ctk.CTkLabel(
            self.sidebar,
            text="🤖 BOTS SAP",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=20)

        # Botones del menú lateral
        self.crear_boton("Carga MIGO", self.panel_migo)
        self.crear_boton("Pallet Altura (LX02)", self.panel_pallet)
        self.crear_boton("Transporte (VT11)", self.panel_transporte)
        self.crear_boton("Visión Pizarra (IA)", self.panel_vision)
        self.crear_boton("Auditor Zombies", self.panel_auditor)

        # --- PANEL CENTRAL ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self.setup_ui_generica()

    def crear_boton(self, texto: str, comando) -> None:
        """Crea un botón en la barra lateral.

        Args:
            texto: Texto que se mostrará en el botón.
            comando: Función que se ejecutará al hacer clic en el botón.
        """
        btn = ctk.CTkButton(self.sidebar, text=texto, command=comando, height=40)
        btn.pack(pady=5, padx=20)

    def setup_ui_generica(self) -> None:
        """Configura los widgets comunes del panel central."""
        self.lbl_title = ctk.CTkLabel(
            self.main_frame,
            text="Selecciona un Bot",
            font=ctk.CTkFont(size=24),
        )
        self.lbl_title.pack(pady=20)

        self.lbl_info = ctk.CTkLabel(self.main_frame, text="", justify="left")
        self.lbl_info.pack(pady=10)

        self.entry_file = ctk.CTkEntry(
            self.main_frame, placeholder_text="Ruta archivo...", width=400
        )
        self.btn_select = ctk.CTkButton(
            self.main_frame,
            text="Buscar Archivo",
            command=self.sel_archivo,
            fg_color="gray",
        )

        self.btn_run = ctk.CTkButton(
            self.main_frame,
            text="EJECUTAR",
            command=self.run_bot_thread,
            height=50,
            fg_color="#2CC985",
            state="disabled",
        )
        self.btn_run.pack(pady=20)

        self.log_box = ctk.CTkTextbox(self.main_frame, width=600, height=250)
        self.log_box.pack(fill="both", expand=True)

    def reset_ui(self, titulo: str, info: str, necesita_archivo: bool = False) -> None:
        """Reinicia la interfaz para el bot seleccionado.

        Args:
            titulo: Título a mostrar en el panel central.
            info: Descripción del bot.
            necesita_archivo: Si el bot necesita un archivo de entrada.
        """
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

    # --- PANELES DE CADA BOT ---
    def panel_migo(self) -> None:
        self.bot_actual_class = SapMigoBotTurbo
        self.reset_ui("Carga MIGO", "Requiere Excel 'carga_migo.xlsx'.", True)

    def panel_pallet(self) -> None:
        self.bot_actual_class = SapBotPallet
        self.reset_ui(
            "Pallet Altura", "Selecciona el Excel donde pegar los datos.", False
        )

    def panel_transporte(self) -> None:
        self.bot_actual_class = SapBotTransporte
        self.reset_ui(
            "Reporte Transporte", "Extrae VT11 de ayer y envía correo.", False
        )

    def panel_vision(self) -> None:
        self.bot_actual_class = BotVisionPizarra
        self.reset_ui(
            "Visión IA", "Analiza la foto más reciente en la carpeta OneDrive.", False
        )

    def panel_auditor(self) -> None:
        self.bot_actual_class = SapBotAuditor
        self.reset_ui(
            "Auditor Stock", "Al ejecutar, te pedirá el Almacén.", False
        )

    # --- LÓGICA DE ACCIONES ---
    def sel_archivo(self) -> None:
        """Abre un diálogo para seleccionar un archivo y lo escribe en la entrada."""
        f = filedialog.askopenfilename()
        if f:
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, f)

    def log(self, msg: str) -> None:
        """Escribe un mensaje en el cuadro de logs."""
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")

    def run_bot_thread(self) -> None:
        """Prepara los argumentos y lanza el bot en un hilo separado."""
        ruta: str | None = None
        arg_extra: str | None = None

        # Determinar si el bot requiere archivo
        if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet]:
            ruta = self.entry_file.get()
            if not ruta:
                self.log("[ERROR] Selecciona un archivo Excel primero.")
                return
        elif self.bot_actual_class == SapBotAuditor:
            # Pedir almacén
            almacenes = [
                "SGVT",
                "CDNW",
                "SGTR",
                "TAVI",
                "SGSD",
                "AVAS",
                "SGBC",
                "SGVE",
                "SGEN",
            ]
            dialog = ctk.CTkInputDialog(
                text=f"Almacenes: {', '.join(almacenes)}\n\nEscribe el Almacén:",
                title="Auditor MM",
            )
            arg_extra = dialog.get_input()
            if not arg_extra:
                self.log("Cancelado.")
                return
            arg_extra = arg_extra.upper().strip()

        # Deshabilitar el botón de ejecución
        self.btn_run.configure(state="disabled", text="Ejecutando...")
        self.log(f"--- Iniciando {self.bot_actual_class.__name__}... ---")

        # Ejecutar en hilo para no congelar la UI
        threading.Thread(
            target=self.execute_bot, args=(ruta, arg_extra), daemon=True
        ).start()

    def execute_bot(self, ruta: str | None, arg_extra: str | None) -> None:
        """Instancia y ejecuta el bot seleccionado manejando errores.

        Args:
            ruta: Ruta del archivo si aplica.
            arg_extra: Argumento adicional (por ejemplo el almacén).
        """
        try:
            if not self.bot_actual_class:
                raise Exception("Bot no configurado.")
            bot_instance = self.bot_actual_class()

            # Bots que reciben ruta
            if self.bot_actual_class in [SapMigoBotTurbo, SapBotPallet]:
                bot_instance.run(ruta)
            elif self.bot_actual_class == SapBotAuditor:
                bot_instance.run(arg_extra)
            else:
                bot_instance.run()
            self.log("✅ PROCESO FINALIZADO CON ÉXITO")
        except Exception as e:
            self.log(f"❌ ERROR: {e}")
        finally:
            # Rehabilitar el botón
            self.btn_run.configure(state="normal", text="EJECUTAR")

if __name__ == "__main__":
    app = App()
    app.mainloop()

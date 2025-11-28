# 🎛️ SAP Logistics Automation Suite (GUI)

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat&logo=python&logoColor=white)
![GUI](https://img.shields.io/badge/Interface-CustomTkinter-blue?style=flat)
![SAP](https://img.shields.io/badge/Automation-SAP_GUI_Scripting-008FD3?style=flat)

> **Centro de Comando Unificado para la automatización de procesos logísticos en SAP.**

![Screenshot de la Interfaz](Logistics-Automation-Suite
/panel_preview.png)


---

## 📋 Descripción del Proyecto

Este proyecto es una aplicación de escritorio moderna desarrollada en **Python** utilizando **CustomTkinter**. Su objetivo es democratizar el uso de scripts de automatización (RPA) en el entorno operativo.

En lugar de ejecutar scripts de consola complejos, los usuarios (supervisores, administrativos y operarios) disponen de un **Panel de Control Centralizado** intuitivo y responsivo para ejecutar tareas críticas de SAP.

### 🎯 Problema que resuelve
Los scripts de automatización suelen ser difíciles de usar para el personal no técnico. Esta interfaz actúa como un **"Wrapper Gráfico"** que gestiona la ejecución, los errores y la configuración de los bots, cerrando la brecha entre el código y la operación diaria.

---

## 🚀 Módulos Integrados (Bots)

La suite orquesta los siguientes módulos, importados dinámicamente desde el paquete local `bots/`:

| Bot | Descripción | Tecnología Clave |
| :--- | :--- | :--- |
| **🧟 Auditor Zombie** | Detecta stock inmovilizado (>30 días) cruzando MB52 vs MB51. | Pandas Merge, Data Cleaning |
| **⚡ MIGO Turbo** | Carga masiva de movimientos interactuando con Excel en tiempo real. | PyWin32, COM Interop |
| **🗺️ Pallet Visual** | Genera mapas de pasillo (LX02) para auditoría física de altura. | Excel Automation, Pandas |
| **👁️ Visión Pizarra** | Digitaliza KPIs escritos a mano en pizarras de andén. | Google Gemini Vision API |
| **🚛 Transporte** | Reportabilidad automática de flotas (VT11/VT03N). | SAP Scripting |

---


## 🛠️ Arquitectura Técnica

El proyecto sigue una arquitectura modular para facilitar el mantenimiento y la escalabilidad:

```text
Logistics-Suite/
├── main.py              # Punto de entrada (Launcher)
├── gui_app.py           # Lógica de la interfaz (CustomTkinter)
├── requirements.txt     # Dependencias
├── assets/              # Iconos e imágenes
└── bots/                # Paquete de Lógica de Negocio
    ├── __init__.py
    ├── sap_bot_auditor.py
    ├── sap_bot_migo.py
    ├── sap_bot_pallet.py
    └── ...

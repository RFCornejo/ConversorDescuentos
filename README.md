# 📊 Conversor de Descuentos Masivo (SAP Z2D1)

Herramienta de escritorio desarrollada en **Python** con **CustomTkinter** para automatizar el procesamiento, cruce y generación de matrices masivas de descuentos listas para cargar en **SAP** bajo la clase de condición `Z2D1`.

---

## 🚀 Características

- **Cruce Automático O(1)**: Cruza catálogos de materiales y descuentos base con tablas de descuento de clientes (`T1`, `T2`, `T3`, `T4`, `T5`) y servicios (`R1`, `R2`, `R3`, `R4`).
- **Soporte Multi-Cliente**: Permite ingresar hasta 6 códigos de cliente (Columna L) en un solo proceso.
- **Modos de Organización**:
  - **Consolidado**: Une todos los clientes en la misma serie de archivos con numeración correlativa en `ITEM`.
  - **Separado por cliente**: Genera lotes independientes identificados con el código de cada cliente, reiniciando `ITEM` desde 1.
- **División en Lotes de 27,000 líneas**: Cumple con los límites de carga de SAP repitiendo encabezados en cada parte.
- **Exportación Flexible**: Generación en **CSV** (UTF-8 con BOM) o **Excel (.xlsx)** nativo con forzado de columnas en formato Texto (`@`).
- **Interfaz Moderna**: Interfaz gráfica en modo oscuro con CustomTkinter y procesamiento asíncrono en segundo plano (sin congelamientos).

---

## 📋 Requisitos Previos

- **Python 3.10** o superior instalado en el sistema.
- Gestor de paquetes `pip`.

---

## 🛠️ Instalación y Uso

1. **Clonar el repositorio**:
   ```bash
   git clone https://github.com/tu-usuario/ConversorDescuentos.git
   cd ConversorDescuentos
   ```

2. **Crear y activar un entorno virtual**:
   - En Windows:
     ```bash
     python -m venv venv
     venv\Scripts\activate
     ```
   - En Linux / macOS:
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Ejecutar la aplicación**:
   ```bash
   python app.py
   ```

---

## 📦 Generar el Ejecutable (.exe)

Si deseas compilar la aplicación en un único archivo ejecutable de Windows:

```bash
pyinstaller --noconsole --onefile --collect-all customtkinter --name "Conversor_Descuentos" app.py
```

El ejecutable generado se ubicará en la carpeta `dist/`.

---

## 📂 Estructura del Proyecto

```text
ConversorDescuentos/
├── app.py                     # Interfaz gráfica de usuario (CustomTkinter)
├── converter_service.py       # Lógica de procesamiento y exportación de datos
├── requirements.txt           # Librerías necesarias para el proyecto
├── Conversor_Descuentos.spec   # Configuración de compilación con PyInstaller
├── .gitignore                 # Reglas de exclusión para Git
└── README.md                  # Documentación del proyecto
```

---

## 📄 Licencia

Este proyecto es de uso interno / privado.

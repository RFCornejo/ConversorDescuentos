import os
import sys
import threading
from datetime import date
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

from converter_service import (
    cargar_datos_excel,
    convertir_datos,
    exportar_en_lotes
)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ConversorGUI(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Conversor Descuentos (SAP Z2D1)")
        self.geometry("860x780")
        self.minsize(840, 720)
        self.resizable(True, True)

        self.ruta_excel = None
        self.entries_clientes = []
        self.crear_interfaz()

    def crear_interfaz(self):
        # Título
        self.label_titulo = ctk.CTkLabel(
            self,
            text="📊 Matriz Descuentos Masivo",
            font=("Segoe UI", 22, "bold")
        )
        self.label_titulo.pack(pady=(12, 6))

        # Sección Selección de Archivo
        self.frame_archivo = ctk.CTkFrame(self)
        self.frame_archivo.pack(pady=4, padx=25, fill="x")

        self.btn_select = ctk.CTkButton(
            self.frame_archivo,
            text="📂 Seleccionar Excel",
            font=("Segoe UI", 13, "bold"),
            command=self.seleccionar_excel
        )
        self.btn_select.pack(pady=(8, 2))

        self.lbl_archivo = ctk.CTkLabel(
            self.frame_archivo,
            text="Ningún archivo seleccionado",
            font=("Segoe UI", 11, "italic"),
            text_color="gray"
        )
        self.lbl_archivo.pack(pady=(0, 8))

        # Sección Parámetros
        self.frame_params = ctk.CTkFrame(self)
        self.frame_params.pack(pady=6, padx=25, fill="x")

        # Fila 1: Parámetros de Datos (Fecha y Acceso)
        self.frame_fila1 = ctk.CTkFrame(self.frame_params, fg_color="transparent")
        self.frame_fila1.pack(pady=(8, 4), padx=15, fill="x")

        # Fecha válida desde
        self.frame_fecha = ctk.CTkFrame(self.frame_fila1, fg_color="transparent")
        self.frame_fecha.pack(side="left", expand=True, padx=10)

        self.lbl_fecha = ctk.CTkLabel(
            self.frame_fecha,
            text="📅 Válido desde:",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_fecha.pack(pady=(0, 2))

        fecha_default = date.today().replace(day=1).strftime("%d.%m.%Y")
        self.entry_fecha = ctk.CTkEntry(
            self.frame_fecha,
            width=140,
            justify="center",
            font=("Segoe UI", 12)
        )
        self.entry_fecha.insert(0, fecha_default)
        self.entry_fecha.pack()

        # Acceso (Columna C '# acces')
        self.frame_acceso = ctk.CTkFrame(self.frame_fila1, fg_color="transparent")
        self.frame_acceso.pack(side="left", expand=True, padx=10)

        self.lbl_acceso = ctk.CTkLabel(
            self.frame_acceso,
            text="🔢 Acceso (Col. C):",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_acceso.pack(pady=(0, 2))

        self.entry_acceso = ctk.CTkEntry(
            self.frame_acceso,
            width=120,
            justify="center",
            font=("Segoe UI", 12)
        )
        self.entry_acceso.insert(0, "35")
        self.entry_acceso.pack()

        # Sección Clientes (Columna L 'Cliente') - Hasta 6 clientes
        self.frame_clientes_section = ctk.CTkFrame(self.frame_params, fg_color="transparent")
        self.frame_clientes_section.pack(pady=(4, 4), padx=15, fill="x")

        self.frame_clientes_header = ctk.CTkFrame(self.frame_clientes_section, fg_color="transparent")
        self.frame_clientes_header.pack(fill="x", pady=(0, 4))

        self.lbl_clientes_titulo = ctk.CTkLabel(
            self.frame_clientes_header,
            text="👤 Clientes (Col. L) - hasta 6 clientes (dejar vacíos los que no se usen):",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_clientes_titulo.pack(side="left")

        self.btn_limpiar_clientes = ctk.CTkButton(
            self.frame_clientes_header,
            text="🧹 Limpiar",
            font=("Segoe UI", 11),
            width=65,
            height=22,
            fg_color="#4A5568",
            hover_color="#2D3748",
            command=self.limpiar_clientes
        )
        self.btn_limpiar_clientes.pack(side="right")

        # Grid de 6 entradas para clientes (2 filas x 3 columnas)
        self.frame_clientes_grid = ctk.CTkFrame(self.frame_clientes_section, fg_color="transparent")
        self.frame_clientes_grid.pack(fill="x", pady=(0, 2))

        self.entries_clientes = []
        for i in range(6):
            r = i // 3
            c = i % 3
            cell = ctk.CTkFrame(self.frame_clientes_grid, fg_color="transparent")
            cell.grid(row=r, column=c, padx=6, pady=2, sticky="ew")
            self.frame_clientes_grid.grid_columnconfigure(c, weight=1)

            lbl_cli = ctk.CTkLabel(
                cell,
                text=f"Cliente {i+1}:",
                font=("Segoe UI", 11, "bold"),
                width=62,
                anchor="e"
            )
            lbl_cli.pack(side="left", padx=(0, 4))

            ent_cli = ctk.CTkEntry(
                cell,
                width=135,
                justify="center",
                font=("Segoe UI", 12),
                placeholder_text="(Vacío)"
            )
            ent_cli.pack(side="left", fill="x", expand=True)
            self.entries_clientes.append(ent_cli)

        # Modo de salida para clientes
        self.frame_modo_clientes = ctk.CTkFrame(self.frame_params, fg_color="transparent")
        self.frame_modo_clientes.pack(pady=(4, 2), padx=15, fill="x")

        self.lbl_modo_clientes = ctk.CTkLabel(
            self.frame_modo_clientes,
            text="📑 Organización para múltiples clientes:",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_modo_clientes.pack(side="left", padx=(0, 8))

        self.var_modo_clientes = ctk.StringVar(value="Consolidado")
        self.seg_modo_clientes = ctk.CTkSegmentedButton(
            self.frame_modo_clientes,
            values=["Consolidado", "Separado por cliente"],
            variable=self.var_modo_clientes,
            font=("Segoe UI", 11, "bold")
        )
        self.seg_modo_clientes.pack(side="left")

        self.lbl_modo_info = ctk.CTkLabel(
            self.frame_params,
            text="• Consolidado: Todos los clientes en los mismos archivos (lotes de 27,000 líneas)\n• Separado: Archivos independientes para cada cliente (ITEM reinicia en 1)",
            font=("Segoe UI", 10, "italic"),
            text_color="#88B4E7"
        )
        self.lbl_modo_info.pack(pady=(0, 4))

        # Fila Carpeta de salida
        self.frame_fila2 = ctk.CTkFrame(self.frame_params, fg_color="transparent")
        self.frame_fila2.pack(pady=(4, 4), padx=15, fill="x")

        self.lbl_carpeta = ctk.CTkLabel(
            self.frame_fila2,
            text="📁 Carpeta de salida:",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_carpeta.pack(side="left", padx=(0, 8))

        nombre_carpeta_default = f"Descuentos_Resultado_{date.today().strftime('%Y%m%d')}"
        self.entry_carpeta = ctk.CTkEntry(
            self.frame_fila2,
            font=("Segoe UI", 12)
        )
        self.entry_carpeta.insert(0, nombre_carpeta_default)
        self.entry_carpeta.pack(side="left", expand=True, fill="x")

        # Fila Formato de Salida
        self.lbl_formato = ctk.CTkLabel(
            self.frame_params,
            text="⚙️ Formato de salida (Columnas en formato Texto):",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_formato.pack(pady=(4, 2))

        self.var_formato = ctk.StringVar(value="CSV (.csv)")
        self.seg_formato = ctk.CTkSegmentedButton(
            self.frame_params,
            values=["CSV (.csv)", "Excel (.xlsx)"],
            variable=self.var_formato,
            font=("Segoe UI", 11, "bold")
        )
        self.seg_formato.pack(pady=(0, 2))

        self.lbl_formato_info = ctk.CTkLabel(
            self.frame_params,
            text="• CSV: Separado por comas (,) con UTF-8 | • Excel: Libro nativo con columnas de Texto ('@')\n• Divididos en lotes de hasta 27,000 líneas con títulos repetidos",
            font=("Segoe UI", 10, "italic"),
            text_color="#88B4E7"
        )
        self.lbl_formato_info.pack(pady=(0, 4))

        # Fila Categorías de cliente
        self.lbl_grupos = ctk.CTkLabel(
            self.frame_params,
            text="👥 Categorías de cliente a procesar:",
            font=("Segoe UI", 12, "bold")
        )
        self.lbl_grupos.pack(pady=(2, 2))

        self.frame_chk = ctk.CTkFrame(self.frame_params, fg_color="transparent")
        self.frame_chk.pack(pady=(0, 6))

        self.grupos_vars = {}
        for grupo in ['T1', 'T2', 'T3', 'T4', 'T5']:
            var = ctk.BooleanVar(value=True)
            chk = ctk.CTkCheckBox(
                self.frame_chk,
                text=grupo,
                variable=var,
                font=("Segoe UI", 12)
            )
            chk.pack(side="left", padx=10)
            self.grupos_vars[grupo] = var

        # Botón de Acción Principal
        self.btn_convertir = ctk.CTkButton(
            self,
            text="▶️ Convertir y Guardar en Carpeta",
            font=("Segoe UI", 14, "bold"),
            height=40,
            command=self.iniciar_conversion
        )
        self.btn_convertir.pack(pady=(8, 4))

        # Mensaje de Estado
        self.estado = ctk.CTkLabel(
            self,
            text="Listo para procesar",
            font=("Segoe UI", 12)
        )
        self.estado.pack(pady=2)

    def limpiar_clientes(self):
        for ent in self.entries_clientes:
            ent.delete(0, "end")

    def obtener_clientes(self):
        clientes = []
        for ent in self.entries_clientes:
            val = ent.get().strip()
            if val and val not in clientes:
                clientes.append(val)
        return clientes

    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(
            parent=self,
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if ruta:
            self.ruta_excel = ruta
            nombre = Path(ruta).name
            stem = Path(ruta).stem
            self.lbl_archivo.configure(
                text=f"Archivo cargado: {nombre}",
                text_color="#3B8ED0"
            )
            carpeta_sugerida = f"Descuentos_{stem}"
            self.entry_carpeta.delete(0, "end")
            self.entry_carpeta.insert(0, carpeta_sugerida)
            self.estado.configure(text="Archivo seleccionado correctamente")

    def obtener_grupos_seleccionados(self):
        return [
            grupo for grupo, var in self.grupos_vars.items()
            if var.get()
        ]

    def iniciar_conversion(self):
        if not self.ruta_excel:
            messagebox.showwarning(
                "Archivo requerido",
                "Por favor selecciona primero un archivo Excel.",
                parent=self
            )
            return

        grupos = self.obtener_grupos_seleccionados()
        if not grupos:
            messagebox.showwarning(
                "Categorías requeridas",
                "Debes seleccionar al menos una categoría de cliente (T1, T2, etc.).",
                parent=self
            )
            return

        fecha_valida = self.entry_fecha.get().strip()
        if not fecha_valida:
            messagebox.showwarning(
                "Fecha requerida",
                "Ingresa una fecha de inicio de validez.",
                parent=self
            )
            return

        acceso = self.entry_acceso.get().strip() or "35"
        clientes = self.obtener_clientes()
        modo_clientes = self.var_modo_clientes.get()

        nombre_carpeta = self.entry_carpeta.get().strip()
        if not nombre_carpeta:
            messagebox.showwarning(
                "Nombre requerido",
                "Ingresa un nombre para la carpeta de salida.",
                parent=self
            )
            return

        ubicacion_base = filedialog.askdirectory(
            parent=self,
            title="Selecciona la ubicación donde crear la carpeta de resultados"
        )

        if not ubicacion_base:
            return

        carpeta_destino = os.path.join(ubicacion_base, nombre_carpeta)
        formato_elegido = "xlsx" if "xlsx" in self.var_formato.get().lower() else "csv"

        self.estado.configure(text="⏳ Procesando datos del archivo Excel...")
        self.btn_convertir.configure(state="disabled")
        self.btn_select.configure(state="disabled")

        threading.Thread(
            target=self._procesar_en_background,
            args=(self.ruta_excel, grupos, fecha_valida, acceso, clientes, modo_clientes, carpeta_destino, nombre_carpeta, formato_elegido),
            daemon=True
        ).start()

    def _procesar_en_background(self, ruta_excel, grupos, fecha_valida, acceso, clientes, modo_clientes, carpeta_destino, prefijo, formato):
        try:
            self.after(0, lambda: self.estado.configure(
                text="⏳ Cargando datos del archivo Excel..."
            ))
            datos, tablas = cargar_datos_excel(ruta_excel, grupos)

            # Lista de clientes (si no se indicó ninguno, se procesa 1 vez con cliente vacío)
            lista_clientes = clientes if clientes else ['']
            total_clientes = len(lista_clientes)
            es_separado = "separado" in modo_clientes.lower()

            archivos_totales = []
            total_registros = 0

            if es_separado:
                for idx, cli in enumerate(lista_clientes, 1):
                    cli_tag = f"'{cli}'" if cli else "(Vacío)"
                    self.after(0, lambda c=cli_tag, i=idx: self.estado.configure(
                        text=f"⏳ Procesando cliente {i}/{total_clientes}: {c}..."
                    ))

                    resultado = convertir_datos(
                        datos,
                        tablas,
                        valido_desde=fecha_valida,
                        num_acceso=acceso,
                        cliente=cli
                    )

                    if resultado.empty:
                        continue

                    total_registros += len(resultado)

                    # Sanitizar nombre del cliente para nombre de archivo en Windows
                    cli_sano = "".join(ch for ch in cli if ch.isalnum() or ch in ('-', '_')).strip()
                    prefijo_cli = f"{prefijo}_Cli_{cli_sano}" if cli_sano else prefijo

                    self.after(0, lambda c=cli_tag, n=len(resultado): self.estado.configure(
                        text=f"⏳ Exportando {n:,} registros para {c}..."
                    ))

                    archivos = exportar_en_lotes(
                        resultado,
                        carpeta_destino,
                        prefijo_archivo=prefijo_cli,
                        formato=formato,
                        limite_filas=27000
                    )
                    archivos_totales.extend(archivos)

            else:
                # Modo Consolidado
                self.after(0, lambda: self.estado.configure(
                    text=f"⏳ Procesando {total_clientes} cliente(s) en modo consolidado..."
                ))

                resultado = convertir_datos(
                    datos,
                    tablas,
                    valido_desde=fecha_valida,
                    num_acceso=acceso,
                    clientes=lista_clientes
                )

                if not resultado.empty:
                    total_registros = len(resultado)
                    self.after(0, lambda: self.estado.configure(
                        text=f"⏳ Exportando {total_registros:,} registros en lotes de 27,000..."
                    ))

                    archivos = exportar_en_lotes(
                        resultado,
                        carpeta_destino,
                        prefijo_archivo=prefijo,
                        formato=formato,
                        limite_filas=27000
                    )
                    archivos_totales.extend(archivos)

            if total_registros == 0:
                self.after(0, lambda: self._mostrar_resultado(
                    "Advertencia",
                    "El proceso finalizó pero no se encontraron coincidencias entre los materiales y las tablas de descuento seleccionadas.",
                    es_error=False,
                    es_warning=True
                ))
                return

            total_archivos = len(archivos_totales)
            formato_str = "Excel (.xlsx)" if formato == "xlsx" else "CSV (.csv)"
            clientes_str = ", ".join([f"'{c}'" if c else "(Vacío)" for c in lista_clientes])
            modo_str = "Separado por cliente" if es_separado else "Consolidado"

            msg = (
                f"¡Conversión exitosa!\n\n"
                f"• Total de registros generados: {total_registros:,}\n"
                f"• Clientes procesados ({total_clientes}): {clientes_str}\n"
                f"• Organización: {modo_str}\n"
                f"• Acceso (# acces): {acceso}\n"
                f"• Formato: {formato_str} (Columnas en formato Texto)\n"
                f"• Archivos creados: {total_archivos} (lotes de hasta 27,000 líneas)\n\n"
                f"Carpeta creada en:\n{carpeta_destino}"
            )
            self.after(0, lambda: self._mostrar_resultado("Éxito", msg, es_error=False))

        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: self._mostrar_resultado("Error", err_msg, es_error=True))

        finally:
            self.after(0, self._restablecer_controles)

    def _mostrar_resultado(self, titulo, mensaje, es_error=False, es_warning=False):
        if es_error:
            messagebox.showerror(titulo, mensaje, parent=self)
            self.estado.configure(text="❌ Ocurrió un error en la conversión")
        elif es_warning:
            messagebox.showwarning(titulo, mensaje, parent=self)
            self.estado.configure(text="⚠️ Finalizado con observaciones (0 filas)")
        else:
            messagebox.showinfo(titulo, mensaje, parent=self)
            self.estado.configure(text="✔️ Archivos generados correctamente en lotes")

    def _restablecer_controles(self):
        self.btn_convertir.configure(state="normal")
        self.btn_select.configure(state="normal")


if __name__ == "__main__":
    app = ConversorGUI()
    app.mainloop()

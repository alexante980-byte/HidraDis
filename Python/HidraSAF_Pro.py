import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import numpy as np
import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ------------------------------------------------------------------------------
# EXCEPCIONES PERSONALIZADAS
# ------------------------------------------------------------------------------
class SAFValidationError(Exception):
    """Excepción para detener el cálculo si se violan los límites hidráulicos del cuenco SAF."""
    pass

# ------------------------------------------------------------------------------
# MOTOR HIDRÁULICO 
# ------------------------------------------------------------------------------
def resolver_ruta(ruta_relativa):
    """Resuelve rutas de recursos tanto en modo .py como en .exe (PyInstaller)."""
    try:
        base_path = sys._MEIPASS          # carpeta temporal del .exe
    except AttributeError:
        base_path = os.path.abspath(".")  # carpeta del script en desarrollo
    return os.path.join(base_path, ruta_relativa)

def calcular_presion_vapor(temperatura_c: float) -> float:
    """Calcula la presión de vapor del agua (Hv) en m.c.a. usando Antoine modificada."""
    p_mmhg = 10**(8.07131 - (1730.63 / (temperatura_c + 233.426)))
    return 0.0136 * p_mmhg

def calcular_presion_atmosferica(altitud_msnm: float) -> float:
    """Calcula la presión atmosférica (H_atm) en m.c.a. según la altitud local."""
    return 10.33 - (altitud_msnm / 900.0)

# ------------------------------------------------------------------------------
# CONFIGURACIÓN DE ESTILO V VISUAL
# ------------------------------------------------------------------------------
ctk.set_appearance_mode("light")
COLOR_FONDO = "#D6E4F0"
COLOR_CONTENEDOR = "#FFFFFF"
COLOR_ACCENTO = "#1A374D"
COLOR_RESULTADO = "#d32f2f"
COLOR_BORDE_TECNICO = "#B1D0E0"

# ------------------------------------------------------------------------------
# INTERFAZ DE USUARIO PRINCIPAL
# ------------------------------------------------------------------------------
class SAF_designer_Final(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("SAF Designer Pro - Hidra_SAF_Pro")
        self.geometry("1420x920")
        self.configure(fg_color=COLOR_FONDO)

        # Rutas de Recursos
        self.ruta_epn       = resolver_ruta(os.path.join("Imagenes", "Escuela_Politécnica_Nacional.png"))
        self.ruta_fica      = resolver_ruta(os.path.join("Imagenes", "logo_fica.png"))
        self.ruta_planta    = resolver_ruta(os.path.join("Imagenes", "Vista en planta SAF.png"))
        self.ruta_elevacion = resolver_ruta(os.path.join("Imagenes", "Vista en elevación_SAF.png"))
        self.ruta_isometrica = resolver_ruta(os.path.join("Imagenes", "Vista isométrica_SAF.png"))

        self._img_refs = []
        self.inputs = {}
        self.input_frames = {}
        self.results = {}
        self.status_labels = {}
        self.data_export = {}

        self.init_ui()

    def init_ui(self):
        # 1. ENCABEZADO
        header = ctk.CTkFrame(self, fg_color=COLOR_FONDO, corner_radius=0)
        header.pack(fill="x", padx=0, pady=0)
        content_head = ctk.CTkFrame(header, fg_color="transparent")
        content_head.pack(pady=5)

        self.load_img(content_head, self.ruta_epn, (50, 50), "left")
        titles_frame = ctk.CTkFrame(content_head, fg_color="transparent")
        titles_frame.pack(side="left", padx=15)
        ctk.CTkLabel(titles_frame, text="DISEÑO DE CUENCO DISIPADOR TIPO SAF - HidraSAF_Pro", font=("Segoe UI", 18, "bold"), text_color=COLOR_ACCENTO).pack()
        ctk.CTkLabel(titles_frame, text="Este programa permite el cálculo y dimensionamiento de la geometría de cuencos disipadores de energía tipo SAF, diseñados para garantizar un desempeño hidráulico óptimo en obras de derivación y descargas de colectores.", font=("Segoe UI", 11, "italic"), text_color="#406882").pack()
        self.load_img(content_head, self.ruta_fica, (120, 38), "right")

        # 2. CUERPO DE LA INTERFAZ
        body = ctk.CTkFrame(self, fg_color=COLOR_FONDO)
        body.pack(fill="both", expand=True, padx=10, pady=2)

        # Tipo de Proyecto
        type_frame = ctk.CTkFrame(body, fg_color=COLOR_FONDO, corner_radius=8, border_width=1, border_color=COLOR_BORDE_TECNICO)
        type_frame.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(type_frame, text="Tipo de Proyecto:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=10, pady=3)
        self.tipo_var = ctk.StringVar(value="Colector")
        ctk.CTkRadioButton(type_frame, text="Azud (Cuenco ancho constante)", variable=self.tipo_var, value="Derivacion", command=self.toggle_inputs, font=("Segoe UI", 10)).pack(side="left", padx=10)
        ctk.CTkRadioButton(type_frame, text="Colector (Cuenco ancho variable)", variable=self.tipo_var, value="Colector", command=self.toggle_inputs, font=("Segoe UI", 10)).pack(side="left", padx=10)

        panels = ctk.CTkFrame(body, fg_color="transparent")
        panels.pack(fill="both", expand=True)

        # Panel Izquierdo
        left_side = ctk.CTkFrame(panels, fg_color=COLOR_FONDO, width=680)
        left_side.pack(side="left", fill="both", expand=False, padx=(0, 5))
        left_side.pack_propagate(False)
        self.left_scroll = ctk.CTkScrollableFrame(left_side, fg_color=COLOR_FONDO, label_text="")
        self.left_scroll.pack(fill="both", expand=True)

        # Panel Derecho
        right_side = ctk.CTkFrame(panels, fg_color=COLOR_FONDO, width=710)
        right_side.pack(side="right", fill="both", expand=True)
        right_side.pack_propagate(False)
        self.right_scroll = ctk.CTkScrollableFrame(right_side, fg_color=COLOR_FONDO)
        self.right_scroll.pack(fill="both", expand=True)

        self.setup_sections()
        self.setup_graphics()

        # 3. PIE DE PÁGINA
        footer = ctk.CTkFrame(self, fg_color=COLOR_FONDO, height=85, corner_radius=0, border_width=1, border_color=COLOR_BORDE_TECNICO)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        info_tic = ctk.CTkFrame(footer, fg_color="transparent")
        info_tic.pack(side="left", padx=20, pady=(6, 12))
        ctk.CTkLabel(info_tic, text="Facultad de Ingeniería Civil y Ambiental - EPN", font=("Segoe UI", 10, "bold"), text_color=COLOR_ACCENTO).pack(anchor="w", pady=(0, 1))
        ctk.CTkLabel(info_tic, text="Trabajo de Integración Curricular", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 1))
        ctk.CTkLabel(info_tic, text=f"Autor: Alex Ante | {datetime.now().strftime('%d/%m/%Y')}", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        btn_frame = ctk.CTkFrame(footer, fg_color="transparent")
        btn_frame.pack(side="right", padx=20, pady=15)
        ctk.CTkButton(btn_frame, text="Calcular", command=self.calcular, fg_color=COLOR_ACCENTO, width=90, height=30, font=("Segoe UI", 11, "bold")).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Limpiar", command=self.limpiar, fg_color="#698396", width=80, height=30, font=("Segoe UI", 11)).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="Exportar Excel", command=self.exportar_excel, fg_color="#2E7D32", width=120, height=30, font=("Segoe UI", 11, "bold")).pack(side="left", padx=4)

    def load_img(self, parent, path, size, side):
        if os.path.exists(path):
            img = ctk.CTkImage(light_image=Image.open(path), size=size)
            self._img_refs.append(img)
            ctk.CTkLabel(parent, image=img, text="").pack(side=side, padx=5)

    def toggle_inputs(self):
        if self.tipo_var.get() == "Derivacion":
            self.input_frames["D0_frame"].grid_forget()
            self.input_frames["Br_frame"].grid(row=0, column=1, padx=6, pady=4)
        else:
            self.input_frames["Br_frame"].grid_forget()
            self.input_frames["D0_frame"].grid(row=0, column=1, padx=6, pady=4)

    def setup_sections(self):
        ctk.CTkLabel(self.left_scroll, text="Datos:", font=("Segoe UI", 13, "bold"), text_color=COLOR_ACCENTO, anchor="w").pack(fill="x", padx=5, pady=(2,0))
        in_f = ctk.CTkFrame(self.left_scroll, fg_color=COLOR_FONDO, corner_radius=6, border_width=1, border_color=COLOR_BORDE_TECNICO)
        in_f.pack(fill="x", pady=(1, 5))

        # Grid para ordenar los datos de entrada
        grid = ctk.CTkFrame(in_f, fg_color="transparent")
        grid.pack(padx=5, pady=4, anchor="w")

        self.input_frames["Q_frame"] = self.create_input(grid, "Caudal (Q):", "m³/s", "Q", row=0, col=0)
        self.input_frames["Br_frame"] = self.create_input(grid, "Ancho azud (Bᵣ):", "m", "Br", row=0, col=1)
        self.input_frames["D0_frame"] = self.create_input(grid, "Diámetro colector (D₀):", "m", "D0", row=0, col=1)
        self.input_frames["y1_frame"] = self.create_input(grid, "Tirante supercrítico (y₁):", "m", "y1", row=1, col=0)
        self.input_frames["Temp_frame"] = self.create_input(grid, "Temperatura (T):", "°C", "Temp", row=1, col=1)
        self.input_frames["Alt_frame"] = self.create_input(grid, "Altitud (A):", "msnm", "Alt", row=2, col=0)

        self.toggle_inputs()

        ctk.CTkLabel(self.left_scroll, text="Resultados:", font=("Segoe UI", 13, "bold"), text_color=COLOR_ACCENTO, anchor="w").pack(fill="x", padx=5, pady=(2,0))
        res_f = ctk.CTkFrame(self.left_scroll, fg_color=COLOR_FONDO, corner_radius=6, border_width=1, border_color=COLOR_BORDE_TECNICO)
        res_f.pack(fill="both", expand=True, pady=(1, 5))

        sections = [
            ("1. Parámetros de Cuenco", [("Ancho Cuenco", "Wb", "Wb", "m"), ("Ancho Rápida", "Wb₁", "Wb1", "m"), ("Ancho Intermedio", "Wb₂", "Wb2", "m")]),
            ("2. Parámetros Hidráulicos Reales", [("Número Froude", "Fr₁", "Fr1_ap", "-"), ("Velocidad Flujo", "V₁", "V1", "m/s"), ("Tirante Crítico", "y꜀", "yc", "m")]),
            ("3. Condiciones del cuenco SAF", [("Tirante Conjugado", "y₂", "d2", "m"), ("Factor Corrección", "C", "C", "-")]),
            ("4. Bloques de Rápida", [("Número Bloques", "n꜀♭", "ncb", "u"), ("Ancho", "w꜀♭", "wcb", "m"), ("Altura", "h꜀♭", "hcb", "m")]),
            ("5. Bloques de Fondo", [("Número Bloques", "nբ♭", "nfb", "u"), ("Ancho", "wբ♭", "wfb", "m"), ("Altura", "hբ♭", "hfb", "m"), ("Espesor Cresta", "e", "e", "m"), ("Distancia Pared", "dᵣₑₐₗ", "dreal", "m")]),
            ("6. Aproximación Geométrica", [("Transición Ent.", "Lₐ", "LA", "m"), ("Longitud Cuenco", "LB", "LB", "m"), ("Ubicación Bloques", "Dₑₛ", "Des", "m")]),
            ("7. Estructura Final", [("Altura Umbral", "h₄", "h4", "m"), ("Altura Pared ", "H₆", "h6", "m"), ("Borde Libre", "z", "z_bl", "m"), ("Ancho Salida", "Wb₃", "Wb3", "m")]),
            ("8. Cavitación (Ingreso de la rápida)", [("Índice de cavitación", "σ", "sigma", "-")])
        ]

        for title, items in sections:
            self.create_result_group(res_f, title, items)

    def create_input(self, parent, label, unit, key, row, col):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=row, column=col, padx=8, pady=4, sticky="w")
        ctk.CTkLabel(f, text=label, width=125, font=("Segoe UI", 10), anchor="w").pack(side="left")
        e = ctk.CTkEntry(f, width=65, height=22, fg_color=COLOR_CONTENEDOR)
        e.pack(side="left", padx=1)
        ctk.CTkLabel(f, text=unit, font=("Segoe UI", 9), text_color="gray", width=30, anchor="w").pack(side="left", padx=2)
        self.inputs[key] = e
        return f

    def create_result_group(self, parent, title, items):
        ctk.CTkLabel(parent, text=title, font=("Segoe UI", 11, "bold"), text_color="#406882", anchor="w").pack(fill="x", padx=10, pady=(4, 1))
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="x", padx=8, pady=1)

        for i, item in enumerate(items):
            r, c = divmod(i, 3)
            cell = ctk.CTkFrame(container, fg_color="transparent")
            cell.grid(row=r, column=c, padx=3, pady=1, sticky="w")

            ctk.CTkLabel(cell, text=f"{item[0]} [{item[1]}]:", width=95, anchor="w", font=("Segoe UI", 9.5)).pack(side="left")
            res_box = ctk.CTkLabel(cell, text="0.00", width=60, height=18, fg_color=COLOR_CONTENEDOR, corner_radius=3, font=("Segoe UI", 9.5, "bold"), text_color=COLOR_RESULTADO)
            res_box.pack(side="left", padx=1)
            self.results[item[2]] = res_box
            ctk.CTkLabel(cell, text=item[3], width=15, text_color="gray", font=("Segoe UI", 9), anchor="w").pack(side="left")

        if "SAF" in title:
            status_fFr1 = ctk.CTkFrame(container, fg_color="transparent")
            status_fFr1.grid(row=2, column=0, columnspan=3, pady=1, sticky="w")
            ctk.CTkLabel(status_fFr1, text="Verificación (1.7 < Fr₁ < 17):", font=("Segoe UI", 9)).pack(side="left", padx=(2,0))
            self.status_labels["Fr1"] = ctk.CTkLabel(status_fFr1, text="--", font=("Segoe UI", 9, "bold"), text_color="gray")
            self.status_labels["Fr1"].pack(side="left", padx=5)
        elif "Fondo" in title:
            status_fOc = ctk.CTkFrame(container, fg_color="transparent")
            status_fOc.grid(row=3, column=0, columnspan=3, sticky="w", pady=1)
            ctk.CTkLabel(status_fOc, text="Ocupación (40-55% Bᵣ):", font=("Segoe UI", 9)).pack(side="left", padx=(2,0))
            self.results["Oc"] = ctk.CTkLabel(status_fOc, text="0.0%", width=45, font=("Segoe UI", 9, "bold"), text_color=COLOR_RESULTADO)
            self.results["Oc"].pack(side="left", padx=3)
            self.status_labels["Oc"] = ctk.CTkLabel(status_fOc, text="--", font=("Segoe UI", 9, "bold"), text_color="gray")
            self.status_labels["Oc"].pack(side="left", padx=3)
        elif "Cavitación" in title:
            status_cav = ctk.CTkFrame(container, fg_color="transparent")
            status_cav.grid(row=1, column=0, columnspan=3, sticky="w", pady=1)
            ctk.CTkLabel(status_cav, text="Diagnóstico Cavitación:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(2,0))
            self.status_labels["Cav"] = ctk.CTkLabel(status_cav, text="Evaluar diseño", font=("Segoe UI", 9, "bold"), text_color="gray")
            self.status_labels["Cav"].pack(side="left", padx=5)

    def setup_graphics(self):
        ctk.CTkLabel(self.right_scroll, text="Esquema:", font=("Segoe UI", 13, "bold"), text_color=COLOR_ACCENTO, anchor="w").pack(fill="x", padx=5, pady=(2, 2))

        formatos = {
            "Vista en Planta": (500, 300),
            "Vista en Elevación": (500, 300),
            "Vista Isométrica": (500, 300)
        }

        for title, path in [
            ("Vista en Planta",    self.ruta_planta),
            ("Vista en Elevación", self.ruta_elevacion),
            ("Vista Isométrica",   self.ruta_isometrica)
        ]:
            f = ctk.CTkFrame(self.right_scroll, fg_color=COLOR_FONDO, corner_radius=6, border_width=1, border_color=COLOR_BORDE_TECNICO)
            f.pack(fill="x", pady=3, padx=2)
            ctk.CTkLabel(f, text=title, font=("Segoe UI", 10, "bold"), text_color=COLOR_ACCENTO).pack(pady=1)

            if os.path.exists(path):
                w, h = formatos[title]
                img = ctk.CTkImage(light_image=Image.open(path), size=(w, h))
                self._img_refs.append(img)
                ctk.CTkLabel(f, image=img, text="").pack(pady=2, padx=10)
            else:
                ctk.CTkLabel(f, text=f"[{title} — archivo no encontrado en recursos]",
                             font=("Segoe UI", 9, "italic"), text_color="gray").pack(pady=20)

            ctk.CTkLabel(f, text="Adaptado Chow (1994)", font=("Segoe UI", 8, "italic"), text_color="#555555").pack(pady=(0, 2), anchor="e", padx=10)

    def calcular(self):
        try:
            g = 9.81
            mode = self.tipo_var.get()

            # Obtención de datos de entrada desde la UI
            Q = float(self.inputs["Q"].get().replace(',', '.'))
            y1 = float(self.inputs["y1"].get().replace(',', '.'))
            temp_agua = float(self.inputs["Temp"].get().replace(',', '.'))
            altitud = float(self.inputs["Alt"].get().replace(',', '.')) # CORRECCIÓN: Captura de Altitud

            if mode == "Derivacion":
                Br = float(self.inputs["Br"].get().replace(',', '.'))
                # Uso del flujo supercrítico con y1
                V1 = Q / (Br * y1)
                Wb1 = Br
                Wb = Br
                LA = 0
            else:
                D0 = float(self.inputs["D0"].get().replace(',', '.'))
                # Pre-dimensionamiento empírico de la solera de entrada
                Wb1 = D0 if (Q / (D0 * y1)) < 6 else 2.5 * D0 if (Q / (D0 * y1)) < 12 else 3.0 * D0
                # Velocidad calculada en el flujo supercrítico usando y1
                V1 = Q / (Wb1 * y1)
                Wb = max(D0, 1.70 * Q / (np.sqrt(g) * D0**1.5))
                LA = (Wb1 - D0) / 2

            Fr1 = V1 / np.sqrt(g * y1)

            # CONTROL DE EXCEPCIONES PERSONALIZADO
            if Fr1 < 1.7 or Fr1 > 17.0:
                raise SAFValidationError(
                    f"El número de Froude calculado (Fr₁ = {Fr1:.2f}) se encuentra fuera del rango empírico de diseño "
                    f"estipulado para un disipador SAF (Límites: 1.7 ≤ Fr₁ ≤ 17).\n\n"
                    f"Modifique el caudal, la sección o el tirante inicial para proceder con el dimensionamiento geométrico."
                )

            yc = ((Q / Wb)**2 / g)**(1/3)
            d2 = (y1 / 2) * (np.sqrt(1 + 8 * Fr1**2) - 1)

            if 1.7 < Fr1 <= 5.5:
                C_val = 1.1 - (Fr1**2 / 120)
            elif 5.5 < Fr1 <= 11:
                C_val = 0.85
            else:
                C_val = 1.0 - (Fr1**2 / 800)

            LB = (4.5 * d2) / (C_val * (Fr1**0.76))
            Wb2 = Wb1 + (Wb - Wb1) / 3 if mode == "Colector" else Wb

            ncb = int(Wb1 / (1.5 * y1)) if (1.5 * y1) > 0 else 0
            wcb = Wb1 / (2 * ncb) if ncb > 0 else 0
            nfb = ncb
            wfb = Wb2 / (2 * nfb) if nfb > 0 else 0

            # Cálculo dinámico de Hatm dependiente de la altitud
            Hatm = calcular_presion_atmosferica(altitud)
            Hv = calcular_presion_vapor(temp_agua)
            carga_vel = (V1**2) / (2 * g)
            sigma = (Hatm - Hv - y1) / carga_vel if carga_vel > 0 else 0

            if sigma > 1.0:
                estado_cav = "NO HAY CAVITACIÓN ✔ (Condición segura)"
                color_cav = "green"
            elif 0.4 <= sigma <= 1.0:
                estado_cav = "RIESGO MODERADO ⚠ (Revisar diseño)"
                color_cav = "orange"
            else:
                estado_cav = "ALTO RIESGO ✘ (Rediseñar estructura)"
                color_cav = "red"

            res_map = {
                "Wb": Wb, "Wb1": Wb1, "Wb2": Wb2, "yc": yc, "V1": V1, "Fr1_ap": Fr1, "Fr1": Fr1, "d2": d2, "C": C_val, "LA": LA, "LB": LB, "Des": LB/3,
                "ncb": ncb, "wcb": wcb, "acb": wcb, "hcb": y1,
                "nfb": nfb, "wfb": wfb, "afb": wfb, "hfb": y1,
                "e": 0.5 * y1, "dreal": (Wb2 - (nfb * wfb) - (nfb - 1) * wfb) / 2 if nfb > 0 else 0,
                "h4": max(1.0, 0.07 * (d2 / C_val)), "h6": d2 * (1 + (1 / (3 * C_val))), "z_bl": d2 / 3, "Wb3": Wb + 2 * LB,
                "sigma": sigma
            }

            for k, v in res_map.items():
                if k in self.results:
                    self.results[k].configure(text=f"{v:.4f}" if k == "sigma" else f"{v:.3f}" if k in ["Fr1_ap", "yc", "V1", "C"] else f"{v:.2f}")

            ocup = (nfb * wfb / Wb2) * 100 if nfb > 0 else 0
            self.results["Oc"].configure(text=f"{ocup:.1f}%")
            self.status_labels["Fr1"].configure(text="CUMPLE", text_color="green")
            self.status_labels["Oc"].configure(text="CUMPLE" if 40 <= ocup <= 55 else "REVISAR", text_color="green" if 40 <= ocup <= 55 else "red")
            self.status_labels["Cav"].configure(text=estado_cav, text_color=color_cav)

            self.data_export = {
                "inputs": {"Tipo Proyecto": mode, "Q": Q, "y1": y1, "Temp": temp_agua, "Alt": altitud, "Br": Br if mode == "Derivacion" else 0.0, "D0": D0 if mode == "Colector" else 0.0},
                "results": res_map,
                "cav_text": estado_cav
            }

        except SAFValidationError as val_err:
            self.limpiar()
            messagebox.showwarning("Límite de Diseño Excedido", str(val_err))

        except Exception as err:
            messagebox.showerror("Error de Cálculo", f"Verifique que los datos de entrada sean numéricos válidos.\nDetalle: {err}")

    def limpiar(self):
        for r in self.results.values():
            r.configure(text="0.00")
        self.status_labels["Fr1"].configure(text="--", text_color="gray")
        self.status_labels["Oc"].configure(text="--", text_color="gray")
        self.status_labels["Cav"].configure(text="Evaluar diseño", text_color="gray")
        self.results["Oc"].configure(text="0.0%")
        self.data_export = {}

    def exportar_excel(self):
        if not self.data_export:
            messagebox.showwarning("Atención", "Por favor, realice el cálculo antes de exportar.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
        if not path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Memoria SAF"
            ws.views.sheetView[0].showGridLines = True

            fill_header = PatternFill(start_color="1A374D", end_color="1A374D", fill_type="solid")
            fill_section = PatternFill(start_color="406882", end_color="406882", fill_type="solid")
            fill_zebra = PatternFill(start_color="F9FBFC", end_color="F9FBFC", fill_type="solid")

            font_title = Font(name="Segoe UI", size=15, bold=True, color="1A374D")
            font_section = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
            font_bold = Font(name="Segoe UI", size=10, bold=True)
            font_regular = Font(name="Segoe UI", size=10)
            font_small = Font(name="Segoe UI", size=8, italic=True, color="555555")

            thin_border = Border(
                left=Side(style='thin', color='B1D0E0'), right=Side(style='thin', color='B1D0E0'),
                top=Side(style='thin', color='B1D0E0'), bottom=Side(style='thin', color='B1D0E0')
            )

            ws["A1"] = "MEMORIA TÉCNICA DE DISEÑO HIDRÁULICO - CUENCO SAF"
            ws["A1"].font = font_title

            ws["A4"] = "1. DATOS DE ENTRADA"
            ws.merge_cells("A4:D4")
            ws["A4"].fill = fill_section
            ws["A4"].font = font_section

            headers = ["Parámetro", "Variable", "Valor", "Unidad"]
            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=5, column=c_idx, value=h)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center")

            inp = self.data_export["inputs"]
            datos_in = [
                ("Tipo de Estructura de Entrada", "Proyecto", inp["Tipo Proyecto"], "-"),
                ("Caudal del Diseño de la Obra", "Q", inp["Q"], "m³/s"),
                ("Tirante en Sección Supercrítica", "y₁", inp["y1"], "m"),
                ("Temperatura del Fluido", "T", inp["Temp"], "°C"),
                ("Altitud del Proyecto", "A", inp["Alt"], "msnm"),
                ("Ancho Solera Azud (si aplica)", "Bᵣ", inp["Br"], "m"),
                ("Diámetro de Conducción Colector", "D₀", inp["D0"], "m")
            ]

            r_curr = 6
            for item in datos_in:
                ws.cell(row=r_curr, column=1, value=item[0]).font = font_regular
                ws.cell(row=r_curr, column=2, value=item[1]).font = font_bold
                v_cell = ws.cell(row=r_curr, column=3, value=item[2])
                v_cell.font = font_regular
                if isinstance(item[2], float): v_cell.number_format = "0.00"
                ws.cell(row=r_curr, column=4, value=item[3]).font = font_regular
                for col in range(1, 5): ws.cell(row=r_curr, column=col).border = thin_border
                r_curr += 1

            r_curr += 1
            ws.cell(row=r_curr, column=1, value="2. RESULTADOS DEL DIMENSIONAMIENTO").font = font_section
            ws.merge_cells(start_row=r_curr, start_column=1, end_row=r_curr, end_column=4)
            ws.cell(row=r_curr, column=1).fill = fill_section

            r_curr += 1
            for c_idx, h in enumerate(headers, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=h)
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center")

            r_curr += 1
            res = self.data_export["results"]
            datos_out = [
                ("Ancho del Cuenco Disipador", "Wb", res["Wb"], "m"),
                ("Ancho Estructura Rápida Entrada", "Wb₁", res["Wb1"], "m"),
                ("Ancho Sección Central de Fondo", "Wb₂", res["Wb2"], "m"),
                ("Ancho de la Sección de Salida", "Wb₃", res["Wb3"], "m"),
                ("Velocidad de Flujo Real (Supercrítica)", "V₁", res["V1"], "m/s"),
                ("Número de Froude de Flujo", "Fr₁", res["Fr1"], "-"),
                ("Tirante Crítico Control", "y꜀", res["yc"], "m"),
                ("Tirante Conjugado", "y₂", res["d2"], "m"),
                ("Factor Corrector por Froude", "C", res["C"], "-"),
                ("Longitud Transición de Entrada", "Lₐ", res["LA"], "m"),
                ("Longitud Total del Cuenco SAF", "L_B", res["LB"], "m"),
                ("Ubicación de Bloques Intermedios", "Dₑₛ", res["Des"], "m"),
                ("Número Bloques de Rápiva", "n꜀♭", res["ncb"], "u"),
                ("Ancho Bloques de Rápida", "w꜀♭", res["wcb"], "m"),
                ("Altura Bloques de Rápida", "h꜀♭", res["hcb"], "m"),
                ("Número Bloques de Fondo", "nբ♭", res["nfb"], "u"),
                ("Ancho Bloques de Fondo", "wբ♭", res["wfb"], "m"),
                ("Altura Bloques de Fondo", "hբ♭", res["hfb"], "m"),
                ("Espesor Cresta de Bloques", "e", res["e"], "m"),
                ("Ubicación Lateral a Pared", "dᵣₑₐₗ", res["dreal"], "m"),
                ("Altura de Umbral de Salida", "h₄", res["h4"], "m"),
                ("Altura de Paredes Laterales", "H₆", res["h6"], "m"),
                ("Margen Borde Libre Sugerido", "z_bl", res["z_bl"], "m"),
                ("Índice de Cavitación Calculado", "σ", res["sigma"], "-"),
                ("Evaluación frente a Cavitación", "Cavitación", self.data_export["cav_text"], "Estado")
            ]

            for item in datos_out:
                ws.cell(row=r_curr, column=1, value=item[0]).font = font_regular
                ws.cell(row=r_curr, column=2, value=item[1]).font = font_bold
                v_cell = ws.cell(row=r_curr, column=3, value=item[2])

                if isinstance(item[2], float):
                    v_cell.font = font_regular
                    v_cell.number_format = "0.0000" if item[1] == "σ" else "0.00"
                elif isinstance(item[2], int):
                    v_cell.font = font_regular
                    v_cell.number_format = "0"
                else:
                    v_cell.font = font_bold
                    v_cell.alignment = Alignment(horizontal="left")

                ws.cell(row=r_curr, column=4, value=item[3]).font = font_regular

                if r_curr % 2 == 0:
                    for col in range(1, 5): ws.cell(row=r_curr, column=col).fill = fill_zebra
                for col in range(1, 5): ws.cell(row=r_curr, column=col).border = thin_border
                r_curr += 1

            ws.cell(row=r_curr+1, column=1, value="Nota: Ecuaciones y coeficientes empíricos adaptados de Chow (1994) y FEMA (2010).").font = font_small

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
            ws.column_dimensions['A'].width = 36
            ws.column_dimensions['C'].width = 45

            wb.save(path)
            messagebox.showinfo("Éxito", "Memoria de cálculo exportada correctamente a Excel por Tablas.")
        except Exception as e:
            messagebox.showerror("Error de Almacenamiento", f"No se pudo guardar el archivo.\nDetalle: {e}")

if __name__ == "__main__":
    SAF_designer_Final().mainloop()

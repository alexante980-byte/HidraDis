import numpy as np

def calcular_presion_vapor(temperatura_c: float) -> float:
    """
    Calcula la presión de vapor del agua (Hv) en m.c.a. utilizando la 
    ecuación de Antoine: Pmca = 0.0136 * 10^(8.07131 - (1730.63 / (T + 233.426)))
    """
    p_mmhg = 10**(8.07131 - (1730.63 / (temperatura_c + 233.426)))
    h_v = 0.0136 * p_mmhg
    return h_v

def resolver_hidraulica_saf(
    opcion: str, 
    Q_max: float, 
    y_1: float, 
    ancho_o_diametro: float,
    temp_agua: float
) -> dict:
    """
    Motor de cálculo simplificado para el cuenco SAF e Índice de Cavitación.
    """
    g = 9.81
    res = {}
    H_atm = 10.33  # Valor fijo estándar en m.c.a. según la hoja técnica del usuario

    # 1. Definición de Parámetros según la opción seleccionada
    if opcion == "1":  # Azud / Ancho Constante
        B_r = ancho_o_diametro
        W_B1 = B_r
        W_B = B_r
        V_1 = Q_max / (B_r * y_1)
        L_A = 0.0
    elif opcion == "2":  # Descarga de Colector / Ancho Variable
        D_0 = ancho_o_diametro
        if np.isclose(y_1, D_0, atol=1e-3) or y_1 > D_0:
            V_1 = Q_max / ((np.pi * D_0**2) / 4)
        else:
            V_1 = Q_max / (D_0 * y_1)
        
        if V_1 < 6.0:
            W_B1 = D_0
        elif 6.0 <= V_1 < 12.0:
            W_B1 = 2.5 * D_0
        else:
            W_B1 = 3.0 * D_0
        
        W_B_calc = 1.70 * Q_max / (np.sqrt(g) * (D_0**1.5))
        W_B = max(D_0, W_B_calc)
        L_A = (W_B1 - D_0) / 2.0
    else:
        raise ValueError("Opción no válida.")

    # Parámetros Hidráulicos Principales
    Fr_1 = V_1 / np.sqrt(g * y_1)
    y_c = ((Q_max / W_B)**2 / g)**(1/3)
    d_2 = (y_1 / 2.0) * (np.sqrt(1.0 + 8.0 * Fr_1**2) - 1.0)

    # 2. Módulo de Cavitación (Ecuaciones de tu Excel)
    H_v = calcular_presion_vapor(temp_agua)
    carga_velocidad = (V_1**2) / (2 * g)
    sigma = (H_atm - H_v) / carga_velocidad
    
    sigma_critico = 1.0 
    estado_cavitacion = "Condición segura, No hay Cavitación" if sigma >= sigma_critico else "RIESGO DE CAVITACIÓN"

    # 3. Factor de Corrección 'C' y Longitud 'L_B'
    if 1.7 < Fr_1 <= 5.5:
        C = 1.1 - (Fr_1**2 / 120.0)
    elif 5.5 < Fr_1 <= 11.0:
        C = 0.85
    else:
        C = 1.0 - (Fr_1**2 / 800.0)
        C = max(0.65, C)

    L_B = (4.5 * d_2) / (C * (Fr_1**0.76))
    Des = L_B / 3.0

    W_B2 = W_B1 + (W_B - W_B1) / 3.0 if opcion == "2" else W_B
    W_B3 = W_B + 2.0 * L_B if opcion == "2" else W_B

    # 4. Bloques de Rápida y Fondo
    n_cb = max(1, int(W_B1 / (1.5 * y_1)))
    w_cb = W_B1 / (2 * n_cb)
    h_cb = y_1

    h_fb = 0.07 * d_2
    w_fb = 0.75 * y_1
    n_fb = max(1, int((W_B2 - w_fb) / (2.0 * w_fb)) + 1)
    w_fb_ajustado = W_B2 / (2 * n_fb)
    d_real = w_fb_ajustado / 2.0
    ocupacion = (n_fb * w_fb_ajustado / W_B2) * 100

    # 5. Estribos y Paredes
    h_4 = max(0.05, 0.07 * d_2)
    h_6_min = d_2 * (1.0 + (1.0 / (3.0 * C)))
    h_6_max = h_6_min * 1.20

    res.update({
        "W_B1": W_B1, "W_B": W_B, "W_B2": W_B2, "W_B3": W_B3, "L_A": L_A,
        "V_1": V_1, "Fr_1": Fr_1, "y_c": y_c, "d_2": d_2, "C": C, "L_B": L_B, "Des": Des,
        "n_cb": n_cb, "w_cb": w_cb, "h_cb": h_cb,
        "n_fb": n_fb, "w_fb": w_fb_ajustado, "h_fb": h_fb, "d_real": d_real, "ocupacion": ocupacion,
        "h_4": h_4, "h_6_min": h_6_min, "h_6_max": h_6_max,
        "H_atm": H_atm, "H_v": H_v, "sigma": sigma, "estado_cavitacion": estado_cavitacion
    })
    return res


def UI_consola():
    print("\n" + "="*75)
    print("      SISTEMA EXPERTO DE DISEÑO SAF & ANÁLISIS DE CAVITACIÓN")
    print("="*75)
    print("Seleccione el tipo de proyecto:")
    print("1. Obras de Derivación (Azud - Ancho constante)")
    print("2. Descargas de Colectores (Tubería - Ancho variable)")
    
    opcion = input("\nIngrese su opción (1 o 2): ").strip()
    if opcion not in ["1", "2"]:
        print("Opción inválida.")
        return

    try:
        # --- ENTRADAS DE USUARIO SOLICITADAS ---
        Q_max = float(input("\nCaudal Máximo Q [m3/s]: "))
        y_1 = float(input("Tirante Supercrítico y₁ [m]: "))
        
        if opcion == "1":
            ancho_o_diametro = float(input("Ancho del Cuenco Br [m]: "))
        else:
            ancho_o_diametro = float(input("Diámetro del colector D₀ [m]: "))
            
        temp_agua = float(input("Temperatura del agua [°C]: "))

        # Procesamiento numérico con los parámetros de entrada exactos
        data = resolver_hidraulica_saf(opcion, Q_max, y_1, ancho_o_diametro, temp_agua)

        # --- REPORTE TÉCNICO DE RESULTADOS ---
        print("\n" + "-"*60)
        print("           RESULTADOS DETALLADOS DEL DISEÑO SAF")
        print("-"*60)
        
        print("1. PARÁMETROS GEOMÉTRICOS DE DISEÑO:")
        print(f"   W_B1 (Ancho Rápida):         {data['W_B1']:.3f} m")
        print(f"   W_B  (Ancho Cuenco):         {data['W_B']:.3f} m")
        print(f"   W_B2 (Ancho Secc. Fondo):    {data['W_B2']:.3f} m")
        print(f"   W_B3 (Ancho Salida Final):   {data['W_B3']:.3f} m")
        
        print("\n2. CONDICIONES DE APROXIMACIÓN:")
        print(f"   V₁   (Velocidad Entrada):    {data['V_1']:.2f} m/s")
        print(f"   Fr₁  (Número de Froude):     {data['Fr_1']:.2f}")
        print(f"   y_c  (Calado Crítico):       {data['y_c']:.3f} m")
        
        print("\n3. MÓDULO CONTROL DE CAVITACIÓN:")
        print(f"   H_atm (Presión Atmosférica): {data['H_atm']:.2f} m.c.a. (Fijo)")
        print(f"   H_v   (Presión de Vapor H₂O): {data['H_v']:.4f} m.c.a. (Antoine)")
        print(f"   σ     (Índice Calculado):    {data['sigma']:.4f}")
        print(f"   VERIFICACIÓN:                >> {data['estado_cavitacion']} <<")

        print("\n4. RESULTADOS DE LA POZA (EMPUJE HIDRÁULICO):")
        print(f"   d₂   (Tirante Conjugado):    {data['d_2']:.3f} m")
        print(f"   C    (Factor Corrección):    {data['C']:.3f}")
        print(f"   L_B  (Longitud Cuenco):      {data['L_B']:.3f} m")

        print("\n5. ELEMENTOS DISIPADORES (BLOQUES):")
        print(f"   n_cb (Bloques de Rápida):    {data['n_cb']} unidades de {data['w_cb']:.3f}m de ancho")
        print(f"   n_fb (Bloques de Fondo):     {data['n_fb']} unidades de {data['w_fb']:.3f}m de ancho")
        print(f"   h_fb (Altura Bloque Fondo):  {data['h_fb']:.3f} m")
        print(f"   Ocupación de Fondo:          {data['ocupacion']:.2f} %")

        print("\n6. ESTRUCTURAS DE SALIDA Y PAREDES:")
        print(f"   h_4  (Altura Umbral):        {data['h_4']:.3f} m")
        print(f"   h_6_max (Muro con Bordolibre):{data['h_6_max']:.3f} m")
        print("-"*60)

    except ValueError as err:
        print(f"\nError en el ingreso de datos: {err}. Recuerda usar números válidos.")

if __name__ == "__main__":
    UI_consola()
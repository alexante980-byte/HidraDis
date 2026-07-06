import numpy as np

def calcular_presion_vapor(temperatura_c: float) -> float:
    """
    Calcula la presión de vapor del agua (Hv) en m.c.a. utilizando la 
    ecuación de Antoine. 
    """
    p_mmhg = 10**(8.07131 - (1730.63 / (temperatura_c + 233.426)))
    h_v = 0.0136 * p_mmhg
    return h_v

def obtener_geometria(y_val, tipo, b_reg, D_reg):
    """Calcula las propiedades geométricas de la sección."""
    if tipo == 1:  # RECTANGULAR
        A = b_reg * y_val
        P = b_reg + 2 * y_val
        T = b_reg
        theta = None
    else:  # CIRCULAR
        y_s = max(1e-6, min(y_val, D_reg - 1e-6))
        theta = 2 * np.arccos(1 - (2 * y_s / D_reg))
        A = (D_reg**2 / 8) * (theta - np.sin(theta))
        P = (D_reg / 2) * theta
        T = D_reg * np.sin(theta / 2)
    
    R = A / P if P > 0 else 0
    return A, P, R, T, theta

def f_manning(y_val, Q_reg, n_reg, S_reg, tipo, b_reg, D_reg):
    """Función de fricción de Manning para el método numérico."""
    A, _, R, _, _ = obtener_geometria(y_val, tipo, b_reg, D_reg)
    if A <= 0: return -Q_reg
    return (1/n_reg) * A * (R**(2/3)) * np.sqrt(S_reg) - Q_reg

def resolver_hidraulica_usbr_vi(
    tipo_sec: int, 
    Q: float, 
    n: float, 
    S: float, 
    ancho_o_diametro: float, 
    temp_agua: float
) -> dict:
    """
    Motor numérico y matemático para el diseño del cuenco USBR VI 
    e Índice de Cavitación.
    """
    g = 9.81
    H_atm = 10.33
    
    b = ancho_o_diametro if tipo_sec == 1 else 0.0
    D = ancho_o_diametro if tipo_sec == 2 else 0.0

    # 1. Algoritmo Newton-Raphson para Tirante Normal
    y_calc = 0.1 * D if tipo_sec == 2 else 0.5
    tol = 1e-9

    for _ in range(200):
        h = 1e-8
        f = f_manning(y_calc, Q, n, S, tipo_sec, b, D)
        df = (f_manning(y_calc + h, Q, n, S, tipo_sec, b, D) - f) / h
        
        paso = 0.6 * (f / df)
        y_new = y_calc - paso
        
        if tipo_sec == 2:
            if y_new > 0.85 * D: y_new = 0.85 * D
            if y_new <= 0: y_new = 1e-4
        else:
            if y_new <= 0: y_new = 1e-3

        if abs(y_new - y_calc) < tol:
            y_calc = y_new
            break
        y_calc = y_new

    # 2. Cálculos Hidráulicos basados en Tirante Redondeado
    y_norm = round(y_calc, 2)
    A, P, R, T, theta = obtener_geometria(y_norm, tipo_sec, b, D)

    V = round(Q / A, 2)
    Fr = round(V / np.sqrt(g * (A / T)), 2)
    Q_cfs = Q * 35.3147

    # 3. Módulo de Índice de Cavitación
    H_v = calcular_presion_vapor(temp_agua)
    carga_velocidad = (V**2) / (2 * g)
    
    # Evitar divisiones por cero si la velocidad es nula
    sigma = (H_atm - H_v) / carga_velocidad if carga_velocidad > 0 else float('inf')
    
    sigma_critico = 1.0
    estado_cavitacion = "Condición segura, No hay Cavitación" if sigma >= sigma_critico else "RIESGO DE CAVITACIÓN (Reducir velocidad o agregar aireación)"

    # 4. Dimensionamiento Geométrico del Cuenco (Unidades Inglesas)
    Winf = 1.47 * (Q_cfs**0.4)
    Wsup = 1.79 * (Q_cfs**0.4)
    W_ft = (Winf + Wsup) / 2.0

    return {
        "y_norm": y_norm, "theta": theta, "A": A, "V": V, "Fr": Fr, "Q_cfs": Q_cfs,
        "H_atm": H_atm, "H_v": H_v, "sigma": sigma, "estado_cavitacion": estado_cavitacion,
        "W_ft": W_ft
    }

def UI_consola():
    print("="*60)
    print("  DISEÑO CUENCO DISIPADOR USBR TIPO VI & CONTROL DE CAVITACIÓN")
    print("="*60)

    print("Seleccione la sección de entrada:")
    print("1 = Rectangular")
    print("2 = Circular (Alcantarilla)")
    
    try:
        tipo_sec = int(input("Opción (1 o 2): "))
        if tipo_sec not in [1, 2]:
            print("Opción inválida.")
            return

        Q = float(input("Caudal de diseño Q [m3/s]: "))
        n = float(input("Coeficiente n de Manning: "))
        S = float(input("Pendiente S [m/m]: "))

        if tipo_sec == 1:
            ancho_o_diametro = float(input("Ancho de solera b [m]: "))
        else:
            ancho_o_diametro = float(input("Diámetro de alcantarilla D [m]: "))

        temp_agua = float(input("Temperatura del agua [°C]: "))

        # Ejecución del backend numérico
        data = resolver_hidraulica_usbr_vi(tipo_sec, Q, n, S, ancho_o_diametro, temp_agua)

        # Imprimir Resultados Hidráulicos
        print("\n" + "-"*50)
        print("                 RESULTADOS HIDRÁULICOS")
        print("-"*50)
        print(f"Tirante normal (y)       = {data['y_norm']:.2f} m")
        if data['theta'] is not None:
            print(f"Ángulo central (theta)   = {data['theta']:.2f} rad")
        print(f"Área mojada (A)          = {data['A']:.2f} m2")
        print(f"Velocidad (V)            = {data['V']:.2f} m/s")
        print(f"Número de Froude (Fr)    = {data['Fr']:.2f}")

        # Imprimir Análisis de Cavitación
        print("\n" + "-"*50)
        print("             ANÁLISIS DE CAVITACIÓN (USBR VI)")
        print("-"*50)
        print(f"H_atm (Presión Atmosférica) : {data['H_atm']:.2f} m.c.a. (Fijo)")
        print(f"H_v   (Presión de Vapor)     : {data['H_v']:.4f} m.c.a. (Antoine)")
        print(f"σ     (Índice Calculado)     : {data['sigma']:.4f}")
        print(f"VERIFICACIÓN:                >> {data['estado_cavitacion']} <<")

        # Imprimir Verificaciones de Diseño Básicas
        print("\n1) VERIFICACIONES DE DISEÑO DE LA NORMA")
        print(f" - Caudal < 400 cfs:      {data['Q_cfs']:.2f} cfs {'(OK)' if data['Q_cfs'] < 400 else '(REVISAR)'}")
        print(f" - Velocidad < 15.24 m/s: {data['V']} m/s {'(OK)' if data['V'] < 15.24 else '(REVISAR)'}")
        print(f" - Froude Fr < 10:        {data['Fr']} {'(OK)' if data['Fr'] < 10 else '(NO RECOMENDADO)'}")

        # Dimensiones Geométricas
        m_conv = 0.3048
        W_ft = data['W_ft']
        
        print("\n2) GEOMETRÍA DEL CUENCO USBR VI")
        print(f"{'Elemento':<25} | {'ft':<10} | {'m':<10}")
        print("-" * 50)

        def mostrar_dim(label, val_ft):
            print(f" {label:<24} | {val_ft:.2f}       | {val_ft * m_conv:.2f}")

        mostrar_dim("Ancho Cuenca (W)", W_ft)
        mostrar_dim("Longitud Cuenca (L)", 1.333 * W_ft)
        mostrar_dim("Espesor frontal (d)", W_ft / 6)
        mostrar_dim("Fillet/Transición (e)", W_ft / 12)
        mostrar_dim("Altura (H)", 0.75 * W_ft)
        mostrar_dim("Lado (a)", 0.50 * W_ft)
        mostrar_dim("Ancho deflector (b)", 0.375 * W_ft)
        mostrar_dim("Altura contrafuerte (c)", 0.50 * W_ft)
        mostrar_dim("Espesor pantalla (t)", W_ft / 12)
        mostrar_dim("Roca protección (Dr)", W_ft / 20)
        print("-"*50)

    except ValueError as err:
        print(f"\nError en el ingreso de datos: {err}. Verifique los valores numéricos.")

if __name__ == "__main__":
    UI_consola()
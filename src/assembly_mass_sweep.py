import win32com.client
import numpy as np
import matplotlib.pyplot as plt

# ---------------- PARÁMETROS ----------------
VARIABLE_NAME = "Sep_X"     # nombre EXACTO de la Global Variable
N_POINTS = 5              # número de puntos
MIN_VALUE = 11.25             # mm
MAX_VALUE = 12.25            # mm


# ---------------- SOLIDWORKS ----------------
sw = win32com.client.Dispatch("SldWorks.Application")
model = sw.ActiveDoc

if model is None:
    raise RuntimeError("No hay documento activo")

if model.GetType != 2:  # swDocASSEMBLY = 2
    raise RuntimeError("El documento activo no es un ensamblaje")

# Equation Manager del ASSEMBLY
eq_mgr = model.GetEquationMgr

# Localizar índice de la Global Variable
var_index = None
for i in range(eq_mgr.GetCount):
    eq = eq_mgr.Equation(i)
    if eq.startswith(f'"{VARIABLE_NAME}"'):
        var_index = i
        break

if var_index is None:
    raise RuntimeError(f'Variable global "{VARIABLE_NAME}" no encontrada')

# Generar array de valores a barrer
values = np.linspace(MIN_VALUE, MAX_VALUE, N_POINTS)
masses = np.zeros_like(values)

print(f"\n📐 Barrido de masa para variable '{VARIABLE_NAME}'\n")

# ---------------- BUCLE PARAMÉTRICO ----------------
for i, value in enumerate(values):

    # 1️⃣ Cambiar variable
    eq_mgr.Equation(
        var_index,
        f'"{VARIABLE_NAME}" = {float(value)}'
    )

    # 2️⃣ Resolver ecuaciones
    eq_mgr.EvaluateAll

    # 3️⃣ Rebuild completo del assembly
    model.ForceRebuild3(True)
    model.ForceRebuild3(True)

    # 4️⃣ Asegurar componentes resueltos
    model.ResolveAllLightWeightComponents(True)

    # 5️⃣ Leer masa
    mass_prop = model.Extension.CreateMassProperty
    mass = mass_prop.Mass
    masses[i] = mass

    print(f"{VARIABLE_NAME} = {value:.3f} m  →  Masa = {mass:.4f} kg")

print("\n✅ Barrido terminado (no se ha guardado ningún archivo)")


print("\n🔢 Arrays NumPy:")
print("Valores:", values)
print("Masas:", masses)


# Ajuste lineal
a, b = np.polyfit(values, masses, 1)

# Línea de regresión
x_fit = np.linspace(values.min(), values.max(), 200)
y_fit = a * x_fit + b

# Gráfico
plt.figure()
plt.plot(values, masses, 'o', label='Datos')
plt.plot(x_fit, y_fit, '-', label=f'Ajuste lineal: masa = {a:.3f}·L + {b:.3f}')
plt.xlabel('Length (m)')
plt.ylabel('Masa (kg)')
plt.title('Masa del ensamblaje vs Length')
plt.legend()
plt.show()

print(f"Ecuación de regresión: masa = {a:.6f} * Length + {b:.6f}")

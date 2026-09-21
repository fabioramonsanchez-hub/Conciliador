import pandas as pd
import random

random.seed(42)

# --- Simular extracto bancario ---
movimientos = [
    {"fecha": "2025-03-02", "descripcion": "PAGO PROVEEDOR JUAN PEREZ SA", "monto": -1500000},
    {"fecha": "2025-03-05", "descripcion": "TRANSF EMPRESA ABC SRL", "monto": -850000},
    {"fecha": "2025-03-07", "descripcion": "DEBITO AUTOMATICO EDESA", "monto": -320000},
    {"fecha": "2025-03-10", "descripcion": "PAGO PROV DISTRIBUIDORA DEL SUR", "monto": -2400000},
    {"fecha": "2025-03-12", "descripcion": "TRANSF JUAN PEREZ SA", "monto": -1500000},
    {"fecha": "2025-03-15", "descripcion": "PAGO SERVICIOS VARIOS", "monto": -450000},
    {"fecha": "2025-03-18", "descripcion": "COOPERATIVA XYZ", "monto": -980000},
]

extracto = pd.DataFrame(movimientos)
extracto["fecha"] = pd.to_datetime(extracto["fecha"])
extracto.to_excel("extracto_banco.xlsx", index=False)

# --- Simular facturas ---
facturas = pd.DataFrame([
    {"fecha": "2025-03-01", "proveedor": "Juan Perez SA", "monto": 1500000, "numero": "F-001"},
    {"fecha": "2025-03-04", "proveedor": "Empresa ABC SRL", "monto": 850000, "numero": "F-002"},
    {"fecha": "2025-03-06", "proveedor": "EDESA", "monto": 320000, "numero": "F-003"},
    {"fecha": "2025-03-09", "proveedor": "Distribuidora del Sur", "monto": 2400000, "numero": "F-004"},
    {"fecha": "2025-03-20", "proveedor": "Proveedor Fantasma", "monto": 500000, "numero": "F-005"},
])

facturas["fecha"] = pd.to_datetime(facturas["fecha"])
facturas.to_excel("facturas.xlsx", index=False)

print("Archivos de prueba creados:")
print("   - extracto_banco.xlsx")
print("   - facturas.xlsx")
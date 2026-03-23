import sqlite3

# Conectar a tu base de datos
conn = sqlite3.connect('ad17solutions.db')
cursor = conn.cursor()

# Crear la tabla
cursor.execute("""
    CREATE TABLE IF NOT EXISTS historial_monto_credito (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        credito_id INTEGER NOT NULL,
        monto_anterior REAL NOT NULL,
        monto_nuevo REAL NOT NULL,
        fecha_cambio DATETIME NOT NULL,
        motivo TEXT,
        usuario TEXT,
        FOREIGN KEY (credito_id) REFERENCES creditos(id) ON DELETE CASCADE
    )
""")

conn.commit()
print("✅ Tabla historial_monto_credito creada exitosamente")

# Verificar
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='historial_monto_credito'")
if cursor.fetchone():
    print("✅ Tabla verificada correctamente")

    # Mostrar estructura
    cursor.execute("PRAGMA table_info(historial_monto_credito)")
    columns = cursor.fetchall()
    print("\n📋 Estructura de la tabla:")
    for col in columns:
        print(f"  - {col[1]} ({col[2]})")
else:
    print("❌ Error: No se pudo crear la tabla")

conn.close()
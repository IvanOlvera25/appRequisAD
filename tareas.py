"""
Tareas programadas para PythonAnywhere (pestaña Tasks).

APScheduler NO funciona dentro de la web app de PythonAnywhere (uWSGI corre sin
--enable-threads y no se puede activar), así que la sincronización, los
recordatorios y las alertas se ejecutan desde este script.

Uso:
    python tareas.py sync           Sincroniza solicitudes y créditos al MySQL remoto
    python tareas.py recurrentes    Crea las solicitudes de pagos recurrentes que vencen en 14 días
                                    y avanza su fecha (agrega --simular para solo ver qué haría)
    python tareas.py recordatorios  Recordatorios de pagos recurrentes
    python tareas.py alertas        Alertas de cambios de estado a creadores
    python tareas.py todo           Todas las anteriores (default si no pasas argumento)
    python tareas.py limpiar        Reporta filas duplicadas viejas en Pagos remoto (no borra)
    python tareas.py limpiar --confirmar   Las borra (correr después de un sync)

Configuración sugerida en la pestaña Tasks de PythonAnywhere:
    Hourly  -> /home/IvanOlvera25/appRequis/venv/bin/python /home/IvanOlvera25/appRequis/tareas.py todo
o por separado:
    Daily 08:00 UTC (02:00 MX) -> ... tareas.py sync
    Daily 13:00 UTC (07:00 MX) -> ... tareas.py recurrentes
    Daily 19:00 UTC (13:00 MX) -> ... tareas.py recordatorios
"""
import sys
from datetime import datetime


def main():
    modo = (sys.argv[1] if len(sys.argv) > 1 else "todo").lower()
    print(f"[tareas] Inicio modo='{modo}' {datetime.now()}")

    # Importar aquí (no arriba) para que un error de import quede registrado en el log de la task
    from app import (
        sync_all_data,
        check_recurring_payment_reminders,
        monitor_state_changes_and_notify,
        limpiar_pagos_remotos_huerfanos,
        generar_solicitudes_recurrentes,
    )

    if modo == "limpiar":
        limpiar_pagos_remotos_huerfanos(confirmar="--confirmar" in sys.argv)
        print(f"[tareas] Fin {datetime.now()}")
        return

    if modo in ("sync", "todo"):
        try:
            sync_all_data()
        except Exception as e:
            print(f"[tareas] Error en sync: {e}")

    # Antes de los recordatorios: generar la solicitud ya avanza la fecha del recurrente
    if modo in ("recurrentes", "todo"):
        try:
            generar_solicitudes_recurrentes(simular="--simular" in sys.argv)
        except Exception as e:
            print(f"[tareas] Error en recurrentes: {e}")

    if modo in ("recordatorios", "todo"):
        try:
            check_recurring_payment_reminders()
        except Exception as e:
            print(f"[tareas] Error en recordatorios: {e}")

    if modo in ("alertas", "todo"):
        try:
            monitor_state_changes_and_notify()
        except Exception as e:
            print(f"[tareas] Error en alertas: {e}")

    print(f"[tareas] Fin {datetime.now()}")


if __name__ == "__main__":
    main()

import flet as ft
from datetime import datetime
import os
import sys
from types import ModuleType
import gspread

# --- MOCK PARA RENDER (Necesario para entornos como Replit/HuggingFace) ---
if "wsgiref" not in sys.modules:
    mock_wsgiref = ModuleType("wsgiref")
    mock_ss = ModuleType("simple_server")
    mock_util = ModuleType("util")

    class MockHandler:
        pass

    mock_ss.WSGIRequestHandler = MockHandler
    mock_wsgiref.simple_server = mock_ss
    mock_wsgiref.util = mock_util
    sys.modules["wsgiref"] = mock_wsgiref
    sys.modules["wsgiref.simple_server"] = mock_ss
    sys.modules["wsgiref.util"] = mock_util


# --- 1. CONEXIÓN ---
def obtener_cliente():
    posibles_rutas = [
        "creds.json",
        "assets/creds.json",
        os.path.join(os.getcwd(), "creds.json"),
    ]
    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            try:
                return gspread.service_account(filename=ruta)
            except:
                pass
    raise FileNotFoundError("No se encontró creds.json.")


# --- 2. LÓGICA DE FORMATO, BORDES Y TOTALES ---
def formatear_y_totalizar(sheet, tarjeta):
    """Aplica bordes, colores, combina celdas y calcula totales dinámicos."""
    data = sheet.get_all_values()
    inicio_bloque = None
    fila_total_ale = None
    fila_total_lu = None

    # Localizar filas clave
    for i, row in enumerate(data):
        row_str = " ".join(row).upper()
        if tarjeta.upper() in row_str and "TOTAL" not in row_str:
            inicio_bloque = i + 2
        if inicio_bloque and "TOTAL ALE" in row_str:
            fila_total_ale = i + 1
        if inicio_bloque and "TOTAL LU" in row_str:
            fila_total_lu = i + 1
            break

    if not (inicio_bloque and fila_total_ale and fila_total_lu):
        return

    requests = []

    # 1. Aplicar Formato (Bordes, Colores y Combinaciones) en Batch
    for r in range(inicio_bloque, fila_total_lu + 1):
        idx = r - 1
        row_data = data[idx] if idx < len(data) else []
        responsable = row_data[3] if len(row_data) > 3 else ""
        detalle = row_data[1] if len(row_data) > 1 else ""

        # Lógica de Color de Fondo
        color = {"red": 1, "green": 1, "blue": 1}  # Blanco
        if "Ale" in responsable:
            color = {"red": 0.85, "green": 0.92, "blue": 0.83}  # Verde suave
        elif "Lu" in responsable:
            color = {"red": 0.82, "green": 0.88, "blue": 1.0}  # Azul suave
        elif "TOTAL" in detalle.upper() or (
            len(row_data) > 0 and "TOTAL" in row_data[0].upper()
        ):
            color = {"red": 0.95, "green": 0.95, "blue": 0.95}  # Gris suave

        # Request de Formato Celda (Fondo y Bordes individuales para evitar error 400)
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": idx,
                        "endRowIndex": r,
                        "startColumnIndex": 0,
                        "endColumnIndex": 21,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": color,
                            "borders": {
                                "top": {"style": "SOLID"},
                                "bottom": {"style": "SOLID"},
                                "left": {"style": "SOLID"},
                                "right": {"style": "SOLID"},
                            },
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,borders)",
                }
            }
        )

        # Combinar B:C y D:E solo si NO es una fila de Total
        if "TOTAL" not in detalle.upper():
            for start_col, end_col in [(1, 3), (3, 5)]:
                requests.append(
                    {
                        "mergeCells": {
                            "range": {
                                "sheetId": sheet.id,
                                "startRowIndex": idx,
                                "endRowIndex": r,
                                "startColumnIndex": start_col,
                                "endColumnIndex": end_col,
                            },
                            "mergeType": "MERGE_ALL",
                        }
                    }
                )

    # Enviar todos los formatos de una vez para evitar Error 429 (Quota Exceeded)
    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})

    # 2. Fórmulas de Totales Dinámicas
    celdas_formulas = []
    for col_idx in range(10, 22):
        letra = chr(64 + col_idx)
        f_ale = f'=SUMAR.SI($D${inicio_bloque}:$D${fila_total_ale-1}; "Ale"; {letra}${inicio_bloque}:{letra}${fila_total_ale-1})'
        f_lu = f'=SUMAR.SI($D${fila_total_ale+1}:$D${fila_total_lu-1}; "Lu"; {letra}${fila_total_ale+1}:{letra}${fila_total_lu-1})'

        celdas_formulas.append(
            {"range": f"{letra}{fila_total_ale}", "values": [[f_ale]]}
        )
        celdas_formulas.append({"range": f"{letra}{fila_total_lu}", "values": [[f_lu]]})

    sheet.batch_update(celdas_formulas, value_input_option="USER_ENTERED")


# --- 3. PROCESO DE CARGA ---
def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    ss = client.open("Gastos 2026 - Tarjetas")

    monto_f = float(
        str(monto).replace("$", "").replace(".", "").replace(",", ".").strip()
    )
    cant_c = int(cuotas)
    val_c = monto_f / cant_c
    meses = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]
    idx_m = meses.index(mes_inicio)

    def procesar_hoja(año, cuotas_restantes, start_idx):
        try:
            sheet = ss.worksheet(f"Gastos {año}")
        except:
            return 0

        data = sheet.get_all_values()
        f_ins = None
        en_bloque = False

        for i, row in enumerate(data):
            row_str = " ".join(row).upper()
            if tarjeta.upper() in row_str and "TOTAL" not in row_str:
                en_bloque = True
            if en_bloque and f"TOTAL {responsable.upper()}" in row_str:
                f_ins = i + 1
                break

        if f_ins:
            sheet.insert_row([], f_ins)
            det_final = (
                detalle.strip().title()
                if año == 2026
                else f"{detalle.strip().title()} (Cont.)"
            )

            fila = [
                datetime.now().strftime("%d/%m/%Y"),
                det_final,
                "",
                responsable,
                "",
                f"{str(año)[2:]}-{meses[start_idx][:3].lower()}",
                monto_f,
                cant_c,
                val_c,
            ]

            for i in range(12):
                if i >= start_idx and cuotas_restantes > 0:
                    fila.append(f"=$I{f_ins}")
                    cuotas_restantes -= 1
                else:
                    fila.append("")

            sheet.update(
                range_name=f"A{f_ins}", values=[fila], value_input_option="USER_ENTERED"
            )
            formatear_y_totalizar(sheet, tarjeta)
            return cuotas_restantes
        return cuotas_restantes

    quedan = procesar_hoja(2026, cant_c, idx_m)
    if quedan > 0:
        procesar_hoja(2027, quedan, 0)


# --- 4. INTERFAZ (FLET) ---
def main(page: ft.Page):
    page.title = "Tarjetita 2.0"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 450
    page.scroll = ft.ScrollMode.ADAPTIVE

    st = ft.Text("Listo para cargar", weight="bold")
    tar = ft.Dropdown(
        label="Tarjeta",
        value="VISA",
        options=[ft.dropdown.Option("VISA"), ft.dropdown.Option("MASTERCARD")],
        expand=True,
    )
    det = ft.TextField(label="Detalle de compra", expand=True)
    mon = ft.TextField(
        label="Monto Total", keyboard_type=ft.KeyboardType.NUMBER, expand=2
    )
    cuo = ft.TextField(
        label="Cuotas", value="1", keyboard_type=ft.KeyboardType.NUMBER, width=80
    )
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
        options=[ft.dropdown.Option("Ale"), ft.dropdown.Option("Lu")],
        expand=1,
    )

    meses_l = [
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]
    mes = ft.Dropdown(
        label="Mes Inicio",
        value=meses_l[datetime.now().month - 1],
        options=[ft.dropdown.Option(m) for m in meses_l],
        expand=2,
    )

    def click(e):
        if not det.value or not mon.value:
            return
        st.value = "⏳ Procesando... No cierres la app"
        st.color = "blue"
        page.update()
        try:
            cargar_gasto(
                det.value, mon.value, cuo.value, res.value, mes.value, tar.value
            )
            st.value = "✅ ¡Cargado y Formateado!"
            st.color = "green"
            det.value = ""
            mon.value = ""
            page.update()
        except Exception as ex:
            st.value = f"❌ Error: {str(ex)}"
            st.color = "red"
            page.update()

    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Text("Tarjetita", size=32, weight="bold", color="blue700"),
                    ft.Row([tar]),
                    ft.Row([det]),
                    ft.Row([mon, cuo]),
                    ft.Row([res, mes]),
                    ft.ElevatedButton(
                        "CARGAR GASTO", on_click=click, width=page.width, height=50
                    ),
                    st,
                ],
                spacing=15,
            ),
            padding=20,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)

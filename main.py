import flet as ft
from datetime import datetime
import os
import sys
from types import ModuleType

# --- MOCK PARA RENDER ---
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

import gspread


# --- 1. CONFIGURACIÓN DE CONEXIÓN ---
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
            except Exception as e:
                print(f"Error en {ruta}: {e}")
    raise FileNotFoundError("No se encontró creds.json.")


# --- 2. FUNCIONES DE GOOGLE SHEETS (ESTILOS Y FORMATO) ---
def obtener_o_crear_pestana(spreadsheet, año):
    nombre = f"Gastos {año}"
    try:
        return spreadsheet.worksheet(nombre)
    except gspread.exceptions.WorksheetNotFound:
        plantilla = spreadsheet.get_worksheet(0)
        nueva = spreadsheet.duplicate_sheet(plantilla.id, new_sheet_name=nombre)
        nueva.batch_clear(["A3:U100"])
        return nueva


def embellecer_fila(sheet, fila_idx):
    requests = [
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": fila_idx - 1,
                    "endRowIndex": fila_idx,
                    "startColumnIndex": 1,
                    "endColumnIndex": 3,
                },
                "mergeType": "MERGE_ALL",
            }
        },
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": fila_idx - 1,
                    "endRowIndex": fila_idx,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5,
                },
                "mergeType": "MERGE_ALL",
            }
        },
        {
            "updateBorders": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": fila_idx - 1,
                    "endRowIndex": fila_idx,
                    "startColumnIndex": 0,
                    "endColumnIndex": 21,
                },
                "bottom": {"style": "SOLID", "width": 1},
                "top": {"style": "SOLID", "width": 1},
                "left": {"style": "SOLID", "width": 1},
                "right": {"style": "SOLID", "width": 1},
                "innerVertical": {"style": "SOLID", "width": 1},
            }
        },
    ]
    sheet.spreadsheet.batch_update({"requests": requests})


def aplicar_estilos_y_totales(sheet, fila_nueva, responsable, tarjeta):
    formato_plata = {"type": "CURRENCY", "pattern": '"$" #,##0.00'}

    # 1. Formatear TODAS las columnas de dinero (G, I y de J a U)
    sheet.format(
        f"G3:G150", {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"}
    )
    sheet.format(
        f"I3:U150", {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"}
    )

    # 2. Color del responsable en la fila nueva
    color = (
        {"red": 0.85, "green": 0.92, "blue": 0.83}
        if responsable == "Ale"
        else {"red": 0.8, "green": 0.88, "blue": 1.0}
    )
    sheet.format(
        f"D{fila_nueva}:E{fila_nueva}",
        {
            "backgroundColor": color,
            "textFormat": {"bold": True},
            "horizontalAlignment": "CENTER",
        },
    )

    # 3. REPARAR TOTALES: Buscar la fila de "TOTAL [Responsable]" y actualizar sus fórmulas
    data = sheet.get_all_values()
    fila_total = None
    en_bloque_tarjeta = False

    for i, row in enumerate(data):
        row_str = " ".join(row).upper()
        if tarjeta.upper() in row_str and "TOTAL" not in row_str:
            en_bloque_tarjeta = True
        if en_bloque_tarjeta and f"TOTAL {responsable.upper()}" in row_str:
            fila_total = i + 1
            break

    if fila_total:
        # Actualizamos las fórmulas de J a U para que sumen todo lo de arriba hasta la fila 2
        for col_idx in range(10, 21):  # Columnas J a U
            letra = chr(65 + col_idx)
            # Ejemplo: =SUMAR.SI($D$2:$D$40; "Ale"; J$2:J$40)
            formula = f'=SUMAR.SI($D$2:$D${fila_total-1}; "{responsable}"; {letra}$2:{letra}${fila_total-1})'
            sheet.update_acell(f"{letra}{fila_total}", formula)


# --- 3. PROCESO DE CARGA ---
def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    ss = client.open("Gastos 2026 - Tarjetas")
    monto_f = float(
        str(monto).replace("$", "").replace(".", "").replace(",", ".").strip()
    )
    cant_c = int(cuotas)
    val_c = monto_f / cant_c
    det_f = detalle.strip().title()
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

    def procesar_hoja(año_hoja):
        sheet = obtener_o_crear_pestana(ss, año_hoja)
        data = sheet.get_all_values()
        f_ins = None
        en_bloque = False
        for i, row in enumerate(data):
            f_str = " ".join(row).upper()
            if tarjeta.upper() in f_str and "TOTAL" not in f_str:
                en_bloque = True
            if en_bloque and f"TOTAL {responsable.upper()}" in f_str:
                f_ins = i + 1
                break
        if f_ins is None:
            f_ins = len(data) + 1

        sheet.insert_row([], f_ins)
        embellecer_fila(sheet, f_ins)
        prefijo = str(año_hoja)[2:]
        fila_datos = [
            datetime.now().strftime("%d/%m/%Y"),
            det_f,
            "",
            responsable,
            "",
            f"{prefijo}-{mes_inicio[:3].lower()}",
            monto_f,
            cant_c,
            val_c,
        ]

        for i in range(12):
            if i >= idx_m and i < idx_m + cant_c:
                fila_datos.append(f"=$I{f_ins}")
            else:
                fila_datos.append("")

        sheet.update(
            range_name=f"A{f_ins}",
            values=[fila_datos],
            value_input_option="USER_ENTERED",
        )
        aplicar_estilos_y_totales(sheet, f_ins, responsable, tarjeta)

    procesar_hoja(2026)
    if idx_m + cant_c > 12:
        procesar_hoja(2027)


# --- 4. INTERFAZ FLET ---
def main(page: ft.Page):
    page.title = "Tarjetita"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.ADAPTIVE
    page.padding = 20

    st = ft.Text("Listo para cargar", weight="bold")
    tar = ft.Dropdown(
        label="Tarjeta",
        value="VISA",
        options=[ft.dropdown.Option("VISA"), ft.dropdown.Option("MASTERCARD")],
        expand=True,
    )
    det = ft.TextField(
        label="Detalle de compra",
        capitalization=ft.TextCapitalization.WORDS,
        expand=True,
    )
    mon = ft.TextField(
        label="Monto Total",
        prefix=ft.Text("$ "),
        keyboard_type=ft.KeyboardType.NUMBER,
        expand=2,
    )
    cuo = ft.TextField(
        label="Cuotas", value="1", keyboard_type=ft.KeyboardType.NUMBER, width=100
    )
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
        options=[ft.dropdown.Option("Ale"), ft.dropdown.Option("Lu")],
        expand=1,
    )
    meses_lista = [
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
        value=meses_lista[datetime.now().month - 1],
        options=[ft.dropdown.Option(m) for m in meses_lista],
        expand=2,
    )

    def click(e):
        if not det.value or not mon.value:
            st.value = "❌ Completá detalle y monto"
            st.color = "red"
            page.update()
            return
        st.value = "⏳ Actualizando planilla y totales..."
        st.color = "blue"
        page.update()
        try:
            cargar_gasto(
                det.value, mon.value, cuo.value, res.value, mes.value, tar.value
            )
            st.value = "✅ ¡Todo actualizado!"
            st.color = "green"
            det.value = ""
            mon.value = ""
        except Exception as ex:
            st.value = f"❌ Error: {str(ex)}"
            st.color = "red"
        page.update()

    page.add(
        ft.SafeArea(
            ft.Container(
                content=ft.Column(
                    [
                        ft.Text("Tarjetita", size=32, weight="bold", color="blue700"),
                        ft.Row([tar]),
                        ft.Row([det]),
                        ft.Row([mon, cuo], spacing=10),
                        ft.Row([res, mes], spacing=10),
                        ft.Divider(height=20, color="transparent"),
                        ft.ElevatedButton(
                            "CARGAR GASTO",
                            on_click=click,
                            width=page.width,
                            height=60,
                            style=ft.ButtonStyle(
                                shape=ft.RoundedRectangleBorder(radius=10)
                            ),
                        ),
                        st,
                    ],
                    spacing=15,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                padding=10,
            )
        )
    )


if __name__ == "__main__":
    ft.app(target=main)

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
            except Exception as e:
                print(f"Error en {ruta}: {e}")
    raise FileNotFoundError("No se encontró creds.json.")


# --- 2. FUNCIONES DE FORMATO Y TOTALES ---
def aplicar_estilos_y_totales(sheet, responsable, tarjeta):
    formato_plata = {"type": "CURRENCY", "pattern": '"$" #,##0.00'}

    # 1. Aplicar formato moneda a columnas clave (G, I hasta U)
    sheet.format(
        "G3:G200", {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"}
    )
    sheet.format(
        "I3:U200", {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"}
    )

    # 2. Actualizar Fórmulas de Totales
    data = sheet.get_all_values()
    fila_total = None
    en_bloque = False

    # Buscamos la fila exacta del total dentro del bloque de la tarjeta
    for i, row in enumerate(data):
        row_str = " ".join(row).upper()
        if tarjeta.upper() in row_str and "TOTAL" not in row_str:
            en_bloque = True
        if en_bloque and f"TOTAL {responsable.upper()}" in row_str:
            fila_total = i + 1
            break

    if fila_total:
        # Actualizamos las fórmulas de J a U (columnas 10 a 21)
        # El rango de suma es desde la fila 2 hasta justo antes del total
        for col_idx in range(10, 22):
            letra = chr(64 + col_idx)
            # SUMAR.SI($D$2:$D$fila_antes; "Nombre"; Col_Actual$2:Col_Actual$fila_antes)
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

    def procesar(año):
        try:
            sheet = ss.worksheet(f"Gastos {año}")
        except:
            return  # Si no existe la hoja del próximo año, ignorar

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

        if f_ins:
            sheet.insert_row([], f_ins)
            # Estructura: Fecha, Detalle, (vacio), Responsable, (vacio), Mes, Total, Cant, Cuota
            fila_datos = [
                datetime.now().strftime("%d/%m/%Y"),
                det_f,
                "",
                responsable,
                "",
                f"{str(año)[2:]}-{mes_inicio[:3].lower()}",
                monto_f,
                cant_c,
                val_c,
            ]

            # Llenar solo los meses que corresponden
            for i in range(12):
                if i >= idx_m and i < idx_m + cant_c:
                    fila_datos.append(f"=$I{f_ins}")
                else:
                    fila_datos.append("")  # Dejar vacío para no ensuciar la planilla

            sheet.update(
                range_name=f"A{f_ins}",
                values=[fila_datos],
                value_input_option="USER_ENTERED",
            )
            aplicar_estilos_y_totales(sheet, responsable, tarjeta)

    procesar(2026)
    if idx_m + cant_c > 12:
        procesar(2027)


# --- 4. INTERFAZ FLET ---
def main(page: ft.Page):
    page.title = "Tarjetita"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.ADAPTIVE
    page.padding = 20

    st = ft.Text("Listo para cargar", weight="bold")

    # Quitamos 'prefix_text' y 'text_capitalize' para evitar el error de Render
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
            st.value = "❌ Completá datos"
            st.color = "red"
            page.update()
            return
        st.value = "⏳ Actualizando planilla..."
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
                            "CARGAR GASTO", on_click=click, width=page.width, height=60
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

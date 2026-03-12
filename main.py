import flet as ft
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import time

# --- 1. CONFIGURACIÓN DE CONEXIÓN ---


# Reemplazá tu función obtener_cliente por esta:
def obtener_cliente():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS"))
        return gspread.authorize(
            ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        )

    # Intentar rutas comunes para el celular y PC
    rutas_posibles = ["assets/creds.json", "creds.json"]
    for ruta in rutas_posibles:
        if os.path.exists(ruta):
            return gspread.authorize(
                ServiceAccountCredentials.from_json_keyfile_name(ruta, scope)
            )

    raise FileNotFoundError(
        "No se encontró el archivo creds.json en ninguna ruta conocida."
    )


# ... (Tus funciones obtener_o_crear_pestana y aplicar_estilos_y_totales quedan igual) ...


def obtener_o_crear_pestana(spreadsheet, año):
    nombre = f"Gastos {año}"
    try:
        return spreadsheet.worksheet(nombre)
    except gspread.exceptions.WorksheetNotFound:
        plantilla = spreadsheet.get_worksheet(0)
        nueva = spreadsheet.duplicate_sheet(plantilla.id, new_sheet_name=nombre)
        nueva.batch_clear(["A3:U100"])
        return nueva


def aplicar_estilos_y_totales(
    sheet, fila_nueva, responsable, ultima_cuota_aca, tarjeta_actual
):
    formato_plata = {"type": "CURRENCY", "pattern": '"$" #,##0.00'}
    sheet.format(f"G{fila_nueva}", {"numberFormat": formato_plata})
    sheet.format(f"I{fila_nueva}", {"numberFormat": formato_plata})
    sheet.format(
        f"J{fila_nueva}:U{fila_nueva}",
        {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"},
    )
    sheet.format(
        f"A{fila_nueva}:U{fila_nueva}",
        {
            "borders": {
                "top": {"style": "SOLID"},
                "bottom": {"style": "SOLID"},
                "left": {"style": "SOLID"},
                "right": {"style": "SOLID"},
            }
        },
    )

    color = (
        {"red": 0.85, "green": 0.92, "blue": 0.83}
        if responsable == "Ale"
        else {"red": 0.8, "green": 0.88, "blue": 1.0}
    )
    sheet.format(
        f"D{fila_nueva}",
        {
            "backgroundColor": color,
            "textFormat": {"bold": True},
            "horizontalAlignment": "CENTER",
        },
    )

    if ultima_cuota_aca is not None:
        col_letra = chr(74 + ultima_cuota_aca)
        sheet.format(f"{col_letra}{fila_nueva}", {"backgroundColor": color})


# --- 3. PROCESO DE CARGA ---


def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    ss = client.open("Gastos 2026 - Tarjetas")

    monto_f = float(
        str(monto).replace("$", "").replace(".", "").replace(",", ".").replace(" ", "")
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

    def procesar_hoja(año):
        sheet = obtener_o_crear_pestana(ss, año)
        data = sheet.get_all_values()
        f_ins = None
        en_bloque_tarjeta = False
        for i, row in enumerate(data):
            f_str = " ".join(row).upper()
            if tarjeta.upper() in f_str and "TOTAL" not in f_str:
                en_bloque_tarjeta = True
            if en_bloque_tarjeta and f"TOTAL {responsable.upper()}" in f_str:
                f_ins = i + 1
                break

        sheet.insert_row([], f_ins)
        sheet.merge_cells(f"B{f_ins}:C{f_ins}")
        sheet.merge_cells(f"D{f_ins}:E{f_ins}")

        prefijo = str(año)[2:]
        fila_datos = [
            datetime.now().strftime("%d/%m/%Y"),
            det_f if año == 2026 else f"{det_f} (Cont.)",
            "",
            responsable,
            "",
            f"{prefijo}-{mes_inicio[:3].lower()}",
            monto_f,
            cant_c,
            val_c,
        ]

        ultima_idx = None
        for i in range(12):
            if año == 2026:
                if i >= idx_m and i < idx_m + cant_c:
                    fila_datos.append(f"=$I{f_ins}")
                    ultima_idx = i
                else:
                    fila_datos.append("")
            else:
                cuotas_rem = (idx_m + cant_c) - 12
                if i < cuotas_rem:
                    fila_datos.append(f"=$I{f_ins}")
                    ultima_idx = i
                else:
                    fila_datos.append("")

        sheet.update(
            range_name=f"A{f_ins}",
            values=[fila_datos],
            value_input_option="USER_ENTERED",
        )
        pintar = (
            ultima_idx
            if (
                (año == 2026 and idx_m + cant_c <= 12)
                or (año == 2027 and idx_m + cant_c > 12)
            )
            else None
        )
        aplicar_estilos_y_totales(sheet, f_ins, responsable, pintar, tarjeta)
        return f_ins

    fila_26 = procesar_hoja(2026)
    if idx_m + cant_c > 12:
        procesar_hoja(2027)
    return fila_26


# --- 4. INTERFAZ FLET CON PANTALLA DE CARGA ---


def main(page: ft.Page):
    page.title = "Tarjetitas"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 450
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    # Componente de la pantalla de carga
    splash = ft.Column(
        [
            ft.Image(src="icon.jpg", width=120, height=120, border_radius=60),
            ft.Text("Iniciando Tarjetitas...", size=20, weight="bold"),
            ft.ProgressBar(width=250, color="blue"),
            ft.Text("Conectando con Google Sheets...", size=12, italic=True),
        ],
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
    )

    page.add(splash)
    page.update()

    # Intentamos la conexión inicial
    try:
        # Esto verifica si las credenciales son válidas antes de mostrar la app
        obtener_cliente()
        time.sleep(1.5)  # Para que se aprecie el logo un momento
    except Exception as e:
        page.clean()
        page.add(ft.Icon(ft.icons.ERROR_OUTLINE, color="red", size=50))
        page.add(ft.Text(f"Error de inicio: {e}", color="red", text_align="center"))
        page.update()
        return

    # Si todo sale bien, limpiamos y armamos la interfaz real
    page.clean()
    page.vertical_alignment = ft.MainAxisAlignment.START

    # --- ELEMENTOS DE LA INTERFAZ ---
    tar = ft.Dropdown(
        label="Tarjeta",
        value="VISA",
        options=[ft.dropdown.Option("VISA"), ft.dropdown.Option("MASTERCARD")],
    )
    det = ft.TextField(label="Detalle de compra")
    mon = ft.TextField(label="Monto Total", prefix=ft.Text("$ "), expand=True)
    cuo = ft.TextField(label="Cuotas", value="1", expand=True)
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
        options=[ft.dropdown.Option("Ale"), ft.dropdown.Option("Lu")],
        expand=True,
    )
    mes = ft.Dropdown(
        label="Mes de Inicio",
        value=datetime.now().strftime("%B").capitalize(),
        options=[
            ft.dropdown.Option(m)
            for m in [
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
        ],
        expand=True,
    )
    st = ft.Text("")

    def click(e):
        st.value = "⏳ Cargando en Google Sheets..."
        st.color = "blue"
        page.update()
        try:
            cargar_gasto(
                det.value, mon.value, cuo.value, res.value, mes.value, tar.value
            )
            st.value = "✅ ¡Gasto registrado!"
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
                    ft.Text("Tarjetitas", size=30, weight="bold", color="blue700"),
                    tar,
                    det,
                    ft.Row([mon, cuo], spacing=10),
                    ft.Row([res, mes], spacing=10),
                    ft.ElevatedButton(
                        "CARGAR GASTO",
                        on_click=click,
                        width=400,
                        height=50,
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=10)
                        ),
                    ),
                    st,
                ],
                spacing=20,
            ),
            padding=20,
        )
    )
    page.update()


if __name__ == "__main__":
    # Importante: assets_dir le dice a Flet dónde están las imágenes y el json
    ft.app(target=main, assets_dir="assets")

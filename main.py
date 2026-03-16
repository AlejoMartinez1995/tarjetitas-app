import flet as ft
from datetime import datetime
import gspread
import os
import json
import time

# --- 1. CONFIGURACIÓN DE CONEXIÓN ---


def obtener_cliente():
    # El scope necesario para Sheets y Drive
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    # Prioridad 1: Render o Entornos con variables de sistema
    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS"))
        # Usamos service_account_from_dict que es más seguro para evitar wsgiref
        return gspread.service_account_from_dict(creds_info)

    # Prioridad 2: Celular/PC (Archivo local)
    # Importante: Asegurate de que tu archivo se llame exatamente 'creds.json'
    rutas_posibles = ["creds.json", "assets/creds.json"]
    for ruta in rutas_posibles:
        if os.path.exists(ruta):
            # Este es el método que NO usa wsgiref
            return gspread.service_account(filename=ruta)

    raise FileNotFoundError("No se encontró el archivo creds.json en el proyecto.")


# --- 2. FUNCIONES DE GOOGLE SHEETS ---


def obtener_o_crear_pestana(spreadsheet, año):
    nombre = f"Gastos {año}"
    try:
        return spreadsheet.worksheet(nombre)
    except gspread.exceptions.WorksheetNotFound:
        # Si no existe, duplica la primera hoja como plantilla
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
        # Calcula la letra de la columna J=74 en ASCII
        col_letra = chr(74 + ultima_cuota_aca)
        sheet.format(f"{col_letra}{fila_nueva}", {"backgroundColor": color})


# --- 3. PROCESO DE CARGA ---


def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    ss = client.open("Gastos 2026 - Tarjetas")

    # Limpieza de monto
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

        if f_ins is None:
            f_ins = len(data) + 1

        sheet.insert_row([], f_ins)

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


# --- 4. INTERFAZ FLET ---


def main(page: ft.Page):
    page.title = "Tarjetitas"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 450
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    st = ft.Text("", weight="bold")

    def click(e):
        if not det.value or not mon.value:
            st.value = "❌ Completá detalle y monto"
            st.color = "red"
            page.update()
            return

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

    # Componentes de la UI
    tar = ft.Dropdown(
        label="Tarjeta",
        value="VISA",
        options=[ft.dropdown.Option("VISA"), ft.dropdown.Option("MASTERCARD")],
    )
    det = ft.TextField(label="Detalle de compra")
    mon = ft.TextField(label="Monto Total", prefix_text="$ ", expand=True)
    cuo = ft.TextField(label="Cuotas", value="1", expand=True)
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
        expand=True,
        options=[ft.dropdown.Option("Ale"), ft.dropdown.Option("Lu")],
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
        expand=True,
        options=[ft.dropdown.Option(m) for m in meses_lista],
    )

    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Text("Tarjetita", size=30, weight="bold", color="blue700"),
                    tar,
                    det,
                    ft.Row([mon, cuo]),
                    ft.Row([res, mes]),
                    ft.ElevatedButton(
                        "CARGAR GASTO", on_click=click, width=400, height=50
                    ),
                    st,
                ],
                spacing=20,
            ),
            padding=20,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)

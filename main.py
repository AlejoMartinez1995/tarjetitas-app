import flet as ft
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json

# --- 1. CONFIGURACIÓN DE CONEXIÓN ---
def obtener_cliente():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    # Si estamos en Render, usamos la variable de entorno
    if "GOOGLE_CREDENTIALS" in os.environ:
        creds_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS"))
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
    else:
        # Si estás en tu PC, sigue buscando el archivo local
        creds = ServiceAccountCredentials.from_json_keyfile_name("assets/creds.json", scope)

    return gspread.authorize(creds)


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


# --- 2. LÓGICA DE ESTILOS Y TOTALES ---
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
    sheet.format(
        f"J{fila_nueva}:U{fila_nueva}",
        {"backgroundColor": {"red": 1, "green": 1, "blue": 1}},
    )

    if ultima_cuota_aca is not None:
        col_letra = chr(74 + ultima_cuota_aca)
        sheet.format(f"{col_letra}{fila_nueva}", {"backgroundColor": color})

    # --- RECALCULAR SUMIF Y RESUMEN GENERAL ---
    data = sheet.get_all_values()
    rango_inicio = 3
    en_bloque = False

    filas_totales_ale = []
    filas_totales_lu = []

    for i, row in enumerate(data):
        num_f = i + 1
        f_str = " ".join(row).upper()

        if ("VISA" in f_str or "MASTERCARD" in f_str) and "TOTAL" not in f_str:
            rango_inicio = num_f + 1
            en_bloque = True

        if "TOTAL ALE" in f_str and "RESUMEN" not in f_str:
            filas_totales_ale.append(num_f)
            if tarjeta_actual.upper() in f_str or (num_f > 10 and en_bloque):
                formulas = [
                    f'=SUMIF($D${rango_inicio}:$D${fila_nueva}; "Ale"; {chr(64+c)}${rango_inicio}:{chr(64+c)}${fila_nueva})'
                    for c in range(10, 22)
                ]
                sheet.update(
                    range_name=f"J{num_f}:U{num_f}",
                    values=[formulas],
                    value_input_option="USER_ENTERED",
                )

        if "TOTAL LU" in f_str and "RESUMEN" not in f_str:
            filas_totales_lu.append(num_f)
            if tarjeta_actual.upper() in f_str or (num_f > 10 and en_bloque):
                formulas = [
                    f'=SUMIF($D${rango_inicio}:$D${fila_nueva}; "Lu"; {chr(64+c)}${rango_inicio}:{chr(64+c)}${fila_nueva})'
                    for c in range(10, 22)
                ]
                sheet.update(
                    range_name=f"J{num_f}:U{num_f}",
                    values=[formulas],
                    value_input_option="USER_ENTERED",
                )

    # ACTUALIZAR EL RESUMEN GENERAL
    for i, row in enumerate(data):
        num_f = i + 1
        f_str = " ".join(row).upper()

        if "RESUMEN GENERAL ALE" in f_str and len(filas_totales_ale) >= 2:
            f_ale = [
                f"={chr(64+c)}{filas_totales_ale[0]}+{chr(64+c)}{filas_totales_ale[1]}"
                for c in range(10, 22)
            ]
            sheet.update(
                range_name=f"J{num_f}:U{num_f}",
                values=[f_ale],
                value_input_option="USER_ENTERED",
            )

        if "RESUMEN GENERAL LU" in f_str and len(filas_totales_lu) >= 2:
            f_lu = [
                f"={chr(64+c)}{filas_totales_lu[0]}+{chr(64+c)}{filas_totales_lu[1]}"
                for c in range(10, 22)
            ]
            sheet.update(
                range_name=f"J{num_f}:U{num_f}",
                values=[f_lu],
                value_input_option="USER_ENTERED",
            )


# --- 3. PROCESO DE CARGA ---
def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    # Asegúrate de que el nombre del Excel en Drive coincida con esto:
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

        if f_ins is None:
            raise Exception(
                f"No encontré la fila de TOTAL para {responsable} en {tarjeta}"
            )

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


# --- 4. INTERFAZ FLET ---
def main(page: ft.Page):
    page.title = "Tarjetitas"  # <-- Nombre corregido
    page.window_width = 450
    page.theme_mode = ft.ThemeMode.LIGHT

    tar = ft.Dropdown(
        label="Tarjeta",
        value="VISA",
        options=[ft.dropdown.Option("VISA"), ft.dropdown.Option("MASTERCARD")],
    )
    det = ft.TextField(label="Detalle de compra")
    mon = ft.TextField(label="Monto Total", prefix=ft.Text("$ "))
    cuo = ft.TextField(label="Cuotas", value="1")
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
        options=[ft.dropdown.Option("Ale"), ft.dropdown.Option("Lu")],
    )
    mes = ft.Dropdown(
        label="Mes de Inicio",
        value="Marzo",
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
    )
    st = ft.Text("")

    def click(e):
        st.value = "⏳ Cargando en Tarjetitas..."
        st.color = "blue"
        page.update()
        try:
            cargar_gasto(
                det.value, mon.value, cuo.value, res.value, mes.value, tar.value
            )
            st.value = "✅ ¡Gasto registrado correctamente!"
            st.color = "green"
            det.value = ""
            mon.value = ""
            page.update()
        except Exception as ex:
            st.value = f"❌ Error: {str(ex)}"
            st.color = "red"
            page.update()

    page.add(
        ft.Text("Tarjetitas", size=25, weight="bold"),  # <-- Nombre corregido
        tar,
        det,
        ft.Row([mon, cuo]),
        ft.Row([res, mes]),
        ft.ElevatedButton("CARGAR GASTO", on_click=click, width=400),
        st,
    )


if __name__ == "__main__":
    ft.app(target=main)

import flet as ft
from datetime import datetime
import os
import sys
from types import ModuleType

# --- TRUCO NIVEL DIOS: Mock completo para engañar a la herencia de clases ---
if "wsgiref" not in sys.modules:
    # Creamos la estructura de carpetas falsa
    mock_wsgiref = ModuleType("wsgiref")
    mock_ss = ModuleType("simple_server")
    mock_util = ModuleType("util")
    
    # Creamos una clase vacía para que la herencia no explote
    class MockHandler: 
        pass

    # Metemos la clase adentro del simple_server falso
    mock_ss.WSGIRequestHandler = MockHandler
    
    # Conectamos todo
    mock_wsgiref.simple_server = mock_ss
    mock_wsgiref.util = mock_util
    
    # Registramos en el sistema
    sys.modules["wsgiref"] = mock_wsgiref
    sys.modules["wsgiref.simple_server"] = mock_ss
    sys.modules["wsgiref.util"] = mock_util

import gspread

# --- 1. CONFIGURACIÓN DE CONEXIÓN ---


def obtener_cliente():
    # Lista de posibles ubicaciones del archivo JSON en Android y PC
    posibles_rutas = [
        "creds.json",  # Raíz del proyecto
        "assets/creds.json",  # Carpeta assets
        os.path.join(os.getcwd(), "creds.json"),
        os.path.join(os.path.dirname(__file__), "creds.json"),
    ]

    for ruta in posibles_rutas:
        if os.path.exists(ruta):
            try:
                return gspread.service_account(filename=ruta)
            except Exception as e:
                print(f"Error cargando {ruta}: {e}")

    # Si llega acá y no lo encontró, tira el error detallado
    raise FileNotFoundError(f"No se encontró creds.json. Busqué en: {posibles_rutas}")


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


def aplicar_estilos_y_totales(sheet, fila_nueva, responsable, ultima_cuota_aca):
    formato_plata = {"type": "CURRENCY", "pattern": '"$" #,##0.00'}

    # Formatear columnas de montos y cuotas
    sheet.format(f"G{fila_nueva}", {"numberFormat": formato_plata})
    sheet.format(f"I{fila_nueva}", {"numberFormat": formato_plata})
    sheet.format(
        f"J{fila_nueva}:U{fila_nueva}",
        {"numberFormat": formato_plata, "horizontalAlignment": "CENTER"},
    )

    # Colores según responsable (Verde para Ale, Azul para Lu)
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
        col_letra = chr(74 + ultima_cuota_aca)  # J=74 en ASCII
        sheet.format(f"{col_letra}{fila_nueva}", {"backgroundColor": color})


# --- 3. PROCESO DE CARGA ---


def cargar_gasto(detalle, monto, cuotas, responsable, mes_inicio, tarjeta):
    client = obtener_cliente()
    # Asegurate de que el nombre coincida EXACTAMENTE con tu planilla
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
        en_bloque_tarjeta = False

        # Buscar la fila donde insertar según la tarjeta y responsable
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

        prefijo = str(año_hoja)[2:]
        fila_datos = [
            datetime.now().strftime("%d/%m/%Y"),
            det_f if año_hoja == 2026 else f"{det_f} (Cont.)",
            "",
            responsable,
            "",
            f"{prefijo}-{mes_inicio[:3].lower()}",
            monto_f,
            cant_c,
            val_c,
        ]

        # Lógica de distribución de cuotas en los meses (J a U)
        ultima_idx_pintar = None
        for i in range(12):
            if año_hoja == 2026:
                if i >= idx_m and i < idx_m + cant_c:
                    fila_datos.append(f"=$I{f_ins}")
                    ultima_idx_pintar = i
                else:
                    fila_datos.append("")
            else:  # Para el año siguiente (2027)
                cuotas_ya_pagadas = 12 - idx_m
                cuotas_restantes = cant_c - cuotas_ya_pagadas
                if i < cuotas_restantes:
                    fila_datos.append(f"=$I{f_ins}")
                    ultima_idx_pintar = i
                else:
                    fila_datos.append("")

        sheet.update(
            range_name=f"A{f_ins}",
            values=[fila_datos],
            value_input_option="USER_ENTERED",
        )

        aplicar_estilos_y_totales(sheet, f_ins, responsable, ultima_idx_pintar)
        return f_ins

    fila_final = procesar_hoja(2026)
    # Si las cuotas desbordan al año siguiente
    if idx_m + cant_c > 12:
        procesar_hoja(2027)

    return fila_final


# --- 4. INTERFAZ FLET ---


def main(page: ft.Page):
    page.title = "Tarjetita"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 450
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.ADAPTIVE

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
    det = ft.TextField(
        label="Detalle de compra", text_capitalize=ft.TextCapitalization.WORDS
    )
    mon = ft.TextField(
        label="Monto Total", prefix_text="$ ", keyboard_type=ft.KeyboardType.NUMBER
    )
    cuo = ft.TextField(label="Cuotas", value="1", keyboard_type=ft.KeyboardType.NUMBER)
    res = ft.Dropdown(
        label="Responsable",
        value="Ale",
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
        options=[ft.dropdown.Option(m) for m in meses_lista],
    )

    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Text("Tarjetita", size=32, weight="bold", color="blue700"),
                    tar,
                    det,
                    ft.Row([mon, cuo], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    ft.Row([res, mes], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
                    ft.ElevatedButton(
                        "CARGAR GASTO",
                        on_click=click,
                        width=400,
                        height=60,
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=10)
                        ),
                    ),
                    st,
                ],
                spacing=20,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=30,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)

import flet as ft
import pandas as pd
import datetime
import os
import webbrowser
import re
import sys

# --- CONFIGURACIÓN DE RUTAS MULTIPLATAFORMA ---
if getattr(sys, 'frozen', False):
    # Si es APK (Android)
    BASE_DIR = os.path.join(sys._MEIPASS, "assets")
else:
    # Si es PC (flet run)
    # Buscamos la carpeta assets que acabas de crear
    BASE_DIR = os.path.join(os.path.dirname(__file__), "assets")

EXCEL_BD = os.path.join(BASE_DIR, "base_datos.xlsx")
PLANTILLA_HTML = os.path.join(BASE_DIR, "plantilla.htm")
# Ruta para el logo dentro de la carpeta assets
RUTA_LOGO = os.path.join(BASE_DIR, "plantilla_archivos")

def cargar_datos():
    try:
        df_emp = pd.read_excel(EXCEL_BD, sheet_name="Hoja1")
        df_ser = pd.read_excel(EXCEL_BD, sheet_name="Hoja2")
        df_art = pd.read_excel(EXCEL_BD, sheet_name="Hoja3")
        return df_emp, df_ser, df_art
    except Exception as e:
        print(f"Error al leer Excel: {e}")
        return None, None, None

df_emp, df_ser, df_art = cargar_datos()

def generar_documento(datos, items):
    try:
        if not os.path.exists(PLANTILLA_HTML):
            return
        
        # SOLUCIÓN SÍMBOLOS: Leer como latin-1 (formato nativo de Word HTML)
        with open(PLANTILLA_HTML, "r", encoding="latin-1") as f:
            html_content = f.read()

        # SOLUCIÓN FOTO: Inyectar la ruta absoluta del sistema para que el navegador la acepte
        ruta_base = BASE_DIR.replace("\\", "/")
        html_content = html_content.replace("<head>", f'<head><base href="file:///{ruta_base}/">')

        # Reemplazos de datos principales
        reemplazos = {
            "{{ciudad}}": str(datos['ciudad']),
            "{{fecha}}": str(datos['fecha']),
            "{{nombre1}}": str(datos['nombre1']),
            "{{cedula1}}": str(datos['cedula1']),
            "{{nombre2}}": str(datos['nombre2']),
            "{{cedula2}}": str(datos['cedula2']),
            "{{numero_serie}}": str(datos['serie']) # Asegura impresión de serie
        }
        
        for clave, valor in reemplazos.items():
            html_content = html_content.replace(clave, valor)

        # SOLUCIÓN ARTÍCULOS: Limpieza total de etiquetas de bucle
        if items:
            # Creamos una lista HTML simple
            lista_html = "<ul>" + "".join([f"<li>{i}</li>" for i in items]) + "</ul>"
        else:
            lista_html = "<i>No se entregaron accesorios adicionales.</i>"

        # Eliminamos el bloque {% for %}...{% endfor %} y cualquier {{ item }} suelto
        html_content = re.sub(r"{%.*?%}", "", html_content, flags=re.DOTALL) 
        html_content = html_content.replace("- {{ item }}", "").replace("{{ item }}", "")
        
        # Buscamos una etiqueta limpia para insertar la lista (debes poner {{lista_final}} en tu Word si esto falla)
        if "{{lista_articulos}}" in html_content:
            html_content = html_content.replace("{{lista_articulos}}", lista_html)
        else:
            # Si no, simplemente la insertamos donde antes estaban los códigos de bucle
            html_content = html_content.replace("{{lista_articulos}}", lista_html)

        # SOLUCIÓN FINAL: Guardar como utf-8-sig para que el navegador entienda las tildes
        nombre_salida = f"Acta_{datos['nombre1'].replace(' ', '_')}.html"
        ruta_salida = os.path.join(BASE_DIR, nombre_salida)
        
        with open(ruta_salida, "w", encoding="utf-8-sig") as f:
            f.write(html_content)

        webbrowser.open(f"file:///{ruta_salida}")
    except Exception as e:
        print(f"Error detallado: {e}")

def main(page: ft.Page):
    page.title = "ISERTEL - Generador Pro"
    page.padding = 20
    page.theme_mode = "light"
    page.scroll = "adaptive"

    accesorios_seleccionados = []
    lista_visual = ft.Column()

    txt_ciudad = ft.TextField(label="Centro de Trabajo", value="Quito", expand=True)
    txt_fecha = ft.TextField(label="Fecha", value=datetime.datetime.now().strftime("%d/%m/%Y"), width=150)
    
    opt_emp = [ft.dropdown.Option(n) for n in df_emp['Nombre'].tolist()] if df_emp is not None else []
    drop_t1 = ft.Dropdown(label="Responsable", options=opt_emp)
    drop_t2 = ft.Dropdown(label="Acompañante", options=opt_emp)
    
    opt_ser = [ft.dropdown.Option(str(s)) for s in df_ser['Serie'].tolist()] if df_ser is not None else []
    drop_ser = ft.Dropdown(label="Serie Fusionadora", options=opt_ser)
    
    opt_art = [ft.dropdown.Option(a) for a in df_art['Nombre_Articulo'].tolist()] if df_art is not None else []
    drop_art = ft.Dropdown(label="Agregar Accesorio", options=opt_art, expand=True)

    def agregar_item(e):
        if drop_art.value:
            accesorios_seleccionados.append(drop_art.value)
            lista_visual.controls.append(ft.Text(f"• {drop_art.value}"))
            page.update()

    def procesar(e):
        if not drop_t1.value or not drop_ser.value: return
        
        c1 = df_emp[df_emp['Nombre'] == drop_t1.value]['Cedula'].values[0]
        c2 = df_emp[df_emp['Nombre'] == drop_t2.value]['Cedula'].values[0] if drop_t2.value else " "
        
        generar_documento({
            "ciudad": txt_ciudad.value, "fecha": txt_fecha.value,
            "nombre1": drop_t1.value, "cedula1": c1,
            "nombre2": drop_t2.value or " ", "cedula2": c2,
            "serie": drop_ser.value
        }, accesorios_seleccionados)

    page.add(
        ft.Text("ACTAS ISERTEL", size=30, weight="bold", color="blue"),
        ft.Row([txt_ciudad, txt_fecha]),
        ft.Row([drop_t1, drop_t2]),
        drop_ser,
        ft.Divider(),
        ft.Row([drop_art, ft.ElevatedButton("Añadir", icon=ft.Icons.ADD, on_click=agregar_item)]),
        lista_visual,
        ft.ElevatedButton("GENERAR ACTA FINAL", on_click=procesar, width=400, height=50)
    )

ft.app(target=main)
import pdfplumber
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def extraer_texto(ruta_archivo):
    extension = os.path.splitext(ruta_archivo)[1].lower()
    if extension == ".pdf":
        texto = ""
        with pdfplumber.open(ruta_archivo) as pdf:
            for pagina in pdf.pages:
                texto += pagina.extract_text() or ""
        return texto
    return ""

def extraer_datos_factura(ruta_archivo):
    datos = {
        "Razón Social": "No encontrado",
        "CIF": "No encontrado",
        "Número Factura": "No encontrado",
        "Fecha": "No encontrado",
        "Base Imponible": "No encontrado",
        "IVA": "No encontrado",
        "Total": "No encontrado"
    }

    texto = extraer_texto(ruta_archivo)
    if not texto:
        return datos

    texto = " ".join(texto.split())

    # ── RAZÓN SOCIAL ──────────────────────────────────────────────
    patrones_razon = [
        r'(?:empresa|raz[oó]n\s*social|emisor|proveedor)[:\s]+([A-ZÁÉÍÓÚÑ][^\n]{3,50}(?:S\.?L\.?|S\.?A\.?|S\.?L\.?U\.?))',
        r'([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñÁÉÍÓÚÑ\s&\.,]{3,50}(?:S\.?L\.?U?\.?|S\.?A\.?))',
    ]
    for patron in patrones_razon:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            valor = m.group(1).strip()
            if len(valor) > 3:
                datos["Razón Social"] = valor[:60]
            break

    # ── CIF / NIF ─────────────────────────────────────────────────
    patrones_cif = [
        r'(?:CIF|NIF|N\.I\.F|C\.I\.F)[:\s\.\-]*([A-Za-z]-?\d{7}-?[A-Za-z0-9])',
        r'\b([A-HJ-NP-SUVW]-?\d{7}-?[0-9A-J])\b',
        r'\b([0-9]{8}[A-Za-z])\b',
        r'\b([XYZ]-?\d{7}[A-Za-z])\b',
    ]
    for patron in patrones_cif:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["CIF"] = m.group(1).upper()
            break

    # ── NÚMERO DE FACTURA ─────────────────────────────────────────
    patrones_numfac = [
        r'\b([A-Z]-\d{4}\/[A-Z]{3}\/\d{4})\b',
        r'CONTINUACI[OÓ]N\s*FACTURA\s+([A-Z0-9][A-Z0-9\-\/]{3,25})',
        r'(?:factura\s*n[uú]m(?:ero)?\.?|n[uú]m(?:ero)?\s*(?:de\s*)?factura|factura\s*n[oº°]?)[:\s\.\-#]*([A-Z0-9][A-Z0-9\-\/]{2,20})',
        r'(?:Ref(?:erencia)?:?\s*)([A-Z]{2,5}[\/\-][0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2,6})',
        r'[Ff]actura[:\s]*([0-9]{4}-[0-9]{3,6})',
        r'[Ff]-(\d{4}-\d{4})',
    ]
    for patron in patrones_numfac:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["Número Factura"] = m.group(1).strip()
            break

    # ── FECHA ─────────────────────────────────────────────────────
    patrones_fecha = [
        r'(?:fecha\s*(?:de\s*)?(?:factura|emisi[oó]n|completa)?)[:\s]*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
        r'(?:fecha\s*completa\s*de\s*emisi[oó]n)[:\s]*(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})',
        r'\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})\b',
    ]
    for patron in patrones_fecha:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["Fecha"] = m.group(1).strip()
            break

    # ── BASE IMPONIBLE ────────────────────────────────────────────
    patrones_base = [
        r'[Bb]ase\s*imponible\s*total[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'[Bb]ase\s*[Ii]mponible[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'[Ss]ubtotal\s*sin\s*IVA[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'[Ss]ubtotal[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'[Nn]eto[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'(?:antes\s*de\s*IVA)[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
    ]
    for patron in patrones_base:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["Base Imponible"] = limpiar_numero(m.group(1)) + " EUR"
            break

    # ── IVA directo ───────────────────────────────────────────────
    patrones_iva = [
        r'IVA\s*\(21%[^)]*\)[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'IVA\s*\(10%[^)]*\)[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'IVA\s*\(4%[^)]*\)[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'IVA\s*(?:21|10|4)\s*%[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'[Cc]uota\s*IVA[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
    ]
    for patron in patrones_iva:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["IVA"] = limpiar_numero(m.group(1)) + " EUR"
            break

    # ── TOTAL ─────────────────────────────────────────────────────
    patrones_total = [
        r'TOTAL\s*A\s*PAGAR\s*\(EUR\)[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)',
        r'TOTAL\s*A\s*PAGAR[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'TOTAL\s*FACTURA[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)?',
        r'TOTAL\s*EUR[:\s]*([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)',
        r'TOTAL[^a-zA-Z0-9]{0,5}([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)\s*(?:€|EUR)',
        r'[Ii]mporte\s*a\s*pagar[^0-9]{0,20}([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})?)',
    ]
    for patron in patrones_total:
        m = re.search(patron, texto, re.IGNORECASE)
        if m:
            datos["Total"] = limpiar_numero(m.group(1)) + " EUR"
            break

    # ── IVA matemático si no se encontró ─────────────────────────
    if datos["IVA"] == "No encontrado" and datos["Base Imponible"] != "No encontrado" and datos["Total"] != "No encontrado":
        try:
            base = float(datos["Base Imponible"].replace(" EUR", "").replace(",", "."))
            total = float(datos["Total"].replace(" EUR", "").replace(",", "."))
            iva_calculado = round(total - base, 2)
            if 0 < iva_calculado < total:
                datos["IVA"] = str(iva_calculado) + " EUR"
        except:
            pass

    return datos


def limpiar_numero(texto):
    texto = texto.strip()
    if re.match(r'^\d{1,3}(\.\d{3})+(,\d{1,2})?$', texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif re.match(r'^\d+(,\d{1,2})$', texto):
        texto = texto.replace(",", ".")
    elif re.match(r'^\d{1,3}(\.\d{3})+$', texto):
        texto = texto.replace(".", "")
    return texto


def guardar_en_excel(datos, ruta_excel="resultados.xlsx", excel_base=None):
    verde_oscuro    = PatternFill("solid", fgColor="1A3C34")
    verde_claro     = PatternFill("solid", fgColor="D6EAE4")
    blanco          = PatternFill("solid", fgColor="FFFFFF")
    fuente_cabecera = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    fuente_normal   = Font(name="Calibri", size=10)
    borde_lado      = Side(style="thin", color="B0C4BE")
    borde           = Border(left=borde_lado, right=borde_lado, top=borde_lado, bottom=borde_lado)
    cabeceras       = ["Razón Social", "CIF", "Número Factura", "Fecha", "Base Imponible", "IVA", "Total"]

    if excel_base and os.path.exists(excel_base):
        libro = openpyxl.load_workbook(excel_base)
        hoja  = libro.active
        es_nuevo = False
    elif os.path.exists(ruta_excel):
        libro = openpyxl.load_workbook(ruta_excel)
        hoja  = libro.active
        es_nuevo = False
    else:
        libro      = openpyxl.Workbook()
        hoja       = libro.active
        hoja.title = "Facturas"
        es_nuevo   = True

    if es_nuevo:
        for col, titulo in enumerate(cabeceras, start=1):
            celda           = hoja.cell(row=1, column=col, value=titulo)
            celda.font      = fuente_cabecera
            celda.fill      = verde_oscuro
            celda.alignment = Alignment(horizontal="center", vertical="center")
            celda.border    = borde
        hoja.row_dimensions[1].height = 22

    fila    = hoja.max_row + 1
    valores = [datos["Razón Social"], datos["CIF"], datos["Número Factura"], datos["Fecha"], datos["Base Imponible"], datos["IVA"], datos["Total"]]
    relleno = verde_claro if fila % 2 == 0 else blanco

    for col, valor in enumerate(valores, start=1):
        celda           = hoja.cell(row=fila, column=col, value=valor)
        celda.font      = fuente_normal
        celda.fill      = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border    = borde

    hoja.row_dimensions[fila].height = 18
    for col in range(1, 8):
        hoja.column_dimensions[get_column_letter(col)].width = 22

    libro.save(ruta_excel)
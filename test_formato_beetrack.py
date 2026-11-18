"""Prueba de punta a punta de formato_beetrack/index.html.

Abre la pagina en un navegador real, le entrega un importable de hela, aprieta
"Descargar corregido" y revisa el .xlsx que sale:

  1. Origen: todo lo de Easy queda 'CD Easy' y Paris queda intacto.
  2. MODO: 'Solo recogida' SOLO en los retiros de Paris (y en ninguna otra fila).
  3. Todo el archivo en formato texto (valores str y number_format '@').
  4. Nada mas cambio: mismas filas, mismas columnas, mismos datos.
  5. El mismo archivo dos veces se agrega una sola vez.

El retiro de Paris del archivo de muestra es la guia 87718446, que en 'BASE DATOS PARIS'
figura con SKU=RETORNO: en el archivo de hela se reconoce porque su CODIGO ITEM es el
LINE_ID del retorno (823636999) y no un codigo de despacho (DD0339..., VEV...).

    python test_formato_beetrack.py [ruta a un beetrack-*.xlsx]
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

AQUI = Path(__file__).resolve().parent
PAGINA = (AQUI / "index.html").as_uri()
ORIGINAL = Path(sys.argv[1] if len(sys.argv) > 1
                else r"C:\Users\Usuario\Downloads\beetrack-14039-0037-20261117.xlsx")
GUIA_REAL = "87718446"            # retiro de Paris que ya trae el archivo de muestra
GUIA_RETIRO = "87724341"          # entrega que el fixture convierte en retiro
SKU_RETIRO = "DD033986405"        # su CODIGO ITEM en el archivo real


def fixture_con_retorno(origen: Path, destino: Path) -> None:
    """Copia el archivo marcando una orden de Paris como RETORNO (como la manda Paris)."""
    with zipfile.ZipFile(origen) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            datos = zin.read(item.filename)
            if item.filename.endswith("sheet1.xml"):
                texto = datos.decode("utf-8")
                assert f"<t>{SKU_RETIRO}</t>" in texto, "cambio el archivo de muestra"
                datos = texto.replace(f"<t>{SKU_RETIRO}</t>", "<t>RETORNO</t>", 1).encode("utf-8")
            zout.writestr(item, datos)


def leer(path: Path) -> tuple[list, dict]:
    """Filas + formato de las celdas con dato y de las columnas.

    Las celdas vacias que hela nunca escribio (F, U, AI..AL en varias filas) no existen
    dentro del .xlsx: su formato lo manda el de la columna, que es lo que se revisa.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    filas = [[c.value for c in f] for f in ws.iter_rows()]
    formatos = {c.coordinate: c.number_format for f in ws.iter_rows() for c in f
                if c.value is not None}
    tipos = {c.coordinate: type(c.value).__name__ for f in ws.iter_rows() for c in f
             if c.value is not None}
    cols = {k: d.number_format for k, d in ws.column_dimensions.items()}
    wb.close()
    return filas, {"fmt": formatos, "tipos": tipos, "cols": cols}


def descargar(page, archivo: Path, salida: Path) -> Path:
    page.goto(PAGINA)
    page.wait_for_function("() => window.XLSX && window.JSZip")
    page.set_input_files("#input", str(archivo))
    page.wait_for_selector(".archivo")
    with page.expect_download() as info:
        page.click("button.principal")
    info.value.save_as(salida)
    return salida


def main() -> int:
    assert ORIGINAL.exists(), f"falta el archivo de muestra: {ORIGINAL}"
    tmp = Path(tempfile.mkdtemp(prefix="fb_"))
    con_retorno = tmp / "beetrack-14039-0037-20261117.xlsx"
    fixture_con_retorno(ORIGINAL, con_retorno)

    base, _ = leer(ORIGINAL)
    cab = [str(v or "").strip() for v in base[0]]
    iO, iU, iCli = cab.index("Origen"), cab.index("MODO"), cab.index("Cliente")
    iGuia = 0
    esperados_origen = sum(1 for f in base[1:] if f[iO] and "easy" in str(f[iO]).lower()
                           and str(f[iO]) != "CD Easy")
    print(f"muestra: {len(base) - 1} filas, {len(cab)} columnas, "
          f"{esperados_origen} origenes de Easy por corregir")

    errores = []
    with sync_playwright() as p:
        # Edge y no chromium: este equipo no tiene bajado el chromium de playwright.
        nav = p.chromium.launch(channel="msedge")
        page = nav.new_page(accept_downloads=True)
        consola = []
        page.on("console", lambda m: consola.append(m.text) if m.type == "error" else None)

        # --- Caso 1: el archivo real, sin ningun retiro de Paris -------------------
        salida = descargar(page, ORIGINAL, tmp / "real-ok.xlsx")
        filas, meta = leer(salida)
        if len(filas) != len(base):
            errores.append(f"cambio la cantidad de filas: {len(filas)} vs {len(base)}")
        if [str(v or "") for v in filas[0]] != [str(v or "") for v in base[0]]:
            errores.append("cambio la cabecera")
        for i, f in enumerate(filas[1:], start=1):
            o = f[iO]
            if o and "easy" in str(o).lower() and o != "CD Easy":
                errores.append(f"fila {i + 1}: Origen quedo en {o!r}")
            if str(base[i][iO] or "").lower().startswith("cd paris") and o != base[i][iO]:
                errores.append(f"fila {i + 1}: toco un Origen de Paris ({o!r})")
            esperado = "Solo recogida" if str(f[iGuia]) == GUIA_REAL else None
            if (f[iU] or None) != esperado:
                errores.append(f"fila {i + 1} ({f[iGuia]}): MODO quedo en {f[iU]!r} "
                               f"y se esperaba {esperado!r}")
        notexto = [c for c, t in meta["tipos"].items() if t != "str"]
        if notexto:
            errores.append(f"celdas que no son texto: {notexto[:6]}")
        nofmt = [c for c, v in meta["fmt"].items() if v != "@"]
        if nofmt:
            errores.append(f"celdas sin formato Texto: {nofmt[:6]}")
        nocol = [k for k, v in meta["cols"].items() if v != "@"]
        if nocol or len(meta["cols"]) != len(cab):
            errores.append(f"columnas sin formato Texto: {nocol or len(meta['cols'])}")
        # El resto de los datos, intacto (comparando como texto)
        for i, f in enumerate(filas):
            for j, v in enumerate(f):
                if j in (iO, iU):
                    continue
                a, b = str(base[i][j] or "").strip(), str(v or "").strip()
                if a != b and not (a.replace(".0", "") == b):
                    errores.append(f"fila {i + 1} col {j + 1}: {a!r} -> {b!r}")
        print(f"caso 1 (archivo real): {len(errores)} problemas")

        # --- Caso 2: una orden de Paris marcada RETORNO ----------------------------
        n1 = len(errores)
        salida2 = descargar(page, con_retorno, tmp / "retorno-ok.xlsx")
        filas2, meta2 = leer(salida2)
        con_modo = sorted((str(f[iGuia]), f[iU]) for f in filas2[1:] if f[iU] not in (None, ""))
        esperado2 = sorted([(GUIA_REAL, "Solo recogida"), (GUIA_RETIRO, "Solo recogida")])
        if con_modo != esperado2:
            errores.append(f"con RETORNO explicito se esperaba {esperado2} y salio {con_modo}")
        if [c for c, t in meta2["tipos"].items() if t != "str"]:
            errores.append("caso 2: hay celdas que no son texto")
        print(f"caso 2 (retiro de Paris): {len(errores) - n1} problemas")

        # --- Caso 3: el mismo archivo dos veces entra una sola vez -----------------
        n2 = len(errores)
        page.goto(PAGINA)
        page.wait_for_function("() => window.XLSX && window.JSZip")
        page.set_input_files("#input", str(ORIGINAL))
        page.wait_for_selector(".archivo")
        page.set_input_files("#input", str(ORIGINAL))
        page.wait_for_selector("#aviso:not(:empty)")
        tarjetas = page.locator(".archivo").count()
        if tarjetas != 1:
            errores.append(f"el archivo repetido se agrego igual: {tarjetas} tarjetas")

        # --- Caso 4: destildar el retiro le saca el MODO ---------------------------
        # dispatch_event y no uncheck(): la descarga anterior deja a Edge esperando una
        # navegacion que nunca termina y uncheck() se queda colgado en esa espera.
        page.locator(".retiros input:checked").first.dispatch_event("click")
        page.wait_for_function("() => !document.querySelector('.retiros input:checked')")
        with page.expect_download() as info:
            page.click("button.principal")
        info.value.save_as(tmp / "destildado-ok.xlsx")
        filas3, _ = leer(tmp / "destildado-ok.xlsx")
        con_modo3 = [(f[iGuia], f[iU]) for f in filas3[1:] if f[iU] not in (None, "")]
        if con_modo3:
            errores.append(f"se destildo el retiro y el MODO quedo igual: {con_modo3}")
        print(f"casos 3 y 4 (duplicado y destildar): {len(errores) - n2} problemas")

        if consola:
            errores.append(f"errores de consola: {consola[:3]}")
        nav.close()

    shutil.rmtree(tmp, ignore_errors=True)
    if errores:
        print("\nFALLO:")
        for e in errores[:25]:
            print("  -", e)
        return 1
    print("\nOK: las tres reglas se cumplen y no se movio nada mas.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

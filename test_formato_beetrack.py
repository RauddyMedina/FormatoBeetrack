"""Prueba de punta a punta de formato_beetrack/index.html.

Abre la pagina en un navegador real, le entrega un importable de hela, aprieta
"Descargar corregido" y revisa el .xlsx que sale:

  1. Origen: todo lo de Easy queda 'CD Easy' y Paris queda intacto.
  2. MODO: 'Solo recogida' SOLO en los retiros de Paris (y en ninguna otra fila).
  3. Todo el archivo en formato texto (valores str y number_format '@').
  4. Nada mas cambio: mismas filas, mismas columnas, mismos datos.

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
GUIA_RETIRO = "87724341"          # fila de Paris que el fixture convierte en retiro
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
            if f[iU] not in (None, ""):
                errores.append(f"fila {i + 1}: MODO quedo en {f[iU]!r} y no hay retiros")
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
        con_modo = [(i + 1, f[iGuia], f[iU]) for i, f in enumerate(filas2[1:], start=1)
                    if f[iU] not in (None, "")]
        if len(con_modo) != 1:
            errores.append(f"se esperaba 1 fila con MODO y hay {len(con_modo)}: {con_modo}")
        else:
            fila, guia, modo = con_modo[0]
            if modo != "Solo recogida":
                errores.append(f"MODO quedo en {modo!r}")
            if str(guia) != GUIA_RETIRO:
                errores.append(f"el MODO cayo en la guia {guia!r} y no en {GUIA_RETIRO}")
        if [c for c, t in meta2["tipos"].items() if t != "str"]:
            errores.append("caso 2: hay celdas que no son texto")
        print(f"caso 2 (retiro de Paris): {len(errores) - n1} problemas")

        # --- Caso 3: el numero pegado a mano marca el retiro -----------------------
        n2 = len(errores)
        page.goto(PAGINA)
        page.wait_for_function("() => window.XLSX && window.JSZip")
        page.evaluate("document.getElementById('manual').open = true")
        page.fill("#manualTxt", "87723382")
        page.set_input_files("#input", str(ORIGINAL))
        page.wait_for_selector(".archivo")
        with page.expect_download() as info:
            page.click("button.principal")
        info.value.save_as(tmp / "manual-ok.xlsx")
        filas3, _ = leer(tmp / "manual-ok.xlsx")
        con_modo3 = [(f[iGuia], f[iU]) for f in filas3[1:] if f[iU] not in (None, "")]
        if con_modo3 != [("87723382", "Solo recogida")]:
            errores.append(f"a mano: se esperaba solo 87723382 y salio {con_modo3}")
        print(f"caso 3 (numero a mano): {len(errores) - n2} problemas")

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

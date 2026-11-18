# Formato Beetrack

Arregla el importable de Beetrack que baja de hela (`beetrack-14039-0037-20261117.xlsx`)
antes de subirlo. Se abre en el navegador, no sube nada a ningún servidor:

**https://rauddymedina.github.io/FormatoBeetrack/**

## Qué le hace al archivo

1. **Origen → `CD Easy`.** Todo origen de Easy (`CD Easy Tienda`, `CD Easy Opl`,
   `CD Easy Cd`) queda parejo en `CD Easy`. Los `CD Paris` no se tocan.
2. **MODO → `Solo recogida`,** sólo en los retiros de Paris. Se reconocen por la marca
   `RETORNO` en el código del ítem —que es como Paris marca sus retornos— o porque el
   nombre del producto los delata. Los que no, se marcan con el tilde de la lista o
   pegando el número de guía.
3. **Todo en formato texto.** Las celdas salen como Texto y las 38 columnas quedan con
   formato Texto, incluidas CANTIDAD y Bultos, que hela manda como número.

## Cómo está hecho

Una sola página sin build. El `.xlsx` se reescribe **celda por celda sobre el archivo
original** (JSZip sobre el XML de la hoja), no se regenera con un escritor de Excel: así
las 38 columnas, sus nombres con tilde y el resto de los datos salen tal cual los mandó
hela. SheetJS se usa sólo para leer.

La columna `MODO` viene sin celda en el archivo de hela, así que hay que crearla dentro
de la fila, en orden, cuando corresponde marcarla.

## Prueba

`test_formato_beetrack.py` abre la página en un navegador real, le entrega un importable,
aprieta Descargar y revisa el archivo que sale: las tres reglas, y que no se haya movido
ninguna otra celda. Necesita `playwright` + `openpyxl` y un `beetrack-*.xlsx` de muestra.

```
python test_formato_beetrack.py
```

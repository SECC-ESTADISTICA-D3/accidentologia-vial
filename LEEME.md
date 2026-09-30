# Accidentología Vial — Provincia de Tucumán

Sitio público de estadísticas de siniestros viales.
Policía de Tucumán · Departamento Operaciones Policiales D-3 · Sección Estadística y Archivo.

## Cómo se actualiza (una vez por mes)

1. Abrí `preparar-base.html` con doble clic y arrastrale la **base completa** (el Excel con todos los datos).
2. Te descarga **dos archivos**, ya depurados (conservan las 23 columnas estadísticas y descartan todos los datos personales):
   - `accidentologia.json`: versión compacta que **lee el sitio** (pesa unas 33 veces menos que el Excel; por eso abre rápido, sobre todo en celulares).
   - `accidentologia.xlsx`: respaldo, por si el `.json` faltara.
3. Copiá **los dos** dentro de la carpeta `datos/`, pisando los anteriores.
4. Si publicás en GitHub Pages: commit y push. En dos o tres minutos el sitio muestra los datos nuevos.

No hay que tocar el HTML ni ejecutar ningún script. Todo el proceso de `preparar-base.html` ocurre en tu equipo: no sube nada a internet.

> **Ojo:** si copiás solo el `.xlsx` y dejás el `.json` viejo, el sitio va a seguir mostrando los datos anteriores
> (lee primero el `.json`). Para detectarlo, al pie de la página figura **"Registros hasta el ..."** con la fecha del
> último registro cargado: verificá que coincida con lo esperado después de cada actualización.

Para probarlo en tu equipo sin servidor, abrí `index.html` con doble clic: como el navegador no puede leer archivos
locales por seguridad, la página te va a pedir que elijas el archivo a mano (sirve el `.json` o el Excel) y funciona igual.

### Comparativo "2026 vs 2025" y mes en curso

El año más reciente de la base suele estar incompleto (la planilla se actualiza mes a mes). Por eso las flechas de
variación de los indicadores comparan **los mismos meses** de ambos años, y lo dicen en cada tarjeta (por ejemplo,
"ene–ago 2026 vs 2025"). En el gráfico "Siniestros por mes" los meses que la base todavía no tiene quedan vacíos (no
se dibujan como cero).

Además, el último mes con datos del año en curso se deja afuera de la comparación y se dibuja punteado, porque suele
estar a medio cargar. Si la base se carga siempre con meses completos, se puede desactivar poniendo
`EXCLUIR_MES_EN_CURSO = false` al principio de `app.js`.

### Qué tiene que respetar el Excel

- La base tiene que estar en la **primera hoja** del libro (hoy es `2024-2025`).
- La **fila 1 son los encabezados**, y los datos arrancan en la fila 2.
- Los nombres de las columnas que se usan tienen que mantenerse. La lectura ignora
  mayúsculas, acentos, espacios y puntos, así que `TIPO DE VIA`, `Tipo de vía` y
  `TIPO DE VÍA` funcionan igual.
- Se pueden agregar años nuevos: la página detecta sola los años presentes y arma la
  comparación entre ellos.
- Estas condiciones las controla `preparar-base.html`; el `.json` que genera es el que usa el sitio.

Si alguna vez se renombra una columna, hay que actualizar la lista `COLUMNAS` que está al
principio de `app.js`.

## Qué datos se publican

La página lee **únicamente** estas columnas:

AÑO · MES · FECHA · ZONA HORARIA · DEPARTAMENTO · LOCALIDAD · U.U.R.R. · DEPENDENCIA · ZONA ·
TIPO DE VIA · CAUSA · CATEGORIA DE SINIESTRO · TIPO DE SINIESTRO · PARTICIPANTE 12 ·
PARTICIPANTE 23 · SEXO · RANGO ETARIO · CONDICION DE LA VICTIMA · ILESO ·
HERIDOS LEVES · HERIDOS GRAVES · FALLECIDOS EN EL LUGAR · FALLECIDOS LUEGO

Todo lo demás queda fuera y nunca llega al navegador: número de sumario, nombres de
víctimas y causantes, edad exacta, diagnóstico, hospitalización, dominio, marca, modelo,
color del vehículo, seguro, carnet y observaciones.

> **Importante:** el archivo que se publica en `datos/` se puede descargar desde el sitio.
> Por eso nunca hay que subir ahí la base completa, sino la copia depurada que genera
> `preparar-base.html`. El archivo que viene en esta carpeta ya está depurado.

## El mapa

La sección "Mapa de siniestros por departamento y localidad" no lee ningún archivo
externo: los límites de los 17 departamentos y las coordenadas de 19 localidades
(las cabeceras de departamento y algunas ciudades más) están incluidos dentro de
`app.js`, simplificados a partir de datos públicos del IGN y de BAHRA. Si en el futuro
hace falta agregar más localidades al mapa, hay que sumarles coordenadas al objeto
`MAPA_DATOS.localidades` dentro de `app.js`.

Cuando en el filtro **Carátula de la causa** quedan seleccionadas únicamente carátulas
que implican muerte (FALLECIMIENTO, HOMICIDIO CULPOSO, LESIONES A FALLECIMIENTO,
LESIONES CULPOSAS A HOMICIDIO CULPOSO), el mapa entra en modo "víctimas fatales": además
de cambiar a la paleta amarillo→rojo, muestra **dos columnas** en las listas de
departamentos, de localidades y en el detalle de cada departamento. La primera es la
cantidad de siniestros y la segunda, en rojo, la cantidad de **personas fallecidas**
(FALLECIDOS EN EL LUGAR + FALLECIDOS LUEGO de la fila cabecera del hecho). Los globos
de ayuda del mapa muestran los dos datos. La lista de carátulas fatales está en la
constante `CAUSAS_FATALES` de `app.js`.

## Cómo cuenta la página

- **Siniestros**: filas que tienen cargada la columna CAUSA, es decir la fila cabecera de
  cada hecho. Es el mismo criterio que usan las tablas dinámicas del Excel ("cuenta de CAUSA").
- **Personas involucradas**: suma de ILESO + HERIDOS LEVES + HERIDOS GRAVES + FALLECIDOS
  EN EL LUGAR + FALLECIDOS LUEGO.
- **Gráficos de víctimas** (sexo, rango etario, condición): cuentan registros de víctimas,
  es decir cada fila con datos de la persona damnificada.

## Etiquetas de los gráficos de barra

Los gráficos dejan un margen extra (arriba en los verticales, a la derecha en los
horizontales) para que el número que acompaña a la barra más alta/larga no quede
cortado por el borde del gráfico. Si en el futuro los valores crecen mucho (muchos
más dígitos), se puede agrandar ese margen en `pintarFicha()` (app.js), en las
propiedades `layout.padding` y `grace` de cada eje.

Al seleccionar el departamento **Capital**, como en la práctica es una sola localidad
(San Miguel de Tucumán), la tarjeta muestra en su lugar el desglose por **jurisdicción
policial** (campo `dependencia` de la base). El resto de los departamentos sigue
mostrando localidades como siempre. Ese comportamiento está en `pintarMapa()` (app.js),
en la constante `DEPTO_JURISDICCION = 'CAPITAL'`: si en algún momento se quisiera aplicar
lo mismo a otro departamento, o revertirlo, se cambia ahí.

## Mapa interactivo

Al tocar o hacer clic en un departamento del mapa se despliega una tarjeta con el listado
de sus localidades y la cantidad de siniestros de cada una (ordenadas de mayor a menor).
Se cierra tocando el mismo departamento de nuevo, el botón ✕, la tecla Escape, o tocando
un área vacía del mapa. Esa relación depto→localidad sale de los mismos registros
filtrados (no depende de que la localidad tenga coordenadas en el mapa), así que aparecen
todas las localidades cargadas para ese departamento, tengan o no burbuja dibujada.

## Botón "Presentación"

Junto a "Descargar PDF" hay un botón "Presentación" que abre una ventana con los datos
institucionales del proyecto (supervisión, ejecución y operadores de datos). Su contenido
está escrito directamente en `index.html`, dentro del bloque `<div class="modal-fondo"
id="modal-presentacion">`. Para actualizar nombres o cargos, editá ese bloque a mano.

## Rendimiento

- **Datos compactos:** el sitio lee `datos/accidentologia.json` (unos 150 KB por la red) en vez de interpretar el Excel
  (5 MB) en el navegador. La descarga de los datos arranca desde el `<head>`, en paralelo con el resto de la página.
- **Sin dependencias externas en el uso normal:** Chart.js y su plugin de etiquetas están en `assets/vendor/` y las
  tipografías en `assets/fonts/`. El lector de Excel (SheetJS) solo se descarga desde internet si hace falta el
  respaldo `.xlsx` o si se elige un Excel a mano.
- **Gráficos por cuadro:** se dibujan de a uno, así los indicadores aparecen primero y los filtros responden sin trabarse.
- Los logos institucionales en `assets/` están recortados al tamaño real en que se muestran; conviene
  redimensionar antes cualquier imagen nueva.
- Si el servidor no comprime archivos (GitHub Pages sí lo hace), el `.json` pesa 1,1 MB en vez de 150 KB: activar la
  compresión gzip para `.json`, `.js` y `.css`.

## Publicación

Subí el contenido de esta carpeta a la raíz del repositorio y activá GitHub Pages.
La estructura tiene que quedar así:

```
index.html
styles.css
app.js
assets/            (logos, fonts/ y vendor/)
datos/accidentologia.json
datos/accidentologia.xlsx
```

`preparar-base.html` y `LEEME.md` son herramientas internas: podés subirlos igual, o dejarlos
solamente en tu equipo.

Para probarlo en tu equipo sin servidor, abrí `index.html` con doble clic: como el
navegador no puede leer archivos locales por seguridad, la página te va a pedir que
elijas el Excel a mano y funciona igual.

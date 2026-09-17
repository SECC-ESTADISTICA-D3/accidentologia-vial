# Accidentología Vial — Provincia de Tucumán

Sitio público de estadísticas de siniestros viales.
Policía de Tucumán · Departamento Operaciones Policiales D-3 · Sección Estadística y Archivo.

## Cómo se actualiza

**Reemplazá un solo archivo:**

```
datos/accidentologia.xlsx
```

Guardá ahí la base actualizada con ese mismo nombre, pisando la anterior. No hay que tocar
el HTML ni ejecutar ningún script: la página lee el Excel cada vez que alguien la abre y
recalcula todos los indicadores y gráficos.

**Antes de copiarlo, pasá la base por `preparar-base.html`.** Abrí ese archivo con doble clic,
arrastrale la base completa y te descarga `accidentologia.xlsx` ya depurado: conserva las 22
columnas estadísticas y descarta todos los datos personales. Ese es el archivo que va en
`datos/`. Trabaja de forma local, no sube nada a internet.

Si publicás en GitHub Pages, el flujo completo es: reemplazar el archivo en la carpeta
`datos/`, hacer commit y push. En dos o tres minutos el sitio ya muestra los datos nuevos.

### Qué tiene que respetar el Excel

- La base tiene que estar en la **primera hoja** del libro (hoy es `2024-2025`).
- La **fila 1 son los encabezados**, y los datos arrancan en la fila 2.
- Los nombres de las columnas que se usan tienen que mantenerse. La lectura ignora
  mayúsculas, acentos, espacios y puntos, así que `TIPO DE VIA`, `Tipo de vía` y
  `TIPO DE VÍA` funcionan igual.
- Se pueden agregar años nuevos: la página detecta sola los años presentes y arma la
  comparación entre ellos.

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

## Cómo cuenta la página

- **Siniestros**: filas que tienen cargada la columna CAUSA, es decir la fila cabecera de
  cada hecho. Es el mismo criterio que usan las tablas dinámicas del Excel ("cuenta de CAUSA").
- **Personas involucradas**: suma de ILESO + HERIDOS LEVES + HERIDOS GRAVES + FALLECIDOS
  EN EL LUGAR + FALLECIDOS LUEGO.
- **Gráficos de víctimas** (sexo, rango etario, condición): cuentan registros de víctimas,
  es decir cada fila con datos de la persona damnificada.

## Botón "Presentación"

Junto a "Descargar PDF" hay un botón "Presentación" que abre una ventana con los datos
institucionales del proyecto (supervisión, ejecución y operadores de datos). Su contenido
está escrito directamente en `index.html`, dentro del bloque `<div class="modal-fondo"
id="modal-presentacion">`. Para actualizar nombres o cargos, editá ese bloque a mano.

## Rendimiento

Los logos institucionales en `assets/` están recortados al tamaño real en que se muestran
(no hace falta que sean más grandes: aunque se peguen fotos de mayor resolución ahí, conviene
redimensionarlas antes, porque el navegador igual las va a mostrar del mismo tamaño chico).
Los scripts externos (`xlsx`, `chart.js`, `chartjs-plugin-datalabels`) y `app.js` se cargan
con el atributo `defer` para no bloquear el primer dibujo de la página mientras se descargan.

## Publicación

Subí el contenido de esta carpeta a la raíz del repositorio y activá GitHub Pages.
La estructura tiene que quedar así:

```
index.html
app.js
assets/logo-d3.png
datos/accidentologia.xlsx
```

`preparar-base.html` y `LEEME.md` son herramientas internas: podés subirlos igual, o dejarlos
solamente en tu equipo.

Para probarlo en tu equipo sin servidor, abrí `index.html` con doble clic: como el
navegador no puede leer archivos locales por seguridad, la página te va a pedir que
elijas el Excel a mano y funciona igual.

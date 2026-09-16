/* =============================================================
   Accidentología Vial — Provincia de Tucumán
   Policía de Tucumán · D-3 · Sección Estadística y Archivo

   La página se actualiza reemplazando el archivo Excel ubicado en
   datos/accidentologia.xlsx  (ver RUTAS_DATOS)
   ============================================================= */

const RUTAS_DATOS = [
  'datos/accidentologia.xlsx',
  'datos/accidentologia.xlsm'
];

const HOJA_BASE = null; // null = primera hoja del libro

/* ---------- columnas que se leen del Excel ----------
   Todo lo que no figure acá NO se carga en la página:
   nombres, edades exactas, dominios, marcas, modelos, colores,
   diagnósticos, sumarios y observaciones quedan fuera.          */
const COLUMNAS = {
  anio:        ['AÑO'],
  mes:         ['MES'],
  fecha:       ['FECHA'],
  zonaHoraria: ['ZONA HORARIA'],
  depto:       ['DEPARTAMENTO'],
  uurr:        ['U.U.R.R.', 'UURR'],
  dependencia: ['DEPENDENCIA'],
  zona:        ['ZONA'],
  tipoVia:     ['TIPO DE VIA'],
  causa:       ['CAUSA'],
  catSin:      ['CATEGORIA DE SINIESTRO'],
  tipoSin:     ['TIPO DE SINIESTRO'],
  p12:         ['PARTICIPANTE 12'],
  p23:         ['PARTICIPANTE 23'],
  sexo:        ['SEXO'],
  rango:       ['RANGO ETARIO'],
  condicion:   ['CONDICION DE LA VICTIMA'],
  ileso:       ['ILESO'],
  leves:       ['HERIDOS LEVES'],
  graves:      ['HERIDOS GRAVES'],
  fallLugar:   ['FALLECIDOS EN EL LUGAR ( 24 HS.)', 'FALLECIDOS EN EL LUGAR'],
  fallLuego:   ['FALLECIDOS LUEGO (+ DE 24 HS.)', 'FALLECIDOS LUEGO']
};

const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
const DIAS  = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
const ORDEN_DIA = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
const ORDEN_RANGO = ['0 a 4','5 a 14','15 a 24','25 a 34','35 a 44','45 a 54','55 a 64','65 a 74','75 a mas','75 a más','s/dato','sin datos'];

/* correcciones de grafía para las etiquetas visibles */
const DICC = {
  'capital':'Capital','este':'Este','oeste':'Oeste','norte':'Norte','sur':'Sur',
  'cruz alta':'Cruz Alta','burruyacu':'Burruyacú','chicligasta':'Chicligasta','famailla':'Famaillá',
  'graneros':'Graneros','juan bautista alberdi':'Juan Bautista Alberdi','la cocha':'La Cocha',
  'leales':'Leales','lules':'Lules','monteros':'Monteros','rio chico':'Río Chico','simoca':'Simoca',
  'tafi del valle':'Tafí del Valle','tafi viejo':'Tafí Viejo','trancas':'Trancas','yerba buena':'Yerba Buena',
  'colision':'Colisión','caida':'Caída','atropello peaton':'Atropello a peatón','atropello animal':'Atropello a animal',
  'choque contra objeto fijo':'Choque contra objeto fijo','despenamiento':'Despeñamiento','despiste':'Despiste',
  'vuelco':'Vuelco','multiple':'Múltiple','incendio':'Incendio','otro':'Otro','otros':'Otros',
  'peaton':'Peatón','conductor':'Conductor','acompanante':'Acompañante','pasajero':'Pasajero',
  'traccion animal':'Tracción animal','automovil':'Automóvil','camion':'Camión',
  'camion con remolque':'Camión con remolque','utilitario - pick up':'Utilitario / Pick-up',
  'transporte pasajeros mas de 8 asientos urbano':'Transporte de pasajeros (+8 asientos)',
  'rastra canera':'Rastra cañera','tractor con semiremolque':'Tractor con semirremolque',
  'urbano':'Urbano','rural':'Rural','diurno':'Diurno','nocturno':'Nocturno',
  'masculino':'Masculino','femenino':'Femenino','s/dato':'Sin datos','sin datos':'Sin datos',
  'con lesionados':'Con lesionados','con fallecidos':'Con fallecidos','75 a mas':'75 y más'
};

/* ---------- utilidades ---------- */
const sinAcento = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');
const clave     = s => sinAcento(String(s||'')).toUpperCase().replace(/[^A-Z0-9]/g,'');
const limpio    = s => String(s==null?'':s).replace(/\s+/g,' ').trim();
const nro       = n => new Intl.NumberFormat('es-AR').format(n||0);
const esc       = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

function lindo(txt){
  const t = limpio(txt);
  if(!t) return '';
  const k = sinAcento(t).toLowerCase();
  if(DICC[k]) return DICC[k];
  const tieneMin = /[a-záéíóúñ]/.test(t), tieneMay = /[A-ZÁÉÍÓÚÑ]/.test(t);
  if(tieneMin && tieneMay) return t;               // ya viene con grafía propia
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase();
}

/* =============================================================
   1. Carga del Excel
   ============================================================= */
const elCarga = document.getElementById('carga');

async function iniciar(){
  for(const ruta of RUTAS_DATOS){
    try{
      const r = await fetch(ruta, {cache:'no-cache'});
      if(!r.ok) continue;
      const buf = await r.arrayBuffer();
      if(buf.byteLength < 1000) continue;
      procesarLibro(buf, ruta);
      return;
    }catch(e){ /* sigue con la ruta siguiente */ }
  }
  pedirArchivo('No se encontró el archivo de datos en la carpeta del sitio.');
}

function pedirArchivo(motivo){
  document.getElementById('carga-t').textContent = 'Cargá el archivo de datos';
  document.getElementById('carga-p').textContent = 'Para ver el tablero hace falta la base de siniestros en formato Excel.';
  document.getElementById('carga-barra').hidden = true;
  document.getElementById('carga-alta').hidden = false;
  document.getElementById('carga-motivo').textContent = motivo || '';
}

document.getElementById('btn-archivo').addEventListener('click', ()=> document.getElementById('archivo').click());
document.getElementById('archivo').addEventListener('change', e=>{
  if(e.target.files[0]) leerArchivoLocal(e.target.files[0]);
});
const zonaSoltar = document.getElementById('soltar');
['dragenter','dragover'].forEach(ev=> zonaSoltar.addEventListener(ev, e=>{e.preventDefault(); zonaSoltar.classList.add('sobre');}));
['dragleave','drop'].forEach(ev=> zonaSoltar.addEventListener(ev, e=>{e.preventDefault(); zonaSoltar.classList.remove('sobre');}));
zonaSoltar.addEventListener('drop', e=>{
  const f = e.dataTransfer.files[0];
  if(f) leerArchivoLocal(f);
});
function leerArchivoLocal(file){
  document.getElementById('carga-t').textContent = 'Leyendo el archivo';
  document.getElementById('carga-alta').hidden = true;
  document.getElementById('carga-barra').hidden = false;
  const fr = new FileReader();
  fr.onload = () => procesarLibro(fr.result, file.name);
  fr.onerror = () => pedirArchivo('No se pudo leer el archivo.');
  fr.readAsArrayBuffer(file);
}

let DATOS = [], ANIOS = [], ACTUALIZADO = '';

function procesarLibro(buf, origen){
  setTimeout(()=>{
    try{
      const wb = XLSX.read(buf, {type:'array', cellDates:true, cellStyles:false});
      const hoja = HOJA_BASE && wb.Sheets[HOJA_BASE] ? HOJA_BASE : wb.SheetNames[0];
      const filas = XLSX.utils.sheet_to_json(wb.Sheets[hoja], {header:1, raw:true, defval:null, blankrows:false});
      armarDatos(filas);
      document.getElementById('pie-act').innerHTML = '<br>Base cargada desde <b>' + origen + '</b>.';
      construirInterfaz();
      elCarga.remove();
    }catch(err){
      console.error(err);
      pedirArchivo('El archivo no pudo interpretarse: ' + err.message);
    }
  }, 30);
}

function armarDatos(filas){
  if(!filas.length) throw new Error('la hoja está vacía');
  const enc = filas[0].map(h => clave(h));
  const idx = {};
  for(const campo in COLUMNAS){
    idx[campo] = -1;
    for(const alias of COLUMNAS[campo]){
      const i = enc.indexOf(clave(alias));
      if(i >= 0){ idx[campo] = i; break; }
    }
  }
  if(idx.anio < 0 || idx.causa < 0)
    throw new Error('no se encontraron las columnas AÑO y CAUSA en la primera hoja');

  const num = v => { const n = Number(v); return isFinite(n) ? n : 0; };
  const val = (f,i) => i < 0 ? '' : limpio(f[i]);

  DATOS = [];
  for(let r = 1; r < filas.length; r++){
    const f = filas[r];
    if(!f || f.every(c => c === null || c === '')) continue;
    const anio = num(f[idx.anio]);
    if(!anio) continue;

    const causa = val(f, idx.causa);
    let mes = val(f, idx.mes).toLowerCase(), dia = '';
    const fch = idx.fecha >= 0 ? f[idx.fecha] : null;
    if(fch instanceof Date && !isNaN(fch)){
      if(!mes) mes = MESES[fch.getMonth()];
      dia = DIAS[fch.getDay()];
    }

    DATOS.push({
      anio: String(anio),
      mes,
      dia,
      zonaHoraria: lindo(val(f, idx.zonaHoraria)),
      depto:       lindo(val(f, idx.depto)),
      uurr:        lindo(val(f, idx.uurr)),
      dependencia: lindo(val(f, idx.dependencia)),
      zona:        lindo(val(f, idx.zona)),
      tipoVia:     lindo(val(f, idx.tipoVia)),
      causa:       lindo(causa),
      catSin:      lindo(val(f, idx.catSin)),
      tipoSin:     lindo(val(f, idx.tipoSin)),
      p12:         lindo(val(f, idx.p12)),
      p23:         lindo(val(f, idx.p23)),
      sexo:        lindo(val(f, idx.sexo)),
      rango:       lindo(val(f, idx.rango)),
      condicion:   lindo(val(f, idx.condicion)),
      hecho:       causa !== '',            // fila cabecera del siniestro
      il:  num(f[idx.ileso]),
      hl:  num(f[idx.leves]),
      hg:  num(f[idx.graves]),
      fl:  num(f[idx.fallLugar]),
      fp:  num(f[idx.fallLuego])
    });
  }
  if(!DATOS.length) throw new Error('no se encontraron registros con año válido');
  ANIOS = [...new Set(DATOS.map(d => d.anio))].sort();
}

/* =============================================================
   2. Filtros
   ============================================================= */
const SEGMENTADORES = [
  {campo:'anio',        etiqueta:'Año',                     orden:'natural'},
  {campo:'mes',         etiqueta:'Mes',                     orden:'mes'},
  {campo:'uurr',        etiqueta:'Unidad Regional',         orden:'cantidad'},
  {campo:'depto',       etiqueta:'Departamento',            orden:'alfa'},
  {campo:'dependencia', etiqueta:'Jurisdicción policial',   orden:'cantidad'},
  {campo:'zona',        etiqueta:'Zona urbana o rural',     orden:'cantidad'},
  {campo:'zonaHoraria', etiqueta:'Franja horaria',          orden:'cantidad'},
  {campo:'tipoVia',     etiqueta:'Tipo de vía',             orden:'cantidad'},
  {campo:'tipoSin',     etiqueta:'Tipo de siniestro',       orden:'cantidad'},
  {campo:'causa',       etiqueta:'Carátula de la causa',    orden:'cantidad'},
  {campo:'sexo',        etiqueta:'Sexo de la víctima',      orden:'cantidad'},
  {campo:'p12',         etiqueta:'Movilidad de la víctima', orden:'cantidad'},
  {campo:'p23',         etiqueta:'Movilidad del causante',  orden:'cantidad'}
];

const seleccion = {};            // campo -> Set de valores
SEGMENTADORES.forEach(s => seleccion[s.campo] = new Set());

function filtrar(){
  const activos = SEGMENTADORES.filter(s => seleccion[s.campo].size);
  if(!activos.length) return DATOS;
  return DATOS.filter(d => activos.every(s => seleccion[s.campo].has(d[s.campo])));
}

function opcionesDe(campo){
  // cuenta con el resto de los filtros aplicados (lógica Y entre categorías)
  const otros = SEGMENTADORES.filter(s => s.campo !== campo && seleccion[s.campo].size);
  const mapa = new Map();
  for(const d of DATOS){
    if(!otros.every(s => seleccion[s.campo].has(d[s.campo]))) continue;
    const v = d[campo];
    if(!v) continue;
    mapa.set(v, (mapa.get(v)||0) + 1);
  }
  return mapa;
}

function ordenar(claves, modo, mapa){
  const arr = [...claves];
  if(modo === 'mes')      return arr.sort((a,b)=> MESES.indexOf(a) - MESES.indexOf(b));
  if(modo === 'natural')  return arr.sort();
  if(modo === 'alfa')     return arr.sort((a,b)=> a.localeCompare(b,'es'));
  if(modo === 'rango')    return arr.sort((a,b)=> idxRango(a) - idxRango(b));
  return arr.sort((a,b)=> (mapa.get(b)||0) - (mapa.get(a)||0) || a.localeCompare(b,'es'));
}
const idxRango = v => {
  const i = ORDEN_RANGO.indexOf(sinAcento(String(v)).toLowerCase());
  return i < 0 ? 99 : i;
};

/* =============================================================
   3. Interfaz
   ============================================================= */
function construirInterfaz(){
  dibujarPanel();
  document.getElementById('secciones').innerHTML = SECCIONES.map(sec => `
    <section class="seccion">
      <h2>${sec.titulo}</h2>
      <p>${sec.bajada}</p>
      <div class="rejilla">
        ${sec.fichas.map(f => `
          <article class="tarjeta${f.ancha ? ' ancha' : ''}">
            <h3>${f.titulo}</h3>
            <p class="sub">${f.sub}</p>
            <div class="lienzo ${f.alto || ''}"><canvas id="cv-${f.id}"></canvas></div>
          </article>`).join('')}
      </div>
    </section>`).join('');

  document.getElementById('limpiar').addEventListener('click', ()=>{
    SEGMENTADORES.forEach(s => seleccion[s.campo].clear());
    actualizar();
  });
  const rail = document.getElementById('rail'), btn = document.getElementById('abrir-filtros');
  btn.addEventListener('click', ()=>{
    const abierto = rail.classList.toggle('abierto');
    btn.setAttribute('aria-expanded', abierto);
    document.body.style.overflow = abierto ? 'hidden' : '';
  });
  rail.addEventListener('click', e=>{
    if(e.target === rail){ rail.classList.remove('abierto'); btn.setAttribute('aria-expanded','false'); document.body.style.overflow=''; }
  });
  document.addEventListener('keydown', e=>{
    if(e.key === 'Escape' && rail.classList.contains('abierto')) rail.click();
  });

  document.getElementById('periodo-txt').textContent =
    ANIOS.length > 1 ? `Período ${ANIOS[0]} – ${ANIOS[ANIOS.length-1]}` : `Período ${ANIOS[0]}`;

  actualizar();
}

function dibujarPanel(){
  document.getElementById('panel-cuerpo').innerHTML = SEGMENTADORES.map(s => `
    <details class="grupo" data-campo="${s.campo}">
      <summary>${s.etiqueta}<span class="fl"></span><span class="cv">⌄</span></summary>
      <div class="opciones"></div>
    </details>`).join('');
}

function refrescarPanel(){
  SEGMENTADORES.forEach(s => {
    const det  = document.querySelector(`.grupo[data-campo="${s.campo}"]`);
    const caja = det.querySelector('.opciones');
    const mapa = opcionesDe(s.campo);
    const sel  = seleccion[s.campo];
    sel.forEach(v => { if(!mapa.has(v)) mapa.set(v, 0); });
    const claves = ordenar(mapa.keys(), s.orden, mapa);

    caja.innerHTML = claves.map(v => `
      <label class="op">
        <input type="checkbox" value="${esc(v)}" ${sel.has(v)?'checked':''}>
        <span class="nom" title="${esc(v)}">${esc(s.campo==='mes'?lindo(v):v)}</span>
        <span class="can">${nro(mapa.get(v))}</span>
      </label>`).join('') || '<p class="can" style="padding:4px">Sin valores para los filtros actuales.</p>';

    det.classList.toggle('activo', sel.size > 0);
    det.querySelector('.fl').textContent = sel.size;

    caja.querySelectorAll('input').forEach(inp => {
      inp.addEventListener('change', ()=>{
        inp.checked ? sel.add(inp.value) : sel.delete(inp.value);
        actualizar();
      });
    });
  });

  const total = SEGMENTADORES.reduce((a,s)=> a + seleccion[s.campo].size, 0);
  document.getElementById('n-filtros').textContent = total;
  document.getElementById('panel-cuenta').textContent = total ? `${total} activo${total>1?'s':''}` : '';

  const chips = document.getElementById('chips');
  chips.innerHTML = SEGMENTADORES.flatMap(s =>
    [...seleccion[s.campo]].map(v =>
      `<span class="chip"><b>${s.etiqueta}:</b> ${esc(s.campo==='mes'?lindo(v):v)}
       <button data-c="${s.campo}" data-v="${esc(v)}" aria-label="Quitar filtro">×</button></span>`)
  ).join('');
  chips.querySelectorAll('button').forEach(b => b.addEventListener('click', ()=>{
    seleccion[b.dataset.c].delete(b.dataset.v);
    actualizar();
  }));
}

/* =============================================================
   4. Indicadores
   ============================================================= */
function totales(filas){
  const t = {sin:0, il:0, hl:0, hg:0, fa:0, vic:0};
  for(const d of filas){
    if(d.hecho){ t.sin++; t.il+=d.il; t.hl+=d.hl; t.hg+=d.hg; t.fa+=d.fl+d.fp; }
    if(d.sexo || d.rango || d.condicion) t.vic++;
  }
  t.per = t.il + t.hl + t.hg + t.fa;
  return t;
}

function pintarCinta(t){
  const partes = [
    ['b-ileso', t.il], ['b-leve', t.hl], ['b-grave', t.hg], ['b-fallecido', t.fa]
  ];
  const tot = t.per || 1;
  document.getElementById('cinta-n').textContent = nro(t.per);
  document.getElementById('barra').innerHTML = partes
    .map(([c,v]) => `<span class="${c}" style="width:${(v/tot*100).toFixed(2)}%"></span>`).join('');
  const eti = [
    ['l-ileso','Ilesas', t.il], ['l-leve','Con lesiones leves', t.hl],
    ['l-grave','Con lesiones graves', t.hg], ['l-fallecido','Fallecidas', t.fa]
  ];
  document.getElementById('leyenda').innerHTML = eti.map(([c,k,v]) =>
    `<div class="${c}"><div class="v">${nro(v)}</div><div class="k">${k} · ${(v/tot*100).toFixed(1)}%</div></div>`).join('');
}

function pintarKpis(filas, t){
  // variación entre los dos años más recientes presentes en la selección
  const presentes = [...new Set(filas.map(d=>d.anio))].sort();
  let cmp = null;
  if(presentes.length >= 2){
    const a = presentes[presentes.length-1], b = presentes[presentes.length-2];
    const ta = totales(filas.filter(d=>d.anio===a)), tb = totales(filas.filter(d=>d.anio===b));
    cmp = {a, b, ta, tb};
  }
  const delta = (k) => {
    if(!cmp || !cmp.tb[k]) return '';
    const p = (cmp.ta[k] - cmp.tb[k]) / cmp.tb[k] * 100;
    const cls = p > 0 ? 'sube' : (p < 0 ? 'baja' : '');
    return `<div class="d ${cls}">${p>0?'▲':(p<0?'▼':'=')} ${Math.abs(p).toFixed(1)}% · ${cmp.a} vs ${cmp.b}</div>`;
  };
  const tarjetas = [
    ['sin','Siniestros registrados'], ['per','Personas involucradas'],
    ['fa','Personas fallecidas'], ['hg','Con lesiones graves'],
    ['hl','Con lesiones leves'], ['il','Ilesas']
  ];
  document.getElementById('kpis').innerHTML = tarjetas.map(([k,et]) =>
    `<div class="kpi"><div class="v">${nro(t[k])}</div><div class="k">${et}</div>${delta(k)}</div>`).join('');
}

/* =============================================================
   5. Gráficos
   ============================================================= */
const SECCIONES = [
  {titulo:'Cuándo ocurren', bajada:'Distribución de los siniestros a lo largo del año, de la semana y del día.', fichas:[
    {id:'mes',  titulo:'Siniestros por mes', sub:'Siniestros registrados en cada mes, comparados por año', tipo:'linea', dim:'mes', orden:'mes', ancha:true},
    {id:'dia',  titulo:'Siniestros por día de la semana', sub:'Según la fecha del hecho', tipo:'barra', dim:'dia', orden:'dia'},
    {id:'fhor', titulo:'Franja horaria', sub:'Siniestros ocurridos en horario diurno y nocturno', tipo:'barra', dim:'zonaHoraria'}
  ]},
  {titulo:'Dónde ocurren', bajada:'Distribución territorial según la jurisdicción policial y las características de la vía.', fichas:[
    {id:'uurr', titulo:'Unidades Regionales', sub:'Siniestros por Unidad Regional', tipo:'barra', dim:'uurr'},
    {id:'zona', titulo:'Zona urbana y rural', sub:'Siniestros según el ámbito donde ocurrieron', tipo:'barra', dim:'zona'},
    {id:'depto',titulo:'Departamentos', sub:'Siniestros por departamento de la provincia', tipo:'barraH', dim:'depto', orden:'alfa', alto:'alto'},
    {id:'via',  titulo:'Tipo de vía', sub:'Siniestros según la vía donde se produjeron', tipo:'barraH', dim:'tipoVia'},
    {id:'dep',  titulo:'Jurisdicciones policiales con más registros', sub:'Las 15 dependencias con mayor cantidad de siniestros', tipo:'barraH', dim:'dependencia', tope:15, alto:'alto', ancha:true}
  ]},
  {titulo:'Cómo ocurren', bajada:'Mecánica del siniestro, carátula de la causa y movilidad de las personas involucradas.', fichas:[
    {id:'tipo', titulo:'Tipo de siniestro', sub:'Mecánica del hecho registrada en la actuación', tipo:'barraH', dim:'tipoSin', alto:'alto'},
    {id:'causa',titulo:'Carátula de la causa', sub:'Calificación legal asignada a la actuación', tipo:'barraH', dim:'causa'},
    {id:'mv',   titulo:'Movilidad de la víctima', sub:'Cómo se desplazaba la persona damnificada', tipo:'barraH', dim:'p12', tope:12, alto:'alto'},
    {id:'mc',   titulo:'Movilidad del causante', sub:'Cómo se desplazaba la otra parte interviniente', tipo:'barraH', dim:'p23', tope:12, alto:'alto'}
  ]},
  {titulo:'A quiénes afectan', bajada:'Perfil agregado de las personas damnificadas. No se publica ningún dato que permita identificarlas.', fichas:[
    {id:'sexo', titulo:'Sexo de las víctimas', sub:'Registros de víctimas según sexo', tipo:'barra', dim:'sexo', victima:true},
    {id:'edad', titulo:'Rango etario de las víctimas', sub:'Registros de víctimas agrupados por franja de edad', tipo:'barra', dim:'rango', orden:'rango', victima:true, ancha:true},
    {id:'cond', titulo:'Condición de la víctima', sub:'Rol de la persona damnificada en el siniestro', tipo:'barra', dim:'condicion', victima:true},
    {id:'cat',  titulo:'Categoría del siniestro', sub:'Siniestros con personas lesionadas y con personas fallecidas', tipo:'barra', dim:'catSin'}
  ]}
];

const PALETA = ['#C3D0DC','#9DB6CB','#6E9AC0','#2E7BB5','#10243A','#C9971C'];
const colorAnio = (i, n) => n >= PALETA.length ? PALETA[i % PALETA.length]
                                               : PALETA.slice(PALETA.length - 1 - n, PALETA.length - 1)[i];

const graficos = {};
Chart.defaults.font.family = "'Archivo', system-ui, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#3C556E';

function pintarGraficos(filas){
  const anios = [...new Set(filas.map(d=>d.anio))].sort();
  SECCIONES.forEach(sec => sec.fichas.forEach(f => pintarFicha(f, filas, anios)));
}

function pintarFicha(f, filas, anios){
  const base = f.victima ? filas.filter(d => d[f.dim]) : filas.filter(d => d.hecho && d[f.dim]);
  const total = new Map(), porAnio = {};
  anios.forEach(a => porAnio[a] = new Map());
  for(const d of base){
    const v = d[f.dim];
    total.set(v, (total.get(v)||0) + 1);
    const m = porAnio[d.anio];
    m.set(v, (m.get(v)||0) + 1);
  }
  let claves = ordenar(total.keys(), f.orden === 'dia' ? null : (f.orden || 'cantidad'), total);
  if(f.orden === 'dia') claves = ORDEN_DIA.filter(d => total.has(d));
  if(f.tope) claves = ordenar(total.keys(), 'cantidad', total).slice(0, f.tope);

  const cv = document.getElementById('cv-' + f.id);
  const cont = cv.parentElement;
  if(!claves.length){
    if(graficos[f.id]){ graficos[f.id].destroy(); delete graficos[f.id]; }
    cont.innerHTML = '<p class="vacio">No hay registros para los filtros seleccionados.</p>';
    return;
  }
  if(!cont.querySelector('canvas')) cont.innerHTML = `<canvas id="cv-${f.id}"></canvas>`;
  const lienzo = document.getElementById('cv-' + f.id);

  const etiquetas = claves.map(v => f.dim === 'mes' ? lindo(v) : v);
  const datasets = anios.map((a,i) => ({
    label: a,
    data: claves.map(v => porAnio[a].get(v) || 0),
    backgroundColor: colorAnio(i, anios.length),
    borderColor: colorAnio(i, anios.length),
    borderWidth: f.tipo === 'linea' ? 2.5 : 0,
    borderRadius: f.tipo === 'linea' ? 0 : 3,
    tension: .32,
    pointRadius: 2.5,
    pointHoverRadius: 5,
    fill: false,
    maxBarThickness: 46
  }));

  if(graficos[f.id]) graficos[f.id].destroy();
  const horizontal = f.tipo === 'barraH';
  graficos[f.id] = new Chart(lienzo, {
    type: f.tipo === 'linea' ? 'line' : 'bar',
    data: {labels: etiquetas, datasets},
    options: {
      responsive:true, maintainAspectRatio:false,
      indexAxis: horizontal ? 'y' : 'x',
      interaction:{mode:'index', intersect:false},
      animation:{duration: 320},
      plugins:{
        legend:{display: anios.length > 1, position:'bottom',
          labels:{boxWidth:10, boxHeight:10, usePointStyle:true, pointStyle:'rectRounded', padding:14}},
        tooltip:{
          backgroundColor:'#10243A', padding:10, cornerRadius:6, titleFont:{weight:'600'},
          callbacks:{ label: c => ` ${c.dataset.label}: ${nro(c.parsed[horizontal?'x':'y'])}` }
        }
      },
      scales:{
        x:{ grid:{display:horizontal, color:'#EEF1F5'}, border:{display:false},
            ticks:{autoSkip:!horizontal, maxRotation: horizontal?0:45, minRotation:0,
                   callback:function(v){ const l = this.getLabelForValue(v);
                     return horizontal ? nro(v) : (String(l).length > 16 ? String(l).slice(0,15)+'…' : l); }}},
        y:{ beginAtZero:true, grid:{display:!horizontal, color:'#EEF1F5'}, border:{display:false},
            ticks:{ callback:function(v){ const l = this.getLabelForValue(v);
                     return horizontal ? (String(l).length > 26 ? String(l).slice(0,25)+'…' : l) : nro(v); }}}
      }
    }
  });
}

/* =============================================================
   6. Ciclo de actualización
   ============================================================= */
function actualizar(){
  const filas = filtrar();
  const t = totales(filas);
  refrescarPanel();
  pintarCinta(t);
  pintarKpis(filas, t);
  pintarGraficos(filas);
  document.getElementById('cinta-et').textContent = t.sin
    ? `personas involucradas en ${nro(t.sin)} siniestros viales`
    : 'no hay siniestros que cumplan con los filtros seleccionados';
}

iniciar();

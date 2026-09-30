/* =============================================================
   Accidentología Vial — Provincia de Tucumán
   Policía de Tucumán · D-3 · Sección Estadística y Archivo

   La página se actualiza reemplazando los archivos de datos/ que
   genera preparar-base.html:
     · datos/accidentologia.json  (versión compacta: es la que se lee)
     · datos/accidentologia.xlsx  (respaldo, si no estuviera el .json)
   ============================================================= */

/* Versión compacta de la base (unas 20 veces más liviana que el Excel y sin
   necesidad de interpretar el .xlsx en el navegador): carga en una fracción
   del tiempo, sobre todo en celulares. */
const RUTA_JSON = 'datos/accidentologia.json';

/* Lector de Excel: solo se descarga si hace falta (respaldo o archivo elegido a mano) */
const URL_SHEETJS = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

/* Comparativo de los indicadores ("2026 vs 2025"): el año en curso se compara
   contra los MISMOS meses del año anterior. Con true, además, se deja afuera el
   último mes con datos del año en curso porque suele estar incompleto (la planilla
   se actualiza una vez por mes). Poné false si la base se carga con meses completos. */
const EXCLUIR_MES_EN_CURSO = true;

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
  localidad:   ['LOCALIDAD'],
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

/* Geometría de departamentos y localidades (simplificada, fuente IGN/BAHRA vía Georef y
   mgaitan/departamentos_argentina). Se usa solo para el mapa. */
const MAPA_DATOS = {"viewBox":[0,0,640,818.5],"deptos":{"LACOCHA":{"d":"M242.4,638.6 L245.2,639.3 L247.9,644.0 L253.3,639.5 L255.4,640.5 L251.5,643.5 L254.9,645.8 L252.2,660.3 L259.6,661.8 L257.5,671.8 L261.5,672.6 L258.0,690.2 L266.4,691.9 L265.4,697.4 L272.2,699.2 L269.1,710.2 L289.7,713.5 L281.7,749.9 L269.7,755.6 L267.5,759.5 L259.3,764.6 L259.4,767.3 L256.8,766.6 L257.1,769.6 L251.4,774.9 L251.9,780.0 L248.3,781.2 L249.4,783.9 L244.9,784.2 L242.9,790.7 L238.6,793.1 L237.2,798.2 L233.0,800.5 L227.7,794.4 L224.5,794.8 L219.4,791.4 L217.5,783.6 L200.2,784.2 L197.2,766.3 L188.5,750.8 L187.8,742.2 L189.3,735.3 L186.0,724.1 L186.8,718.1 L181.1,721.3 L180.5,713.9 L167.9,702.6 L167.1,693.9 L163.2,691.4 L165.9,689.8 L166.0,683.4 L167.5,683.2 L169.2,673.6 L171.5,672.0 L171.6,666.5 L173.3,666.6 L169.4,657.7 L170.5,650.5 L173.2,648.9 L172.3,647.2 L176.6,639.4 L184.5,638.8 L189.6,642.3 L191.3,647.3 L198.1,649.4 L209.8,646.9 L217.1,647.8 L219.8,646.7 L220.0,644.7 L224.5,644.6 L225.0,640.8 L228.8,644.2 L234.0,640.6 L234.9,641.7 L242.4,638.6 Z","cab":"LA COCHA","cx":224.5,"cy":712.3},"GRANEROS":{"d":"M427.8,629.7 L442.5,691.6 L427.5,696.0 L431.6,711.4 L423.6,746.6 L399.3,752.2 L380.8,745.8 L377.1,761.9 L374.0,761.3 L376.6,773.5 L307.6,733.2 L297.1,737.7 L293.1,744.0 L281.7,749.9 L289.7,713.5 L269.1,710.2 L272.2,699.2 L265.4,697.4 L266.4,691.9 L258.0,690.2 L261.5,672.6 L257.5,671.8 L259.6,661.8 L252.2,660.3 L254.9,645.8 L251.5,643.5 L255.4,640.5 L253.3,639.5 L248.8,643.8 L246.9,643.5 L253.4,612.1 L263.1,611.2 L264.6,612.8 L265.3,610.4 L266.7,610.3 L266.6,610.6 L266.5,612.1 L266.8,612.1 L266.7,610.5 L270.1,611.5 L276.1,609.5 L276.8,610.9 L292.4,611.0 L292.9,609.5 L295.2,610.0 L295.0,607.9 L305.6,606.0 L308.8,607.0 L313.5,604.6 L315.8,609.7 L318.5,609.3 L321.9,611.7 L327.7,610.1 L333.0,612.0 L332.9,613.7 L337.1,613.3 L338.3,616.2 L344.9,620.5 L355.3,618.2 L359.9,620.3 L366.3,619.8 L368.1,621.9 L377.4,621.1 L379.6,627.5 L383.2,630.0 L395.2,625.4 L396.5,622.2 L402.1,623.2 L404.1,627.8 L409.1,627.7 L412.2,629.7 L417.4,624.7 L424.9,626.1 L427.8,629.7 Z","cab":"GRANEROS","cx":346.6,"cy":680.2},"JUANBAUTISTAALBERDI":{"d":"M243.3,611.5 L253.4,612.2 L246.9,643.5 L245.2,639.3 L241.7,638.5 L228.8,644.2 L225.0,640.8 L224.5,644.6 L220.0,644.7 L219.8,646.7 L217.1,647.8 L209.8,646.9 L198.1,649.4 L191.3,647.3 L189.6,642.3 L184.5,638.8 L175.9,640.1 L172.3,647.2 L173.2,648.9 L170.5,650.6 L169.4,657.7 L173.3,666.6 L171.6,666.5 L171.5,672.0 L169.2,673.6 L167.5,683.2 L166.0,683.4 L165.9,689.8 L163.2,691.4 L167.0,693.8 L167.1,701.3 L168.6,703.1 L152.8,709.8 L154.4,717.6 L153.1,717.2 L151.1,721.6 L149.4,721.2 L149.5,723.1 L145.2,723.1 L147.2,725.7 L145.5,725.5 L145.8,728.1 L141.0,735.4 L142.3,720.6 L130.8,684.3 L129.2,670.5 L125.7,663.9 L111.4,662.5 L105.5,650.2 L101.2,632.9 L97.4,625.5 L99.0,619.6 L93.4,612.3 L91.7,605.0 L93.1,602.0 L89.9,592.3 L90.3,571.0 L92.7,566.6 L90.8,563.0 L106.2,570.8 L105.6,582.5 L107.4,585.5 L113.6,593.3 L124.5,600.7 L161.8,608.2 L165.5,607.8 L167.3,605.6 L169.7,606.1 L170.7,610.5 L178.6,611.5 L183.6,608.9 L193.9,609.2 L201.4,603.0 L201.4,605.8 L206.0,605.7 L208.2,601.9 L213.6,601.6 L217.3,603.6 L220.6,601.2 L226.2,601.4 L231.6,602.5 L232.1,604.8 L243.3,611.5 Z","cab":"JUAN BAUTISTA ALBERDI","cx":158.5,"cy":637.2},"RIOCHICO":{"d":"M242.0,567.2 L238.6,586.9 L239.6,589.6 L243.6,589.2 L244.8,591.1 L245.8,587.0 L247.9,588.8 L243.3,611.5 L232.1,604.8 L231.2,602.3 L223.3,601.2 L218.0,602.0 L217.3,603.6 L213.6,601.6 L208.2,601.9 L206.0,605.7 L201.4,605.8 L201.4,603.0 L193.9,609.2 L183.6,608.9 L178.6,611.5 L170.7,610.5 L169.7,606.1 L167.3,605.6 L165.5,607.8 L161.8,608.2 L124.5,600.7 L113.6,593.3 L107.4,585.5 L105.6,582.5 L106.2,570.8 L90.8,563.0 L90.4,560.8 L92.8,554.5 L96.2,555.7 L99.8,554.0 L104.7,556.1 L104.3,552.2 L107.2,550.3 L106.1,548.2 L109.8,548.1 L116.7,541.7 L118.2,543.3 L119.4,541.7 L126.7,543.3 L133.1,539.6 L139.6,540.6 L148.1,537.7 L154.4,537.0 L164.2,539.0 L178.2,537.0 L195.4,547.3 L204.2,557.9 L208.2,555.5 L208.8,560.3 L210.0,558.1 L214.2,559.1 L217.7,563.0 L220.5,558.6 L220.3,563.3 L222.7,558.8 L224.1,559.2 L224.0,562.5 L230.5,561.4 L232.7,563.8 L239.7,561.3 L242.0,567.2 Z","cab":"AGUILARES","cx":169.6,"cy":574.9},"CHICLIGASTA":{"d":"M262.0,514.1 L272.2,516.1 L270.1,524.3 L273.3,525.7 L267.1,547.5 L270.2,553.3 L273.0,552.4 L269.5,563.5 L265.6,563.7 L266.8,558.3 L263.9,556.7 L260.0,568.0 L256.9,569.1 L255.4,566.7 L253.1,568.1 L252.0,566.2 L251.3,565.7 L249.6,565.1 L248.3,568.1 L242.6,567.7 L239.7,561.3 L232.7,563.8 L230.5,561.4 L224.0,562.5 L224.1,559.2 L222.7,558.8 L220.3,563.3 L220.5,558.6 L217.7,563.0 L214.2,559.1 L210.0,558.1 L208.8,560.3 L208.2,555.5 L204.2,557.9 L195.4,547.3 L178.2,537.0 L164.2,539.0 L154.4,537.0 L148.1,537.7 L139.6,540.6 L133.1,539.6 L126.7,543.3 L119.4,541.7 L118.2,543.3 L116.7,541.7 L109.8,548.1 L106.1,548.2 L107.2,550.3 L104.3,552.2 L104.7,556.1 L99.8,554.0 L95.4,555.7 L92.8,554.5 L93.7,539.1 L91.3,528.0 L83.3,520.8 L76.8,518.6 L71.5,514.3 L55.3,510.7 L39.5,501.4 L31.1,499.3 L34.0,496.0 L34.1,490.8 L37.5,489.5 L37.9,485.8 L42.0,481.2 L53.5,478.5 L51.9,474.2 L56.5,470.3 L61.8,452.5 L66.5,448.9 L68.5,444.9 L73.5,443.1 L73.8,440.5 L81.2,436.2 L85.5,436.1 L91.5,422.6 L90.9,416.4 L108.7,433.1 L117.7,435.6 L121.0,439.3 L130.7,442.1 L144.3,449.6 L148.0,456.7 L151.7,457.9 L153.4,460.5 L152.5,467.2 L154.0,472.9 L160.0,472.3 L166.2,475.6 L172.5,473.3 L189.2,475.5 L194.0,485.4 L201.3,489.9 L209.0,491.6 L214.8,496.8 L227.2,501.5 L230.1,505.8 L244.0,509.0 L247.7,505.2 L249.4,510.0 L252.9,510.0 L253.2,511.8 L256.0,510.3 L258.4,513.6 L261.0,511.8 L262.0,514.1 Z","cab":"CONCEPCION","cx":148.6,"cy":502.6},"SIMOCA":{"d":"M302.2,468.1 L312.2,467.5 L315.0,473.7 L318.3,475.7 L318.2,481.4 L321.9,481.1 L323.9,483.4 L327.4,483.0 L330.6,485.0 L328.5,493.0 L331.4,497.4 L330.3,501.7 L335.2,502.7 L339.9,507.9 L347.7,512.4 L355.3,521.5 L360.0,520.9 L359.2,527.9 L361.5,531.8 L366.1,534.8 L369.8,540.4 L375.1,540.2 L377.9,543.5 L381.5,543.6 L386.2,549.3 L386.5,552.8 L392.5,558.2 L395.6,558.3 L397.8,565.4 L402.6,566.3 L404.5,563.3 L405.9,570.3 L401.0,572.4 L400.7,574.0 L402.9,573.4 L404.2,576.7 L406.7,575.9 L414.9,579.5 L415.0,571.8 L424.0,586.6 L434.3,588.4 L435.4,590.2 L447.2,594.4 L449.3,598.6 L449.9,612.2 L446.9,620.0 L440.4,623.4 L433.6,619.6 L432.9,624.2 L428.8,626.1 L427.8,629.7 L424.9,626.1 L417.4,624.7 L412.2,629.7 L409.1,627.7 L404.1,627.8 L402.1,623.2 L396.5,622.2 L395.2,625.4 L383.2,630.0 L379.6,627.5 L377.4,621.1 L368.1,621.9 L366.3,619.8 L359.9,620.3 L355.3,618.2 L344.9,620.5 L338.3,616.2 L337.1,613.3 L332.9,613.7 L333.0,612.0 L327.7,610.1 L321.9,611.7 L318.5,609.3 L315.8,609.7 L313.5,604.6 L308.8,607.0 L305.6,606.0 L295.0,607.9 L295.2,610.0 L292.9,609.5 L292.4,611.0 L276.8,610.9 L276.1,609.5 L270.1,611.5 L266.7,610.5 L266.8,612.1 L266.7,610.3 L265.3,610.4 L264.6,612.8 L263.1,611.2 L251.4,613.1 L243.3,611.5 L247.9,588.8 L245.8,587.0 L244.8,591.1 L243.6,589.2 L239.6,589.6 L238.6,586.9 L242.0,567.2 L248.3,568.1 L249.6,565.1 L253.1,568.1 L255.4,566.7 L257.3,569.0 L260.0,568.0 L261.0,563.4 L262.4,563.8 L263.9,556.7 L266.8,558.3 L265.6,563.7 L269.5,563.5 L271.8,557.4 L273.0,552.4 L270.2,553.3 L267.1,547.5 L273.3,525.7 L270.1,524.3 L272.2,516.1 L270.0,514.8 L274.1,502.2 L272.1,499.2 L275.0,488.6 L277.0,488.6 L278.6,481.6 L281.2,482.9 L282.1,478.6 L287.0,481.7 L287.4,479.3 L289.1,479.3 L292.2,468.1 L295.0,464.7 L302.2,468.1 Z","cab":"SIMOCA","cx":333.4,"cy":564.6},"LULES":{"d":"M283.7,318.2 L283.5,320.9 L303.0,325.4 L304.7,325.4 L305.5,322.8 L324.7,326.3 L325.3,329.9 L327.9,330.4 L328.3,328.0 L341.1,329.2 L342.0,332.2 L339.6,335.2 L341.1,337.7 L355.5,342.2 L355.5,349.4 L363.4,360.9 L361.0,362.5 L363.4,366.2 L360.9,367.2 L362.4,369.6 L361.1,371.1 L359.2,369.6 L360.7,373.2 L359.6,374.6 L361.0,375.3 L356.0,382.3 L353.9,382.1 L354.5,386.6 L352.9,387.3 L351.5,385.4 L351.3,389.4 L348.6,391.2 L350.2,393.2 L348.1,394.8 L350.6,396.1 L344.8,400.8 L346.1,403.8 L341.3,402.7 L342.8,401.9 L341.8,399.7 L343.7,400.1 L343.0,397.1 L344.6,398.8 L345.9,397.8 L343.1,389.6 L344.9,389.0 L345.6,385.5 L347.9,386.1 L348.0,384.1 L345.0,383.0 L345.7,378.7 L343.1,376.6 L344.7,376.0 L336.4,371.2 L335.9,367.1 L333.8,367.4 L334.9,367.0 L332.5,363.3 L328.0,381.7 L329.9,383.1 L329.4,386.5 L327.3,386.3 L325.6,389.4 L331.5,392.6 L332.5,388.5 L336.4,389.3 L330.4,402.9 L331.1,399.3 L325.4,398.8 L327.3,400.8 L327.3,404.6 L322.0,408.1 L319.6,403.0 L314.2,403.3 L311.4,398.1 L310.4,398.6 L310.3,395.4 L307.5,395.2 L308.3,389.3 L304.6,387.1 L303.9,381.3 L295.8,375.8 L292.1,377.2 L291.6,373.3 L288.5,372.4 L286.6,369.2 L285.1,369.8 L283.1,364.5 L280.1,366.4 L277.3,360.7 L277.6,355.1 L273.3,351.9 L271.4,347.8 L266.2,348.2 L264.3,344.3 L260.1,341.2 L252.6,340.9 L239.6,334.6 L234.4,336.5 L231.3,333.3 L222.0,332.3 L217.2,329.9 L217.0,325.0 L206.3,325.3 L204.3,327.6 L204.7,319.6 L195.3,319.9 L184.3,290.3 L196.9,291.3 L198.6,295.1 L205.5,297.5 L209.5,295.8 L209.0,293.1 L211.9,290.1 L221.3,286.2 L233.2,292.5 L236.5,297.4 L243.3,300.4 L253.8,294.0 L254.7,296.2 L263.1,297.5 L273.6,303.6 L277.1,312.9 L283.7,318.2 Z","cab":"LUNES","cx":283.4,"cy":339.6},"MONTEROS":{"d":"M285.6,445.9 L291.7,449.8 L289.9,453.1 L293.1,452.8 L295.3,456.1 L296.3,454.6 L300.0,455.2 L298.3,455.9 L303.0,457.5 L303.1,460.3 L304.2,458.4 L304.3,459.9 L306.1,459.7 L305.1,463.6 L308.4,465.4 L302.4,468.1 L295.0,464.7 L287.0,481.7 L282.1,478.6 L281.2,482.9 L278.6,481.6 L277.0,488.6 L275.0,488.6 L273.3,493.6 L272.1,499.2 L274.2,501.8 L271.7,512.6 L270.0,514.8 L265.8,515.4 L262.0,514.1 L261.3,511.9 L258.4,513.6 L256.0,510.3 L253.2,511.8 L252.9,510.0 L249.4,510.0 L247.7,505.2 L244.0,509.0 L230.1,505.8 L227.2,501.5 L214.8,496.8 L209.0,491.6 L201.3,489.9 L194.0,485.4 L189.2,475.5 L172.5,473.3 L166.2,475.6 L160.0,472.3 L154.0,472.9 L152.5,467.2 L153.4,460.5 L151.7,457.9 L148.0,456.7 L144.3,449.6 L130.7,442.1 L121.0,439.3 L117.7,435.6 L108.7,433.1 L90.9,416.4 L93.7,414.2 L94.0,408.1 L96.9,407.8 L103.2,402.6 L110.8,400.8 L112.1,397.5 L110.5,391.3 L112.0,387.9 L118.3,386.1 L119.8,383.9 L130.9,388.4 L132.9,391.4 L133.2,401.4 L144.4,410.7 L143.5,412.4 L146.3,414.9 L145.9,420.0 L151.6,420.1 L152.4,424.4 L154.3,424.7 L153.4,428.8 L155.8,430.9 L154.9,433.0 L159.3,434.7 L163.0,431.8 L161.1,430.2 L162.4,427.4 L160.6,424.7 L159.6,413.7 L182.4,407.5 L181.2,404.9 L184.1,398.6 L188.4,391.7 L192.5,388.7 L184.9,383.5 L189.3,381.9 L195.0,385.9 L198.1,376.9 L197.9,371.7 L201.5,371.4 L201.5,375.9 L206.8,376.9 L203.5,366.8 L207.4,368.4 L209.6,363.7 L219.8,368.3 L221.4,366.2 L226.5,368.2 L235.3,377.9 L243.9,382.1 L242.3,383.1 L245.0,387.5 L248.6,389.5 L249.6,395.5 L253.9,396.8 L261.3,404.8 L261.3,401.2 L262.8,402.7 L267.0,412.3 L267.7,419.1 L263.5,427.8 L265.6,431.8 L264.7,433.1 L271.2,435.5 L269.3,438.0 L274.8,440.1 L272.7,441.1 L276.7,441.2 L278.1,443.3 L279.6,442.0 L279.3,444.0 L281.7,445.4 L284.3,444.1 L284.6,447.4 L285.6,445.9 Z","cab":"MONTEROS","cx":208.9,"cy":439.9},"LEALES":{"d":"M349.8,393.8 L394.3,405.2 L395.1,401.3 L419.8,406.6 L420.9,400.9 L503.6,420.5 L503.2,422.9 L515.4,425.5 L509.3,455.1 L504.7,454.3 L489.5,521.8 L460.9,532.6 L459.7,538.0 L446.6,535.1 L434.3,588.4 L427.4,588.2 L423.1,585.9 L415.0,571.8 L414.9,579.5 L406.7,575.9 L404.2,576.7 L403.3,573.6 L400.4,573.6 L405.9,570.3 L404.5,563.3 L402.6,566.3 L397.8,565.4 L395.6,558.3 L392.5,558.2 L386.5,552.8 L386.2,549.3 L381.5,543.6 L377.9,543.5 L376.4,540.6 L371.3,541.2 L368.6,539.6 L367.8,536.8 L359.2,528.1 L360.1,521.0 L355.3,521.5 L347.7,512.4 L339.9,507.9 L335.2,502.7 L330.3,501.7 L331.4,497.4 L328.5,493.1 L330.9,486.7 L330.2,484.5 L323.9,483.4 L321.9,481.1 L318.2,481.4 L318.3,475.7 L315.0,473.7 L312.2,467.5 L307.7,466.8 L308.5,465.1 L305.1,463.6 L306.1,459.7 L304.3,459.9 L304.2,458.4 L303.1,460.3 L303.0,457.5 L298.3,455.9 L300.0,455.2 L296.3,454.6 L295.3,456.1 L293.1,452.8 L289.9,453.1 L291.7,449.8 L285.6,445.9 L287.3,445.0 L287.2,439.7 L292.1,442.5 L294.0,441.1 L297.4,436.6 L298.6,430.3 L311.8,438.5 L311.9,435.8 L315.3,436.1 L312.3,429.8 L317.3,422.7 L317.8,418.9 L319.3,418.6 L318.3,416.9 L319.7,417.9 L320.4,415.0 L322.7,414.3 L321.6,413.3 L322.7,411.4 L320.9,410.3 L322.0,408.1 L326.8,405.8 L327.3,400.8 L325.4,398.8 L331.1,399.3 L330.4,402.9 L336.4,389.3 L332.5,388.5 L331.5,392.6 L325.6,389.4 L327.3,386.3 L329.4,386.5 L329.9,383.1 L328.0,381.7 L332.5,363.3 L334.9,367.0 L333.8,367.4 L335.9,367.1 L336.4,371.2 L344.7,376.0 L343.1,376.6 L345.7,378.7 L345.0,383.0 L348.0,384.1 L347.9,386.1 L345.6,385.5 L344.9,389.0 L343.1,389.6 L345.9,397.8 L344.6,398.8 L343.0,397.1 L343.7,400.1 L341.6,399.9 L342.8,401.9 L341.4,403.0 L346.1,403.8 L344.7,401.0 L350.5,396.3 L348.1,394.8 L349.8,393.8 Z","cab":"BELLA VISTA","cx":407.3,"cy":471.6},"FAMAILLA":{"d":"M322.0,408.1 L320.9,410.3 L322.7,411.4 L321.6,413.3 L322.7,414.3 L320.4,415.0 L319.7,417.9 L318.3,416.9 L319.3,418.6 L317.8,418.9 L317.3,422.7 L312.3,429.8 L314.9,436.6 L311.9,435.8 L311.8,438.5 L298.6,430.3 L297.4,436.6 L294.0,441.1 L292.1,442.5 L287.2,439.7 L287.3,445.0 L285.0,447.4 L284.3,444.1 L281.7,445.4 L279.3,444.0 L279.6,442.0 L278.1,443.3 L276.7,441.2 L272.7,441.1 L274.8,440.1 L269.6,438.6 L271.2,435.5 L264.8,433.2 L265.6,431.8 L263.8,429.9 L263.7,427.0 L267.7,419.0 L267.0,412.3 L262.8,402.7 L261.3,401.2 L261.3,404.8 L253.9,396.8 L249.6,395.5 L248.6,389.5 L245.0,387.5 L242.3,383.1 L243.9,382.1 L235.3,377.9 L226.5,368.2 L221.4,366.2 L219.8,368.3 L209.6,363.7 L207.4,368.4 L203.3,366.7 L201.3,358.9 L205.2,351.1 L209.9,352.4 L210.7,347.1 L213.5,345.3 L213.1,334.7 L215.2,324.9 L217.4,325.6 L217.2,329.9 L222.0,332.3 L231.3,333.3 L234.4,336.5 L239.6,334.6 L252.6,340.9 L260.1,341.2 L264.3,344.3 L266.2,348.2 L271.4,347.8 L273.3,351.9 L277.6,355.1 L277.3,360.7 L280.1,366.4 L283.1,364.5 L285.1,369.8 L286.6,369.2 L288.5,372.4 L291.6,373.3 L292.1,377.2 L295.8,375.8 L303.9,381.3 L304.6,387.1 L308.3,389.3 L307.5,395.2 L310.3,395.4 L310.4,398.6 L311.4,398.1 L314.2,403.3 L319.6,403.0 L322.0,408.1 Z","cab":"FAMAILLA","cx":266.8,"cy":385.1},"YERBABUENA":{"d":"M348.6,308.3 L346.0,311.7 L341.0,329.3 L328.3,328.0 L327.9,330.4 L325.3,329.9 L324.7,326.3 L305.5,322.8 L304.7,325.4 L303.0,325.4 L283.5,320.9 L286.6,289.8 L291.0,302.7 L294.6,307.5 L303.7,299.0 L310.9,300.2 L308.8,289.8 L310.0,288.3 L345.1,295.1 L341.1,305.2 L348.6,308.3 Z","cab":"YERBA BUENA","cx":317.4,"cy":310.1},"BURRUYACU":{"d":"M617.3,82.3 L614.8,90.3 L622.0,117.5 L615.0,128.9 L608.0,135.9 L604.8,143.8 L603.3,159.0 L607.0,171.3 L606.1,180.5 L609.4,189.3 L608.0,212.3 L612.4,214.3 L616.0,264.2 L584.4,258.0 L573.6,315.9 L554.1,312.4 L552.3,324.6 L541.8,322.6 L537.3,348.7 L533.1,347.5 L530.8,348.9 L531.3,341.6 L515.4,337.9 L516.5,332.5 L505.7,330.7 L506.9,323.8 L492.3,320.8 L490.2,331.2 L475.2,328.0 L480.3,304.4 L460.0,300.2 L456.0,317.5 L400.2,305.5 L401.9,297.1 L391.2,294.9 L391.2,292.4 L390.4,294.2 L386.4,294.3 L381.4,297.3 L382.4,285.5 L380.0,284.1 L379.8,274.7 L374.6,268.9 L376.4,265.3 L375.8,257.3 L369.2,251.2 L370.1,246.7 L370.9,247.6 L371.9,245.6 L371.0,242.0 L378.1,236.7 L378.6,234.5 L377.0,233.2 L376.3,226.1 L373.9,225.1 L371.3,228.1 L369.3,227.1 L369.4,216.2 L377.9,216.2 L380.3,208.6 L385.2,203.1 L384.8,199.9 L387.3,197.9 L390.9,199.5 L397.1,197.5 L397.6,193.0 L395.2,188.8 L398.6,179.1 L398.2,167.5 L400.9,161.0 L400.9,155.1 L405.1,151.5 L406.0,144.4 L412.4,137.8 L413.3,131.4 L419.2,127.0 L421.6,122.8 L424.4,110.4 L423.1,104.3 L426.0,103.6 L437.2,107.3 L446.9,105.5 L448.7,106.4 L448.3,109.4 L451.4,116.3 L454.9,117.5 L459.3,113.7 L461.0,114.5 L461.5,112.6 L470.0,110.9 L474.2,106.3 L477.6,107.7 L482.6,104.7 L488.0,95.8 L490.5,96.2 L489.8,94.7 L491.9,93.9 L492.6,90.1 L497.2,89.4 L500.6,91.8 L503.3,87.3 L506.0,87.9 L512.1,84.2 L522.9,84.5 L526.7,89.5 L531.0,91.1 L537.0,91.0 L540.4,88.5 L546.6,89.4 L547.4,91.3 L552.5,89.5 L557.7,92.4 L564.8,90.1 L564.0,87.5 L569.6,88.0 L572.0,84.4 L575.1,88.1 L577.2,85.3 L580.2,87.0 L580.8,84.5 L582.7,84.9 L583.7,82.9 L585.7,84.3 L587.4,82.2 L592.7,85.5 L596.9,83.7 L597.3,85.5 L598.2,83.1 L600.8,82.5 L608.6,84.5 L611.6,80.6 L617.3,82.3 Z","cab":"BURRUYACU","cx":501.7,"cy":207.1},"TAFIVIEJO":{"d":"M185.1,201.8 L193.0,205.6 L196.3,203.9 L205.6,204.5 L216.0,216.8 L218.3,213.9 L224.1,212.6 L227.3,214.3 L238.1,208.7 L258.4,214.0 L260.4,204.5 L263.9,205.5 L263.4,203.7 L264.7,204.3 L265.4,201.4 L271.5,202.4 L265.1,209.9 L263.7,215.0 L265.9,216.6 L260.9,221.3 L266.7,226.4 L268.9,236.1 L274.0,228.7 L283.4,232.1 L288.6,237.6 L293.6,238.5 L297.7,242.7 L300.7,243.8 L302.5,242.0 L312.4,244.5 L319.8,238.9 L329.7,239.7 L337.4,235.7 L344.2,239.1 L350.2,236.8 L352.8,233.7 L356.9,235.2 L354.1,227.0 L358.7,223.3 L356.1,218.4 L358.4,215.1 L358.0,212.3 L360.8,210.2 L358.7,204.2 L372.2,208.1 L370.0,209.2 L367.1,217.3 L369.4,216.2 L370.0,218.2 L369.3,227.1 L371.3,228.1 L373.9,225.1 L376.3,226.1 L377.0,233.2 L378.6,234.7 L378.1,236.7 L371.0,242.0 L371.9,245.6 L370.9,247.6 L370.1,246.7 L369.2,251.2 L375.8,257.2 L376.4,265.3 L374.7,269.4 L379.8,274.7 L380.0,284.1 L382.4,285.5 L381.4,289.4 L382.4,295.2 L379.6,300.3 L381.3,302.6 L380.1,311.2 L376.1,310.4 L370.5,312.9 L342.3,306.5 L341.1,304.4 L345.1,295.1 L310.0,288.3 L308.8,289.8 L310.9,300.2 L303.7,299.0 L294.6,307.5 L291.0,302.7 L286.6,289.8 L283.7,318.2 L277.1,312.9 L273.6,303.6 L263.1,297.5 L254.7,296.2 L253.8,294.0 L243.3,300.4 L236.5,297.4 L233.1,292.4 L220.8,286.2 L211.9,290.1 L209.0,293.1 L209.5,295.8 L205.5,297.5 L198.6,295.1 L196.9,291.3 L186.0,291.2 L186.3,284.2 L189.9,281.9 L186.4,273.7 L180.6,268.2 L167.3,261.9 L166.5,257.9 L164.6,257.0 L165.4,253.7 L161.2,250.7 L163.0,248.6 L162.0,246.1 L164.5,244.7 L168.3,235.2 L170.9,233.2 L174.4,234.3 L176.9,237.9 L178.3,234.5 L180.5,211.7 L174.8,205.4 L171.2,199.4 L171.6,197.3 L175.7,198.1 L180.2,203.3 L185.1,201.8 Z","cab":"TAFI VIEJO","cx":272.6,"cy":259.2},"TAFIDELVALLE":{"d":"M81.3,100.3 L184.3,137.3 L179.1,149.5 L183.2,158.8 L180.9,163.8 L184.4,173.1 L184.1,179.9 L188.3,188.0 L187.5,196.1 L185.1,201.8 L180.2,203.3 L175.7,198.1 L171.6,197.3 L171.2,199.4 L174.8,205.4 L180.5,211.7 L178.3,234.5 L176.9,237.9 L174.4,234.3 L170.9,233.2 L168.3,235.2 L164.5,244.7 L162.0,246.1 L163.0,248.6 L161.2,251.0 L165.4,253.7 L164.6,257.0 L166.5,257.9 L167.3,261.9 L180.6,268.2 L186.4,273.7 L189.9,281.9 L186.3,284.2 L185.5,290.2 L184.3,290.3 L195.3,319.9 L204.7,319.6 L204.3,327.6 L206.3,325.3 L215.5,325.5 L213.1,334.7 L213.5,345.3 L210.7,347.1 L209.9,352.4 L205.2,351.1 L201.3,358.9 L206.8,376.9 L201.5,375.9 L201.5,371.4 L197.9,371.7 L198.1,376.9 L195.0,385.9 L189.3,381.9 L184.9,383.5 L192.5,388.7 L188.4,391.7 L184.1,398.6 L181.2,404.9 L182.4,407.5 L159.6,413.7 L160.6,424.7 L162.4,427.4 L161.1,430.2 L163.0,431.8 L157.2,435.1 L154.8,432.7 L155.7,430.8 L153.5,429.5 L154.3,424.9 L152.4,424.4 L152.6,421.3 L151.2,419.8 L146.0,420.1 L146.3,414.9 L143.5,412.4 L144.4,410.7 L133.2,401.4 L132.5,390.7 L130.6,388.1 L126.0,387.6 L119.7,383.5 L125.8,368.0 L131.9,366.5 L132.2,356.5 L125.9,351.4 L124.4,347.4 L121.1,347.1 L116.7,340.3 L120.4,337.2 L118.9,334.6 L119.6,330.2 L126.0,322.1 L123.5,317.2 L129.1,313.0 L133.3,298.6 L141.2,294.7 L140.9,289.9 L66.3,234.9 L20.3,247.2 L21.6,235.3 L18.6,225.9 L21.2,221.8 L20.6,212.7 L22.5,209.4 L19.8,199.0 L22.7,195.8 L18.4,189.8 L18.0,183.4 L23.3,178.3 L26.1,172.2 L26.6,167.6 L23.8,162.4 L25.8,153.7 L36.2,150.4 L38.9,145.5 L40.6,132.2 L36.9,117.7 L38.1,105.1 L39.3,100.4 L44.6,98.7 L46.3,95.3 L45.8,82.9 L55.9,90.2 L64.6,91.2 L67.0,95.1 L69.5,94.7 L81.3,100.3 Z","cab":"TAFI DEL VALLE","cx":121.0,"cy":232.6},"CAPITAL":{"d":"M380.1,311.2 L380.7,320.8 L371.0,342.0 L367.9,342.0 L368.8,343.8 L366.7,345.6 L365.8,350.8 L363.4,352.0 L362.8,356.2 L364.4,357.3 L363.0,360.5 L355.5,349.4 L355.5,342.2 L341.1,337.7 L339.7,335.7 L342.0,332.2 L340.9,327.7 L346.0,311.7 L348.6,308.3 L351.5,308.3 L370.5,312.9 L376.1,310.4 L380.1,311.2 Z","cab":"SAN MIGUEL DE TUCUMAN","cx":360.2,"cy":328.1},"CRUZALTA":{"d":"M537.3,348.7 L529.2,396.6 L526.4,396.0 L525.2,401.5 L520.5,400.5 L515.4,425.5 L503.2,422.9 L503.6,420.5 L420.9,400.9 L419.8,406.6 L395.1,401.3 L394.3,405.2 L349.9,393.8 L348.6,391.5 L351.3,389.4 L351.5,385.4 L352.9,387.3 L354.5,386.6 L353.9,382.1 L356.0,382.3 L361.0,375.4 L359.2,369.7 L361.1,371.1 L362.4,369.6 L360.9,367.2 L363.4,366.2 L361.0,362.5 L363.3,361.6 L364.4,357.7 L362.8,356.2 L363.4,352.0 L365.8,350.8 L366.7,345.6 L368.8,343.8 L367.9,342.0 L371.0,342.0 L380.7,320.8 L379.7,312.9 L381.3,302.6 L379.7,300.0 L381.5,297.3 L386.3,294.3 L390.4,294.2 L391.2,292.4 L391.2,294.9 L401.9,297.1 L400.2,305.5 L456.0,317.5 L460.0,300.2 L480.3,304.4 L475.2,328.0 L490.2,331.2 L492.3,320.8 L506.9,323.8 L505.7,330.7 L516.5,332.5 L515.4,337.9 L531.3,341.6 L530.8,348.9 L533.1,347.5 L537.3,348.7 Z","cab":"BANDA DEL RIO SALI","cx":446.2,"cy":362.1},"TRANCAS":{"d":"M420.9,97.6 L424.4,110.4 L421.6,122.8 L419.2,127.0 L413.3,131.4 L412.4,137.8 L406.0,144.4 L405.1,151.5 L400.9,155.1 L400.9,161.0 L398.2,167.5 L398.6,179.1 L395.2,188.8 L397.6,193.0 L396.7,198.2 L390.9,199.5 L387.3,197.9 L384.8,199.9 L385.2,203.1 L380.3,208.6 L377.9,216.2 L367.1,217.3 L370.0,209.2 L372.2,208.1 L358.7,204.1 L360.8,210.2 L358.0,212.3 L358.4,215.1 L356.1,218.4 L358.7,223.3 L354.1,227.0 L356.9,235.2 L352.8,233.7 L350.2,236.8 L344.2,239.1 L337.4,235.7 L329.7,239.7 L319.8,238.9 L312.4,244.5 L302.3,242.0 L300.5,243.8 L293.6,238.4 L288.6,237.6 L283.3,232.1 L274.0,228.7 L268.9,236.1 L266.7,226.4 L260.9,221.3 L265.9,216.6 L263.7,215.0 L265.1,209.9 L271.5,202.4 L265.4,201.4 L264.7,204.3 L263.4,203.7 L263.9,205.5 L260.4,204.5 L258.4,214.0 L238.1,208.7 L227.3,214.3 L224.1,212.6 L218.3,213.9 L216.0,216.8 L205.6,204.5 L196.3,203.9 L193.0,205.6 L185.0,201.6 L187.5,196.1 L188.3,188.0 L184.1,179.9 L184.4,173.1 L180.9,163.8 L183.2,158.8 L179.1,149.5 L184.3,137.3 L181.2,135.8 L184.3,130.3 L183.8,126.2 L179.9,119.8 L185.6,113.1 L179.5,106.3 L185.3,98.0 L183.7,96.0 L187.1,90.7 L186.9,84.8 L184.1,83.3 L184.5,78.3 L186.7,75.4 L187.1,70.8 L189.9,69.0 L190.9,60.6 L188.2,57.2 L192.2,41.6 L200.8,41.6 L209.5,39.1 L213.1,44.0 L218.9,44.1 L223.2,43.0 L225.3,40.1 L228.3,40.5 L234.7,33.6 L242.2,29.7 L251.7,32.8 L257.7,30.7 L267.2,32.8 L271.7,31.2 L274.6,33.9 L274.2,35.9 L280.7,36.5 L281.8,38.7 L286.3,38.3 L292.2,31.8 L299.3,28.7 L303.1,20.7 L305.4,21.7 L308.9,18.0 L313.9,18.9 L317.7,22.0 L322.3,22.7 L325.1,27.7 L331.4,27.4 L330.0,30.9 L334.7,39.3 L333.1,45.2 L336.2,49.5 L335.4,51.8 L337.9,53.6 L337.3,56.0 L339.3,56.9 L338.3,59.7 L340.0,63.5 L358.1,67.9 L356.5,70.5 L368.6,71.7 L423.2,86.6 L420.9,97.6 Z","cab":"TRANCAS","cx":293.1,"cy":132.0}},"localidades":{"SANMIGUELDETUCUMAN":{"x":359.6,"y":328.7},"LASTALITAS":{"x":371.2,"y":299.8},"ALDERETES":{"x":389.6,"y":311.6},"CONCEPCION":{"x":219.2,"y":532.8},"AGUILARES":{"x":213.6,"y":570.1},"SIMOCA":{"x":310.7,"y":498.5},"LULES":{"x":316.0,"y":376.3},"JUANBAUTISTAALBERDI":{"x":218.1,"y":630.8},"GRANEROS":{"x":285.7,"y":653.2},"BELLAVISTA":{"x":329.7,"y":417.4},"LACOCHA":{"x":224.5,"y":712.3},"MONTEROS":{"x":208.9,"y":439.9},"FAMAILLA":{"x":266.8,"y":385.1},"YERBABUENA":{"x":317.4,"y":310.1},"BURRUYACU":{"x":501.7,"y":207.1},"TAFIVIEJO":{"x":272.6,"y":259.2},"TAFIDELVALLE":{"x":121.0,"y":232.6},"BANDADELRIOSALI":{"x":446.2,"y":362.1},"TRANCAS":{"x":293.1,"y":132.0}}}
;

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
/* devuelve una versión de fn que recuerda sus resultados (solo para textos y números) */
function memoizar(fn){
  const cache = new Map();
  return function(x){
    if(x !== null && typeof x === 'object') return fn(x);
    let r = cache.get(x);
    if(r === undefined){ r = fn(x); cache.set(x, r); }
    return r;
  };
}
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

/* localidades: título por palabra + corrección de acentos por palabra suelta */
const PALABRA_FIX = {
  tucuman:'Tucumán', andres:'Andrés', jose:'José', tafi:'Tafí', lucia:'Lucía',
  delfin:'Delfín', marti:'Martí', rio:'Río', sali:'Salí', ines:'Inés', maria:'María',
  concepcion:'Concepción', guemes:'Güemes', mojon:'Mojón', reduccion:'Reducción',
  leon:'León', benjamin:'Benjamín', caceres:'Cáceres', capitan:'Capitán',
  estacion:'Estación', gomez:'Gómez', arboles:'Árboles', colon:'Colón'
};
const MINUSC_ES = new Set(['de','del','la','las','los']);
function tituloEs(txt){
  const t = limpio(txt);
  if(!t) return '';
  const k = sinAcento(t).toLowerCase();
  if(DICC[k]) return DICC[k];
  const tieneMin = /[a-záéíóúñ]/.test(t), tieneMay = /[A-ZÁÉÍÓÚÑ]/.test(t);
  if(tieneMin && tieneMay) return t;
  return t.toLowerCase().split(' ').map((p,i) => {
    const base = sinAcento(p).toLowerCase();
    if(PALABRA_FIX[base]) return PALABRA_FIX[base];
    if(i > 0 && MINUSC_ES.has(p)) return p;
    return p.charAt(0).toUpperCase() + p.slice(1);
  }).join(' ');
}

/* =============================================================
   1. Carga del Excel
   ============================================================= */
const elCarga = document.getElementById('carga');

/* Reconstruye la matriz de filas (encabezados + registros) desde el formato compacto:
   cada columna guarda una lista de valores distintos y, por registro, la posición
   de su valor en esa lista. */
function decodificarBase(j){
  if(!j || j.v !== 1 || !Array.isArray(j.cols) || !Array.isArray(j.cod) || !j.cod.length)
    throw new Error('el archivo de datos compacto no tiene el formato esperado');
  const nc = j.cols.length, n = j.cod[0].length;
  const dic = j.dic.map(lista => lista.map(e =>
    (e && typeof e === 'object') ? new Date(+e.d.slice(0,4), +e.d.slice(5,7)-1, +e.d.slice(8,10)) : e));
  const filas = new Array(n + 1);
  filas[0] = j.cols.slice();
  for(let r = 0; r < n; r++){
    const f = new Array(nc);
    for(let c = 0; c < nc; c++) f[c] = dic[c][j.cod[c][r]];
    filas[r + 1] = f;
  }
  return filas;
}

let _promesaSheetJS = null;
function cargarSheetJS(){
  if(window.XLSX) return Promise.resolve();
  if(!_promesaSheetJS){
    _promesaSheetJS = new Promise((ok, mal) => {
      const sc = document.createElement('script');
      sc.src = URL_SHEETJS;
      sc.onload = ok;
      sc.onerror = () => { _promesaSheetJS = null; mal(new Error('no se pudo descargar el lector de Excel (¿sin conexión?)')); };
      document.head.appendChild(sc);
    });
  }
  return _promesaSheetJS;
}

async function iniciar(){
  // 1) versión compacta
  let filas = null;
  try{
    // index.html ya inició la descarga (window.__datos); si no, se pide acá
    let j;
    if(window.__datos) j = await window.__datos;
    else { const r = await fetch(RUTA_JSON, {cache:'no-cache'}); j = r.ok ? await r.json() : null; }
    if(j) filas = decodificarBase(j);
  }catch(e){ console.warn('No se pudo usar ' + RUTA_JSON + ', se prueba con el Excel.', e); }
  if(filas){ procesarFilas(filas, RUTA_JSON); return; }

  // 2) respaldo: el Excel
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
  fr.onerror = () => pedirArchivo('No se pudo leer el archivo.');
  if(/\.json$/i.test(file.name)){
    fr.onload = () => {
      try{ procesarFilas(decodificarBase(JSON.parse(fr.result)), file.name); }
      catch(err){ pedirArchivo('El archivo no pudo interpretarse: ' + err.message); }
    };
    fr.readAsText(file);
  } else {
    fr.onload = () => procesarLibro(fr.result, file.name);
    fr.readAsArrayBuffer(file);
  }
}

/* ---------- modal "Presentación" ---------- */
(function(){
  const btnAbrir = document.getElementById('btn-presentacion');
  const btnCerrar = document.getElementById('modal-presentacion-cerrar');
  const modal = document.getElementById('modal-presentacion');
  if(!btnAbrir || !modal) return;
  let ultimoFoco = null;
  function abrir(){
    ultimoFoco = document.activeElement;
    modal.classList.add('visible');
    document.body.style.overflow = 'hidden';
    btnCerrar.focus();
  }
  function cerrar(){
    modal.classList.remove('visible');
    document.body.style.overflow = '';
    if(ultimoFoco) ultimoFoco.focus();
  }
  btnAbrir.addEventListener('click', abrir);
  btnCerrar.addEventListener('click', cerrar);
  modal.addEventListener('click', e => { if(e.target === modal) cerrar(); });
  document.addEventListener('keydown', e => { if(e.key === 'Escape' && modal.classList.contains('visible')) cerrar(); });
})();

/* ---------- detalle de departamento en el mapa (estado compartido) ---------- */
let mapaCerrarActivo = null;
document.addEventListener('keydown', e => {
  if(e.key === 'Escape' && mapaCerrarActivo) mapaCerrarActivo();
});

let DATOS = [], ANIOS = [], ACTUALIZADO = '', FECHA_MAX = null;
const DEPTO_NOMBRE = {}, LOC_NOMBRE = {};

async function procesarLibro(buf, origen){
  try{
    await cargarSheetJS();
    await new Promise(r => setTimeout(r, 30));      // deja dibujar la pantalla de carga
    const wb = XLSX.read(buf, {type:'array', cellDates:true, cellStyles:false});
    const hoja = HOJA_BASE && wb.Sheets[HOJA_BASE] ? HOJA_BASE : wb.SheetNames[0];
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[hoja], {header:1, raw:true, defval:null, blankrows:false});
    procesarFilas(filas, origen);
  }catch(err){
    console.error(err);
    pedirArchivo('El archivo no pudo interpretarse: ' + err.message);
  }
}

function procesarFilas(filas, origen){
  setTimeout(()=>{
    try{
      armarDatos(filas);
      const hasta = FECHA_MAX ? FECHA_MAX.toLocaleDateString('es-AR', {day:'numeric', month:'long', year:'numeric'}) : '';
      document.getElementById('pie-act').innerHTML =
        (hasta ? '<br>Registros hasta el <b>' + hasta + '</b>.' : '') + '<br>Base cargada desde <b>' + origen + '</b>.';
      construirInterfaz();
      elCarga.remove();
    }catch(err){
      console.error(err);
      pedirArchivo('El archivo no pudo interpretarse: ' + err.message);
    }
  }, 0);
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

  // los textos se repiten miles de veces con pocos valores distintos: se calcula cada uno una sola vez
  const mLimpio = memoizar(limpio), mLindo = memoizar(lindo), mTitulo = memoizar(tituloEs), mClave = memoizar(clave);
  const num = v => { const n = Number(v); return isFinite(n) ? n : 0; };
  const val = (f,i) => i < 0 ? '' : mLimpio(f[i]);

  DATOS = []; FECHA_MAX = null;
  for(const k in _cacheMeses) delete _cacheMeses[k];
  for(let r = 1; r < filas.length; r++){
    const f = filas[r];
    if(!f || f.every(c => c === null || c === '')) continue;
    const anio = num(f[idx.anio]);
    if(!anio) continue;

    const causa = val(f, idx.causa);
    let mes = val(f, idx.mes).toLowerCase(), dia = '';
    const fch = idx.fecha >= 0 ? f[idx.fecha] : null;
    if(fch instanceof Date && !isNaN(fch)){
      if(!FECHA_MAX || fch > FECHA_MAX) FECHA_MAX = fch;
      if(!mes) mes = MESES[fch.getMonth()];
      dia = DIAS[fch.getDay()];
    }

    const dNombre = mLindo(val(f, idx.depto)), lNombre = mTitulo(val(f, idx.localidad));
    if(dNombre && !DEPTO_NOMBRE[mClave(dNombre)]) DEPTO_NOMBRE[mClave(dNombre)] = dNombre;
    if(lNombre && !LOC_NOMBRE[mClave(lNombre)]) LOC_NOMBRE[mClave(lNombre)] = lNombre;

    DATOS.push({
      anio: String(anio),
      mes,
      dia,
      zonaHoraria: mLindo(val(f, idx.zonaHoraria)),
      depto:       mLindo(val(f, idx.depto)),
      localidad:   mTitulo(val(f, idx.localidad)),
      uurr:        mLindo(val(f, idx.uurr)),
      dependencia: mLindo(val(f, idx.dependencia)),
      zona:        mLindo(val(f, idx.zona)),
      tipoVia:     mLindo(val(f, idx.tipoVia)),
      causa:       mLindo(causa),
      catSin:      mLindo(val(f, idx.catSin)),
      tipoSin:     mLindo(val(f, idx.tipoSin)),
      p12:         mLindo(val(f, idx.p12)),
      p23:         mLindo(val(f, idx.p23)),
      sexo:        mLindo(val(f, idx.sexo)),
      rango:       mLindo(val(f, idx.rango)),
      condicion:   mLindo(val(f, idx.condicion)),
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
  {campo:'causa',       etiqueta:'Carátula de la causa',    orden:'cantidad'},
  {campo:'anio',        etiqueta:'Año',                     orden:'natural'},
  {campo:'mes',         etiqueta:'Mes',                     orden:'mes'},
  {campo:'uurr',        etiqueta:'Unidad Regional',         orden:'cantidad'},
  {campo:'depto',       etiqueta:'Departamento',            orden:'alfa'},
  {campo:'dependencia', etiqueta:'Jurisdicción policial',   orden:'cantidad'},
  {campo:'zona',        etiqueta:'Zona urbana o rural',     orden:'cantidad'},
  {campo:'zonaHoraria', etiqueta:'Franja horaria',          orden:'cantidad'},
  {campo:'tipoVia',     etiqueta:'Tipo de vía',             orden:'cantidad'},
  {campo:'tipoSin',     etiqueta:'Tipo de siniestro',       orden:'cantidad'},
  {campo:'sexo',        etiqueta:'Sexo de la víctima',      orden:'cantidad'},
  {campo:'p12',         etiqueta:'Movilidad de la víctima', orden:'cantidad'},
  {campo:'p23',         etiqueta:'Movilidad del causante',  orden:'cantidad'}
];

/* Carátulas de causa que implican víctimas fatales: cuando son las únicas
   seleccionadas en el filtro "Carátula de la causa", el mapa cambia a una
   paleta amarillo→rojo para destacar que se trata de víctimas fatales. */
const CAUSAS_FATALES = ['FALLECIMIENTO', 'HOMICIDIO CULPOSO', 'LESIONES A FALLECIMIENTO',
  'LESIONES CULPOSAS A HOMICIDIO CULPOSO'].map(lindo);

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
            ${f.tipo === 'mapa'
              ? `<div class="mapa-caja" id="mapa-${f.id}"></div>`
              : `<div class="lienzo ${f.alto || ''}"><canvas id="cv-${f.id}"></canvas></div>`}
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

/* ---- períodos comparables ----------------------------------------------------
   El año en curso (el más reciente de la base) tiene solo algunos meses cargados;
   compararlo contra el año anterior completo daba una baja falsa. Se comparan los
   mismos meses. Con EXCLUIR_MES_EN_CURSO se omite además el último mes con datos,
   que suele estar a medio cargar.                                                  */
const MES_IDX = Object.fromEntries(MESES.map((m,i) => [m, i]));
const MES_CORTO = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
const _cacheMeses = {};

/* meses (0-11) con al menos un siniestro cargado en el año, ordenados */
function mesesConDatos(anio){
  if(!_cacheMeses[anio]){
    const set = new Set();
    for(const d of DATOS) if(d.anio === anio && d.hecho && d.mes in MES_IDX) set.add(MES_IDX[d.mes]);
    _cacheMeses[anio] = [...set].sort((x,y) => x-y);
  }
  return _cacheMeses[anio];
}
const anioEnCurso = () => ANIOS[ANIOS.length - 1];

/* último mes con datos del año en curso, o null si el año está completo */
function mesEnCurso(){
  const m = mesesConDatos(anioEnCurso());
  return (m.length && m.length < 12) ? m[m.length - 1] : null;
}

/* meses que se usan para comparar el año `a` con el anterior */
function mesesComparables(a){
  let m = mesesConDatos(a).slice();
  if(a === anioEnCurso() && EXCLUIR_MES_EN_CURSO && mesEnCurso() !== null && m.length > 1) m.pop();
  return m;
}

function etiquetaPeriodo(meses){
  if(!meses.length || meses.length === 12) return '';
  const primero = MES_CORTO[meses[0]], ultimo = MES_CORTO[meses[meses.length-1]];
  return (primero === ultimo ? primero : primero + '–' + ultimo) + ' ';
}

function pintarKpis(filas, t){
  // variación entre los dos años más recientes presentes en la selección,
  // sobre los mismos meses de ambos
  const presentes = [...new Set(filas.map(d=>d.anio))].sort();
  let cmp = null;
  if(presentes.length >= 2){
    const a = presentes[presentes.length-1], b = presentes[presentes.length-2];
    const meses = mesesComparables(a);
    const set = new Set(meses);
    const enPeriodo = (d) => set.has(MES_IDX[d.mes]);
    const ta = totales(filas.filter(d => d.anio === a && enPeriodo(d)));
    const tb = totales(filas.filter(d => d.anio === b && enPeriodo(d)));
    cmp = {a, b, ta, tb, rotulo: etiquetaPeriodo(meses)};
  }
  const delta = (k) => {
    if(!cmp || !cmp.tb[k]) return '';
    const p = (cmp.ta[k] - cmp.tb[k]) / cmp.tb[k] * 100;
    const cls = p > 0 ? 'sube' : (p < 0 ? 'baja' : '');
    return `<div class="d ${cls}">${p>0?'▲':(p<0?'▼':'=')} ${Math.abs(p).toFixed(1)}% · ${cmp.rotulo}${cmp.a} vs ${cmp.b}</div>`;
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
    {id:'mes',  titulo:'Siniestros por mes', sub:'Siniestros registrados en cada mes, comparados por año.' + (EXCLUIR_MES_EN_CURSO ? ' El tramo punteado es el mes en curso, todavía incompleto.' : ''), tipo:'linea', dim:'mes', orden:'mes', ancha:true},
    {id:'dia',  titulo:'Siniestros por día de la semana', sub:'Según la fecha del hecho', tipo:'barra', dim:'dia', orden:'dia'},
    {id:'fhor', titulo:'Franja horaria', sub:'Siniestros ocurridos en horario diurno y nocturno', tipo:'barra', dim:'zonaHoraria'}
  ]},
  {titulo:'Dónde ocurren', bajada:'Distribución territorial según la jurisdicción policial y las características de la vía.', fichas:[
    {id:'uurr', titulo:'Unidades Regionales', sub:'Siniestros por Unidad Regional', tipo:'barra', dim:'uurr'},
    {id:'zona', titulo:'Zona urbana y rural', sub:'Siniestros según el ámbito donde ocurrieron', tipo:'barra', dim:'zona'},
    {id:'depto',titulo:'Departamentos', sub:'Siniestros por departamento de la provincia', tipo:'barraH', dim:'depto', orden:'alfa', alto:'alto'},
    {id:'via',  titulo:'Tipo de vía', sub:'Siniestros según la vía donde se produjeron', tipo:'barraH', dim:'tipoVia'},
    {id:'mapa', titulo:'Mapa de siniestros por departamento y localidad', sub:'Filtrado según los criterios seleccionados a la izquierda', tipo:'mapa', ancha:true},
    {id:'loc',  titulo:'Localidades con más registros', sub:'Las 15 localidades con mayor cantidad de siniestros', tipo:'barraH', dim:'localidad', tope:15, alto:'alto', ancha:true},
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

/* colores fijos por año (no por posición) para que cada año se identifique siempre
   con el mismo color aunque cambien los filtros; se eligen tonos bien diferenciables
   entre sí (evita la rampa de un solo color que resultaba confusa) */
const COLOR_ANIO = {'2023':'#7FA6C9', '2024':'#2E5F8A', '2025':'#C9971C'};
const RESERVA_ANIO = ['#5B8C6A','#A3232B','#7A5CB0','#3C556E'];
const colorPorAnio = (a) => {
  if(COLOR_ANIO[a]) return COLOR_ANIO[a];
  const extras = [...new Set(DATOS.map(d=>d.anio))].filter(y=>!COLOR_ANIO[y]).sort();
  return RESERVA_ANIO[extras.indexOf(a) % RESERVA_ANIO.length];
};

const graficos = {};
Chart.defaults.font.family = "'Archivo', system-ui, sans-serif";
Chart.defaults.font.size = 12;
Chart.defaults.color = '#3C556E';
if(window.ChartDataLabels) Chart.register(ChartDataLabels);

/* Los gráficos se dibujan de a uno por cuadro (no los ~15 de golpe): así la pantalla
   responde enseguida y los filtros no se sienten trabados. Si llega una actualización
   nueva mientras se dibujaba, la tanda anterior se abandona. */
let _tandaGraficos = 0;
function pintarGraficos(filas){
  const tanda = ++_tandaGraficos;
  const anios = [...new Set(filas.map(d=>d.anio))].sort();
  const fichas = SECCIONES.flatMap(sec => sec.fichas);
  let i = 0;
  const paso = () => {
    if(tanda !== _tandaGraficos) return;
    const t0 = performance.now();
    do{
      const f = fichas[i++];
      if(f.tipo === 'mapa') pintarMapa(f, filas); else pintarFicha(f, filas, anios);
    }while(i < fichas.length && performance.now() - t0 < 12);
    if(i < fichas.length) setTimeout(paso, 0);
  };
  paso();
}

/* =============================================================
   5-b. Mapa de Tucumán (departamentos + localidades)
   ============================================================= */
function pintarMapa(f, filas){
  const cont = document.getElementById('mapa-' + f.id);
  if(!cont) return;
  const base = filas.filter(d => d.hecho);

  const porDepto = new Map();
  const porLoc = new Map();
  const porLocPorDepto = new Map(); // deptoKey -> Map(locKey -> n)
  const porDepPorDepto = new Map(); // deptoKey -> Map(dependencia -> n), usado solo para CAPITAL
  // Mismas aperturas, pero contando PERSONAS FALLECIDAS (fallecidos en el lugar
  // + fallecidos luego) en vez de cantidad de siniestros.
  const falDepto = new Map();
  const falLoc = new Map();
  const falLocPorDepto = new Map();
  const falDepPorDepto = new Map();
  const sumar = (mapa, k, n) => mapa.set(k, (mapa.get(k)||0) + n);
  const sumarEn = (contenedor, dk, k, n) => {
    if(!contenedor.has(dk)) contenedor.set(dk, new Map());
    sumar(contenedor.get(dk), k, n);
  };
  const DEPTO_JURISDICCION = 'CAPITAL'; // departamento que se desglosa por jurisdicción en vez de por localidad
  for(const d of base){
    const dk = d.depto ? clave(d.depto) : null;
    const fa = (d.fl||0) + (d.fp||0);   // personas fallecidas del siniestro
    if(dk){ sumar(porDepto, dk, 1); sumar(falDepto, dk, fa); }
    if(d.localidad){
      const lk = clave(d.localidad);
      sumar(porLoc, lk, 1);
      sumar(falLoc, lk, fa);
    }
    if(dk && d.localidad){
      const lk = clave(d.localidad);
      sumarEn(porLocPorDepto, dk, lk, 1);
      sumarEn(falLocPorDepto, dk, lk, fa);
    }
    // Capital es una sola localidad (San Miguel de Tucumán), así que ahí no tiene
    // sentido desglosar por localidad: se guarda también por jurisdicción policial
    // para poder mostrar eso en su lugar al seleccionar el departamento.
    if(dk && d.dependencia){
      sumarEn(porDepPorDepto, dk, d.dependencia, 1);
      sumarEn(falDepPorDepto, dk, d.dependencia, fa);
    }
  }

  if(!base.length){
    cont.innerHTML = '<p class="vacio" style="height:220px">No hay registros para los filtros seleccionados.</p>';
    return;
  }

  // Modo "víctimas fatales": si el filtro de Carátula de la causa tiene
  // seleccionadas únicamente carátulas que implican fallecimiento, el mapa
  // usa una paleta amarillo→rojo en lugar de la paleta azul habitual.
  const modoFatal = seleccion.causa.size > 0 &&
    [...seleccion.causa].every(v => CAUSAS_FATALES.includes(v));
  const escala = modoFatal
    ? ['var(--mapa-fatal-0)','var(--mapa-fatal-1)','var(--mapa-fatal-2)','var(--mapa-fatal-3)','var(--mapa-fatal-4)']
    : ['var(--mapa-0)','var(--mapa-1)','var(--mapa-2)','var(--mapa-3)','var(--mapa-4)'];

  const maxDepto = Math.max(1, ...[...porDepto.values()]);
  const escalaColor = n => {
    const p = n / maxDepto;
    if(n === 0) return escala[0];
    if(p < .15) return escala[1];
    if(p < .40) return escala[2];
    if(p < .70) return escala[3];
    return escala[4];
  };

  /* Texto auxiliar: en modo víctimas fatales se agrega la cantidad de personas
     fallecidas al lado de la cantidad de siniestros. */
  const txtSin = n => `${nro(n)} siniestro${n===1?'':'s'}`;
  const txtFal = n => `${nro(n)} persona${n===1?'':'s'} fallecida${n===1?'':'s'}`;
  const txtAmbos = (n, fa) => modoFatal ? `${txtSin(n)} · ${txtFal(fa)}` : txtSin(n);
  // celda de valores de cada renglón de lista
  const celdas = (n, fa) => modoFatal
    ? `<span class="vl">${nro(n)}</span><span class="vl-fal" title="Personas fallecidas">${nro(fa)}</span>`
    : `<span class="vl">${nro(n)}</span>`;
  const encabezadoLista = modoFatal
    ? `<div class="mapa-lista-cab"><span class="et-sin">Siniestros</span><span class="et-fal">Fallecidos</span></div>`
    : '';

  const totalSinSel = base.length;
  const totalFalSel = base.reduce((a,d) => a + (d.fl||0) + (d.fp||0), 0);

  const [vx,vy,vw,vh] = MAPA_DATOS.viewBox;
  const paths = Object.entries(MAPA_DATOS.deptos).map(([k,v]) => {
    const n = porDepto.get(k) || 0;
    const fa = falDepto.get(k) || 0;
    const nombre = DEPTO_NOMBRE[k] || v.cab || k;
    const verQue = k === DEPTO_JURISDICCION ? 'jurisdicciones policiales' : 'localidades';
    return `<path d="${v.d}" fill="${escalaColor(n)}" data-depto="${k}" tabindex="0" role="button"
      aria-label="${esc(nombre)}: ${txtAmbos(n, fa)}. Tocar para ver ${verQue}."><title>${esc(nombre)}: ${txtAmbos(n, fa)}</title></path>`;
  }).join('');

  const deptosConDato = Object.entries(MAPA_DATOS.deptos)
    .map(([k,v]) => ({k, n: porDepto.get(k) || 0, fa: falDepto.get(k) || 0, nombre: DEPTO_NOMBRE[k] || v.cab || k}))
    .sort((a,b) => b.n - a.n);

  const localesConDato = Object.entries(MAPA_DATOS.localidades)
    .map(([k,v]) => ({k, v, n: porLoc.get(k) || 0, fa: falLoc.get(k) || 0}))
    .filter(o => o.n > 0)
    .sort((a,b) => b.n - a.n);
  const maxLoc = Math.max(1, ...localesConDato.map(o=>o.n));
  const rMin = 3, rMax = 22;
  const radio = n => rMin + (rMax-rMin) * Math.sqrt(n / maxLoc);

  const burbujas = localesConDato
    .map(o => {
      const nombre = LOC_NOMBRE[o.k] || o.k;
      return `<circle class="burbuja" cx="${o.v.x}" cy="${o.v.y}" r="${radio(o.n).toFixed(1)}"><title>${esc(nombre)}: ${txtAmbos(o.n, o.fa)}</title></circle>`;
    }).join('');

  cont.innerHTML = `
    <div class="mapa-svg">
      <svg viewBox="${vx} ${vy} ${vw} ${vh}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Mapa de siniestros viales por departamento y localidad. Tocá un departamento para ver sus localidades.">
        ${paths}${burbujas}
      </svg>
      <div class="mapa-detalle" id="mapa-detalle-${f.id}"></div>
    </div>
    <div class="mapa-panel">
      ${modoFatal ? `<p class="mapa-badge-fatal">
        <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><path d="M12 17h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>
        Víctimas fatales</p>` : ''}
      ${modoFatal ? `<p class="mapa-total-fatal">
        <b>${nro(totalSinSel)}</b> siniestro${totalSinSel===1?'':'s'} ·
        <b>${nro(totalFalSel)}</b> persona${totalFalSel===1?'':'s'} fallecida${totalFalSel===1?'':'s'} en la selección</p>` : ''}
      <h4>${modoFatal ? 'Siniestros y fallecidos por departamento' : 'Siniestros por departamento'}</h4>
      <div class="mapa-escala">
        <span style="background:${escala[0]}"></span><span style="background:${escala[1]}"></span>
        <span style="background:${escala[2]}"></span><span style="background:${escala[3]}"></span>
        <span style="background:${escala[4]}"></span>
      </div>
      <div class="mapa-escala-et"><span>Menos</span><span>Más (máx. ${nro(maxDepto)})</span></div>
      ${encabezadoLista}
      <ul class="mapa-lista mapa-lista-scroll">
        ${deptosConDato.map(o =>
          `<li><span class="pt" style="background:${escalaColor(o.n)}"></span><span class="nm">${esc(o.nombre)}</span>${celdas(o.n, o.fa)}</li>`
        ).join('')}
      </ul>
      <p class="mapa-ayuda">
        <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 16v-5"/><path d="M12 8h.01"/></svg>
        Tocá o hacé clic en un departamento del mapa para ver sus localidades
        (en Capital, se muestran las jurisdicciones policiales).
      </p>
      <h4 style="margin-top:18px">Localidades</h4>
      ${encabezadoLista}
      <ul class="mapa-lista mapa-lista-scroll">
        ${localesConDato.map(o =>
          `<li><span class="pt"></span><span class="nm">${esc(LOC_NOMBRE[o.k]||o.k)}</span>${celdas(o.n, o.fa)}</li>`
        ).join('') || '<li>Sin localidades geolocalizadas en esta selección.</li>'}
      </ul>
      ${modoFatal ? `<p class="mapa-nota-fatal">La primera columna es la cantidad de siniestros y la segunda, en rojo,
      la cantidad de personas fallecidas (fallecidos en el lugar + fallecidos luego).</p>` : ''}
      <p class="mapa-nota">El color de cada departamento indica su cantidad de siniestros; los círculos marcan las
      localidades con coordenadas disponibles, a mayor tamaño más registros. Límites simplificados con fines
      ilustrativos (fuente: IGN / BAHRA).</p>
    </div>`;

  /* ---- interacción: click/tap en un departamento despliega sus localidades ---- */
  const svgCont = cont.querySelector('.mapa-svg');
  const svgEl = svgCont.querySelector('svg');
  const detalle = document.getElementById('mapa-detalle-' + f.id);
  let deptoActivo = null;

  function nombreDepto(k){
    const v = MAPA_DATOS.deptos[k];
    return DEPTO_NOMBRE[k] || (v && v.cab) || k;
  }

  function cerrarDetalle(){
    deptoActivo = null;
    detalle.classList.remove('visible');
    detalle.innerHTML = '';
    svgEl.querySelectorAll('path.seleccionado').forEach(p => p.classList.remove('seleccionado'));
  }

  function abrirDetalle(k, path){
    deptoActivo = k;
    svgEl.querySelectorAll('path.seleccionado').forEach(p => p.classList.remove('seleccionado'));
    if(path) path.classList.add('seleccionado');

    const nombre = nombreDepto(k);
    const total = porDepto.get(k) || 0;
    const totalFal = falDepto.get(k) || 0;
    const esJurisdiccion = k === DEPTO_JURISDICCION;
    const falItems = (esJurisdiccion ? falDepPorDepto : falLocPorDepto).get(k) || new Map();
    const items = esJurisdiccion
      ? [...(porDepPorDepto.get(k) || new Map()).entries()]
          .map(([nombreDep, n]) => ({n, fa: falItems.get(nombreDep) || 0, nombre: nombreDep}))
      : [...(porLocPorDepto.get(k) || new Map()).entries()]
          .map(([lk, n]) => ({n, fa: falItems.get(lk) || 0, nombre: LOC_NOMBRE[lk] || lk}));
    items.sort((a,b) => b.n - a.n || a.nombre.localeCompare(b.nombre,'es'));
    const etiqueta = esJurisdiccion ? 'jurisdicción policial' : 'localidad';

    detalle.innerHTML = `
      <div class="mapa-detalle-cab">
        <h5>${esc(nombre)}</h5>
        <button type="button" class="mapa-detalle-cerrar" aria-label="Cerrar">
          <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18"/><path d="M6 6l12 12"/></svg>
        </button>
      </div>
      <p class="mapa-detalle-total">${nro(total)} siniestro${total===1?'':'s'} en el departamento${
        modoFatal ? ` · <b class="fal">${nro(totalFal)}</b> persona${totalFal===1?'':'s'} fallecida${totalFal===1?'':'s'}` : ''}</p>
      ${esJurisdiccion ? '<p class="mapa-detalle-sub">Por jurisdicción policial</p>' : ''}
      ${items.length ? encabezadoLista : ''}
      <ul class="mapa-lista mapa-lista-scroll">
        ${items.length
          ? items.map(o => `<li><span class="pt"></span><span class="nm">${esc(o.nombre)}</span>${celdas(o.n, o.fa)}</li>`).join('')
          : ''}
      </ul>
      ${items.length ? '' : `<p class="vacio">No hay ${etiqueta} registrada para los siniestros de este departamento.</p>`}
    `;
    detalle.classList.add('visible');
    detalle.querySelector('.mapa-detalle-cerrar').addEventListener('click', cerrarDetalle);
  }

  svgEl.addEventListener('click', e => {
    const path = e.target.closest('path[data-depto]');
    if(!path){ if(deptoActivo) cerrarDetalle(); return; }
    const k = path.dataset.depto;
    if(deptoActivo === k) cerrarDetalle();
    else abrirDetalle(k, path);
  });
  svgEl.addEventListener('keydown', e => {
    if(e.key !== 'Enter' && e.key !== ' ') return;
    const path = e.target.closest('path[data-depto]');
    if(!path) return;
    e.preventDefault();
    const k = path.dataset.depto;
    if(deptoActivo === k) cerrarDetalle();
    else abrirDetalle(k, path);
  });
  mapaCerrarActivo = cerrarDetalle;
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

  const horizontal = f.tipo === 'barraH';
  if(horizontal){
    // alto proporcional a la cantidad de filas y de series, para que cada barra
    // tenga lugar suficiente y se pueda leer/tocar sin apretar varias a la vez
    const alto = Math.min(920, Math.max(220, claves.length * (20 + 13 * anios.length)));
    cont.style.height = alto + 'px';
  } else {
    cont.style.height = '';
  }

  const etiquetas = claves.map(v => f.dim === 'mes' ? lindo(v) : v);
  const muchasCategorias = claves.length > 9;
  const esMensual = f.tipo === 'linea' && f.dim === 'mes';
  const mesActual = esMensual ? mesEnCurso() : null;                  // mes (0-11) en curso, si hay
  const ixActual = mesActual === null ? -1 : claves.indexOf(MESES[mesActual]);
  const datasets = anios.map((a,i) => ({
    label: a,
    // un mes que la base todavía no tiene para ese año queda vacío (no es un "0")
    data: claves.map(v => {
      if(esMensual && !mesesConDatos(a).includes(MES_IDX[v])) return null;
      return porAnio[a].get(v) || 0;
    }),
    segment: (esMensual && a === anioEnCurso() && ixActual >= 0)
      ? {borderDash: ctx => ctx.p1DataIndex === ixActual ? [5,4] : undefined} : undefined,
    pointBackgroundColor: (esMensual && a === anioEnCurso() && ixActual >= 0)
      ? (ctx => ctx.dataIndex === ixActual ? '#fff' : colorPorAnio(a)) : colorPorAnio(a),
    backgroundColor: colorPorAnio(a),
    borderColor: colorPorAnio(a),
    borderWidth: f.tipo === 'linea' ? 2.5 : 0,
    borderRadius: f.tipo === 'linea' ? 0 : 3,
    tension: .32,
    pointRadius: 2.5,
    pointHoverRadius: 5,
    fill: false,
    maxBarThickness: 46,
    // en barras horizontales el número siempre se muestra: ahí es lo que se
    // pidió poder leer sin tener que acertarle al mouse sobre la barra
    datalabels: {display: horizontal ? true : (muchasCategorias ? false : 'auto')}
  }));

  if(graficos[f.id]) graficos[f.id].destroy();
  graficos[f.id] = new Chart(lienzo, {
    type: f.tipo === 'linea' ? 'line' : 'bar',
    data: {labels: etiquetas, datasets},
    options: {
      responsive:true, maintainAspectRatio:false,
      indexAxis: horizontal ? 'y' : 'x',
      interaction:{mode:'index', intersect:false},
      animation:{duration: 320},
      // deja aire de sobra alrededor de las barras para que la etiqueta con el
      // número (que se dibuja pegada a la punta de la barra) siempre entre
      // completa y no quede cortada por el borde del gráfico
      layout:{padding: horizontal ? {top:4, right:40, bottom:4, left:4} : {top:24, right:8, bottom:4, left:4}},
      plugins:{
        legend:{display: anios.length > 1, position:'bottom',
          labels:{boxWidth:10, boxHeight:10, usePointStyle:true, pointStyle:'rectRounded', padding:14}},
        tooltip:{
          backgroundColor:'#10243A', padding:10, cornerRadius:6, titleFont:{weight:'600'},
          callbacks:{ label: c => ` ${c.dataset.label}: ${nro(c.parsed[horizontal?'x':'y'])}` }
        },
        datalabels:{
          color:'#3C556E', font:{size:9.5, weight:'600'}, formatter: v => v > 0 ? nro(v) : '',
          anchor:'end', align: horizontal ? 'end' : (f.tipo==='linea' ? 'top' : 'end'),
          offset:3, clip:false
        }
      },
      scales:{
        x:{ grid:{display:horizontal, color:'#EEF1F5'}, border:{display:false},
            // "grace" agranda el máximo del eje un poco más allá del dato más alto,
            // así la etiqueta de la barra más larga tiene lugar antes del borde
            ...(horizontal ? {grace:'14%'} : {}),
            ticks:{autoSkip:!horizontal, maxRotation: horizontal?0:45, minRotation:0,
                   callback:function(v){ const l = this.getLabelForValue(v);
                     return horizontal ? nro(v) : (String(l).length > 16 ? String(l).slice(0,15)+'…' : l); }}},
        y:{ beginAtZero:true, grid:{display:!horizontal, color:'#EEF1F5'}, border:{display:false},
            ...(horizontal ? {} : {grace:'14%'}),
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
  pintarKpis(filas, t);
  pintarGraficos(filas);
}

const btnPdf = document.getElementById('btn-pdf');
if(btnPdf) btnPdf.addEventListener('click', () => {
  const activos = SEGMENTADORES
    .filter(s => seleccion[s.campo].size)
    .map(s => `${s.etiqueta}: ${[...seleccion[s.campo]].join(', ')}`);
  const fecha = new Date().toLocaleDateString('es-AR', {day:'2-digit', month:'long', year:'numeric'});
  document.getElementById('pie-impresion').textContent =
    `Generado el ${fecha}.` + (activos.length ? ` Filtros aplicados — ${activos.join(' · ')}.` : ' Sin filtros aplicados (datos de todo el período).');
  window.print();
});

iniciar();

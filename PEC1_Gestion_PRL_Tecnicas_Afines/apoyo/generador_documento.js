// PEC 01 (Gestión de la PRL y Técnicas Afines): evaluación de riesgos del puesto en el domicilio. Formato del enunciado: A4, 2,5/2,5/3/3 cm, Calibri 11, interlineado 1,5.
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, HeadingLevel,
  Header, Footer, PageNumber, ShadingType, BorderStyle, PageBreak, VerticalAlign, ImageRun, TabStopType, LeaderType,
} = require('docx');

const OUT = process.argv[2];
const PAGES = process.argv[3] ? JSON.parse(fs.readFileSync(process.argv[3], 'utf8')) : {};
const FONT = 'Calibri', SZ = 22; // 11 pt
const GRAY = '595959', FILL = 'DCE6F1';
const PAGE_W = 11906, PAGE_H = 16838, M_TB = 1418, M_LR = 1701; // A4; 2,5 cm arriba/abajo; 3 cm izquierda/derecha
const TEXT_W = PAGE_W - 2 * M_LR; // 8504 twips
const FECHA = '24 de noviembre de 2026';

function runs(text, o = {}) {
  const out = [];
  for (const tok of text.split(/(\*\*.*?\*\*|__.*?__)/)) {
    if (!tok) continue;
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), bold: true, font: FONT, size: o.size || SZ, color: o.color }));
    else if (tok.startsWith('__')) out.push(new TextRun({ text: tok.slice(2, -2), italics: true, font: FONT, size: o.size || SZ, color: o.color }));
    else out.push(new TextRun({ text: tok, font: FONT, size: o.size || SZ, bold: o.bold, italics: o.italics, color: o.color }));
  }
  return out;
}
function P(text, o = {}) {
  return new Paragraph({ alignment: o.align || AlignmentType.JUSTIFIED, spacing: { line: 360, after: o.after === undefined ? 120 : o.after, before: o.before || 0 }, keepNext: o.keepNext, indent: o.indent, children: runs(text, o) });
}
const ENTRIES = [];
function H1(text) { ENTRIES.push([1, text]); return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text, font: FONT, size: 28, bold: true })] }); }
function H2(text) { ENTRIES.push([2, text]); return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text, font: FONT, size: 24, bold: true })] }); }
function cell(text, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: [new Paragraph({ alignment: o.align || AlignmentType.LEFT, spacing: { after: 0, line: 252 }, children: runs(text, { size: o.size || 19, bold: o.bold }) })],
  });
}
function table(widths, header, rows, o = {}) {
  const w = widths.reduce((a, b) => a + b, 0);
  const trs = [new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, widths[i], { bold: true, fill: FILL, align: AlignmentType.CENTER })) })];
  rows.forEach(r => trs.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { align: (o.aligns || [])[i] })) })));
  return new Table({ width: { size: w, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}
function caption(n, text) { return new Paragraph({ spacing: { before: 160, after: 80, line: 276 }, keepNext: true, children: [new TextRun({ text: `Tabla ${n}. `, bold: true, font: FONT, size: 20 }), new TextRun({ text, italics: true, font: FONT, size: 20 })] }); }
function note(text) { return new Paragraph({ spacing: { before: 60, after: 200, line: 252 }, alignment: AlignmentType.JUSTIFIED, children: runs(text, { size: 18, color: GRAY }) }); }
function tocLine(level, text) {
  const pg = PAGES[text] !== undefined ? String(PAGES[text]) : '';
  return new Paragraph({ spacing: { after: 60, line: 300 }, indent: { left: level === 2 ? 567 : 0 }, tabStops: [{ type: TabStopType.RIGHT, position: TEXT_W, leader: LeaderType.DOT }], children: [new TextRun({ text, font: FONT, size: SZ, bold: level === 1 }), new TextRun({ text: '\t' + pg, font: FONT, size: SZ })] });
}

// ---------------- contenido ----------------
const body = [];
const add = (...e) => e.forEach(x => body.push(x));

add(H1('1. Introducción'));
add(H2('1.1. Objeto, alcance y justificación'));
add(P('En esta PEC evalúo los riesgos del puesto de trabajo que utilizo en mi domicilio para la parte administrativa de mi actividad profesional. Trabajo como técnica en un centro de menores con medidas judiciales. La normativa interna de seguridad del centro y la protección de datos de los menores prohíben introducir dispositivos de captura de imágenes en las instalaciones, de modo que no puedo fotografiar el puesto presencial. El enunciado admite en ese caso evaluar el espacio habilitado en el domicilio, que es donde redacto los informes socioeducativos, registro expedientes, preparo documentación y participo en reuniones por videoconferencia.'));
add(P('La evaluación se limita a la zona habilitada para el trabajo, que es el alcance que fija el artículo 16 de la Ley 10/2021, de trabajo a distancia, para esta modalidad (Ley 10/2021, 2021), y sigue el procedimiento general de evaluación del artículo 16 de la Ley 31/1995 (Ley 31/1995, 1995) y de los artículos 3 a 7 del Reglamento de los Servicios de Prevención (Real Decreto 39/1997, 1997): identificar los peligros, estimar el riesgo con un método reconocido, valorarlo y planificar las medidas. He elegido dos disciplinas. Seguridad en el trabajo, porque los riesgos de accidente de un despacho doméstico (cables, instalación eléctrica, estanterías) suelen pasar desapercibidos precisamente por estar en casa. Y ergonomía, porque el trabajo con pantallas de visualización de datos es la tarea principal del puesto y el Real Decreto 488/1997 obliga a evaluar sus condiciones (Real Decreto 488/1997, 1997).'));
add(H2('1.2. Descripción del puesto y del entorno'));
add(P('El despacho ocupa una habitación del domicilio destinada a ese uso, con ventana lateral que aporta luz natural y cortina para regularla. El mobiliario se compone de una mesa de trabajo de madera con superficie mate, una silla giratoria de red con base de cinco ruedas, reposacabezas, regulación de altura y respaldo reclinable, pero sin reposabrazos regulables, y dos estanterías auxiliares para la documentación. El equipo informático es un ordenador portátil conectado a un monitor externo de gran formato, con teclado y ratón independientes, una lámpara flexo orientable y una impresora multifunción. Todos los equipos se alimentan desde una única toma de corriente mediante una regleta múltiple situada bajo la mesa, y parte del cableado discurre por el suelo entre la mesa y la pared.'));
add(P('Las tareas son la redacción de informes y registros, la gestión documental, la impresión de documentación y las reuniones por videoconferencia. Se realizan en sesiones de entre una y tres horas, varios días por semana, fuera del horario presencial del centro, lo que sitúa el uso diario del puesto por debajo de las cuatro horas. Durante la redacción atiendo con frecuencia llamadas de teléfono móvil, sin manos libres. No existe un horario fijo de teletrabajo ni un registro de las horas dedicadas, aspecto que anoto porque el artículo 16 de la Ley 10/2021 pide atender también a los factores organizativos, aunque la disciplina de psicosociología queda fuera de esta evaluación.'));
add(H2('1.3. Fotografías del puesto de trabajo'));
add(P('La Figura 1 recoge cuatro vistas del puesto: la mesa con el monitor, el teclado y la lámpara; la silla y su respaldo; la zona de la impresora y las estanterías, y el conjunto desde la puerta. Las fotografías se tomaron en el domicilio y se retocaron con una herramienta de inteligencia artificial para eliminar documentos y objetos personales visibles, sin alterar el mobiliario ni los equipos, como se declara en el anexo.', { keepNext: true }));
const img = fs.readFileSync('foto-000.jpg');
add(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 60 }, keepNext: true, children: [new ImageRun({ type: 'jpg', data: img, transformation: { width: 600, height: 295 } })] }));
add(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200, line: 276 }, children: [new TextRun({ text: 'Figura 1. ', bold: true, font: FONT, size: 20 }), new TextRun({ text: 'Vistas del puesto de trabajo en el domicilio: mesa con monitor externo, teclado y lámpara; silla giratoria con reposacabezas; zona de impresora y estanterías; vista general.', italics: true, font: FONT, size: 20 })] }));

add(H1('2. Metodología'));
add(H2('2.1. Disciplinas y herramientas del INSST elegidas'));
add(P('Para la seguridad en el trabajo utilizo el sistema simplificado de evaluación de riesgos de accidente de la NTP 330 (Instituto Nacional de Seguridad e Higiene en el Trabajo [INSHT], 1993), que es el método que propone el enunciado y el más extendido en los servicios de prevención por su sencillez: permite ordenar los riesgos por niveles y fijar prioridades de intervención sin necesidad de datos estadísticos propios. Para la ergonomía utilizo el método ROSA (Rapid Office Strain Assessment), desarrollado por Sonne et al. (2012) para puestos de oficina con pantallas, en la versión que el INSST publica en la NTP 1173 (Instituto Nacional de Seguridad y Salud en el Trabajo [INSST], 2022) y en su calculador y aplicación de 2023 (INSST, 2023). Lo elijo frente a otras herramientas del catálogo porque evalúa justo los elementos del puesto doméstico (silla, monitor, teléfono, teclado y ratón) y porque tiene en cuenta el tiempo de uso diario, que en mi caso es moderado. Como referencia normativa de las condiciones del puesto con pantallas utilizo el anexo del Real Decreto 488/1997.'));
add(H2('2.2. Sistema simplificado de la NTP 330'));
add(P('El método estima el nivel de riesgo (NR) como producto del nivel de probabilidad (NP) y del nivel de consecuencias (NC). El nivel de probabilidad se obtiene a su vez multiplicando el nivel de deficiencia (ND), que mide la magnitud de los factores de riesgo observados, por el nivel de exposición (NE), que mide la frecuencia con la que el trabajador está expuesto. La Tabla 1 recoge los valores de ND, NE y NP, y la Tabla 2 los de NC y los niveles de riesgo e intervención resultantes (INSHT, 1993).', { keepNext: true }));
add(caption(1, 'Niveles de deficiencia, de exposición y de probabilidad de la NTP 330'));
add(table([2500, 1100, 2500, 1100, 1304], ['Nivel de deficiencia (ND)', 'Valor', 'Nivel de exposición (NE)', 'Valor', 'NP = ND × NE'], [
  ['Muy deficiente', '10', 'Continuada', '4', 'Muy alta: 24 a 40'],
  ['Deficiente', '6', 'Frecuente', '3', 'Alta: 10 a 20'],
  ['Mejorable', '2', 'Ocasional', '2', 'Media: 6 a 8'],
  ['Aceptable', 'No se valora', 'Esporádica', '1', 'Baja: 2 a 4'],
], { aligns: [null, AlignmentType.CENTER, null, AlignmentType.CENTER, AlignmentType.CENTER] }));
add(note('Fuente: NTP 330 (INSHT, 1993), cuadros 3, 4 y 5.'));
add(caption(2, 'Niveles de consecuencias, de riesgo y de intervención de la NTP 330'));
add(table([2500, 900, 1500, 1300, 2304], ['Nivel de consecuencias (NC)', 'Valor', 'Nivel de riesgo (NR = NP × NC)', 'Nivel de intervención', 'Significado'], [
  ['Mortal o catastrófico', '100', '4.000 a 600', 'I', 'Situación crítica. Corrección urgente'],
  ['Muy grave', '60', '500 a 150', 'II', 'Corregir y adoptar medidas de control'],
  ['Grave', '25', '120 a 40', 'III', 'Mejorar si es posible. Justificar la intervención'],
  ['Leve', '10', '20', 'IV', 'No intervenir, salvo que un análisis más preciso lo justifique'],
], { aligns: [null, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, null] }));
add(note('Fuente: NTP 330 (INSHT, 1993), cuadros 6, 7 y 8.'));
add(H2('2.3. Método ROSA'));
add(P('El método ROSA puntúa tres grupos de elementos del puesto a partir de la observación de la postura habitual: la silla (sección A: altura del asiento, profundidad del asiento, reposabrazos y respaldo), el monitor y el teléfono (sección B) y el ratón y el teclado (sección C). Cada elemento recibe una puntuación de partida de 1 cuando coincide con la postura de referencia y sumas de uno o dos puntos por cada desviación observada, por ejemplo un teléfono sujeto entre el cuello y el hombro. A cada sección se le añade el tiempo de uso diario: más de cuatro horas suma un punto, entre una y cuatro no modifica la puntuación y menos de una hora resta uno. Las tablas A a E del método combinan las puntuaciones parciales hasta una puntuación final entre 1 y 10. El valor 1 corresponde a la postura de referencia, los valores de 2 a 4 indican un nivel bajo de riesgo y una puntuación igual o superior a 5 es el nivel de acción, que exige cambios en el puesto (INSST, 2022). Conviene subrayar, como hace la NTP 1173, que una puntuación inferior a 5 no significa ausencia de riesgo, sino un riesgo menor.'));

add(H1('3. Evaluación de seguridad en el trabajo'));
add(H2('3.1. Identificación de peligros'));
add(P('Recorrí el despacho con la lista de formas de accidente del INSST y observé cinco situaciones con deficiencias. La primera es el cableado del portátil, el monitor, la lámpara y la impresora, que cruza por el suelo la zona de paso entre la silla y la puerta (caída de personas al mismo nivel, código 020). La segunda y la tercera están en la regleta múltiple bajo la mesa, que alimenta cinco equipos desde una sola toma y a la que se conecta en ocasiones un segundo ladrón: posibilidad de contacto eléctrico por deterioro de las conexiones (código 160) y de incendio por sobrecarga o calentamiento (código 210). La cuarta es una de las estanterías de documentación, cargada en la parte alta y sin anclaje a la pared (caída de objetos por desplome, código 030). La quinta es el espacio reducido entre la mesa, la silla y la estantería, con cajones que quedan abiertos (choques contra objetos inmóviles, código 070).'));
add(H2('3.2. Valoración de los riesgos'));
add(P('La Tabla 3 aplica la NTP 330 a cada uno de los cinco riesgos. En todos los casos el nivel de deficiencia es mejorable (2), porque existen factores de riesgo de menor importancia que no anulan la eficacia de las medidas existentes, y lo que cambia es la exposición y la gravedad de las consecuencias.', { keepNext: true }));
add(caption(3, 'Evaluación de los riesgos de accidente del puesto según la NTP 330'));
add(table([2404, 900, 900, 900, 900, 900, 1600], ['Riesgo (código INSST)', 'ND', 'NE', 'NP', 'NC', 'NR', 'Nivel de intervención'], [
  ['Caída al mismo nivel por cables en la zona de paso (020)', '2', '3', '6 (media)', '10', '60', 'III. Mejorar si es posible'],
  ['Contacto eléctrico por regleta y conexiones en cadena bajo la mesa (160)', '2', '4', '8 (media)', '25', '200', 'II. Corregir y adoptar medidas de control'],
  ['Incendio por sobrecarga de la regleta con cinco equipos conectados (210)', '2', '4', '8 (media)', '60', '480', 'II. Corregir y adoptar medidas de control'],
  ['Caída de objetos por desplome de la estantería sin anclar (030)', '2', '2', '4 (baja)', '25', '100', 'III. Mejorar si es posible'],
  ['Choques contra cajones abiertos y mobiliario en espacio reducido (070)', '2', '3', '6 (media)', '10', '60', 'III. Mejorar si es posible'],
], { aligns: [null, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, null] }));
add(note('ND: nivel de deficiencia; NE: nivel de exposición; NP: nivel de probabilidad; NC: nivel de consecuencias; NR: nivel de riesgo. Valores según la NTP 330 (INSHT, 1993).'));
add(P('Los dos riesgos eléctricos alcanzan el nivel de intervención II y son los que hay que corregir primero: la exposición es continuada mientras el puesto está en uso y las consecuencias de un contacto o de un incendio en una habitación del domicilio son graves o muy graves. Los otros tres riesgos quedan en el nivel III, que la NTP 330 asocia a mejorar si es posible; en este caso la mejora es tan barata que no necesita más justificación.'));
add(H2('3.3. Medidas preventivas de seguridad'));
add(P('Para el cableado, recoger y fijar los cables con canaletas adhesivas en la pared y bridas bajo la mesa, de forma que ningún cable cruce la zona de paso, y reubicar la impresora junto a la toma de corriente. Para la instalación eléctrica, sustituir la regleta por una base múltiple con interruptor y protección contra sobretensiones, con potencia suficiente para los equipos conectados, eliminar el segundo ladrón y comprobar que el cuadro de la vivienda dispone de interruptor diferencial y que funciona al pulsar el botón de prueba; si hace falta otra toma, que la instale un instalador autorizado. Para la estantería, anclarla a la pared con escuadras y trasladar el peso a las baldas inferiores. Para el espacio, mantener los cajones cerrados, retirar las cajas del suelo y dejar libre un paso de al menos un metro entre la silla y la puerta. Estas medidas, con su prioridad y plazo, se recogen en la Tabla 5.'));

add(H1('4. Evaluación ergonómica del puesto con pantalla'));
add(H2('4.1. Puntuaciones parciales del método ROSA'));
add(P('Apliqué el método observando mi postura habitual durante una sesión de redacción y una videollamada, y comprobé las puntuaciones con el calculador del INSST. La Tabla 4 recoge la condición observada de cada elemento y la puntuación asignada.', { keepNext: true }));
add(caption(4, 'Aplicación del método ROSA al puesto: condiciones observadas y puntuaciones'));
add(table([1300, 1500, 4304, 1400], ['Sección', 'Elemento', 'Condición observada', 'Puntuación'], [
  ['A. Silla', 'Altura del asiento', 'Regulable; rodillas a unos 90 grados y pies apoyados en el suelo', '1'],
  ['A. Silla', 'Profundidad del asiento', 'Regulable; espacio de unos ocho centímetros entre el borde y la rodilla', '1'],
  ['A. Silla', 'Reposabrazos', 'No regulables y algo bajos: los antebrazos no quedan apoyados al teclear', '2'],
  ['A. Silla', 'Respaldo', 'Reclinable, con apoyo lumbar adecuado', '1'],
  ['A. Silla', 'Tabla A y tiempo de uso', 'Combinación (2, 3) = 2; uso entre una y cuatro horas diarias: sin ajuste', '2'],
  ['B. Monitor y teléfono', 'Monitor', 'Borde superior a la altura de los ojos, a la distancia de un brazo, sin reflejos', '1'],
  ['B. Monitor y teléfono', 'Teléfono', 'Móvil sujeto entre el cuello y el hombro en llamadas largas (+2) y sin manos libres (+1)', '4'],
  ['B. Monitor y teléfono', 'Tabla B', 'Combinación (1, 4)', '4'],
  ['C. Ratón y teclado', 'Ratón', 'Alineado con el hombro, sobre la misma superficie', '1'],
  ['C. Ratón y teclado', 'Teclado', 'Muñecas en ligera extensión, con la superficie algo alta', '2'],
  ['C. Ratón y teclado', 'Tabla C', 'Combinación (1, 2)', '2'],
  ['Periféricos', 'Tabla D', 'Combinación de B (4) y C (2)', '4'],
  ['Puntuación final', 'Tabla E', 'Combinación de la silla (2) y los periféricos (4)', '4'],
], { aligns: [null, null, null, AlignmentType.CENTER] }));
add(note('Puntuaciones asignadas conforme a la NTP 1173 (INSST, 2022) y al calculador del método ROSA (INSST, 2023).'));
add(H2('4.2. Resultado y comparación con el Real Decreto 488/1997'));
add(P('La puntuación final es 4: nivel de riesgo bajo, por debajo del nivel de acción de 5, pero con dos factores que elevan la puntuación de forma clara, el uso del teléfono sin manos libres y los reposabrazos. Si las sesiones superasen las cuatro horas diarias, el ajuste por tiempo de uso llevaría la puntuación al nivel de acción, de modo que el margen es pequeño. Comparado con el anexo del Real Decreto 488/1997, el puesto cumple en lo esencial: la pantalla es orientable y está a la altura y distancia adecuadas, el teclado es independiente e inclinable, la mesa tiene superficie mate y espacio suficiente, la silla es regulable y estable, y la iluminación combina luz natural regulable con una lámpara orientable que evita reflejos. Las carencias están en el apoyo de los antebrazos, en el teléfono y en la organización de las pausas, que no sigo de forma sistemática.'));
add(H2('4.3. Medidas preventivas ergonómicas'));
add(P('Utilizar auriculares con micrófono para las llamadas y las videollamadas, que es la medida de mayor efecto y menor coste. Sustituir la silla por otra con reposabrazos regulables en altura o, como alternativa inmediata, acercar la silla y bajar la altura del teclado con una bandeja extraíble para que los antebrazos descansen. Mantener el portátil sobre un soporte y el monitor externo como pantalla principal, con el borde superior a la altura de los ojos. Organizar las pausas que exige el artículo 3.3 del Real Decreto 488/1997 con cambios de actividad o descansos breves, por ejemplo diez minutos cada hora y media de trabajo continuado con la pantalla, con estiramientos de cuello y muñecas, y fijar un horario orientativo de teletrabajo con un final de jornada claro, en línea con lo que pide la Ley 10/2021 sobre descansos y desconexión.'));

add(H1('5. Planificación de la actividad preventiva'));
add(P('La Tabla 5 ordena las medidas de las dos disciplinas por prioridad, con plazo y forma de comprobación, como pide el artículo 16.2.b) de la Ley 31/1995 para la planificación de la actividad preventiva. Al tratarse de mi propio puesto en el domicilio, la ejecución y el seguimiento son responsabilidad mía, sin perjuicio de la obligación de la empresa de evaluar el puesto de trabajo a distancia conforme a la Ley 10/2021.', { keepNext: true }));
add(caption(5, 'Planificación de las medidas preventivas'));
add(table([2900, 1900, 1000, 1100, 1604], ['Medida', 'Riesgo que trata', 'Prioridad', 'Plazo', 'Comprobación'], [
  ['Base múltiple con interruptor y protección, sin conexiones en cadena; prueba del diferencial', 'Contacto eléctrico e incendio (nivel II)', '1', 'Una semana', 'Inspección visual mensual; botón de prueba del diferencial'],
  ['Canalización del cableado y reubicación de la impresora junto a la toma', 'Caída al mismo nivel (nivel III)', '1', 'Una semana', 'Zona de paso libre de cables'],
  ['Auriculares con micrófono para llamadas y videollamadas', 'Postura del cuello (ROSA, teléfono)', '1', 'Inmediato', 'Uso en todas las llamadas largas'],
  ['Anclaje de la estantería y redistribución del peso en las baldas bajas', 'Caída de objetos por desplome (nivel III)', '2', 'Un mes', 'Estantería fijada; baldas altas ligeras'],
  ['Silla con reposabrazos regulables o bandeja para el teclado', 'Apoyo de antebrazos (ROSA, silla y teclado)', '2', 'Tres meses', 'Nueva aplicación del método ROSA'],
  ['Pausas de diez minutos cada hora y media y horario orientativo de teletrabajo', 'Fatiga visual y postural; organización', '2', 'Inmediato', 'Registro semanal de sesiones y pausas'],
  ['Orden del espacio: cajones cerrados, cajas fuera del suelo, paso libre', 'Choques contra objetos (nivel III)', '3', 'Inmediato', 'Revisión al inicio de cada sesión'],
  ['Reevaluación completa del puesto', 'Todos', '3', 'Un año o ante cambios', 'Nueva aplicación de la NTP 330 y del método ROSA'],
], { aligns: [null, null, AlignmentType.CENTER, AlignmentType.CENTER, null] }));
add(note('Prioridad 1: medidas que corrigen riesgos de nivel II o de coste nulo; prioridad 2: inversión pequeña; prioridad 3: hábitos y seguimiento.'));

add(H1('6. Conclusiones'));
add(P('La evaluación muestra un puesto razonablemente bien resuelto en lo ergonómico y descuidado en lo eléctrico, que es lo habitual en los despachos domésticos: la silla, la mesa y el monitor se eligieron pensando en la postura, pero la instalación eléctrica se fue improvisando a medida que se añadían equipos. Los dos riesgos de nivel II, contacto eléctrico e incendio por la regleta, se corrigen en una semana con una base múltiple adecuada y un cableado ordenado. En ergonomía, la puntuación ROSA de 4 queda por debajo del nivel de acción, pero el teléfono sujeto con el hombro y la falta de reposabrazos son dos factores concretos y fáciles de eliminar, y el margen frente al nivel de acción desaparecería si aumentase el tiempo de uso.'));
add(P('Lo que más me ha aportado la actividad es aplicar las herramientas del INSST a un espacio que conozco de memoria y descubrir que la familiaridad esconde deficiencias que en un centro de trabajo habría detectado a la primera. También he comprobado que los dos métodos se complementan: la NTP 330 ordena bien los riesgos de accidente y obliga a razonar la gravedad de las consecuencias, mientras que el método ROSA convierte en números una postura que de otro modo se describiría con adjetivos. La planificación de la Tabla 5 fija plazos cortos porque ninguna medida requiere una inversión importante, y la reevaluación anual servirá para comprobar que se mantienen.'));

add(H1('7. Bibliografía'));
const REFS = [
  'Instituto Nacional de Seguridad e Higiene en el Trabajo. (1993). __Sistema simplificado de evaluación de riesgos de accidente__ (Nota Técnica de Prevención 330). Recuperado el ' + FECHA + ', de https://www.insst.es/documentacion/colecciones-tecnicas/ntp-notas-tecnicas-de-prevencion/9-serie-ntp-numeros-296-a-330-ano-1994/ntp-330-sistema-simplificado-de-evaluacion-de-riesgos-de-accidente-1993',
  'Instituto Nacional de Seguridad y Salud en el Trabajo. (2022). __Modelo para la evaluación de puestos de trabajo en oficina: método ROSA (Rapid Office Strain Assessment)__ (Nota Técnica de Prevención 1173). Recuperado el ' + FECHA + ', de https://www.insst.es/documentacion/colecciones-tecnicas/ntp-notas-tecnicas-de-prevencion/35-serie-ntp-numeros-1169-a-1175-ano-2022/ntp-1173-modelo-para-la-evaluacion-de-puestos-de-trabajo-en-oficina-metodo-rosa-2022',
  'Instituto Nacional de Seguridad y Salud en el Trabajo. (2023). __Calculador y App: Método ROSA (Rapid Office Strain Assessment)__ [Herramienta PRL]. Recuperado el ' + FECHA + ', de https://www.insst.es/documentacion/herramientas-de-prl/calculadores/calculador-metodo-rosa',
  'Ley 10/2021, de 9 de julio, de trabajo a distancia. __Boletín Oficial del Estado__, 164, de 10 de julio de 2021. https://www.boe.es/buscar/act.php?id=BOE-A-2021-11472',
  'Ley 31/1995, de 8 de noviembre, de Prevención de Riesgos Laborales. __Boletín Oficial del Estado__, 269, de 10 de noviembre de 1995. https://www.boe.es/buscar/act.php?id=BOE-A-1995-24292',
  'Real Decreto 39/1997, de 17 de enero, por el que se aprueba el Reglamento de los Servicios de Prevención. __Boletín Oficial del Estado__, 27, de 31 de enero de 1997. https://www.boe.es/buscar/act.php?id=BOE-A-1997-1853',
  'Real Decreto 488/1997, de 14 de abril, sobre disposiciones mínimas de seguridad y salud relativas al trabajo con equipos que incluyen pantallas de visualización. __Boletín Oficial del Estado__, 97, de 23 de abril de 1997. https://www.boe.es/buscar/act.php?id=BOE-A-1997-8671',
  'Sonne, M., Villalta, D. L., y Andrews, D. M. (2012). Development and evaluation of an office ergonomic risk checklist: ROSA (Rapid Office Strain Assessment). __Applied Ergonomics, 43__(1), 98-108. https://doi.org/10.1016/j.apergo.2011.03.008',
];
REFS.forEach(r => body.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120, line: 300 }, indent: { left: 567, hanging: 567 }, children: runs(r) })));

add(H1('8. Declaración sobre el uso de inteligencia artificial'));
add(P('Declaro que **sí** he utilizado inteligencia artificial generativa como herramienta de apoyo en la elaboración de esta PEC, en dos tareas concretas: la revisión de la redacción y de la estructura del documento, y el retoque de las fotografías del puesto para eliminar documentos y objetos personales. La identificación de los peligros, la aplicación de la NTP 330 y del método ROSA, las puntuaciones, la valoración y la planificación de las medidas son trabajo propio. El detalle figura en el Anexo I.'));
add(new Paragraph({ children: [new PageBreak()] }));
add(H1('Anexo I. Uso de inteligencia artificial: herramienta, consultas, resultados y reflexión'));
add(P('**Herramienta empleada (cita APA).** OpenAI. (2026). ChatGPT (versión de noviembre de 2026) [Modelo de lenguaje de gran tamaño]. https://chatgpt.com/'));
add(P('**Consultas realizadas (prompts).** Primera: "Revisa la redacción de este apartado de la PEC 1 para mejorar la fluidez y comprueba que la estructura sigue los apartados recomendados para una evaluación de riesgos: introducción, metodología, evaluación, medidas, conclusiones y bibliografía". Segunda: "Elimina de estas fotografías de mi despacho los documentos y objetos personales que aparecen sobre la mesa y en las estanterías, sin cambiar el mobiliario, los equipos ni la disposición de la habitación".'));
add(P('**Instrucciones utilizadas y resultados obtenidos.** En la primera consulta pedí que no modificara los datos técnicos ni las puntuaciones y que se limitara al estilo. La herramienta devolvió sugerencias de redacción (frases más cortas, enlace entre apartados) y un esquema orientativo de la memoria que coincidía con el del enunciado; adopté parte de las sugerencias de estilo y descarté las propuestas de contenido, porque introducían afirmaciones generales sin fuente. En la segunda consulta, la herramienta devolvió las imágenes con los documentos y objetos personales eliminados; comprobé que el mobiliario, los equipos y la disposición coincidían con el puesto real antes de incluirlas en la Figura 1.'));
add(P('**Reflexión sobre la aportación propia.** El uso de la herramienta fue útil para la forma del documento y para proteger datos personales en las fotografías, y poco útil para el fondo: la evaluación exige observar el puesto, medir y decidir, y eso no lo puede hacer un modelo de lenguaje. La elección de las disciplinas y de las herramientas del INSST, la identificación de los cinco riesgos de accidente, la valoración con la NTP 330, la observación de la postura y las puntuaciones del método ROSA, la comparación con el Real Decreto 488/1997, las medidas y la planificación son aportación personal, hecha a partir de mi propio puesto de trabajo y de las fuentes citadas en la bibliografía.'));

// ---------------- portada e índice ----------------
const cover = [
  new Paragraph({ spacing: { after: 1200 }, children: [] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: 'UNIVERSIDAD NACIONAL DE EDUCACIÓN A DISTANCIA (UNED)', font: FONT, size: 26, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: 'Máster Universitario en Prevención de Riesgos Laborales', font: FONT, size: 24 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1400 }, children: [new TextRun({ text: 'Asignatura: Gestión de la PRL y Técnicas Afines', font: FONT, size: 24, italics: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: 'Bloque de actividades 01 (PEC 1)', font: FONT, size: 24, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160, line: 320 }, children: [new TextRun({ text: 'Evaluación de riesgos del puesto de trabajo en el domicilio: despacho para tareas administrativas y de redacción de informes', font: FONT, size: 30, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1400 }, children: [new TextRun({ text: 'Disciplinas evaluadas: seguridad en el trabajo (NTP 330) y ergonomía (método ROSA)', font: FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 }, children: [new TextRun({ text: 'Estudiante: Milagros Moncada Roa', font: FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 }, children: [new TextRun({ text: 'Puesto evaluado: despacho en el domicilio habitual', font: FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 }, children: [new TextRun({ text: 'Perfil profesional: técnica en centro de menores con medidas judiciales', font: FONT, size: 22 })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 }, children: [new TextRun({ text: 'Curso académico 2026-2027', font: FONT, size: 22 })] }),
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: 'Índice', font: FONT, size: 28, bold: true })] }),
];
const toc = ENTRIES.map(([lvl, t]) => tocLine(lvl, t));
toc.push(new Paragraph({ children: [new PageBreak()] }));

const doc = new Document({
  creator: 'Milagros Moncada Roa', title: 'PEC 1. Evaluación de riesgos del puesto de trabajo en el domicilio',
  styles: { default: {
    document: { run: { font: FONT, size: SZ } },
    heading1: { run: { font: FONT, size: 28, bold: true, color: '000000' }, paragraph: { spacing: { before: 360, after: 160, line: 300 }, keepNext: true, outlineLevel: 0 } },
    heading2: { run: { font: FONT, size: 24, bold: true, color: '000000' }, paragraph: { spacing: { before: 240, after: 120, line: 300 }, keepNext: true, outlineLevel: 1 } },
  } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: M_TB, bottom: M_TB, left: M_LR, right: M_LR } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'PEC 01 · Gestión de la PRL y Técnicas Afines · Milagros Moncada Roa', font: FONT, size: 16, color: GRAY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Página ', font: FONT, size: 18, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GRAY }), new TextRun({ text: ' de ', font: FONT, size: 18, color: GRAY }), new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: GRAY })] })] }) },
    children: [...cover, ...toc, ...body],
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); fs.writeFileSync(OUT + '.entries.json', JSON.stringify(ENTRIES)); console.log('escrito', OUT, b.length, 'bytes'); });

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
add(P('El despacho ocupa una habitación del domicilio destinada a ese uso, con una ventana situada delante de la mesa y una persiana de lamas para regular la luz natural. El mobiliario se compone de una mesa de trabajo de madera con superficie mate y estructura metálica fija, una silla giratoria de red con base de cinco ruedas, reposacabezas, regulación de altura, respaldo reclinable y reposabrazos no regulables, un mueble de almacenamiento con estanterías abiertas para la documentación y una cajonera bajo la mesa. El equipo informático es un ordenador portátil conectado a un monitor externo de gran formato, con teclado y ratón independientes, una lámpara flexo orientable y una impresora multifunción. Los equipos se alimentan mediante regletas múltiples, parte del cableado queda suelto en la zona de paso y bajo la silla hay una alfombra de fibra natural (Figura 1).'));
add(P('Las tareas son la redacción de informes y registros, la gestión documental, la impresión de documentación y las reuniones por videoconferencia. El trabajo con la pantalla se realiza en periodos continuados de más de una hora, y de forma esporádica atiendo llamadas con el teléfono móvil sujeto entre el hombro y la cabeza, sin manos libres. El artículo 16 de la Ley 10/2021 pide atender también a los factores psicosociales y organizativos, que quedan fuera de esta evaluación porque las disciplinas elegidas son la seguridad y la ergonomía.'));
add(H2('1.3. Fotografías del puesto de trabajo'));
add(P('La Figura 1 recoge cuatro vistas del puesto: la mesa con el monitor, el teclado y la lámpara; la silla y su respaldo; la zona de la impresora y las estanterías, y el conjunto desde la puerta. Las fotografías se tomaron en el domicilio y después se modificaron con ChatGPT, como se declara en el apartado 8.', { keepNext: true }));
const img = fs.readFileSync('foto-000.jpg');
add(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 60 }, keepNext: true, children: [new ImageRun({ type: 'jpg', data: img, transformation: { width: 600, height: 295 } })] }));
add(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 200, line: 276 }, children: [new TextRun({ text: 'Figura 1. ', bold: true, font: FONT, size: 20 }), new TextRun({ text: 'Vistas del puesto de trabajo en el domicilio: mesa con monitor externo, teclado y lámpara; silla giratoria con reposacabezas; zona de impresora y estanterías; vista general.', italics: true, font: FONT, size: 20 })] }));

add(H1('2. Metodología'));
add(H2('2.1. Disciplinas y herramientas del INSST elegidas'));
add(P('Para la seguridad en el trabajo utilizo el sistema simplificado de evaluación de riesgos de accidente de la NTP 330 (Instituto Nacional de Seguridad e Higiene en el Trabajo [INSHT], 1993), que es el método que propone el enunciado y el más extendido en los servicios de prevención por su sencillez: permite ordenar los riesgos por niveles y fijar prioridades de intervención sin necesidad de datos estadísticos propios. Para la ergonomía utilizo el método ROSA (Rapid Office Strain Assessment), desarrollado por Sonne et al. (2012) para puestos de oficina con pantallas, en la versión que el INSST publica en la NTP 1173 (Instituto Nacional de Seguridad y Salud en el Trabajo [INSST], 2022) y en su calculador y aplicación de 2023 (INSST, 2023). Lo elijo frente a otras herramientas del catálogo porque evalúa justo los elementos del puesto doméstico (silla, monitor, teléfono, teclado y ratón) y porque tiene en cuenta el tiempo de uso, que en este puesto es relevante por el trabajo continuado con la pantalla. Como referencia normativa de las condiciones del puesto con pantallas utilizo el anexo del Real Decreto 488/1997.'));
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
add(P('El método ROSA puntúa tres grupos de elementos del puesto a partir de la observación de la postura habitual: la silla (sección A: altura del asiento, profundidad del asiento, reposabrazos y respaldo), el monitor y el teléfono (sección B) y el ratón y el teclado (sección C). Cada elemento recibe una puntuación de partida de 1 cuando coincide con la postura de referencia y sumas de uno o dos puntos por cada desviación observada, por ejemplo un teléfono sujeto entre el cuello y el hombro. El tiempo de uso se incorpora con la tabla F: suma un punto si el uso es continuo durante más de una hora o supera las cuatro horas diarias, y resta uno si el uso continuo no llega a treinta minutos o el total diario no alcanza una hora. En la silla se añade al resultado de la tabla A; en la pantalla, el teléfono, el ratón y el teclado, a la puntuación de cada elemento. Las tablas A a E combinan después las puntuaciones parciales hasta una puntuación final entre 1 y 10. El valor 1 corresponde a la postura de referencia, los valores de 2 a 4 indican un nivel bajo de riesgo y una puntuación igual o superior a 5 es el nivel de acción, que exige cambios en el puesto (INSST, 2022). Conviene subrayar, como hace la NTP 1173, que una puntuación inferior a 5 no significa ausencia de riesgo, sino un riesgo menor.'));

add(H1('3. Evaluación de seguridad en el trabajo'));
add(H2('3.1. Identificación de peligros'));
add(P('A partir de la observación del despacho, de las fotografías y de la lista de formas de accidente del INSST se identifican cinco situaciones con deficiencias. La primera es el cableado que queda suelto en la zona de paso, junto al borde de la alfombra situada bajo la silla, donde es fácil tropezar al levantarse o al desplazar la silla (caída de personas al mismo nivel, código 020). La segunda y la tercera dependen de las regletas múltiples que alimentan los equipos informáticos, la lámpara y la impresora, con posibilidad de conexiones en cadena: contacto eléctrico por deterioro de las conexiones (código 160) e incendio por sobrecarga o calentamiento (código 210). La cuarta es el mueble de estanterías, con carpetas y objetos en las baldas altas y sin que conste su fijación a la pared (caída de objetos por desplome, código 030). La quinta es el espacio de giro de la silla entre la mesa, la cajonera y el mueble, donde son posibles los golpes contra el mobiliario (choques contra objetos inmóviles, código 070).'));
add(H2('3.2. Valoración de los riesgos'));
add(P('La Tabla 3 aplica la NTP 330 a cada uno de los cinco riesgos. En todos los casos el nivel de deficiencia es mejorable (2), porque existen factores de riesgo de menor importancia que no anulan la eficacia de las medidas existentes, y lo que cambia es la exposición y la gravedad de las consecuencias.', { keepNext: true }));
add(caption(3, 'Evaluación de los riesgos de accidente del puesto según la NTP 330'));
add(table([2404, 900, 900, 900, 900, 900, 1600], ['Riesgo (código INSST)', 'ND', 'NE', 'NP', 'NC', 'NR', 'Nivel de intervención'], [
  ['Caída al mismo nivel por cables sueltos y borde de la alfombra en la zona de paso (020)', '2', '3', '6 (media)', '10', '60', 'III. Mejorar si es posible'],
  ['Contacto eléctrico por regletas múltiples y conexiones en cadena (160)', '2', '4', '8 (media)', '25', '200', 'II. Corregir y adoptar medidas de control'],
  ['Incendio por sobrecarga o calentamiento de las regletas (210)', '2', '4', '8 (media)', '60', '480', 'II. Corregir y adoptar medidas de control'],
  ['Caída de objetos por desplome del mueble de estanterías, sin fijación comprobada (030)', '2', '2', '4 (baja)', '25', '100', 'III. Mejorar si es posible'],
  ['Choques contra la cajonera y el mobiliario en el giro de la silla (070)', '2', '3', '6 (media)', '10', '60', 'III. Mejorar si es posible'],
], { aligns: [null, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, AlignmentType.CENTER, null] }));
add(note('ND: nivel de deficiencia; NE: nivel de exposición; NP: nivel de probabilidad; NC: nivel de consecuencias; NR: nivel de riesgo. Valores según la NTP 330 (INSHT, 1993).'));
add(P('Los dos riesgos eléctricos alcanzan el nivel de intervención II y son los que hay que corregir primero: la exposición es continuada mientras el puesto está en uso y las consecuencias de un contacto o de un incendio en una habitación del domicilio son graves o muy graves. Los otros tres riesgos quedan en el nivel III, que la NTP 330 asocia a mejorar si es posible; en este caso la mejora es tan barata que no necesita más justificación.'));
add(H2('3.3. Medidas preventivas de seguridad'));
add(P('Para el cableado y la alfombra, recoger y fijar los cables con canaletas adhesivas en la pared y bridas bajo la mesa, de forma que ningún cable cruce la zona de paso, y fijar los bordes de la alfombra con cinta antideslizante o sustituirla por una alfombrilla rígida para sillas con ruedas. Para la instalación eléctrica, sustituir las regletas por una base múltiple con interruptor y protección contra sobretensiones, con potencia suficiente para los equipos conectados, evitar las conexiones en cadena y comprobar que el cuadro de la vivienda dispone de interruptor diferencial y que funciona al pulsar el botón de prueba; si hace falta otra toma, que la instale un instalador autorizado. Para el mueble de estanterías, comprobar su fijación y, si no la tiene, anclarlo a la pared con escuadras, trasladando el peso a las baldas inferiores. Para el espacio, mantener cerrados los cajones de la cajonera y libre el radio de giro de la silla. Estas medidas, con su prioridad y plazo, se recogen en la Tabla 5.'));

add(H1('4. Evaluación ergonómica del puesto con pantalla'));
add(H2('4.1. Puntuaciones parciales del método ROSA'));
add(P('El método se aplica a la postura habitual durante el trabajo con la pantalla, con las tablas A a F de la NTP 1173. Como ese trabajo se realiza en periodos continuados de más de una hora, la tabla F suma un punto a la silla, la pantalla, el ratón y el teclado; el teléfono, de uso esporádico, resta uno. La Tabla 4 recoge la condición de cada elemento y la puntuación resultante.', { keepNext: true }));
add(table([1300, 1500, 4304, 1400], ['Sección', 'Elemento', 'Condición observada', 'Puntuación'], [
  ['A. Silla', 'Altura del asiento (A-1)', 'Regulable; rodillas a 90 grados y pies apoyados', '1'],
  ['A. Silla', 'Profundidad del asiento (A-2)', 'Postura de referencia', '1'],
  ['A. Silla', 'Reposabrazos (A-3)', 'Los codos quedan sin apoyo al teclear (2) y no son regulables (+1)', '3'],
  ['A. Silla', 'Respaldo (A-4)', 'Respaldo de red reclinable, con apoyo lumbar', '1'],
  ['A. Silla', 'Tabla A y tabla F', 'Tabla A con asiento 2 y reposabrazos más respaldo 4: 3; uso continuado de más de una hora (+1)', '4'],
  ['B. Teléfono y pantalla', 'Teléfono (B-1)', 'Postura de referencia (1), sujeto entre el hombro y la cabeza (+2), sin manos libres (+1); uso esporádico (resta 1)', '3'],
  ['B. Teléfono y pantalla', 'Pantalla (B-2)', 'A la altura de los ojos y entre 40 y 75 cm (1); uso continuado (+1)', '2'],
  ['B. Teléfono y pantalla', 'Tabla B', 'Teléfono 3 y pantalla 2', '3'],
  ['C. Ratón y teclado', 'Ratón (C-1)', 'Alineado con el hombro, en la misma superficie que el teclado (1); uso continuado (+1)', '2'],
  ['C. Ratón y teclado', 'Teclado (C-2)', 'Muñecas rectas (1); sobre la mesa fija, sin soporte ajustable (+1); uso continuado (+1)', '3'],
  ['C. Ratón y teclado', 'Tabla C', 'Ratón 2 y teclado 3', '3'],
  ['Pantalla y periféricos', 'Tabla D', 'Tabla B (3) y tabla C (3)', '3'],
  ['Puntuación final', 'Tabla E', 'Silla con tiempo de uso (4) y pantalla y periféricos (3)', '4'],
], { aligns: [null, null, null, AlignmentType.CENTER] }));
add(note('Puntuaciones obtenidas con las tablas A a F de la NTP 1173 (INSST, 2022).'));
add(H2('4.2. Resultado y comparación con el Real Decreto 488/1997'));
add(P('La puntuación final es 4: nivel de riesgo bajo, por debajo del nivel de acción de 5, con un margen de un solo punto. La determina la silla, porque los reposabrazos no regulables dejan los codos sin apoyo y el uso continuado de más de una hora suma un punto a todos los elementos. Con reposabrazos regulables a la altura de los codos, la silla bajaría a 3 y la puntuación final también a 3. El teléfono sujeto con el hombro es la postura más forzada que se observa, aunque pesa poco en el resultado porque su uso es esporádico. Comparado con el anexo del Real Decreto 488/1997, el puesto cumple en lo esencial: la pantalla es orientable y está a la altura y distancia adecuadas, el teclado es independiente, la mesa tiene superficie mate y espacio suficiente, y la silla es estable, regulable en altura y con respaldo reclinable. La carencia más clara, junto al apoyo de los antebrazos, es la orientación: la pantalla está delante de la ventana, de modo que la luz natural entra de frente en el campo visual. La persiana de lamas permite regularla, pero el anexo pide que las ventanas no provoquen deslumbramiento directo, y lo recomendable es colocar la mesa perpendicular a la ventana.'));
add(H2('4.3. Medidas preventivas ergonómicas'));
add(P('Dotar a la silla de reposabrazos regulables en altura, o sustituirla, para que los codos descansen a 90 grados al teclear: es la medida que más reduce la puntuación. Utilizar auriculares con micrófono en las llamadas y videollamadas, que elimina la postura más forzada del cuello con un coste mínimo. Girar la mesa para que la ventana quede a un lado de la pantalla y no delante, y mantener mientras tanto la persiana regulada. Organizar las pausas que exige el artículo 3.3 del Real Decreto 488/1997 con cambios de actividad o descansos breves, por ejemplo diez minutos cada hora y media de trabajo continuado con la pantalla, con estiramientos de cuello y muñecas, y fijar un horario orientativo de teletrabajo con un final de jornada claro, en línea con lo que pide la Ley 10/2021 sobre descansos y desconexión.'));

add(H1('5. Planificación de la actividad preventiva'));
add(P('La Tabla 5 ordena las medidas de las dos disciplinas por prioridad, con plazo y forma de comprobación, como pide el artículo 16.2.b) de la Ley 31/1995 para la planificación de la actividad preventiva. Al tratarse de mi propio puesto en el domicilio, la ejecución y el seguimiento son responsabilidad mía, sin perjuicio de la obligación de la empresa de evaluar el puesto de trabajo a distancia conforme a la Ley 10/2021.', { keepNext: true }));
add(caption(5, 'Planificación de las medidas preventivas'));
add(table([2900, 1900, 1000, 1100, 1604], ['Medida', 'Riesgo que trata', 'Prioridad', 'Plazo', 'Comprobación'], [
  ['Base múltiple con interruptor y protección, sin conexiones en cadena; prueba del diferencial', 'Contacto eléctrico e incendio (nivel II)', '1', 'Una semana', 'Inspección visual mensual; botón de prueba del diferencial'],
  ['Canalización del cableado y fijación o sustitución de la alfombra', 'Caída al mismo nivel (nivel III)', '1', 'Una semana', 'Zona de paso sin cables ni bordes sueltos'],
  ['Auriculares con micrófono para llamadas y videollamadas', 'Postura del cuello (ROSA, teléfono)', '1', 'Inmediato', 'Uso en todas las llamadas largas'],
  ['Comprobación de la fijación del mueble de estanterías y anclaje a la pared; peso en las baldas bajas', 'Caída de objetos por desplome (nivel III)', '2', 'Un mes', 'Mueble fijado; baldas altas ligeras'],
  ['Reposabrazos regulables en altura o sustitución de la silla', 'Apoyo de antebrazos (ROSA, silla)', '2', 'Tres meses', 'Nueva aplicación del método ROSA'],
  ['Reorientación de la mesa con la ventana a un lado de la pantalla', 'Deslumbramiento (RD 488/1997)', '2', 'Un mes', 'Sin luz natural de frente en el campo visual'],
  ['Pausas de diez minutos cada hora y media y horario orientativo de teletrabajo', 'Fatiga visual y postural; organización', '2', 'Inmediato', 'Registro semanal de sesiones y pausas'],
  ['Orden del espacio: cajones cerrados y radio de giro de la silla libre', 'Choques contra objetos (nivel III)', '3', 'Inmediato', 'Revisión al inicio de cada sesión'],
  ['Reevaluación completa del puesto', 'Todos', '3', 'Un año o ante cambios', 'Nueva aplicación de la NTP 330 y del método ROSA'],
], { aligns: [null, null, AlignmentType.CENTER, AlignmentType.CENTER, null] }));
add(note('Prioridad 1: medidas que corrigen riesgos de nivel II o de coste nulo; prioridad 2: inversión pequeña; prioridad 3: hábitos y seguimiento.'));

add(H1('6. Conclusiones'));
add(P('La evaluación muestra un puesto razonablemente bien resuelto en lo ergonómico y mejorable en lo eléctrico, que es lo habitual en los despachos domésticos: la silla, la mesa y el monitor responden a criterios posturales, mientras que la alimentación de los equipos depende de regletas múltiples. Los dos riesgos de nivel II, contacto eléctrico e incendio, se corrigen en una semana con una base múltiple adecuada y un cableado ordenado. En ergonomía, la puntuación ROSA de 4 queda por debajo del nivel de acción, a un solo punto, y la determinan los reposabrazos no regulables y el uso continuado; con reposabrazos regulables bajaría a 3. El teléfono sujeto con el hombro y la posición de la pantalla frente a la ventana son las otras dos mejoras, sencillas y baratas.'));
add(P('Lo que más me ha aportado la actividad es aplicar las herramientas del INSST a un espacio que conozco de memoria y descubrir que la familiaridad esconde deficiencias que en un centro de trabajo llamarían la atención enseguida. También he comprobado que los dos métodos se complementan: la NTP 330 ordena bien los riesgos de accidente y obliga a razonar la gravedad de las consecuencias, mientras que el método ROSA convierte en números una postura que de otro modo se describiría con adjetivos. La planificación de la Tabla 5 fija plazos cortos porque ninguna medida requiere una inversión importante, y la reevaluación anual servirá para comprobar que se mantienen.'));

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
add(P('Declaro que en la elaboración de esta PEC **sí** se ha utilizado inteligencia artificial generativa, en tres momentos: la revisión de la redacción de la primera versión y la modificación de las fotografías del puesto, con ChatGPT, y la revisión completa del documento para ajustarlo al enunciado, con Claude. El detalle de las herramientas, las consultas, los resultados y mi aportación figura en el Anexo I.'));
add(new Paragraph({ children: [new PageBreak()] }));
add(H1('Anexo I. Uso de inteligencia artificial: herramientas, consultas, resultados y reflexión'));
add(P('**Herramientas empleadas (cita APA).**', { after: 60 }));
add(P('OpenAI. (2026). ChatGPT [Modelo de lenguaje de gran tamaño]. https://chatgpt.com/', { indent: { left: 567 }, after: 40, align: AlignmentType.LEFT }));
add(P('Anthropic. (2026). Claude [Modelo de lenguaje de gran tamaño]. https://claude.ai/', { indent: { left: 567 }, after: 80, align: AlignmentType.LEFT }));
add(P('**Consultas realizadas (prompts).** A ChatGPT, sobre la primera versión del trabajo: "Revisa la redacción de este apartado de la PEC 01 para mejorar la fluidez y comprueba la estructura recomendada para la evaluación de riesgos". A ChatGPT, con las fotografías tomadas en el despacho: se le pidió que las arreglara para incluirlas en el trabajo. A Claude, para la versión final: se le facilitaron el enunciado de la PEC 1 y la primera versión del trabajo, con la petición de revisarla y corregirla para su entrega.'));
add(P('**Instrucciones utilizadas y resultados obtenidos.** ChatGPT devolvió correcciones de estilo, un esquema orientativo de la memoria y las fotografías modificadas que se incluyen en la Figura 1. Claude ajustó el formato al del enunciado (A4, márgenes, Calibri 11 e interlineado 1,5), amplió la evaluación de seguridad de dos a cinco riesgos a partir de la descripción del despacho y de las fotografías, corrigió los códigos de las formas de accidente, rehízo el cálculo del método ROSA con las tablas de la NTP 1173, añadió las tablas de la NTP 330, la comparación con el Real Decreto 488/1997, la planificación de las medidas y la bibliografía en APA 7, y redactó de nuevo la mayor parte del texto.'));
add(P('**Reflexión sobre la aportación propia.** La elección del puesto, de las disciplinas y de los métodos, las fotografías, la descripción del despacho y la primera evaluación, con los riesgos de cableado y de regletas, las medidas iniciales y la puntuación ROSA de 4, son míos. La revisión con inteligencia artificial ha sido útil para completar la metodología, las tablas y las referencias y para corregir errores de la primera versión, pero no puede observar el puesto ni comprobar las condiciones reales. Por eso he revisado la versión final y he comprobado que lo que describe se corresponde con mi despacho.'));

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

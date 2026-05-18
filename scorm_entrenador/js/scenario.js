/* ============================================================
   SCENARIO – Dimensiones del Entrenador Profesional · C-ENT-001
   ~30 escenas · Quiz + Decisiones · Puntuación máxima: 100 pts
   Idioma: castellano
   ============================================================ */

const CHARACTERS = {
  alex:      { name: 'Alex',     color: '#2563EB', initials: 'A',  shape: 'circle' },
  carmen:    { name: 'Carmen',   color: '#059669', initials: 'Ca', shape: 'circle' },
  ruben:     { name: 'Rubén',    color: '#7C3AED', initials: 'R',  shape: 'circle' },
  marcos:    { name: 'Marcos',   color: '#D97706', initials: 'M',  shape: 'circle' },
  narracio:  { name: 'Narrador', color: '#6B7280', initials: '✦',  shape: 'square' }
};

const JOURNEY_STAGES = [
  { id: 'act1',  label: 'Acto I',    title: 'El Mundo Ordinario',  scenes: ['scene_01','scene_02','scene_03','scene_04','scene_05','scene_06','scene_07','scene_08'], color: '#2563EB' },
  { id: 'act2a', label: 'Acto II-A', title: 'El Primer Umbral',    scenes: ['scene_09','scene_10','scene_11','scene_12','scene_13','scene_14'],                      color: '#0891B2' },
  { id: 'act2b', label: 'Acto II-B', title: 'La Prueba Suprema',   scenes: ['scene_15','scene_16','scene_17','scene_18','scene_19','scene_20'],                      color: '#7C3AED' },
  { id: 'act2c', label: 'Acto II-C', title: 'La Recompensa',       scenes: ['scene_21','scene_22','scene_23'],                                                        color: '#059669' },
  { id: 'act3',  label: 'Acto III',  title: 'El Retorno',          scenes: ['scene_24','scene_25','scene_26','scene_26b','scene_27','scene_28','scene_29'],           color: '#334155' },
  { id: 'final', label: 'Epílogo',   title: 'El Nuevo Alex',       scenes: ['scene_30'],                                                                              color: '#2563EB' }
];

const CANONICAL_SCENE_COUNT = 28;

/* ============================================================
   ESCENAS
   ============================================================ */
const scenes = {

  /* ══════════════════════════════════
     ACTO I – EL MUNDO ORDINARIO
  ══════════════════════════════════ */

  scene_01: {
    id: 'scene_01', etapa: 'El Mundo Ordinario', personatge: 'alex',
    tipus: 'text_block', titol: 'El primer día en CD Olimpia',
    imatge: null,
    narracio: `Alex Soria tiene 28 años, el título UEFA B recién obtenido y una carpeta llena de planillas de periodización. Hoy es su primer día como entrenador principal de CD Olimpia, un club de Segunda División B situado en una ciudad de tamaño mediano.

El campo de entrenamiento es sencillo pero bien mantenido. Dos canchas de hierba, un gimnasio funcional y un vestuario que huele a décadas de historia. Rubén Molina, el director deportivo, le da la bienvenida en la puerta: "Te hemos contratado porque conoces el juego. Ahora tienes que aprender a gestionar personas."`,
    dialeg: { personatge: 'alex', text: '"Sé cómo diseñar una sesión de entrenamiento. Lo que no sé todavía es cómo hacer que treinta personas distintas quieran seguir el mismo plan."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Las cuatro dimensiones del entrenador profesional',
      text: 'Un entrenador de alto rendimiento no es solo un **técnico especializado**: es también responsable del **proyecto deportivo** del equipo, actúa como **gestor de recursos humanos** y es el principal embajador del club en las **relaciones públicas**. Estas cuatro dimensiones funcionan de forma integrada y se refuerzan mutuamente.'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02', etapa: 'El Mundo Ordinario', personatge: 'narracio',
    tipus: 'text_block', titol: 'El equipo técnico: Carmen y Marcos',
    imatge: null,
    narracio: `Rubén presenta a Alex al cuerpo técnico. Carmen Rivas lleva 15 años en clubes de esta categoría. Cuando estrecha la mano de Alex, le mira a los ojos y dice con calma: "El primer mes es para escuchar, no para demostrar." Su reputación en el vestuario es de directa, justa y siempre disponible.

Marcos Vidal, capitán y centrocampista de 34 años, representa otro tipo de referente. Lleva ocho temporadas en CD Olimpia. Cuando Rubén lo presenta a Alex, lo mira durante unos segundos antes de decir: "Espero que te guste el trabajo duro." No es hostilidad. Es una prueba.`,
    dialeg: { personatge: 'narracio', text: '"Dos estilos. Carmen representa el liderazgo relacional y la escucha activa. Marcos, la autoridad construida sobre el rendimiento. Alex tendrá que aprender a integrar ambos modelos."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Liderazgo técnico vs. liderazgo relacional',
      text: 'El entrenador **técnico** domina el contenido: sistemas de juego, ejercicios, análisis táctico. El entrenador **relacional** construye confianza, gestiona emociones y crea cultura de equipo. El entrenador profesional completo integra ambas dimensiones sin sacrificar ninguna de las dos.'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03', etapa: 'El Mundo Ordinario', personatge: 'narracio',
    tipus: 'quiz', titol: 'Comprensión: las dimensiones del entrenador',
    imatge: null,
    narracio: `Reflexiona sobre lo que acabas de leer sobre las cuatro dimensiones del entrenador profesional.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: '¿Cuál de las siguientes afirmaciones describe mejor la dimensión de "gestor de recursos humanos" del entrenador profesional?',
    opcions: [
      { id: 'A', text: 'Es la capacidad de diseñar sesiones de entrenamiento adaptadas al nivel técnico del equipo', correcta: false, feedback: 'La planificación técnica pertenece a la dimensión de **técnico especializado**. La dimensión de RRHH tiene que ver con las personas, no con el contenido de las sesiones.' },
      { id: 'B', text: 'Es la capacidad de gestionar las relaciones con medios de comunicación y patrocinadores del club', correcta: false, feedback: 'Esa es la dimensión de **relaciones públicas**. La gestión de RRHH se enfoca en las personas del equipo: jugadores, cuerpo técnico y staff.' },
      { id: 'C', text: 'Es la capacidad de gestionar la dinámica humana del equipo: motivación, conflictos, cohesión y desarrollo individual de cada jugador', correcta: true, feedback: 'Exacto. La dimensión de gestor de RRHH incluye la **motivación individualizada**, la **gestión de conflictos**, el **desarrollo personal** de cada jugador y la **cohesión del colectivo**. Es la dimensión más compleja y la menos trabajada en la formación técnica convencional.' }
    ],
    seguent: 'scene_04'
  },

  scene_04: {
    id: 'scene_04', etapa: 'El Mundo Ordinario', personatge: 'ruben',
    tipus: 'text_block', titol: 'El proyecto deportivo del club',
    imatge: null,
    narracio: `En el despacho de Rubén, antes de la primera semana de pretemporada. El director abre un documento en el ordenador: "Este es el proyecto deportivo de CD Olimpia para los próximos tres años. Necesito que lo interiorices antes de hablar con el equipo."

El proyecto incluye tres ejes: mantener la categoría este año, desarrollar tres jugadores de la cantera para el primer equipo en dos temporadas, y generar una identidad de juego reconocible. "No te pido resultados inmediatos. Te pido un proceso con criterio y coherencia."`,
    dialeg: { personatge: 'ruben', text: '"Un entrenador sin proyecto es como un equipo sin sistema: puede funcionar bien durante un partido, pero no puede sostenerse en el tiempo."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'La dimensión de proyecto deportivo',
      text: 'El proyecto deportivo es el **marco estratégico** que da coherencia a todas las decisiones del entrenador: a quién fichar, cómo entrenar, qué valores promover y cómo comunicar los objetivos. Sin proyecto, las decisiones son reactivas. Con proyecto, son **estratégicas** y sostenibles a largo plazo.'
    },
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05', etapa: 'El Mundo Ordinario', personatge: 'alex',
    tipus: 'decision_scenario', titol: 'La primera reunión con el equipo',
    imatge: null,
    narracio: `Primera reunión de pretemporada. Treinta jugadores en el vestuario. Alex tiene la palabra. Tres opciones.`,
    dialeg: { personatge: 'alex', text: '"Esto es real. Treinta personas esperan que les diga quién soy y qué vamos a hacer. La impresión del primer día no se puede repetir."' },
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Presentar directamente su metodología y el sistema de juego sin preguntar a los jugadores',
        feedback: 'Decisión subóptima. Empezar con contenido técnico sin escucha activa genera una comunicación de **una sola vía**. Los jugadores perciben que el entrenador ya tiene todo decidido antes de conocerlos. La primera reunión debe construir **relación**, no solo transmitir información técnica.',
        punts: 0, seguent: 'scene_06'
      },
      {
        id: 'B', text: 'Iniciar escuchando: presentarse brevemente y preguntar al equipo qué espera de la temporada',
        feedback: '**Decisión óptima.** Empezar escuchando demuestra que el equipo importa como sujeto activo, no solo como receptor de instrucciones. Activa la dimensión de RRHH desde el primer momento y crea las bases de un liderazgo legítimo. La escucha en la primera reunión es el gesto de liderazgo más potente que puede hacer un entrenador nuevo.',
        punts: 10, seguent: 'scene_07'
      },
      {
        id: 'C', text: 'Delegar la reunión a Carmen para observar cómo reacciona el equipo antes de intervenir',
        feedback: 'Posición pasiva. Aunque observar puede parecer estratégico, delegar la primera reunión envía el mensaje de que el entrenador no está preparado para liderar. La primera reunión define quién asume la responsabilidad del equipo.',
        punts: 5, seguent: 'scene_07'
      }
    ]
  },

  scene_06: {
    id: 'scene_06', etapa: 'El Mundo Ordinario', personatge: 'narracio',
    tipus: 'text_block', titol: 'Las consecuencias de no escuchar',
    imatge: null,
    narracio: `Alex habla durante cuarenta minutos sobre su metodología, los sistemas tácticos que quiere implementar y las exigencias físicas de la pretemporada. Los jugadores escuchan en silencio.

Cuando termina, Marcos se inclina hacia el jugador de al lado y dice en voz baja: "No ha preguntado ni una vez cómo estamos." Carmen lo observa desde el fondo del vestuario y toma nota.`,
    dialeg: { personatge: 'narracio', text: '"La comunicación unidireccional activa la resistencia pasiva: los jugadores no protestan, pero tampoco se comprometen. El **compromiso** nace de sentirse escuchado, no solo informado."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Barreras comunicativas en el deporte de equipo',
      text: 'Las barreras más frecuentes en la primera comunicación entrenador–equipo: **monólogo técnico** (solo habla el entrenador), **asimetría de poder** (no se invita a responder) y **ausencia de feedback** (no hay confirmación de comprensión ni de acuerdo emocional).'
    },
    seguent: 'scene_07'
  },

  scene_07: {
    id: 'scene_07', etapa: 'La Llamada a la Aventura', personatge: 'carmen',
    tipus: 'worked_example', titol: 'El modelo de Carmen: comunicación eficaz desde el primer día',
    imatge: null,
    narracio: `Después de la reunión, Carmen lleva a Alex a la sala de vídeo. "El problema no es lo que has dicho. Es lo que no has preguntado." Le explica su modelo de las cinco fases para la primera comunicación con un equipo nuevo.

Hace un role-play: actúa como el capitán Marcos mientras Alex practica la apertura empática. Cuando Alex termina, Carmen dice: "Bien. Ahora di su nombre. Pregunta. Luego habla."`,
    dialeg: { personatge: 'carmen', text: '"Un entrenador que escucha en la primera reunión no muestra debilidad. Muestra que sabe que el equipo sabe cosas que él todavía no sabe."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: '5 fases para la primera reunión con el equipo',
      text: `**1. Presentación personal:** brevedad y autenticidad. Quién eres y por qué estás aquí.
**2. Escucha colectiva:** "¿Qué espera el equipo de esta temporada?" Cede la palabra al capitán.
**3. Reconocimiento:** valora lo que el equipo ha construido antes de tu llegada.
**4. Marco compartido:** presenta el proyecto deportivo como propuesta, no como imposición.
**5. Cierre con acción:** termina con un compromiso concreto y los próximos pasos claros.`
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08', etapa: 'La Llamada a la Aventura', personatge: 'carmen',
    tipus: 'quiz', titol: 'Comprensión: comunicación eficaz',
    imatge: null,
    narracio: `Carmen te hace una pregunta para asegurarse de que has entendido la fase 3 del modelo.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: '¿Por qué es importante reconocer lo que el equipo ha construido antes de tu llegada como nuevo entrenador?',
    opcions: [
      { id: 'A', text: 'Para parecer humilde y ganarse la simpatía del vestuario', correcta: false, feedback: 'La humildad tiene valor, pero no es el propósito central del reconocimiento. La función estratégica va más allá de la simpatía.' },
      { id: 'B', text: 'Porque sin ese reconocimiento, el equipo interpretará que el nuevo entrenador desvaloriza su historia, lo que genera resistencia al cambio', correcta: true, feedback: 'Exacto. Cuando un entrenador llega con un proyecto nuevo sin reconocer el pasado del equipo, activa un mecanismo de **defensa colectiva**: el equipo siente que debe proteger su identidad. El reconocimiento reduce esa resistencia y abre el espacio para el cambio.' },
      { id: 'C', text: 'Para retrasar la presentación de la metodología y ganar tiempo de adaptación', correcta: false, feedback: 'El reconocimiento no es una táctica de dilación: es una necesidad comunicativa real que tiene función propia, independientemente del tiempo que ocupe.' }
    ],
    seguent: 'scene_09'
  },

  /* ══════════════════════════════════
     ACTO II-A – EL PRIMER UMBRAL
  ══════════════════════════════════ */

  scene_09: {
    id: 'scene_09', etapa: 'Cruce del Primer Umbral', personatge: 'alex',
    tipus: 'text_block', titol: 'El primer entrenamiento',
    imatge: null,
    narracio: `Tercera semana de pretemporada. Alex ha diseñado un bloque de trabajo defensivo muy intenso: pressing alto, recuperación rápida y trabajo posicional en bloque. El volumen es elevado pero justificado técnicamente.

Pablo Sanz, mediapunta titular y uno de los más creativos del equipo, para el ejercicio y dice en voz alta: "Llevamos veinte minutos de presión sin tocar el balón. ¿Cuándo entrenamos el juego?" Tres jugadores asienten. El entrenamiento se ha detenido.`,
    dialeg: { personatge: 'alex', text: '"Necesito responder ahora. Treinta personas me observan. Lo que diga en este momento definirá mi autoridad durante meses."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Gestión de quejas individuales vs. necesidades colectivas',
      text: 'Una queja individual en un entrenamiento no es solo el problema de esa persona: es una señal del **estado emocional del grupo**. El entrenador debe evaluar si la queja es legítima —y responder al contenido— o si es una prueba de autoridad —y responder al proceso, no al contenido—.'
    },
    seguent: 'scene_10'
  },

  scene_10: {
    id: 'scene_10', etapa: 'Cruce del Primer Umbral', personatge: 'ruben',
    tipus: 'decision_scenario', titol: 'El consejo de Rubén vs. el modelo de Carmen',
    imatge: null,
    narracio: `Antes del entrenamiento, Rubén le había dicho a Alex: "No dejes que te cuestionen en público. Aquí el que manda eres tú." Carmen le había dicho algo distinto: "Si un jugador para el ejercicio, es que todavía confía en que le escucharás."`,
    dialeg: { personatge: 'ruben', text: '"Que quede claro quién decide desde el primer día. Si cedes un centímetro, perderás el vestuario en una semana."' },
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Seguir el consejo de Rubén: responder con firmeza que el plan está diseñado por razones técnicas y que se continúa',
        feedback: 'Aplicar el enfoque autoritario genera cumplimiento inmediato pero cierra el diálogo. Pablo y los jugadores que asintieron aprenderán que preguntar tiene coste. Eso no elimina la duda: la suprime, y las dudas suprimidas reaparecen con más fuerza.',
        punts: 0, seguent: 'scene_11'
      },
      {
        id: 'B', text: 'Aplicar el modelo de Carmen: agradecer la pregunta, explicar el propósito del bloque y preguntar si tienen alguna propuesta',
        feedback: '**Decisión óptima.** Responder con transparencia sobre el **por qué** del diseño técnico activa la dimensión técnica y la relacional simultáneamente. Los jugadores entienden el razonamiento y sienten que su opinión importa. La autoridad no se impone: se construye con criterio y escucha.',
        punts: 10, seguent: 'scene_12'
      }
    ]
  },

  scene_11: {
    id: 'scene_11', etapa: 'Cruce del Primer Umbral', personatge: 'narracio',
    tipus: 'text_block', titol: 'La autoridad sin diálogo',
    imatge: null,
    narracio: `Alex responde con firmeza: "El plan está diseñado por razones técnicas. Seguimos." Pablo encoge los hombros. El entrenamiento continúa, pero el ritmo baja visiblemente. A los quince minutos, varios jugadores ejecutan los ejercicios de forma mecánica, sin intensidad.

Carmen anota en su bloc: "Cumplimiento sin compromiso. El equipo ejecuta pero no entiende por qué."`,
    dialeg: { personatge: 'narracio', text: '"La autoridad sin explicación genera **cumplimiento**, no **compromiso**. Los jugadores hacen lo que se les pide, pero sin la energía y la intención que requiere el rendimiento de alto nivel."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Autoridad y legitimidad en el deporte de equipo',
      text: 'La autoridad del entrenador puede ser **formal** (viene del cargo) o **legítima** (viene del respeto ganado). La autoridad formal funciona a corto plazo. La legitimidad construida sobre la transparencia y la escucha genera rendimiento sostenido y cohesión real.'
    },
    seguent: 'scene_12'
  },

  scene_12: {
    id: 'scene_12', etapa: 'Pruebas, Aliados y Enemigos', personatge: 'alex',
    tipus: 'text_block', titol: 'La rueda de prensa',
    imatge: null,
    narracio: `Primer partido de pretemporada. Resultado: 1-1. En la sala de prensa, un periodista hace la pregunta que Alex llevaba días anticipando:

"¿Cuál es el objetivo real de CD Olimpia esta temporada? ¿Salvación, playoffs o algo más?" Dos cámaras, cuatro periodistas, el presidente del club al fondo de la sala. La respuesta saldrá en tres medios locales y en las redes del club.`,
    dialeg: { personatge: 'alex', text: '"Esta pregunta no es solo deportiva. Es una pregunta sobre el proyecto. Lo que diga ahora afectará a los patrocinadores, a los jugadores y a la directiva."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'La dimensión de relaciones públicas del entrenador',
      text: 'El entrenador es el **primer portavoz** de la identidad del club ante los medios. Sus declaraciones afectan a la **imagen institucional**, a las **expectativas de los aficionados** y a la **moral del vestuario**. Una comunicación mediática eficaz requiere: mensaje alineado con el proyecto, honestidad sin exceso de promesas y coherencia entre lo que se dice y lo que se hace.'
    },
    seguent: 'scene_13'
  },

  scene_13: {
    id: 'scene_13', etapa: 'Pruebas, Aliados y Enemigos', personatge: 'alex',
    tipus: 'decision_scenario', titol: '¿Cómo gestiona Alex la presión mediática?',
    imatge: null,
    narracio: `Alex tiene la palabra. Cuatro periodistas esperan. El silencio dura tres segundos.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"El objetivo es llegar lo más alto posible. Tenemos equipo para pelear por el playoff."',
        feedback: 'Promesa sin respaldo en el análisis realista del plantel. Genera expectativas que el equipo puede no cumplir, creando presión externa contraproducente. Si el equipo no llega al playoff, el entrenador habrá dado herramientas públicas para cuestionarlo.',
        punts: 0, seguent: 'scene_14'
      },
      {
        id: 'B', text: '"Nuestro objetivo es construir un proyecto sólido. El primer año es de consolidación: quiero que el equipo entienda mi metodología y compita con identidad clara. Los resultados vendrán cuando el proceso esté bien asentado."',
        feedback: '**Comunicación institucional óptima.** Alinea el mensaje con el proyecto real del club, gestiona las expectativas sin prometer lo que no se puede garantizar y transmite criterio profesional. Esta respuesta protege al entrenador, al equipo y al club.',
        punts: 10, seguent: 'scene_14'
      },
      {
        id: 'C', text: '"Todavía es pronto para hablar de objetivos. Prefiero no adelantar nada hasta ver cómo evoluciona la pretemporada."',
        feedback: 'Respuesta evasiva que genera incertidumbre. Los medios interpretarán la evasión como falta de proyecto o de confianza. Una respuesta sin contenido puede ser más dañina que una respuesta honestamente cautelosa.',
        punts: 5, seguent: 'scene_14'
      }
    ]
  },

  scene_14: {
    id: 'scene_14', etapa: 'Pruebas, Aliados y Enemigos', personatge: 'carmen',
    tipus: 'quiz', titol: 'Comprensión: relaciones públicas y comunicación',
    imatge: null,
    narracio: `Carmen te hace una pregunta sobre lo que ha pasado en la rueda de prensa.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: '¿Por qué una respuesta evasiva ante los medios puede ser más perjudicial que una respuesta honesta y cautelosa?',
    opcions: [
      { id: 'A', text: 'Porque los periodistas siempre interpretan el silencio como algo negativo y lo publican así', correcta: false, feedback: 'El silencio no siempre es negativo. El problema no es el silencio, sino la evasión estructural: la ausencia de cualquier contenido proyectual.' },
      { id: 'B', text: 'Porque la evasión transmite ausencia de proyecto, mientras que una respuesta honestamente cautelosa transmite criterio y control de la situación', correcta: true, feedback: '**Exacto.** "No sé todavía" puede comunicar honestidad. "Prefiero no decir nada" comunica que hay algo que no se quiere revelar o que aún no existe. La distinción es sutil pero decisiva en la percepción mediática e institucional.' },
      { id: 'C', text: 'Porque los medios tienen más poder que el entrenador y siempre van a buscar lo negativo', correcta: false, feedback: 'El poder mediático es real, pero no es absoluto. Un entrenador con criterio claro puede gestionar la narrativa mediática sin necesidad de temer a los periodistas.' }
    ],
    seguent: 'scene_15'
  },

  /* ══════════════════════════════════
     ACTO II-B – LA PRUEBA SUPREMA
  ══════════════════════════════════ */

  scene_15: {
    id: 'scene_15', etapa: 'Aproximación a la Cueva', personatge: 'carmen',
    tipus: 'worked_example', titol: 'Preparación de la conversación difícil',
    imatge: null,
    narracio: `Semana cinco de competición. Marcos Vidal ha visto reducido su tiempo de juego sin que Alex le haya explicado el motivo. En el siguiente entrenamiento, Marcos hace un comentario en voz alta que Alex no puede ignorar: "Si el entrenador no confía en los que llevamos aquí ocho años, ¿en quién confía?"

Carmen lleva a Alex a la sala de reuniones: "Tienes que hablar con Marcos hoy. Y tienes que hacerlo bien."`,
    dialeg: { personatge: 'carmen', text: '"Una conversación difícil mal preparada puede destruir una relación que tardaste meses en construir. Bien preparada, puede convertirse en la base de una confianza mucho más sólida."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: 'Estructura de una conversación difícil con un jugador',
      text: `**1. Objetivo:** ¿Qué quiero conseguir? (entendernos, no ganar)
**2. Anticipación:** ¿Cómo puede estar sintiéndose Marcos en este momento?
**3. Espacio y momento:** sala privada, sin interrupciones, suficiente tiempo.
**4. Apertura empática:** reconocer su perspectiva antes de explicar la tuya.
**5. Escucha activa:** dejar hablar, reformular, confirmar que has entendido.`
    },
    seguent: 'scene_16'
  },

  scene_16: {
    id: 'scene_16', etapa: 'Aproximación a la Cueva', personatge: 'alex',
    tipus: 'checklist', titol: 'Prepara la conversación con Marcos',
    imatge: null,
    narracio: `Alex se sienta con un folio en blanco. Carmen le ha dicho que hay cinco elementos esenciales para preparar cualquier conversación difícil con un jugador. Marca los que crees que deberías incluir en tu preparación. Selecciona todos los correctos.`,
    dialeg: { personatge: 'alex', text: '"No es una conversación de cinco minutos. Es la conversación que puede cambiar cómo Marcos me ve como entrenador."' },
    contingut_pedagogic: null,
    checklistItems: [
      { id: 'cl1', text: 'Definir cuál es mi objetivo real: entendernos, no demostrar que tengo razón', correcta: true },
      { id: 'cl2', text: 'Elegir un momento tranquilo y un espacio privado, fuera del vestuario', correcta: true },
      { id: 'cl3', text: 'Preparar una apertura que reconozca la perspectiva de Marcos', correcta: true },
      { id: 'cl4', text: 'Practicar la escucha activa: dejar hablar sin interrumpir y reformular', correcta: true },
      { id: 'cl5', text: 'Establecer un tono asertivo: decir lo que pienso sin atacar ni ceder innecesariamente', correcta: true },
      { id: 'cl6', text: 'Preparar una lista de todos los errores que he cometido para disculparme por cada uno', correcta: false },
      { id: 'cl7', text: 'Pedir a Rubén que esté presente como árbitro por si Marcos se enfada', correcta: false }
    ],
    punts_per_item: 2,
    seguent: 'scene_17'
  },

  scene_17: {
    id: 'scene_17', etapa: 'La Prueba Suprema', personatge: 'marcos',
    tipus: 'text_block', titol: 'La conversación: el contexto',
    imatge: null,
    narracio: `Sala de reuniones. Puerta cerrada. Dos sillas, una mesa estrecha. Alex ha convocado a Marcos con un mensaje directo: "Quiero explicarte mis decisiones y escuchar lo que piensas. ¿Mañana después del entrenamiento?"

Marcos llega puntual pero tenso. Se sienta con los brazos cruzados. Alex respira. Ahora empieza la prueba suprema de la dimensión de gestor de personas.`,
    dialeg: { personatge: 'narracio', text: '"La prueba suprema no es táctica ni física: es emocional. Cuando estás bajo presión y la otra persona está enfadada, tu capacidad de regulación emocional determina si la conversación construye o destruye."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Regulación emocional bajo presión competitiva',
      text: 'En situaciones de tensión, el cerebro activa la respuesta de amenaza, dificultando el pensamiento racional. Técnicas: **respiración profunda**, pausa consciente ("necesito un segundo") y **reformulación cognitiva** ("su enfado no es un ataque personal: es una necesidad no satisfecha").'
    },
    seguent: 'scene_18'
  },

  scene_18: {
    id: 'scene_18', etapa: 'La Prueba Suprema – Inicio', personatge: 'marcos',
    tipus: 'decision_scenario', titol: 'Decisión 1: ¿Cómo inicia Alex la conversación?',
    imatge: null,
    narracio: `Marcos espera. Alex tiene la palabra. Primer movimiento.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"Marcos, entiendo que puede ser frustrante no tener el rol que esperabas. Quiero explicarte el razonamiento de mis decisiones y escuchar tu perspectiva."',
        feedback: '**Apertura empática perfecta.** Alex reconoce la emoción de Marcos antes de explicar el contenido. Esta secuencia —primero la emoción, luego el razonamiento— es la estructura central de la comunicación asertiva en situaciones de tensión interpersonal.',
        punts: 10, seguent: 'scene_19'
      },
      {
        id: 'B', text: '"Marcos, quiero explicarte por qué he tomado estas decisiones técnicas. Tienen una lógica clara."',
        feedback: 'Correcto pero incompleto. Explicar sin reconocer primero la perspectiva de Marcos pone el razonamiento por delante de la emoción. La comunicación efectiva en conflictos requiere que la otra persona se sienta escuchada antes de que puedas ser escuchado tú.',
        punts: 5, seguent: 'scene_19'
      },
      {
        id: 'C', text: '"Marcos, sé que tu comentario en el entrenamiento no fue el más profesional, pero estoy aquí para aclarar las cosas."',
        feedback: 'Inicio **defensivo y acusatorio**. Empezar señalando el comportamiento de Marcos activa su mecanismo de defensa antes de que haya comenzado el diálogo real. La conversación difícil se habrá convertido en un enfrentamiento.',
        punts: 0, seguent: 'scene_19'
      }
    ]
  },

  scene_19: {
    id: 'scene_19', etapa: 'La Prueba Suprema – Tensión', personatge: 'marcos',
    tipus: 'decision_scenario', titol: 'Decisión 2: Marcos se tensa',
    imatge: null,
    narracio: `Marcos se incorpora en la silla: "Llevo ocho años en este club. He jugado con cinco entrenadores distintos. Nunca ninguno me ha quitado minutos sin decirme por qué. Eso no es respeto." Su tono ha subido. Alex nota la tensión.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Pausa, respiración, y decir: "Tienes razón en el fondo. El cambio era necesario técnicamente, pero el proceso ha fallado. Tendría que haberlo hablado contigo antes. ¿Cómo hubieras preferido que lo gestionara?"',
        feedback: '**Regulación emocional excelente.** Alex no reacciona a la intensidad emocional: hace una pausa, valida el contenido (no la forma), asume la responsabilidad real y redirige hacia la solución. Esto es **empatía profesional**: comprender sin perder el rol de entrenador.',
        punts: 10, seguent: 'scene_20'
      },
      {
        id: 'B', text: 'Decir: "Lo siento mucho, tienes razón en todo, no lo volvería a hacer..."',
        feedback: 'Disculpa excesiva y pasiva. Asumir toda la culpa sin matices no resuelve el problema y crea un precedente negativo: la presión emocional funciona como herramienta de control. La respuesta asertiva implica asumir la responsabilidad real, no toda la responsabilidad imaginable.',
        punts: 5, seguent: 'scene_20'
      },
      {
        id: 'C', text: 'Decir: "Marcos, el entrenador soy yo y las decisiones técnicas son mías. Tienes que confiar en el criterio profesional."',
        feedback: 'Respuesta **agresiva** que escala el conflicto. Invocar la autoridad cuando la otra persona ha expresado una necesidad legítima (ser informado) cierra la conversación. La autoridad profesional no se impone: se gana.',
        punts: 0, seguent: 'scene_20'
      }
    ]
  },

  scene_20: {
    id: 'scene_20', etapa: 'La Prueba Suprema – Resolución', personatge: 'marcos',
    tipus: 'decision_scenario', titol: 'Decisión 3: Llegar a un acuerdo',
    imatge: null,
    narracio: `Marcos acepta que las decisiones técnicas tienen una lógica. Pero quiere saber qué puede esperar para el resto de la temporada. Alex tiene que proponer cómo será la relación de aquí en adelante.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"Marcos, de ahora en adelante, cuando vaya a hacer cambios relevantes en tu rol, te lo comunicaré con 48 horas de antelación y te explicaré el motivo. Si tienes dudas, hablamos. ¿Te parece bien?"',
        feedback: '**Acuerdo concreto y mutuamente válido.** Un protocolo específico (48 horas, comunicación, explicación) convierte el conflicto en un acuerdo de trabajo sostenible. Esto es **negociación efectiva**: una solución que respeta las necesidades de ambas partes sin que ninguna renuncie a lo esencial.',
        punts: 10, seguent: 'scene_21'
      },
      {
        id: 'B', text: '"De acuerdo, de ahora en adelante te consultaré antes de cualquier cambio que te afecte."',
        feedback: 'Cesión excesiva. Renunciar al criterio técnico por evitar el conflicto no es un buen acuerdo: es rendición. El entrenador perderá la autoridad necesaria para tomar decisiones independientes cuando sea necesario.',
        punts: 5, seguent: 'scene_21'
      },
      {
        id: 'C', text: '"Bien, intentaré comunicarme mejor." (sin proponer nada concreto)',
        feedback: '"Intentaré" no es un compromiso: es una forma de cerrar la conversación sin resolverla. Marcos saldrá con la sensación de que nada cambiará. Las conversaciones difíciles deben terminar con acuerdos claros, no con buenas intenciones.',
        punts: 0, seguent: 'scene_21'
      }
    ]
  },

  /* ══════════════════════════════════
     ACTO II-C – LA RECOMPENSA
  ══════════════════════════════════ */

  scene_21: {
    id: 'scene_21', etapa: 'La Recompensa', personatge: 'carmen',
    tipus: 'worked_example', titol: 'La confianza ganada',
    imatge: null,
    narracio: `Marcos sale de la sala y le estrecha la mano a Alex. Para un capitán que raramente expresa aprobación, el gesto es muy significativo. "De acuerdo. Así trabajamos."

Carmen, que había esperado en el pasillo, se acerca y dice simplemente: "Has crecido diez años en una hora." Sentados en la sala, le explica los seis estilos de liderazgo de Goleman y por cuál ha optado Alex en este proceso.`,
    dialeg: { personatge: 'carmen', text: '"El conflicto no es el problema. El conflicto es la oportunidad. Cómo lo gestionas determina si la relación sale reforzada o rota."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: 'Los 6 estilos de liderazgo (Goleman)',
      text: `**Visionario:** moviliza hacia una visión compartida. Útil en cambios de dirección estratégica.
**Coaching:** desarrolla a las personas para el futuro. Útil con talentos con potencial.
**Afiliativo:** crea armonía y refuerza vínculos emocionales. Útil en momentos de tensión grupal.
**Democrático:** genera consenso y participación. Útil para aprovechar el conocimiento del equipo.
**Timonel:** exige excelencia y lo modela él mismo. Eficaz a corto plazo. Riesgo: agota al equipo.
**Coercitivo:** obediencia inmediata. Solo para crisis reales. Alta toxicidad si se abusa de él.`
    },
    seguent: 'scene_22'
  },

  scene_22: {
    id: 'scene_22', etapa: 'La Recompensa', personatge: 'carmen',
    tipus: 'text_block', titol: 'El conflicto como aprendizaje',
    imatge: null,
    narracio: `Carmen le muestra a Alex lo que han conseguido: Marcos ha pasado de cuestionar en público a proponer mejoras en privado. La comunicación en los entrenamientos ha mejorado de forma notable.

"¿Ves el patrón?" le dice Carmen. "Cuando escuchas de verdad, las personas pasan de resistirse a co-crear. El jugador deja de ser un obstáculo y se convierte en un recurso."`,
    dialeg: { personatge: 'carmen', text: '"La asertividad no es dureza. Es claridad. Y la claridad, cuando va acompañada de respeto, genera confianza."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Empatía profesional: comprender sin perder el rol',
      text: 'La empatía profesional es la capacidad de comprender la perspectiva y las emociones de un jugador **sin fundirte** con ellas. No es decir "sé exactamente cómo te sientes" (nadie lo sabe), sino "entiendo por qué te sientes así y me parece legítimo". Esta distinción es clave para mantener el rol de entrenador.'
    },
    seguent: 'scene_23'
  },

  scene_23: {
    id: 'scene_23', etapa: 'La Recompensa', personatge: 'carmen',
    tipus: 'quiz', titol: 'Comprensión: estilos de liderazgo',
    imatge: null,
    narracio: `Comprobemos que has entendido los seis estilos de liderazgo de Goleman.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: '¿Qué estilo de liderazgo predominó en la gestión de Alex con Marcos (escucha, asume responsabilidad y establece un protocolo conjunto de trabajo)?',
    opcions: [
      { id: 'A', text: 'Estilo Timonel (exige excelencia y lo modela él mismo)', correcta: false, feedback: 'El estilo timonel se caracteriza por la exigencia de alto rendimiento sin énfasis relacional. Lo que hizo Alex fue lo contrario: construyó relación antes de establecer exigencias.' },
      { id: 'B', text: 'Estilo Coaching (desarrolla a las personas con un enfoque relacional y de crecimiento mutuo)', correcta: true, feedback: 'Correcto. Alex utilizó el **estilo coaching**: escuchó, reconoció la perspectiva de Marcos, asumió su responsabilidad y estableció un acuerdo de trabajo conjunto. No solo resolvió el conflicto: invirtió en la relación a largo plazo.' },
      { id: 'C', text: 'Estilo Afiliativo (crea armonía evitando el conflicto directo)', correcta: false, feedback: 'El estilo afiliativo tiende a evitar el conflicto. Alex, en cambio, lo gestionó frontalmente con respeto y criterio. Eso no es afiliar: es hacer coaching.' }
    ],
    seguent: 'scene_24'
  },

  /* ══════════════════════════════════
     ACTO III – EL RETORNO
  ══════════════════════════════════ */

  scene_24: {
    id: 'scene_24', etapa: 'El Camino de Regreso', personatge: 'ruben',
    tipus: 'text_block', titol: 'La tensión en el cuerpo técnico',
    imatge: null,
    narracio: `Dos semanas después. El ambiente en el equipo ha mejorado. Pero en el cuerpo técnico hay una tensión creciente: el preparador físico, Iván, y el analista de vídeo, Santi, llevan días sin coordinarse. Iván diseñó una sesión de recuperación que contradecía directamente el análisis táctico de Santi. Los dos se han acusado mutuamente ante Alex.

Rubén llama a Alex al despacho: "El cuerpo técnico no puede trabajar desalineado. Tienes que gestionar esto."`,
    dialeg: { personatge: 'ruben', text: '"No te pido que elijas quién tiene razón. Te pido que facilites que vuelvan a trabajar juntos. Eso también es entrenar."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Mediación informal en contextos deportivos',
      text: 'La mediación es un proceso donde una **tercera persona neutral** facilita la comunicación para ayudar a las partes a encontrar un acuerdo. En el deporte, el entrenador principal frecuentemente asume este rol. Requiere: imparcialidad real, intención declarada ("quiero ayudar, no juzgar") y escuchar a ambas partes por separado antes de reunirlas.'
    },
    seguent: 'scene_25'
  },

  scene_25: {
    id: 'scene_25', etapa: 'El Camino de Regreso', personatge: 'alex',
    tipus: 'decision_scenario', titol: '¿Mediar o no mediar?',
    imatge: null,
    narracio: `Alex tiene la oportunidad de intervenir. Pero interviniendo se arriesga a cometer errores. No interviniendo, permite que el conflicto siga su curso.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Intervenir como mediador: hablar por separado con Iván y Santi, y facilitar un diálogo entre ellos',
        feedback: '**Decisión madura y profesional.** Hablar por separado primero permite que cada parte se sienta escuchada sin la presión de la otra. Cuando las personas se sienten escuchadas, bajan las defensas y están más dispuestas al diálogo. Esta es la base de la mediación efectiva en contextos deportivos.',
        punts: 10, seguent: 'scene_26'
      },
      {
        id: 'B', text: 'No intervenir y dejar que Rubén lo gestione directamente desde dirección',
        feedback: 'En conflictos de cuerpo técnico que afectan al rendimiento del equipo, no intervenir cuando puedes ayudar es una oportunidad perdida. Iván y Santi han demostrado (días de silencio) que no pueden resolverlo solos. Verás lo que pasa cuando el conflicto escala.',
        punts: 0, seguent: 'scene_26b'
      }
    ]
  },

  scene_26: {
    id: 'scene_26', etapa: 'El Camino de Regreso', personatge: 'alex',
    tipus: 'text_block', titol: 'La mediación: el proceso',
    imatge: null,
    narracio: `Alex habla primero a solas con Iván. Escucha sin juzgar, reformula, y le pregunta: "¿Entiendes por qué Santi pudo sentirse contradecido?" Iván, sorprendido, admite que no había pensado en el impacto de su decisión.

Luego habla con Santi. Le explica que Iván no era consciente del efecto que había tenido. Le pregunta: "¿Estarías dispuesto a tener una conversación estructurada donde ambos puedan expresar lo que necesitan?" Santi acepta.`,
    dialeg: { personatge: 'alex', text: '"He aprendido que la mediación no consiste en decidir quién tiene razón. Consiste en ayudar a las dos partes a entender por qué actúan como actúan."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Toma de decisiones con criterio: cuándo intervenir',
      text: 'La decisión de intervenir en conflictos ajenos requiere evaluar: (1) El **impacto** del conflicto en el rendimiento colectivo. (2) La **capacidad** propia para intervenir de forma neutral. (3) El **momento** oportuno. (4) El **consentimiento** de las partes. Intervenir sin criterio es intrusión; no intervenir cuando puedes ayudar es negligencia.'
    },
    seguent: 'scene_27'
  },

  scene_26b: {
    id: 'scene_26b', etapa: 'El Camino de Regreso', personatge: 'narracio',
    tipus: 'text_block', titol: 'El conflicto escala',
    imatge: null,
    narracio: `Alex decide no intervenir. Una semana después, durante una sesión de análisis de vídeo con todo el equipo, Iván interrumpe la presentación de Santi con un comentario sarcástico. Santi abandona la sala. Los jugadores se miran entre sí.

Rubén cita a Alex inmediatamente: "Cuando te dije que lo gestionaras, lo decía en serio. Ahora el cuerpo técnico ha perdido credibilidad delante de los jugadores."`,
    dialeg: { personatge: 'ruben', text: '"El silencio ante un conflicto no es neutralidad. Quien ve un problema y no hace nada cuando puede hacerlo, es parte del problema."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'La evitación y sus consecuencias',
      text: 'Los conflictos no resueltos no desaparecen: **se acumulan y escalan**. Las consecuencias: deterioro del clima del cuerpo técnico, pérdida de credibilidad ante los jugadores y necesidad de una gestión correctiva mucho más costosa.'
    },
    seguent: 'scene_27'
  },

  scene_27: {
    id: 'scene_27', etapa: 'La Resurrección', personatge: 'ruben',
    tipus: 'text_block', titol: 'Rubén pide una opinión profesional',
    imatge: null,
    narracio: `Final de la primera vuelta. Las encuestas de satisfacción del equipo revelan que varios jugadores mencionan a un miembro del cuerpo técnico como fuente de tensión sin nombrarlo. Rubén tiene sus sospechas.

Cita a Alex en privado: "Necesito tu opinión profesional y honesta sobre el cuerpo técnico. No la que crees que quiero escuchar. Basada en lo que has observado."`,
    dialeg: { personatge: 'ruben', text: '"Te pido la opinión de alguien que ha visto de cerca cómo trabaja el equipo técnico. No acusar, no defender: información real y respetuosa para tomar una decisión justa."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Dar una opinión profesional difícil',
      text: 'Una opinión profesional sobre un colega debe: (1) Basarse en **hechos observables**, no en juicios personales. (2) **Separar persona y comportamiento** ("el comportamiento X ha generado el efecto Y"). (3) Reconocer el contexto. (4) Proponer vías constructivas cuando sea posible.'
    },
    seguent: 'scene_28'
  },

  scene_28: {
    id: 'scene_28', etapa: 'La Resurrección', personatge: 'ruben',
    tipus: 'decision_scenario', titol: '¿Qué estilo comunicativo usa Alex?',
    imatge: null,
    narracio: `Alex tiene la palabra. Rubén espera su opinión honesta sobre el cuerpo técnico.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"Iván es un problema. Lleva meses generando mal ambiente. Tendría que haberse ido hace tiempo."',
        feedback: 'Respuesta **agresiva** basada en juicios de valor. "Un problema" es una etiqueta global que no aporta información útil ni respeta la dignidad de Iván. Una opinión asertiva separa hechos observables de conclusiones globales.',
        punts: 0, seguent: 'scene_29'
      },
      {
        id: 'B', text: '"Honestamente, no me siento cómodo opinando sobre un compañero. No creo que sea mi lugar."',
        feedback: 'Respuesta **pasiva** que evita la responsabilidad. Rubén ha pedido explícitamente tu opinión como profesional que ha observado al cuerpo técnico de cerca. Evitar opinar cuando tienes información relevante y se te pregunta directamente no contribuye a la toma de decisiones del club.',
        punts: 5, seguent: 'scene_29'
      },
      {
        id: 'C', text: '"Lo que he observado es que Iván tiende a tomar decisiones de forma unilateral sin coordinar con el resto del cuerpo técnico. He visto dos situaciones concretas donde esto ha generado tensión. Su trabajo físico es sólido. La pregunta es si puede ajustar ese estilo de trabajo con la formación adecuada."',
        feedback: '**Asertividad avanzada ejemplar.** Alex se basa en hechos observables, reconoce lo positivo, separa persona y comportamiento, y propone una vía constructiva. Expresar la verdad de forma respetuosa, útil y constructiva es el máximo exponente de la asertividad profesional.',
        punts: 10, seguent: 'scene_29'
      }
    ]
  },

  scene_29: {
    id: 'scene_29', etapa: 'El Retorno con el Elixir', personatge: 'alex',
    tipus: 'checklist', titol: 'El decálogo del entrenador profesional completo',
    imatge: null,
    narracio: `Final de temporada. CD Olimpia ha mantenido la categoría y ha dado paso a tres jugadores de la cantera. Rubén propone que Alex lidere las sesiones trimestrales de formación del cuerpo técnico.

Alex prepara el **decálogo de buenas prácticas del entrenador profesional**. Marca todos los principios que crees que deberían formar parte de él.`,
    dialeg: { personatge: 'alex', text: '"Cuando llegué pensaba que mi trabajo era diseñar entrenamientos. Ahora sé que el trabajo real es crear las condiciones para que treinta personas puedan dar lo mejor de sí mismas."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Las seis competencias del entrenador profesional completo',
      text: '**Proyecto deportivo** · **Técnico especializado** · **Gestión de personas** · **Relaciones públicas** · **Liderazgo situacional** · **Comunicación efectiva**. Estas seis competencias no se enseñan solo en el campo de entrenamiento: se construyen en cada conversación, cada decisión y cada relación profesional.'
    },
    checklistItems: [
      { id: 'dc1',  text: 'Definir y comunicar un proyecto deportivo claro antes del inicio de cada temporada',                      correcta: true },
      { id: 'dc2',  text: 'Adaptar el estilo de liderazgo al estado emocional y la situación táctica del equipo',                   correcta: true },
      { id: 'dc3',  text: 'Comunicar los cambios de rol a los jugadores con antelación y explicando el motivo',                     correcta: true },
      { id: 'dc4',  text: 'Gestionar las críticas públicas con calma y proponer un espacio privado para el diálogo',                correcta: true },
      { id: 'dc5',  text: 'Expresar opiniones profesionales de forma asertiva: hechos observables, no juicios globales',            correcta: true },
      { id: 'dc6',  text: 'Preparar y estructurar las conversaciones difíciles con jugadores antes de tenerlas',                    correcta: true },
      { id: 'dc7',  text: 'Alinear las declaraciones mediáticas con el proyecto deportivo real del club',                           correcta: true },
      { id: 'dc8',  text: 'Coordinar el trabajo del cuerpo técnico con reuniones periódicas y objetivos compartidos',               correcta: true },
      { id: 'dc9',  text: 'Reconocer el trabajo previo del equipo cuando se llega como nuevo entrenador',                          correcta: true },
      { id: 'dc10', text: 'Solicitar y dar feedback constructivo dentro del cuerpo técnico de forma regular',                       correcta: true },
      { id: 'dc11', text: 'Usar la comunicación agresiva con los jugadores cuando hay que poner límites firmes',                    correcta: false },
      { id: 'dc12', text: 'Evitar siempre los conflictos dentro del equipo para mantener un ambiente positivo',                     correcta: false },
      { id: 'dc13', text: 'No pedir nunca consejo a los jugadores para no mostrar inseguridad profesional',                         correcta: false }
    ],
    punts_per_item: 1,
    seguent: 'scene_30'
  },

  /* ══════════════════════════════════
     EPÍLOGO
  ══════════════════════════════════ */

  scene_30: {
    id: 'scene_30', etapa: 'Epílogo: El Nuevo Alex', personatge: 'carmen',
    tipus: 'epilogue', titol: 'El retorno con el elixir',
    imatge: null,
    narracio: `Seis meses desde el primer día. Campo de entrenamiento de CD Olimpia, lunes por la mañana.

Carmen le ha dejado un libro sobre la mesa: "La dirección del talento en el deporte de equipo." Con una nota: "Ya no necesitas que te enseñen. Necesitas practicar lo que ya sabes."`,
    dialeg: { personatge: 'carmen', text: '"Ya no eres el Alex del primer día. Eres alguien que ha aprendido que el rendimiento deportivo no lo producen los ejercicios: lo producen las personas que confían entre sí. Bienvenido al equipo de los que entrenan de verdad."' },
    contingut_pedagogic: null
  }
};

/* ============================================================
   ENGINE
   ============================================================ */
const Engine = {
  state: {
    currentScene: 'scene_01',
    score: 0,
    decisions: {},
    checklistData: {},
    quizData: {},
    visitedScenes: [],
    completedScenes: [],
    finished: false
  },

  init: function () {
    SCORM.init();
    var saved = SCORM.loadSuspendData();
    if (saved && saved.currentScene && scenes[saved.currentScene]) {
      this.state = saved;
    }
    this._markVisited(this.state.currentScene);
    this._render(this.state.currentScene);
  },

  _save: function () { SCORM.saveSuspendData(this.state); },
  _allIds: function () { return Object.keys(scenes); },

  _markVisited: function (id) {
    if (this.state.visitedScenes.indexOf(id) === -1) this.state.visitedScenes.push(id);
  },

  _markCompleted: function (id) {
    if (this.state.completedScenes.indexOf(id) === -1) this.state.completedScenes.push(id);
  },

  goToScene: function (id) {
    if (!scenes[id]) return;
    this._markCompleted(this.state.currentScene);
    this.state.currentScene = id;
    this._markVisited(id);
    this._save();
    this._render(id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  },

  recordDecision: function (sceneId, optionId, punts) {
    if (this.state.decisions[sceneId]) return;
    this.state.decisions[sceneId] = { optionId: optionId, punts: punts };
    this.state.score += punts;
    this._save();
    UI.updateScore(this.state.score);
  },

  recordChecklist: function (sceneId, selectedIds, punts) {
    this.state.checklistData[sceneId] = { selectedIds: selectedIds, punts: punts };
    this.state.score += punts;
    this._save();
    UI.updateScore(this.state.score);
  },

  recordQuiz: function (sceneId, optionId) {
    this.state.quizData[sceneId] = { optionId: optionId };
    this._save();
  },

  finish: function () {
    this.state.finished = true;
    this._markCompleted(this.state.currentScene);
    this._save();
    SCORM.finish(this.state.score);
  },

  _render: function (id) {
    UI.render(scenes[id], this.state);
    UI.updateProgressBar(this.state.completedScenes.length, CANONICAL_SCENE_COUNT);
    UI.updateJourneyMap(JOURNEY_STAGES, this.state.visitedScenes, this.state.completedScenes);
  }
};

/* ============================================================
   UI
   ============================================================ */
const UI = {
  _timerHandle: null,

  render: function (scene, state) {
    clearTimeout(this._timerHandle);
    document.getElementById('stage-label').textContent  = scene.etapa || '';
    document.getElementById('scene-title').textContent  = scene.titol || '';
    this.updateScore(state.score);
    this._renderImage(scene);
    this._renderChar(scene);
    this._renderNarrative(scene);
    this._renderDialogue(scene);
    this._renderPedagogic(scene);
    this._renderInteraction(scene, state);
  },

  updateScore: function (score) {
    document.getElementById('score-display').textContent = score + ' pts';
  },

  updateProgressBar: function (done, total) {
    var pct = Math.min(100, Math.round((done / total) * 100));
    document.getElementById('progress-bar-fill').style.width = pct + '%';
    document.getElementById('progress-label').textContent = done + '/' + total;
  },

  updateJourneyMap: function (stages, visited, completed) {
    var el = document.getElementById('journey-map-content');
    if (!el) return;
    el.innerHTML = '';
    stages.forEach(function (stage) {
      var hasVisited   = stage.scenes.some(function (s) { return visited.indexOf(s)   !== -1; });
      var hasCompleted = stage.scenes.some(function (s) { return completed.indexOf(s) !== -1; });
      var div = document.createElement('div');
      div.className = 'map-stage ' + (hasCompleted ? 'map-done' : hasVisited ? 'map-active' : 'map-pending');
      div.style.borderLeftColor = stage.color;
      var badge = document.createElement('span');
      badge.className = 'map-badge';
      badge.style.background = stage.color;
      badge.textContent = stage.label;
      var title = document.createElement('span');
      title.className = 'map-title';
      title.textContent = stage.title;
      var status = document.createElement('span');
      status.className = 'map-status';
      status.textContent = hasCompleted ? '✓ Completado' : hasVisited ? '● En curso' : '○ Pendiente';
      div.appendChild(badge); div.appendChild(title); div.appendChild(status);
      el.appendChild(div);
    });
  },

  _renderImage: function (scene) {
    var container = document.getElementById('scene-image-container');
    var img = document.getElementById('scene-image');
    if (scene.imatge) {
      img.src = scene.imatge;
      img.alt = scene.titol || '';
      container.style.display = 'block';
      img.onerror = function () { container.style.display = 'none'; };
    } else {
      container.style.display = 'none';
    }
  },

  _renderChar: function (scene) {
    var char = CHARACTERS[scene.personatge] || CHARACTERS.narracio;
    var av = document.getElementById('character-avatar');
    av.style.background = char.color;
    av.textContent = char.initials;
    av.style.borderRadius = char.shape === 'square' ? '8px' : '50%';
    var nm = document.getElementById('character-name');
    nm.textContent = char.name;
    nm.style.color = char.color;
  },

  _renderNarrative: function (scene) {
    document.getElementById('narrative-text').innerHTML = this._md(scene.narracio || '');
  },

  _renderDialogue: function (scene) {
    var box = document.getElementById('dialogue-box');
    if (!scene.dialeg) { box.style.display = 'none'; return; }
    var char = CHARACTERS[scene.dialeg.personatge] || CHARACTERS.narracio;
    box.style.display = 'block';
    box.style.borderLeftColor = char.color;
    var nm = document.getElementById('dialogue-char-name');
    nm.textContent = char.name;
    nm.style.color = char.color;
    document.getElementById('dialogue-text').innerHTML = this._md(scene.dialeg.text || '');
  },

  _renderPedagogic: function (scene) {
    var block = document.getElementById('pedagogic-block');
    if (!scene.contingut_pedagogic) { block.style.display = 'none'; return; }
    block.style.display = 'block';
    document.getElementById('pedagogic-title').textContent = scene.contingut_pedagogic.titol || '';
    document.getElementById('pedagogic-text').innerHTML    = this._md(scene.contingut_pedagogic.text || '');
  },

  _renderInteraction: function (scene, state) {
    var area = document.getElementById('interaction-area');
    area.innerHTML = '';
    if (scene.tipus === 'text_block' || scene.tipus === 'worked_example' || scene.tipus === 'concept_definition') {
      this._renderTextBlock(scene, state, area);
    } else if (scene.tipus === 'decision_scenario') {
      this._renderDecision(scene, state, area);
    } else if (scene.tipus === 'checklist') {
      this._renderChecklist(scene, state, area);
    } else if (scene.tipus === 'quiz') {
      this._renderQuiz(scene, state, area);
    } else if (scene.tipus === 'epilogue') {
      this._renderEpilogue(scene, state, area);
    }
  },

  _renderTextBlock: function (scene, state, area) {
    var timerSecs = 15;
    var timerDiv = document.createElement('div');
    timerDiv.className = 'timer-container';
    var timerTrack = document.createElement('div');
    timerTrack.className = 'timer-bar-track';
    var timerFill = document.createElement('div');
    timerFill.className = 'timer-bar-fill';
    timerTrack.appendChild(timerFill);
    var timerLabel = document.createElement('span');
    timerLabel.textContent = timerSecs + 's';
    timerDiv.appendChild(timerTrack);
    timerDiv.appendChild(timerLabel);
    area.appendChild(timerDiv);

    var btn = document.createElement('button');
    btn.className = 'btn btn-primary btn-disabled';
    btn.textContent = 'Continuar →';
    btn.addEventListener('click', function () {
      if (!btn.classList.contains('btn-disabled')) Engine.goToScene(scene.seguent);
    });
    area.appendChild(btn);

    var remaining = timerSecs;
    var handle = setInterval(function () {
      remaining--;
      timerFill.style.width = (remaining / timerSecs * 100) + '%';
      timerLabel.textContent = remaining > 0 ? remaining + 's' : '';
      if (remaining <= 0) {
        clearInterval(handle);
        btn.classList.remove('btn-disabled');
        btn.classList.add('btn-enabled');
        timerDiv.style.display = 'none';
      }
    }, 1000);
    this._timerHandle = handle;
  },

  _renderDecision: function (scene, state, area) {
    var self = this;
    var prev = state.decisions[scene.id];
    var label = document.createElement('p');
    label.className = 'decision-label';
    label.textContent = 'Toma una decisión:';
    area.appendChild(label);

    scene.opcions.forEach(function (op) {
      var btn = document.createElement('button');
      btn.className = 'btn btn-option' + (prev && prev.optionId === op.id ? ' btn-selected' : '') + (prev ? ' btn-disabled' : '');
      btn.textContent = op.id + ') ' + op.text;
      btn.addEventListener('click', function () {
        if (prev) return;
        area.querySelectorAll('.btn-option').forEach(function (b) { b.classList.add('btn-disabled'); });
        Engine.recordDecision(scene.id, op.id, op.punts);
        self._showFeedback(op, function () { Engine.goToScene(op.seguent); });
      });
      area.appendChild(btn);
    });

    if (prev) {
      var op = scene.opcions.find(function (o) { return o.id === prev.optionId; });
      if (op) {
        var nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-primary btn-enabled';
        nextBtn.textContent = 'Continuar →';
        nextBtn.addEventListener('click', function () { Engine.goToScene(op.seguent); });
        area.appendChild(nextBtn);
      }
    }
  },

  _renderQuiz: function (scene, state, area) {
    var self = this;
    var prev = state.quizData[scene.id];
    var qEl = document.createElement('p');
    qEl.className = 'quiz-question';
    qEl.textContent = scene.pregunta;
    area.appendChild(qEl);

    scene.opcions.forEach(function (op) {
      var btn = document.createElement('button');
      btn.className = 'btn btn-quiz';
      btn.textContent = op.id + ') ' + op.text;
      btn.dataset.qid = op.id;
      if (prev) {
        btn.classList.add('btn-disabled');
        if (op.correcta) btn.classList.add('btn-quiz-correct');
        else if (op.id === prev.optionId) btn.classList.add('btn-quiz-wrong');
      }
      btn.addEventListener('click', function () {
        if (prev) return;
        area.querySelectorAll('.btn-quiz').forEach(function (b) { b.classList.add('btn-disabled'); });
        Engine.recordQuiz(scene.id, op.id);
        scene.opcions.forEach(function (o) {
          var b = area.querySelector('[data-qid="' + o.id + '"]');
          if (b) {
            if (o.correcta) b.classList.add('btn-quiz-correct');
            else if (o.id === op.id) b.classList.add('btn-quiz-wrong');
          }
        });
        var fbDiv = document.createElement('div');
        fbDiv.className = 'quiz-feedback ' + (op.correcta ? 'qf-correct' : 'qf-wrong');
        fbDiv.innerHTML = (op.correcta ? '<strong>✓ Correcto.</strong> ' : '<strong>✗ No exactamente.</strong> ') + self._md(op.feedback);
        area.appendChild(fbDiv);
        var nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-primary btn-enabled';
        nextBtn.style.marginTop = '4px';
        nextBtn.textContent = 'Continuar →';
        nextBtn.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
        area.appendChild(nextBtn);
      });
      area.appendChild(btn);
    });

    if (prev) {
      var prevOp = scene.opcions.find(function (o) { return o.id === prev.optionId; });
      if (prevOp) {
        var fbDiv = document.createElement('div');
        fbDiv.className = 'quiz-feedback ' + (prevOp.correcta ? 'qf-correct' : 'qf-wrong');
        fbDiv.innerHTML = (prevOp.correcta ? '<strong>✓ Correcto.</strong> ' : '<strong>✗ No exactamente.</strong> ') + self._md(prevOp.feedback);
        area.appendChild(fbDiv);
      }
      var nextBtn = document.createElement('button');
      nextBtn.className = 'btn btn-primary btn-enabled';
      nextBtn.textContent = 'Continuar →';
      nextBtn.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
      area.appendChild(nextBtn);
    }
  },

  _renderChecklist: function (scene, state, area) {
    var self = this;
    var prev = state.checklistData[scene.id];
    var selected = prev ? prev.selectedIds.slice() : [];
    var label = document.createElement('p');
    label.className = 'decision-label';
    label.textContent = 'Marca los elementos importantes:';
    area.appendChild(label);

    var listDiv = document.createElement('div');
    listDiv.className = 'checklist-items';

    scene.checklistItems.forEach(function (item) {
      var div = document.createElement('div');
      div.className = 'checklist-item' + (selected.indexOf(item.id) !== -1 ? ' checked' : '');
      div.dataset.id = item.id;
      var icon = document.createElement('span');
      icon.className = 'checkbox-icon';
      icon.textContent = selected.indexOf(item.id) !== -1 ? '☑' : '☐';
      var txt = document.createElement('span');
      txt.className = 'checkbox-text';
      txt.textContent = item.text;
      div.appendChild(icon); div.appendChild(txt);
      if (!prev) {
        div.addEventListener('click', function () {
          var idx = selected.indexOf(item.id);
          if (idx === -1) { selected.push(item.id); div.classList.add('checked'); icon.textContent = '☑'; }
          else { selected.splice(idx, 1); div.classList.remove('checked'); icon.textContent = '☐'; }
        });
      }
      listDiv.appendChild(div);
    });
    area.appendChild(listDiv);

    var _applyResults = function () {
      scene.checklistItems.forEach(function (item) {
        var d = listDiv.querySelector('[data-id="' + item.id + '"]');
        if (!d) return;
        d.style.pointerEvents = 'none';
        var was = selected.indexOf(item.id) !== -1;
        d.classList.remove('checked');
        if      (item.correcta && was)  d.classList.add('cl-correct');
        else if (!item.correcta && was) d.classList.add('cl-incorrect');
        else if (item.correcta && !was) d.classList.add('cl-missed');
      });
    };

    if (!prev) {
      var confirmBtn = document.createElement('button');
      confirmBtn.className = 'btn btn-primary btn-enabled';
      confirmBtn.textContent = 'Confirmar selección';
      confirmBtn.addEventListener('click', function () {
        var punts = 0;
        selected.forEach(function (id) {
          var it = scene.checklistItems.find(function (i) { return i.id === id; });
          if (it && it.correcta) punts += scene.punts_per_item;
        });
        Engine.recordChecklist(scene.id, selected, punts);
        _applyResults();
        confirmBtn.remove();
        var maxPts = scene.checklistItems.filter(function (i) { return i.correcta; }).length * scene.punts_per_item;
        var fbDiv = document.createElement('div');
        fbDiv.className = 'checklist-feedback';
        fbDiv.innerHTML = '<strong>' + punts + ' / ' + maxPts + ' puntos.</strong>  ' +
          '<span style="color:var(--good)">■ Verde = correcto marcado</span>  ' +
          '<span style="color:var(--bad)">■ Rojo = incorrecto marcado</span>  ' +
          '<span style="color:var(--ok)">■ Naranja = correcto no marcado</span>';
        area.appendChild(fbDiv);
        var nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-primary btn-enabled';
        nextBtn.textContent = 'Continuar →';
        nextBtn.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
        area.appendChild(nextBtn);
      });
      area.appendChild(confirmBtn);
    } else {
      _applyResults();
      var nextBtn2 = document.createElement('button');
      nextBtn2.className = 'btn btn-primary btn-enabled';
      nextBtn2.textContent = 'Continuar →';
      nextBtn2.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
      area.appendChild(nextBtn2);
    }
  },

  _renderEpilogue: function (scene, state, area) {
    var score = state.score;
    var pct   = Math.round((score / 100) * 100);

    var mentorMsg, mentorClass;
    if (score >= 85) {
      mentorMsg   = '"Has demostrado una comprensión excepcional del modelo competencial del entrenador. CD Olimpia tiene suerte de tenerte. Sigue creciendo: cada equipo es un nuevo viaje."';
      mentorClass = 'epilogue-excellent';
    } else if (score >= 60) {
      mentorMsg   = '"Has avanzado mucho en poco tiempo. La base es sólida. Las competencias comunicativas y de liderazgo necesitan práctica diaria: cada interacción es una oportunidad de aprendizaje."';
      mentorClass = 'epilogue-good';
    } else if (score >= 35) {
      mentorMsg   = '"Has dado los primeros pasos. Las competencias del entrenador profesional son como el entrenamiento físico: requieren constancia. Revisa las escenas donde has tenido más dificultades."';
      mentorClass = 'epilogue-ok';
    } else {
      mentorMsg   = '"El camino del entrenador completo empieza con errores. Lo que importa no es dónde empiezas, sino la dirección en que vas. Revisa el recorrido y reflexiona sobre cada decisión."';
      mentorClass = 'epilogue-needs-work';
    }

    var dsHtml = '<ul class="decisions-summary">';
    var scored = ['scene_05','scene_10','scene_13','scene_16','scene_18','scene_19','scene_20','scene_25','scene_28','scene_29'];
    scored.forEach(function (sid) {
      var sc  = scenes[sid];
      var dec = (sid === 'scene_16' || sid === 'scene_29') ? state.checklistData[sid] : state.decisions[sid];
      if (!sc || !dec) return;
      var pts   = dec.punts !== undefined ? dec.punts : 0;
      var icon  = pts >= 10 ? '✓' : pts >= 5 ? '⚠' : '✗';
      var cls   = pts >= 10 ? 'ds-good' : pts >= 5 ? 'ds-ok' : 'ds-bad';
      dsHtml += '<li class="' + cls + '"><span class="ds-icon">' + icon + '</span><span>' + sc.titol + '</span><span class="ds-pts">' + pts + ' pts</span></li>';
    });
    dsHtml += '</ul>';

    var skillsHtml = '<ul class="skills-list">' +
      '<li>✓ Proyecto deportivo: visión, valores y planificación estratégica</li>' +
      '<li>✓ Técnico especializado: metodología, diseño y análisis del rendimiento</li>' +
      '<li>✓ Gestión de personas: motivación, conflictos y cohesión de equipo</li>' +
      '<li>✓ Relaciones públicas: comunicación mediática e imagen institucional</li>' +
      '<li>✓ Liderazgo situacional: los seis estilos de Goleman en contexto</li>' +
      '<li>✓ Comunicación efectiva: asertividad, empatía profesional y escucha activa</li>' +
      '</ul>';

    var circ = 339.292;
    var dash = circ * pct / 100;

    var container = document.createElement('div');
    container.className = 'epilogue-container';
    container.innerHTML =
      '<div class="epilogue-score-ring">' +
        '<svg viewBox="0 0 120 120" class="score-ring-svg">' +
          '<circle cx="60" cy="60" r="54" fill="none" stroke="#E2E8F0" stroke-width="9"/>' +
          '<circle cx="60" cy="60" r="54" fill="none" stroke="#2563EB" stroke-width="9" ' +
            'stroke-dasharray="' + dash + ' ' + circ + '" stroke-dashoffset="84.823" stroke-linecap="round"/>' +
        '</svg>' +
        '<div class="score-ring-text">' +
          '<span class="score-big">' + score + '</span>' +
          '<span class="score-max">/100</span>' +
        '</div>' +
      '</div>' +
      '<div class="epilogue-mentor ' + mentorClass + '">' +
        '<div class="epilogue-mentor-avatar" style="background:#059669">Ca</div>' +
        '<div class="epilogue-mentor-msg">' + mentorMsg + '</div>' +
      '</div>' +
      '<p class="epilogue-section-title">Resumen de tus decisiones clave</p>' +
      dsHtml +
      '<p class="epilogue-section-title">Competencias trabajadas</p>' +
      skillsHtml +
      '<div class="epilogue-buttons"><button id="finish-btn" class="btn btn-finish">Finalizar y enviar puntuación al LMS</button></div>';

    area.appendChild(container);

    document.getElementById('finish-btn').addEventListener('click', function () {
      Engine.finish();
      this.disabled = true;
      this.textContent = '✓ Puntuación enviada';
    });
  },

  _showFeedback: function (op, onContinue) {
    var icon = op.punts === 10 ? '✓' : op.punts === 5 ? '⚠' : '✗';
    var cls  = op.punts === 10 ? 'fb-good' : op.punts === 5 ? 'fb-ok' : 'fb-bad';
    document.getElementById('feedback-icon').textContent = icon;
    document.getElementById('feedback-icon').className   = 'feedback-icon ' + cls;
    document.getElementById('feedback-pts').textContent  = '+' + op.punts + ' puntos';
    document.getElementById('feedback-text').innerHTML   = this._md(op.feedback);
    document.getElementById('feedback-overlay').style.display = 'flex';
    var continueBtn = document.getElementById('feedback-continue');
    continueBtn.onclick = function () {
      document.getElementById('feedback-overlay').style.display = 'none';
      onContinue();
    };
  },

  _md: function (t) {
    if (!t) return '';
    function inl(s) {
      return s
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/\*([^*]+)\*/g, '<em>$1</em>');
    }
    return t.split('\n\n').map(function (para) {
      para = para.trim();
      if (!para) return '';
      var lines = para.split('\n');
      var first = lines[0].trim();
      if (/^(\*\*)?[1-9]\d*[.)]\s/.test(first)) {
        return '<ol class="step-list">' +
          lines.filter(Boolean).map(function (l) { return '<li>' + inl(l.trim()) + '</li>'; }).join('') +
          '</ol>';
      }
      if (/^[-•]\s/.test(first)) {
        return '<ul>' +
          lines.filter(Boolean).map(function (l) { return '<li>' + inl(l.trim().replace(/^[-•]\s*/, '')) + '</li>'; }).join('') +
          '</ul>';
      }
      if (lines.length > 1 && /^\*\*[^*]/.test(first)) {
        return '<ul class="def-list">' +
          lines.filter(Boolean).map(function (l) { return '<li>' + inl(l.trim()) + '</li>'; }).join('') +
          '</ul>';
      }
      return '<p>' + inl(para) + '</p>';
    }).join('');
  }
};

window.Engine = Engine;

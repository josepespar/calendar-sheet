/* ============================================================
   SCENARIO – Liderazgo Deportivo: El Viaje del Héroe
   10 escenas canónicas · Decisiones + Quiz · Máx: 100 pts
   Idioma: español
   Basado en: Coma Bau et al. (2019) – Efectos de la Competición
   en las Autopercepciones Conductuales de Entrenadores ASOBAL
   ============================================================ */

const CHARACTERS = {
  narracion: { name: 'Narrador',    color: '#4B6584', initials: '✦', shape: 'square' },
  jordi:     { name: 'Jordi Coma',  color: '#F59E0B', initials: 'JC', shape: 'circle' },
  jack:      { name: 'Jack',        color: '#DC2626', initials: 'J',  shape: 'circle' },
  gery:      { name: 'Gery',        color: '#8B5CF6', initials: 'G',  shape: 'circle' },
  ambrosio:  { name: 'Ambrosio',    color: '#10B981', initials: 'A',  shape: 'circle' },
  marco:     { name: 'Marco',       color: '#60A5FA', initials: 'M',  shape: 'circle' }
};

const JOURNEY_STAGES = [
  {
    id: 'act1',
    label: 'Acto I',
    title: 'El Llamado',
    scenes: ['scene_01', 'scene_02', 'scene_03'],
    color: '#F59E0B'
  },
  {
    id: 'act2',
    label: 'Acto II',
    title: 'La Prueba',
    scenes: ['scene_04', 'scene_04b', 'scene_05', 'scene_06', 'scene_06b'],
    color: '#DC2626'
  },
  {
    id: 'act3',
    label: 'Acto III',
    title: 'La Transformación',
    scenes: ['scene_07', 'scene_08', 'scene_09', 'scene_09b'],
    color: '#10B981'
  },
  {
    id: 'final',
    label: 'Epílogo',
    title: 'El Regreso',
    scenes: ['scene_10'],
    color: '#8B5CF6'
  }
];

const CANONICAL_SCENE_COUNT = 10;

const scenes = {

  /* ═══════════════════════════════════════════════════════════
     ACTO I – El Llamado a la Aventura
  ═══════════════════════════════════════════════════════════ */

  scene_01: {
    id: 'scene_01',
    tipus: 'text_block',
    titol: 'La Sala en Silencio',
    personatge: 'narracion',
    narracio: 'Barcelona, octubre. La sala de conferencias del Institut Nacional d\'Educació Física de Catalunya huele a café frío y expectación contenida. Doce entrenadores de élite de la ASOBAL —los mejores del balonmano español— miran sus teléfonos, sus notas, o simplemente al frente.\n\nHan venido porque tienen que venir. O eso creen.\n\nAl fondo de la sala, de pie junto a la pizarra, Jordi Coma —entrenador, investigador, profesor de la Universitat de Vic— deja pasar tres segundos antes de hablar.',
    dialeg: {
      personatge: 'jordi',
      text: '"Tengo una pregunta para vosotros. Y quiero que la respondáis con honestidad.\n\n¿Cuántos de vosotros entrenáis exactamente igual en el partido 30 que en el partido 1?"\n\n(Silencio.)\n\n"Ni uno. Y ese cambio... lo estáis haciendo solos, sin saberlo, sin quererlo. Hoy vamos a hablar de eso."'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02',
    tipus: 'text_block',
    titol: 'El Modelo',
    personatge: 'jordi',
    narracio: 'Jordi proyecta un esquema limpio sobre la pizarra. Cinco palabras. Cinco dimensiones. Una teoría que ha resistido décadas de escrutinio científico en el deporte de élite.',
    dialeg: {
      personatge: 'jordi',
      text: '"En los años 80, Chelladurai diseñó un modelo para medir lo que los entrenadores realmente hacemos —no lo que creemos que hacemos. Lo llamó Modelo Multidimensional de Liderazgo. Y la herramienta que lo mide, el LSS-3, sigue siendo la referencia mundial."'
    },
    contingut_pedagogic: {
      titol: 'Modelo Multidimensional de Liderazgo · Chelladurai',
      text: 'El **LSS-3** mide 5 dimensiones del comportamiento del entrenador:\n\n**1. Instrucción y Entrenamiento** — Orientar la técnica, organizar el equipo, optimizar el rendimiento.\n**2. Comportamiento Democrático** — Involucrar a los deportistas en las decisiones del equipo.\n**3. Comportamiento Autocrático** — Tomar decisiones de forma independiente, sin consultar.\n**4. Apoyo Social** — Preocuparte por el bienestar personal del deportista, más allá del rendimiento.\n**5. Feedback Positivo** — Reconocer y reforzar las actuaciones bien ejecutadas.\n\nLa eficacia del liderazgo depende de la **congruencia** entre el comportamiento preferido del deportista, el requerido por la situación, y el comportamiento real del entrenador.'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03',
    tipus: 'text_block',
    titol: 'El Desafío',
    personatge: 'jack',
    narracio: 'Jack Müller lleva 17 años en la élite. Dos Copas del Rey, una Liga. Sabe exactamente dónde está el límite entre la teoría y el vestuario. Y no teme decirlo.\n\nSe recuesta en la silla, cruza los brazos y espera a que Jordi termine la frase.',
    dialeg: {
      personatge: 'jack',
      text: '"Con todo el respeto, Jordi. Esto son cuestionarios. Yo soy entrenador, no psicólogo. ¿Qué me va a decir un papel sobre lo que pasa en mi equipo que yo no sepa ya?\n\nLa presión no se mide en encuestas. Se mide en el marcador del domingo."\n\n(Doce pares de ojos se giran hacia Jordi. El momento ha llegado.)'
    },
    seguent: 'scene_04'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTO II – La Prueba
  ═══════════════════════════════════════════════════════════ */

  scene_04: {
    id: 'scene_04',
    tipus: 'decisio',
    titol: 'Tu Respuesta',
    personatge: 'jordi',
    narracio: 'La sala aguanta la respiración. Jack ha cuestionado públicamente la metodología delante de los mejores entrenadores del país. El silencio dura exactamente dos segundos.',
    pregunta: '¿Cómo responde Jordi a la provocación de Jack?',
    opcions: [
      {
        id: 'A',
        text: 'Saca los datos del estudio ASOBAL. "Jack, tienes razón: el marcador no miente. Por eso usamos datos de entrenadores exactamente como tú."',
        punts: 25,
        feedback: '**La evidencia como espejo.**\n\nCuando la resistencia viene del miedo a que los datos revelen algo incómodo, la mejor respuesta es precisamente los datos. Jordi invita a Jack a mirarse, no a defenderse.\n\n*"El feedback positivo no es un premio. Es oxígeno."*',
        seguent: 'scene_05'
      },
      {
        id: 'B',
        text: 'Sonríe y sigue la presentación. "Cada uno tiene su experiencia. Sigamos."',
        punts: 0,
        feedback: '**El conflicto evitado no desaparece.**\n\nDesviar el desafío de Jack mantiene la paz superficial, pero pierde una oportunidad de oro: convertir la resistencia en aprendizaje compartido. El aula lo nota.',
        seguent: 'scene_04b'
      }
    ]
  },

  scene_04b: {
    id: 'scene_04b',
    tipus: 'text_block',
    titol: 'El Silencio que Habla',
    personatge: 'narracion',
    narracio: 'La respuesta de Jordi evita el conflicto. Jack se recuesta en su silla. Gery toma nota en su cuaderno. El aula pierde temperatura.',
    dialeg: {
      personatge: 'marco',
      text: '"(Susurrando a Ambrosio) Creía que íbamos a debatir de verdad..."\n\n(Una oportunidad perdida de convertir la resistencia en aprendizaje. El momento pasa, pero deja huella.)'
    },
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05',
    tipus: 'text_block',
    titol: 'La Evidencia que Cambia Todo',
    personatge: 'jordi',
    narracio: 'Jordi proyecta el primer resultado del estudio. Una tabla simple. Cinco dimensiones. Dos mediciones: inicio y final de temporada. Y un asterisco junto a una sola fila.\n\nLa sala se inclina hacia delante.',
    dialeg: {
      personatge: 'jordi',
      text: '"Seguimos a entrenadores de la ASOBAL a lo largo de cinco meses de competición. El resultado más sorprendente no fue lo que cambiaron. Fue lo que **todos** cambiaron por igual.\n\nIndependientemente de si su equipo ganaba o perdía, todos los entrenadores redujeron significativamente su **Comportamiento Democrático**.\n\nTodos. Sin excepción. Sin ser conscientes de ello."'
    },
    seguent: 'scene_06'
  },

  scene_06: {
    id: 'scene_06',
    tipus: 'decisio',
    titol: 'El Consejo a Gery',
    personatge: 'gery',
    narracio: 'Pausa. Gery levanta la mano. Dirige a un equipo que lleva tres semanas sin ganar. Tiene cara de quien necesita una respuesta real, no una académica.',
    dialeg: {
      personatge: 'gery',
      text: '"Jordi, mi equipo lleva tres semanas sin ganar. Noto que los jugadores están distantes. ¿Qué me recomiendas?"'
    },
    pregunta: '¿Cuál es el consejo que Jordi da a Gery?',
    opcions: [
      {
        id: 'A',
        text: '"Gery, antes de cambiar nada: ¿has preguntado a tus jugadores qué creen que está fallando?"',
        punts: 25,
        feedback: '**La pregunta como liderazgo.**\n\nEl Comportamiento Democrático no es un lujo que te puedes permitir solo cuando vas ganando. Es precisamente en la adversidad donde más se necesita la voz del equipo.\n\n*"No hay democracia que valga si solo la ejerces cuando vas ganando."*',
        seguent: 'scene_07'
      },
      {
        id: 'B',
        text: '"Dale más directrices claras, más estructura. En los momentos difíciles, el equipo necesita saber que tú tienes el control."',
        punts: 0,
        feedback: '**La trampa del control.**\n\nEn la adversidad, el instinto autocrático se dispara. Pero dar más órdenes no reconstruye la confianza. Refuerza la dependencia. Y cuando el equipo pierde la voz, pierde también la responsabilidad.',
        seguent: 'scene_06b'
      }
    ]
  },

  scene_06b: {
    id: 'scene_06b',
    tipus: 'text_block',
    titol: 'La Trampa del Control',
    personatge: 'gery',
    narracio: 'Gery implementa la recomendación. Más estructura, más directrices, menos espacio para la iniciativa. El equipo obedece.',
    dialeg: {
      personatge: 'gery',
      text: '"(Al día siguiente) Los chicos hacen lo que digo. Pero algo falta. Un equipo puede seguir órdenes y aun así perder el alma."\n\n(Jordi escucha. Hay una lección aquí que todavía puede rescatarse.)'
    },
    seguent: 'scene_07'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTO III – La Transformación
  ═══════════════════════════════════════════════════════════ */

  scene_07: {
    id: 'scene_07',
    tipus: 'text_block',
    titol: 'La Grieta en el Equipo',
    personatge: 'narracion',
    narracio: 'Tres semanas después del seminario. Jordi recibe un mensaje de Gery a las 22 horas. Y luego, de forma inesperada, Ambrosio toma la palabra delante del grupo.',
    dialeg: {
      personatge: 'ambrosio',
      text: '"Eso es lo que pasa cuando llevas años mandando sin escuchar. No es que no hablen. Es que aprendieron que no tiene sentido hablar.\n\n*El silencio del vestuario no es neutralidad. Es el sonido de un equipo que dejó de creer.*"'
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08',
    tipus: 'quiz',
    titol: 'La Pregunta Clave',
    personatge: 'jordi',
    narracio: 'Jordi detiene la conversación. Señala la tabla que sigue proyectada en la pizarra. Es el momento de comprobar si el aprendizaje ha calado.',
    pregunta: 'Según el estudio de Coma Bau et al. (2019) con entrenadores de élite de la ASOBAL, ¿cuál fue el único cambio conductual significativo en **todos** los entrenadores tras cinco meses de competición, independientemente de los resultados de su equipo?',
    punts: 25,
    opcions: [
      {
        id: 'A',
        text: 'Disminución del Comportamiento Democrático',
        correcta: true,
        feedback: '**Exacto.** La presión competitiva reduce el espacio de participación de los deportistas en la toma de decisiones, incluso en los entrenadores con mejor rendimiento. Es un efecto sistemático, no individual.'
      },
      {
        id: 'B',
        text: 'Aumento del Feedback Positivo',
        correcta: false,
        feedback: 'No exactamente. El Feedback Positivo varió en función de los resultados del equipo: los entrenadores con mejores resultados lo aumentaron; los que no alcanzaban sus expectativas lo redujeron. No fue universal.'
      },
      {
        id: 'C',
        text: 'Reducción de la Instrucción y Entrenamiento',
        correcta: false,
        feedback: 'No. La Instrucción y el Entrenamiento no mostraron cambios significativos generalizados. La dimensión que cayó en todos los entrenadores —con y sin buenos resultados— fue el Comportamiento Democrático.'
      },
      {
        id: 'D',
        text: 'Incremento del Comportamiento Autocrático',
        correcta: false,
        feedback: 'No. Aunque la presión suele asociarse con más autocracia, el estudio no encontró un cambio significativo y universal en esa dimensión. El hallazgo clave fue la caída del Comportamiento Democrático.'
      }
    ],
    seguent: 'scene_09'
  },

  scene_09: {
    id: 'scene_09',
    tipus: 'decisio',
    titol: 'El Momento de Ambrosio',
    personatge: 'ambrosio',
    narracio: 'Ambrosio lleva 22 años en el banquillo. Su equipo está en semifinales. En el descanso del partido decisivo, su capitán le busca con la mirada y se acerca.',
    dialeg: {
      personatge: 'ambrosio',
      text: '"Capitán: \'Míster, necesitamos más libertad para leer el partido. Confíe en nosotros.\'\n\n(Ambrosio tiene 5 minutos para decidir. Marcador: 14-14.)"'
    },
    pregunta: '¿Cuál es la decisión de Ambrosio?',
    opcions: [
      {
        id: 'A',
        text: '"De acuerdo. Vosotros conocéis al rival en la pista mejor que yo desde aquí. Os doy el marco, vosotros tomáis las decisiones dentro."',
        punts: 25,
        feedback: '**Confianza construida en el trabajo diario.**\n\nEl liderazgo democrático no se improvisa en el marcador 14-14. Se construye en cada entrenamiento. Cuando llega el momento, el equipo sabe que su voz importa.\n\n*"El silencio del vestuario no es neutralidad. Es el sonido de un equipo que dejó de creer."*',
        seguent: 'scene_10'
      },
      {
        id: 'B',
        text: '"Ahora mismo no. Hay demasiado en juego. Cuando terminemos hablamos."',
        punts: 0,
        feedback: '**Demasiado en juego para ser democrático.**\n\nEsta es la trampa más común: el liderazgo democrático como algo que se permite solo en la victoria. Pero los equipos que solo participan en las decisiones fáciles aprenden que su voz no tiene peso real.',
        seguent: 'scene_09b'
      }
    ]
  },

  scene_09b: {
    id: 'scene_09b',
    tipus: 'text_block',
    titol: 'Demasiado en Juego',
    personatge: 'narracion',
    narracio: 'El segundo tiempo comienza. El equipo ejecuta el plan de Ambrosio con precisión. Pierden por dos goles.',
    dialeg: {
      personatge: 'ambrosio',
      text: '"(En el túnel) Jugaron bien. Pero algo faltó.\n\n(El capitán no dice nada. Solo recoge su bolsa y sale. Hay conversaciones que no se tienen en el momento exacto, y después ya no tienen el mismo efecto.)"'
    },
    seguent: 'scene_10'
  },

  /* ═══════════════════════════════════════════════════════════
     EPÍLOGO – El Regreso con el Elixir
  ═══════════════════════════════════════════════════════════ */

  scene_10: {
    id: 'scene_10',
    tipus: 'epilog',
    titol: 'El Legado',
    personatge: 'jordi',
    narracio: 'El seminario llega a su fin. Fuera, Barcelona continúa su ritmo. Dentro, doce entrenadores guardan silencio durante un instante antes de levantarse.',
    dialeg: {
      personatge: 'jordi',
      text: '"Os propongo un ejercicio para mañana. Solo uno.\n\nAntes del entrenamiento, pregunta a uno de tus jugadores qué cree que podría mejorar el equipo. Escucha sin interrumpir. Sin opinar. Solo escucha.\n\nY después observa qué cambia."'
    },
    mentorChar: 'jordi',
    mentorMsgs: {
      excellent: '**Maestro del Liderazgo Deportivo.**\n\nHas demostrado una comprensión profunda del Modelo de Chelladurai y de cómo la presión competitiva afecta el comportamiento del entrenador. Tus decisiones han sido consistentes con la evidencia: escuchar, preguntar, confiar.\n\nEso no es teoría. Eso es liderazgo.',
      good: '**Líder en Evolución.**\n\nTienes las bases sólidas. Comprendes el modelo, identificas los patrones. El siguiente paso es aplicarlos cuando la presión aprieta: precisamente cuando el instinto dice lo contrario.\n\n*"La presión no cambia quiénes somos. Revela quiénes somos cuando nadie nos mira."*',
      ok: '**El Camino Continúa.**\n\nHas absorbido los conceptos clave. Algunas decisiones han sido costosas. El aprendizaje del liderazgo deportivo no es lineal: es exactamente eso, un viaje.\n\nRelee los resultados del estudio ASOBAL. Los datos no mienten.',
      needsWork: '**El Inicio del Viaje.**\n\nEste módulo ha tocado terreno nuevo. El Modelo de Chelladurai y los hallazgos del estudio ASOBAL son el punto de partida. Te recomendamos revisar las escenas con decisión y volver a intentarlo con más contexto.\n\nEl liderazgo, como el entrenamiento, mejora con la repetición consciente.'
    }
  }

};

/* ============================================================
   ENGINE
   Gestiona estado, SCORM, renderizado de escenas y mapa
   ============================================================ */

const Engine = (function () {

  var state = {
    currentScene: 'scene_01',
    score: 0,
    visitedScenes: [],
    decisions: []
  };

  function _save() { SCORM.saveSuspendData(state); }
  function _qs(id) { return document.getElementById(id); }

  function _md(t) {
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

  function _updateHeader(scene) {
    var stage = JOURNEY_STAGES.find(function (s) { return s.scenes.indexOf(scene.id) !== -1; }) || JOURNEY_STAGES[0];
    _qs('stage-label').textContent = stage.label + ' · ' + stage.title;
    _qs('scene-title').textContent = scene.titol || '';
    _qs('score-display').textContent = state.score + ' pts';
    var visited = state.visitedScenes.filter(function (s) { return s.indexOf('b') === -1; }).length;
    var pct = Math.min(100, Math.round(visited / CANONICAL_SCENE_COUNT * 100));
    _qs('progress-bar-fill').style.width = pct + '%';
    _qs('progress-label').textContent = visited + '/' + CANONICAL_SCENE_COUNT;
  }

  function _renderJourneyMap() {
    var c = _qs('journey-map-content');
    if (!c) return;
    c.innerHTML = JOURNEY_STAGES.map(function (stage) {
      var done   = stage.scenes.every(function (s) { return state.visitedScenes.indexOf(s) !== -1; });
      var active = stage.scenes.indexOf(state.currentScene) !== -1;
      var cls    = done ? 'map-done' : active ? 'map-active' : 'map-pending';
      var status = done ? '✓ Completado' : active ? '▶ En curso' : '○ Pendiente';
      return '<div class="map-stage ' + cls + '" style="border-left-color:' + stage.color + '">' +
        '<span class="map-badge" style="background:' + stage.color + '">' + stage.label + '</span>' +
        '<span class="map-title">' + stage.title + '</span>' +
        '<span class="map-status">' + status + '</span>' +
        '</div>';
    }).join('');
  }

  function _showFeedback(icon, pts, text, onContinue) {
    var ov = _qs('feedback-overlay');
    _qs('feedback-icon').className = 'feedback-icon ' + icon;
    _qs('feedback-icon').textContent = icon === 'fb-good' ? '✓' : icon === 'fb-ok' ? '~' : '✗';
    _qs('feedback-pts').textContent = pts > 0 ? '+' + pts + ' pts' : '';
    _qs('feedback-text').innerHTML = _md(text);
    ov.style.display = 'flex';
    var btn = _qs('feedback-continue');
    btn.onclick = function () { ov.style.display = 'none'; onContinue(); };
  }

  function _renderScene(scene) {
    _qs('scene-image-container').style.display = 'none';
    _qs('dialogue-box').style.display = 'none';
    _qs('pedagogic-block').style.display = 'none';
    _qs('interaction-area').innerHTML = '';

    var ch = CHARACTERS[scene.personatge] || CHARACTERS.narracion;
    var r = ch.shape === 'square' ? '4px' : '50%';
    _qs('character-avatar').style.background   = ch.color;
    _qs('character-avatar').style.borderRadius = r;
    _qs('character-avatar').textContent        = ch.initials;
    _qs('character-name').textContent          = ch.name;

    if (scene.imatge) {
      _qs('scene-image').src = scene.imatge;
      _qs('scene-image-container').style.display = 'block';
    }

    _qs('narrative-text').innerHTML = _md(scene.narracio || '');

    if (scene.dialeg) {
      var dc = CHARACTERS[scene.dialeg.personatge] || ch;
      _qs('dialogue-char-name').textContent  = dc.name;
      _qs('dialogue-char-name').style.color  = dc.color;
      _qs('dialogue-text').innerHTML         = _md(scene.dialeg.text);
      _qs('dialogue-box').style.display      = 'block';
    }

    if (scene.contingut_pedagogic) {
      var pb = scene.contingut_pedagogic;
      _qs('pedagogic-title').textContent = pb.titol || '';
      _qs('pedagogic-text').innerHTML    = _md(pb.text || '');
      _qs('pedagogic-block').style.display = 'block';
    }

    var ia = _qs('interaction-area');
    if (scene.tipus === 'text_block') {
      _renderTextBlock(scene, ia);
    } else if (scene.tipus === 'quiz') {
      _renderQuiz(scene, ia);
    } else if (scene.tipus === 'decisio') {
      _renderDecisio(scene, ia);
    } else if (scene.tipus === 'epilog') {
      _renderEpilog(scene, ia);
    }
  }

  function _renderTextBlock(scene, ia) {
    if (scene.seguent) {
      var btn = document.createElement('button');
      btn.className = 'btn btn-primary btn-enabled';
      btn.textContent = 'Continuar →';
      btn.onclick = function () { Engine.showScene(scene.seguent); };
      ia.appendChild(btn);
    } else {
      var restartBtn = document.createElement('button');
      restartBtn.className = 'btn btn-primary btn-enabled';
      restartBtn.textContent = 'Reiniciar curso ↺';
      restartBtn.onclick = function () {
        state = { currentScene: 'scene_01', score: 0, visitedScenes: [], decisions: [] };
        _save();
        Engine.showScene('scene_01');
      };
      ia.appendChild(restartBtn);
    }
  }

  function _renderQuiz(scene, ia) {
    var label = document.createElement('div');
    label.className = 'decision-label';
    label.textContent = 'Pregunta de comprensión';
    ia.appendChild(label);

    var q = document.createElement('div');
    q.className = 'quiz-question';
    q.innerHTML = _md(scene.pregunta);
    ia.appendChild(q);

    var answered = false;
    scene.opcions.forEach(function (op) {
      var btn = document.createElement('button');
      btn.className = 'btn btn-quiz btn-enabled';
      btn.innerHTML = '<strong>' + op.id + ')</strong> ' + op.text;
      btn.onclick = function () {
        if (answered) return;
        answered = true;
        var pts = op.correcta ? (scene.punts || 10) : 0;
        state.score += pts;
        state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
        scene.opcions.forEach(function (o, i) {
          var b = ia.querySelectorAll('.btn-quiz')[i];
          if (o.correcta)       b.className = 'btn btn-quiz btn-quiz-correct btn-disabled';
          else if (o.id === op.id) b.className = 'btn btn-quiz btn-quiz-wrong btn-disabled';
          else                   b.className = 'btn btn-quiz btn-disabled';
        });
        var fb = document.createElement('div');
        fb.className = 'quiz-feedback ' + (op.correcta ? 'qf-correct' : 'qf-wrong');
        fb.innerHTML = _md(op.feedback);
        ia.appendChild(fb);
        _save();
        var contBtn = document.createElement('button');
        contBtn.className = 'btn btn-primary btn-enabled';
        contBtn.style.marginTop = '10px';
        contBtn.textContent = 'Continuar →';
        contBtn.onclick = function () { Engine.showScene(scene.seguent); };
        ia.appendChild(contBtn);
        _qs('score-display').textContent = state.score + ' pts';
        setTimeout(function () { ia.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 100);
      };
      ia.appendChild(btn);
    });
  }

  function _renderDecisio(scene, ia) {
    var label = document.createElement('div');
    label.className = 'decision-label';
    label.textContent = 'Decisión';
    ia.appendChild(label);

    if (scene.pregunta) {
      var q = document.createElement('div');
      q.className = 'quiz-question';
      q.innerHTML = _md(scene.pregunta);
      ia.appendChild(q);
    }

    var decided = false;
    scene.opcions.forEach(function (op) {
      var btn = document.createElement('button');
      btn.className = 'btn btn-option btn-enabled';
      btn.innerHTML = '<strong>' + op.id + ')</strong> ' + op.text;
      btn.onclick = function () {
        if (decided) return;
        decided = true;
        var pts = op.punts !== undefined ? op.punts : 0;
        state.score += pts;
        state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
        scene.opcions.forEach(function (o, i) {
          ia.querySelectorAll('.btn-option')[i].className = 'btn btn-option btn-disabled';
        });
        btn.className = 'btn btn-option btn-selected btn-disabled';
        _save();
        var icon = pts >= 20 ? 'fb-good' : pts >= 10 ? 'fb-ok' : 'fb-bad';
        var nextScene = op.seguent || scene.seguent;
        _showFeedback(icon, pts, op.feedback, function () {
          Engine.showScene(nextScene);
        });
        _qs('score-display').textContent = state.score + ' pts';
      };
      ia.appendChild(btn);
    });
  }

  function _renderEpilog(scene, ia) {
    var maxScore    = 100;
    var circumference = 2 * Math.PI * 54;
    var offset      = circumference * (1 - state.score / maxScore);
    var grade, gradeClass, mentorMsg;

    if (state.score >= 90) {
      grade      = 'Excelente';
      gradeClass = 'epilogue-excellent';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.excellent  : '';
    } else if (state.score >= 70) {
      grade      = 'Muy bien';
      gradeClass = 'epilogue-good';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.good       : '';
    } else if (state.score >= 50) {
      grade      = 'Correcto';
      gradeClass = 'epilogue-ok';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.ok         : '';
    } else {
      grade      = 'Necesitas repasar';
      gradeClass = 'epilogue-needs-work';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.needsWork  : '';
    }

    var mentor = CHARACTERS[scene.mentorChar] || CHARACTERS.jordi;
    var html   = '<div class="epilogue-container">';

    html += '<div class="epilogue-score-ring">';
    html += '<svg class="score-ring-svg" viewBox="0 0 120 120">';
    html += '<circle cx="60" cy="60" r="54" fill="none" stroke="var(--surface3)" stroke-width="8"/>';
    html += '<circle cx="60" cy="60" r="54" fill="none" stroke="var(--gold)" stroke-width="8"';
    html += ' stroke-dasharray="' + circumference.toFixed(2) + '" stroke-dashoffset="' + offset.toFixed(2) + '"';
    html += ' stroke-linecap="round" transform="rotate(-90 60 60)"/>';
    html += '</svg>';
    html += '<div class="score-ring-text"><span class="score-big">' + state.score + '</span><span class="score-max">/ 100</span></div>';
    html += '</div>';

    html += '<div class="epilogue-grade ' + gradeClass + '">' + grade + '</div>';

    html += '<div class="epilogue-mentor ' + gradeClass + '">';
    html += '<div class="epilogue-mentor-avatar" style="background:' + mentor.color + ';color:#000">' + mentor.initials + '</div>';
    html += '<div class="epilogue-mentor-msg">' + _md(mentorMsg) + '</div>';
    html += '</div>';

    if (state.decisions.length > 0) {
      html += '<div><div class="epilogue-section-title">Resumen de decisiones</div>';
      html += '<ul class="decisions-summary">';
      state.decisions.forEach(function (d) {
        var cls  = d.pts >= 20 ? 'ds-good' : d.pts >= 10 ? 'ds-ok' : 'ds-bad';
        var icon = d.pts >= 20 ? '✓' : d.pts >= 10 ? '~' : '✗';
        html += '<li class="' + cls + '">';
        html += '<span class="ds-icon">' + icon + '</span>';
        html += '<span>' + (d.label || d.scene) + '</span>';
        html += '<span class="ds-pts">' + d.pts + ' pts</span>';
        html += '</li>';
      });
      html += '</ul></div>';
    }

    html += '<div style="margin-top:16px;">';
    html += '<button class="btn btn-primary btn-enabled" onclick="';
    html += 'var s={currentScene:\'scene_01\',score:0,visitedScenes:[],decisions:[]};';
    html += 'SCORM.saveSuspendData(s);Engine.showScene(\'scene_01\');">';
    html += 'Reiniciar curso ↺</button></div>';
    html += '</div>';

    ia.innerHTML = html;
    SCORM.finish(state.score);
    _save();
  }

  return {
    init: function () {
      SCORM.init();
      var saved = SCORM.loadSuspendData();
      if (saved && saved.currentScene) { state = saved; }
      Engine.showScene(state.currentScene);
      _renderJourneyMap();
    },

    showScene: function (id) {
      var scene = scenes[id];
      if (!scene) { console.warn('[Engine] Escena no encontrada:', id); return; }
      state.currentScene = id;
      if (state.visitedScenes.indexOf(id) === -1) state.visitedScenes.push(id);
      _save();
      _updateHeader(scene);
      _renderScene(scene);
      _renderJourneyMap();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    },

    renderJourneyMap: _renderJourneyMap
  };
})();

window.Engine = Engine;

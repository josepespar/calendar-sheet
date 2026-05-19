/* ============================================================
   SCENARIO – Percepció i Creativitat en el Procés d'Aprenentatge de l'Handbol
   30 escenes · Quiz + Decisions + Checklists · Puntuació màxima: 100 pts
   Idioma: català
   Basat en l'obra de Pinaud Philippe i Enrique Díez
   ============================================================ */

const CHARACTERS = {
  marc: {
    name: 'Marc Vidal',
    color: '#06C798',
    initials: 'M',
    shape: 'circle'
  },
  sofia: {
    name: 'Sofia Massaguer',
    color: '#FF6B35',
    initials: 'S',
    shape: 'circle'
  },
  jordi: {
    name: 'Jordi',
    color: '#4A90D9',
    initials: 'J',
    shape: 'circle'
  },
  laia: {
    name: 'Laia',
    color: '#A855F7',
    initials: 'L',
    shape: 'circle'
  },
  narracio: {
    name: 'Narrador',
    color: '#5A7A90',
    initials: '✦',
    shape: 'square'
  }
};

const JOURNEY_STAGES = [
  { id: 'act1', label: 'Acte I',   title: 'El Punt de Partida',      scenes: ['scene_01','scene_02','scene_03','scene_04','scene_05','scene_06','scene_07','scene_08'], color: '#06C798' },
  { id: 'act2', label: 'Acte II',  title: 'La Visió i la Ment',      scenes: ['scene_09','scene_10','scene_11','scene_12','scene_13','scene_14'],                      color: '#4A90D9' },
  { id: 'act3', label: 'Acte III', title: 'Les Estratègies Visuals', scenes: ['scene_15','scene_16','scene_17','scene_18','scene_19','scene_20'],                      color: '#A855F7' },
  { id: 'act4', label: 'Acte IV',  title: 'La Pedagogia Creativa',   scenes: ['scene_21','scene_22','scene_23'],                                                        color: '#FF6B35' },
  { id: 'act5', label: 'Acte V',   title: 'Metodologia Pràctica',    scenes: ['scene_24','scene_25','scene_26','scene_26b','scene_27','scene_28','scene_29'],           color: '#27AE60' },
  { id: 'final',label: 'Epíleg',   title: 'La Nova Mirada',          scenes: ['scene_30'],                                                                              color: '#06C798' }
];

const CANONICAL_SCENE_COUNT = 28;

const scenes = {

  /* ═══════════════════════════════════════════════════════════
     ACTE I – El Punt de Partida
  ═══════════════════════════════════════════════════════════ */

  scene_01: {
    id: 'scene_01',
    tipus: 'text_block',
    titol: 'El primer dia al Garraf',
    personatge: 'narracio',
    narracio: 'Marc Vidal arriba al pavelló del Club Handbol Garraf un dimarts a les sis de la tarda. Té 30 anys, la llicència d\'entrenador de nivell II i molta il·lusió. Avui és el seu primer dia com a entrenador principal de les categories inferiors.\n\nA la pista, un exercici de 3x3 és en curs. Marc s\'atura i observa. Tots els jugadors miren la pilota tot el temps: quan Laia rep, gira el cap cap a la pilota; quan Jordi s\'apropa, torna a girar. Ningú no "llegeix" el joc; tots reaccionen, però sempre amb un pas de retard.\n\nUna desmarcada evident de Laia davant la porteria passa completament desapercebuda per als seus companys. La pilota arriba al defensor.',
    dialeg: {
      personatge: 'narracio',
      text: '**El cervell humà és capaç de processar en paral·lel milers d\'informacions al mateix temps. Però la majoria dels entrenadors, i dels jugadors, ignoren com aprofitar aquest potencial.**\n\nLa pregunta que Marc haurà d\'aprendre a respondre és: com es transforma un jugador que "mira" en un jugador que "veu"?'
    },
    contingut_pedagogic: {
      titol: 'Les quatre claus de la percepció esportiva',
      text: '**1. El 95% de l\'activitat cerebral és inconscient** — la major part de la percepció i de la decisió tàctica passa fora de la consciència del jugador.\n**2. Menys d\'un 1% dels estímuls visuals arriben a la consciència** — el cervell filtra activament la informació rebuda per la retina.\n**3. Els processos paral·lels inconscients són el motor de la creativitat** — el processament seqüencial conscient és massa lent per a les decisions tàctiques urgents.\n**4. L\'aprenentatge perceptiu ha de començar aviat** — la neuroplasticitat màxima dura fins a la pubertat; cada any comptat des dels 6 és crucial.'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02',
    tipus: 'text_block',
    titol: 'El problema: mirar sense veure',
    personatge: 'jordi',
    narracio: 'Jordi, el millor jugador del grup, rep la pilota en posició privilegiada. Laia es desmarca amb claredat davant la porteria, completament lliure. Jordi passa directament al defensor.',
    dialeg: {
      personatge: 'jordi',
      text: '"La he vista pero no tenía tiempo de reaccionar." Jordi es queda quiet, desconcertat per la seva pròpia resposta.'
    },
    contingut_pedagogic: {
      titol: 'Percepció conscient i inconscient',
      text: 'Sofia apareix al costat de Marc i explica en veu baixa:\n\n"No és un problema de velocitat, Jordi. És un problema de percepció. Mirava però no veia."\n\nLa distinció és fonamental. **El processament conscient és seqüencial**: el cervell analitza un element a la vegada. Però en handbol, els temps de reacció requerits (entre 50 i 200 ms) no permeten aquest procés seqüencial.\n\n**El motor de la creativitat tàctica és el processament inconscient en paral·lel**: el cervell tracta simultàniament totes les informacions disponibles, sense que el jugador en sigui conscient. El 95% de l\'activitat cerebral és inconscient. És aquest mecanisme que el bon entrenament ha de desenvolupar.'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03',
    tipus: 'quiz',
    titol: 'Comprensió: percepció conscient',
    personatge: 'sofia',
    punts: 0,
    narracio: 'Sofia mira Marc i li fa una pregunta per verificar que ha entès el mecanisme fonamental que acaba d\'explicar.',
    pregunta: 'Quan un jugador d\'handbol reacciona instantàniament a un desmarcat del pivot sense ser conscient de la decisió, quin mecanisme cerebral predomina principalment?',
    opcions: [
      {
        id: 'A',
        text: 'El processament conscient seqüencial del sistema visual central',
        correcta: false,
        feedback: 'El processament conscient és seqüencial: analitza un element a la vegada. Quan els temps de reacció són molt curts (com en el desmarcat d\'un pivot), el cervell no té temps per a processos seqüencials. **La resposta instantània prové dels processos inconscients paral·lels**, que analitzen simultàniament totes les informacions disponibles.'
      },
      {
        id: 'B',
        text: 'El processament inconscient en paral·lel dels processos viso-motors',
        correcta: true,
        feedback: 'Exacte. En situacions de decisió tàctica "urgent", el cervell elabora la resposta a través de múltiples processos inconscients en paral·lel. **El jugador no ha participat conscientment en la tria de l\'acció**: el cervell "ha decidit per ell", i el component conscient es limita a una vigilància i auto-control del procés en curs.'
      },
      {
        id: 'C',
        text: 'Un reflex espinal automàtic independent del cervell',
        correcta: false,
        feedback: 'Els reflexos espinals (com retirar la mà d\'un objecte calent) no intervenen en decisions tàctiques complexes. **La presa de decisions tàctica és un procés cerebral**, però en situacions urgents es fa a través dels circuits inconscients del cervell, no dels circuits conscients.'
      }
    ],
    seguent: 'scene_04'
  },

  scene_04: {
    id: 'scene_04',
    tipus: 'text_block',
    titol: 'Aprendre sense saber que aprens',
    personatge: 'sofia',
    narracio: 'Sofia i Marc s\'asseuen a la graderia mentre els jugadors fan pausa. Sofia explica la diferència entre aprenentatge conscient i inconscient amb una analogia musical.',
    dialeg: {
      personatge: 'sofia',
      text: '"Pensa en un pianista experimentat. Quan toca una sonata de Beethoven, no pensa en cada tecla que premeu. El seu cervell inconscient gestiona centenars de variables simultànies —pressió dels dits, tempo, dinàmica, posició corporal— sense consumir atenció conscient. Aprendre handbol és exactament el mateix.\n\nL\'aprenentatge inconscient és global: processa totes les variables del joc en paral·lel sense consumir atenció. Però hi ha un perill enorme: si obstruïm la vivència mental del nen obligant-lo a concentrar-se en percepcions que podrien romandre en l\'àmbit inconscient, bloquejem el camí cap a operacions mentals de major nivell tàctic."'
    },
    contingut_pedagogic: {
      titol: 'Aprenentatge conscient vs inconscient',
      text: '**Conscient** → seqüencial, lent, un element a la vegada. Útil per aprendre regles tècniques bàsiques en primers estadis.\n**Inconscient** → paral·lel, ràpid, global. Motor de la creativitat i de la decisió tàctica avançada.\n\nEl pas del piano → l\'handbol:\n- El pianista novell llegeix nota a nota (conscient, seqüencial).\n- El concert master interpreta frascs i emocions (inconscient, global).\n- El jove jugador que repeteix una acció tècnica aïllada queda "enganxat" al nivell conscient.\n- El jugador creatiu ha après a "no pensar" durant l\'execució.'
    },
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05',
    tipus: 'decisio',
    titol: 'La primera decisió de Marc',
    personatge: 'marc',
    punts: 10,
    narracio: 'Laia acaba de fer una passada directament al defensor en lloc d\'al company que estava clarament desmarcat. Marc té l\'impuls de aturar l\'exercici i explicar-li el que ha de fer. Sofia l\'observa en silenci des del lateral.',
    pregunta: 'Laia ha comès l\'error. Quina és la millor reacció de Marc com a entrenador?',
    opcions: [
      {
        id: 'A',
        punts: 0,
        text: 'Atura l\'exercici, crida Laia i explica-li detalladament: "Has de mirar primer si el teu company de l\'esquerra estava desmarcant-se..."',
        feedback: 'Aturar l\'exercici per donar instruccions verbals detallades **interromp el processament inconscient** que és precisament el que volem desenvolupar. El cervell de la Laia estava en mig d\'un procés d\'aprenentatge que la teva intervenció ha tallat. Sofia t\'hauria dit: "Deixa que el cervell processi."',
        seguent: 'scene_06'
      },
      {
        id: 'B',
        punts: 5,
        text: 'Deixa continuar, i quan acaba la seqüència preguntes a Laia: "Què has intentat fer?" (objectiu) sense dir-li "com" hauria d\'haver-ho fet',
        feedback: 'Bona aproximació! Preguntar per l\'**objectiu** (no pel mitjà) preserva l\'aprenentatge inconscient i estimula la reflexió autèntica. **Però manquen 5 punts** perquè l\'ideal seria no intervenir en absolut en els primers estadis: deixar que la pràctica repetida consolidi els nous patrons perceptius sense guia verbal.',
        seguent: 'scene_06'
      },
      {
        id: 'C',
        punts: 10,
        text: 'Continues l\'exercici sense cap intervenció verbal. Modifiques subtilment el disseny del proper exercici per afegir un company addicional en posició de desmarcat evident, que "atraurà" l\'atenció periférica de la Laia de forma natural',
        feedback: 'Perfecte. El principi clau és que **el cervell aprèn millor quan selecciona per si mateix les conductes més eficaces**. En lloc d\'instruir verbalment, modifiques les condicions de pràctica perquè l\'entorn "guiï" l\'aprenentatge. Això aprofita al màxim els mecanismes d\'aprenentatge inconscient.',
        seguent: 'scene_06'
      }
    ],
    seguent: 'scene_06'
  },

  scene_06: {
    id: 'scene_06',
    tipus: 'text_block',
    titol: 'Visió central i visió periférica',
    personatge: 'sofia',
    narracio: 'Acabat l\'entrenament, Sofia i Marc es queden al pavelló. Sofia agafa un retolador i dibuixa dos cercles en la pissarra: un petit i un gran.',
    dialeg: {
      personatge: 'sofia',
      text: '"El sistema visual humà té dos subsistemes completament diferents. La **visió central** cobreix tan sols uns 5° d\'arc (l\'equivalent al diàmetre aparent d\'una pilota a dos metres). És precisa, detecta colors i detalls fins, però és costosa i intermitent: el cervell ha de desplaçar l\'ull d\'un punt a un altre per cada nova fixació.\n\nLa **visió periférica** cobreix 180° en total, fins a 60° en el camp binocular (que permet càlculs de profunditat) i fins a 90° en el camp monocular. Treballa en contrast de lluminàncies —escala de grisos— no en color. Menys precisa, sí. Però ràpida i contínua: mai no s\'atura.\n\nConseqüència pràctica: les ulleres redueixen el camp periféric. Per això, sempre que sigui possible, els jugadors haurien de portar lentilles."\n\nQuant frenem el cotxe per evitar una col·lisió, **no intervé la visió central**: és la visió periférica que detecta el creixement ràpid de la imatge retiniana del vehicle i activa la resposta reflexa.'
    },
    contingut_pedagogic: {
      titol: 'Per a un jugador central amb pilota',
      text: 'El jugador central no ha de focalitzar successivament tots els elements de l\'escena. Ha de mirar constantment **lluny** (vers la porteria) per percebre el màxim d\'informació per unitat de temps.\n\n**Visió central** → precisa, en color, 5°, costosa, intermitent.\n**Visió periférica** → menys precisa, en grisos, 180°, ràpida, contínua.\n\nUna taca de color en moviment és suficient per identificar el pivot; no cal veure el número de la samarreta. El jugador que mira directament a cada company "gasta" sacudides i temps que el camp no li permet.'
    },
    seguent: 'scene_07'
  },

  scene_07: {
    id: 'scene_07',
    tipus: 'text_block',
    titol: 'La dansa de la mirada',
    personatge: 'sofia',
    narracio: 'Sofia continua l\'explicació. Marc pren notes a la seva llibreta.',
    dialeg: {
      personatge: 'sofia',
      text: '"Quan l\'ull es desplaça d\'un punt a un altre, fa el que s\'anomenen **sacudides** (saccades). La velocitat màxima d\'una sacudida augmenta amb l\'amplitud, però l\'acceleració i la frenada consumeixen temps. L\'amplitud màxima d\'una sacudida única és d\'uns 45°. Per a objectius més lluny calen dues sacudides o un moviment de cap, cosa molt més lenta.\n\nEl processament de la informació visual central **no comença fins que la fixació és efectiva**: durant la sacudida, el cervell és pràcticament cec. Cada transició entre dues fixacions és temps perdut."'
    },
    contingut_pedagogic: {
      titol: 'Visió central vs visió periférica en situació de joc',
      text: '**Visió central** → precisa / puntual / costosa / intermitent / sacudida mínima ~200ms.\n**Visió periférica** → menys precisa / ràpida / contínua / sempre activa.\n\nMarc: "Llavors si el jugador mira tot el temps a la pilota, perd tota la informació dels companys..."\n\nSofia: "I el que és pitjor: fins i tot perd precisió en la pilota, perquè consumeix sacudides constants. La **dansa de la mirada** no és aleatòria: és el resultat d\'un procés d\'atenció selectiva inconscient que guia cada fixació. Entrenar els jugadors significa, en part, optimitzar aquesta dansa."'
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08',
    tipus: 'quiz',
    titol: 'Comprensió: visió central i periférica',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc reflexiona sobre el que acaba de veure al 3x3 inicial. Sofia li planteja una pregunta directament relacionada amb el comportament de la Laia.',
    pregunta: 'En un exercici d\'1x1, Marc observa que la seva jugadora Laia sempre focalitza la mirada directament a l\'adversari defensora. Quina implicació pràctica principal té aquest comportament visual?',
    opcions: [
      {
        id: 'A',
        text: 'Perd informació sobre els companys i el porter, i consumeix temps en sacudides constants que alenteixen la seva presa de decisió',
        correcta: true,
        feedback: 'Exacte. **La visió central és precisa però costosa i intermitent**: cada fixació dura un temps i mentre fixa un element, els altres es mouen. En handbol, els temps de reacció requerits (50-200 ms) no permeten el "barratge central" continu. La Laia necessita aprendre a explotar la visió periférica, que li permetrà percebre el moviment dels companys i adversaris **sense desplaçar la mirada**.'
      },
      {
        id: 'B',
        text: 'No té cap implicació especial, ja que el que importa és la velocitat de les seves sacudides',
        correcta: false,
        feedback: 'La velocitat de les sacudides no és suficient per compensar la limitació de la visió central. **El problema no és la velocitat de l\'ull sinó l\'estructura del sistema visual**: mentre fixa a l\'adversari, el cervell de la Laia no pot processar amb precisió el que passa en la perifèria. Cada sacudida cap a un company costa mínim 200ms addicionals.'
      },
      {
        id: 'C',
        text: 'Hauria de mirar alternativament a la pilota i al portador contrari per màxima eficàcia',
        correcta: false,
        feedback: 'Alternar la mirada entre dos objectius (central alternat) ja és una millora, però segueix sent una estratègia **limitada als dos objectius seleccionats**. El que busquem és que la jugadora pugui percebre **tots** els elements de la situació simultàniament, cosa que requereix les estratègies perifériques.'
      }
    ],
    seguent: 'scene_09'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE II – La Visió i la Ment
  ═══════════════════════════════════════════════════════════ */

  scene_09: {
    id: 'scene_09',
    tipus: 'text_block',
    titol: 'Atenció i percepció',
    personatge: 'sofia',
    narracio: 'A la sessió següent, Sofia explica a Marc la distinció clau entre atenció difusa i atenció focal. Marc escolta amb creixent interès.',
    dialeg: {
      personatge: 'sofia',
      text: '"Hi ha dos grans modes d\'atenció: l\'**atenció difusa** i l\'**atenció focal** (o selectiva). L\'atenció difusa és igualitària: no privilegia cap punt del camp visual. L\'atenció focal selecciona un punt d\'interès específic.\n\nPerò l\'atenció focal té tres formes molt interessants:\n\n1. La mirada i l\'atenció van al mateix punt — la més freqüent, la que tots els jugadors fan per defecte.\n2. L\'atenció es **dissocia** de la mirada (\'mirada mental\'): l\'ull fixa un punt, però l\'atenció ja ha saltat al punt següent, preparant la sacudida.\n3. La dissociació és **persistent**: vigilar \'de reüll\', mantenir l\'atenció en un punt mentre la mirada és en un altre. Exigeix esforç sostingut però és molt valuosa per al defensor.\n\nLa \'**dansa de la mirada**\' és l\'expressió d\'una atenció selectiva inconscient: el cervell tria quins events visuals mereixen una fixació."'
    },
    contingut_pedagogic: {
      titol: 'La dissociació atenció-mirada: el secret de les estratègies visuals avançades',
      text: 'La clau de les estratègies visuals avançades en handbol és precisament la **dissociació entre on mirem i on posem l\'atenció**.\n\nUn jugador que ha après a dissocie l\'atenció de la mirada pot:\n- Fixar la mirada en la pilota i percebre simultàniament el moviment d\'un company en perifèria.\n- Fixar la mirada en un punt neutral i distribuir l\'atenció per tot el camp.\n- Preparar la propera sacudida mentre la mirada encara no s\'ha mogut.\n\nAquesta dissociació no és innata: s\'aprèn, però ha d\'aprendre\'s en el moment biològic adequat.'
    },
    seguent: 'scene_10'
  },

  scene_10: {
    id: 'scene_10',
    tipus: 'text_block',
    titol: 'Les sis estratègies visuals – Panorama general',
    personatge: 'narracio',
    narracio: 'Marc rep de Sofia un document amb el mapa complet de les estratègies visuals. Les llegeix lentament, assegut a la graderia mentre els jugadors fan escalfament.',
    dialeg: {
      personatge: 'sofia',
      text: '"Les estratègies visuals no són categories rígides. Representen un **contínuum evolutiu** dels estats atencionals del jugador. Les fluctuacions en el joc real poden ser extremadament ràpides, alternant diverses estratègies en el mateix segon.\n\nUn principi pedagògic fonamental: **s\'aborden primer en defensa** (on la urgència de la decisió és menor) i es transfereixen a l\'atac quan ja s\'han assolit. No al revés.\n\nAixí construïm de la base cap al vèrtex, sense saltar-nos etapes del calendari biològic."'
    },
    contingut_pedagogic: {
      titol: 'Les 6 estratègies visuals (contínuum)',
      text: '**1. Central iteratiu** → salta de fixació en fixació amb llargues pauses. Rudimentari, espontani.\n**2. Central continu** → mirada fixada en un sol objectiu, sacudides inhibides voluntàriament. Transicional.\n**3. Central alternat** → vaivé voluntari entre dos objectius (pilota + adversari).\n**4. Perifèric intermitent** → mirada a la pilota, adversari controlat en perifèria. Primera dissociació conscient.\n**5. Perifèric iteratiu** → mirada en punt no informatiu; atenció circula lliurement pel camp. Màxima eficiència.\n**6. Perifèric difús** → atenció igualitària a tot el camp. Forçat (pressió extrema) o espontani (talent).'
    },
    seguent: 'scene_11'
  },

  scene_11: {
    id: 'scene_11',
    tipus: 'text_block',
    titol: 'Del central iteratiu al central alternat',
    personatge: 'sofia',
    narracio: 'Sofia explica els tres primers tipus d\'estratègia visual amb exemples concrets de situacions de joc que Marc ha observat durant la setmana.',
    dialeg: {
      personatge: 'sofia',
      text: '"El **central iteratiu** és el punt de partida: el jugador salta de fixació en fixació amb llargues pauses. L\'ull identifica el pròxim objectiu en visió periférica, però l\'atenció se centra en el que fixa. Espontani en prebenjamins; el més rudimentari.\n\nEl **central continu** és una etapa transitòria: el jugador inhíbeix voluntàriament les sacudides i fixa un sol objectiu. Útil en defensa al principi de la fase de benjamins per establir la noció de responsabilitat defensiva —\'aquell és el meu adversari\'— però no es pot quedar aquí.\n\nEl **central alternat** és la primera millora real: l\'ull controla dos objectius, anant i venint. El nen espontàniament intenta acostar-se a l\'adversari per reduir l\'amplitud del desplaçament entre els dos objectius. El camp es 'comprimeix\' perceptivament."'
    },
    contingut_pedagogic: {
      titol: 'Evolució dels tres tipus "centrals"',
      text: '**Central iteratiu** → prebenjamins. Espontani i natural. Punt de partida.\n**Central continu** → inici de benjamins, en defensa. Etapa TRANSITÒRIA curta: el jugador aprèn que té un adversari "propi".\n**Central alternat** → progressió natural. El jugador "tria" instintivament escurçar les pauses per ser més eficaç. Busca una posició que redueixi l\'amplitud necessària entre els dos objectius d\'interès.\n\nEn defensa, el nen que ha assolit el central alternat comença a posicionar-se per tenir pilota i adversari al màxim de proper al centre del seu camp visual: primer senyal que s\'encamina cap al perifèric intermitent.'
    },
    seguent: 'scene_12'
  },

  scene_12: {
    id: 'scene_12',
    tipus: 'quiz',
    titol: 'El perifèric intermitent',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc observa durant l\'entrenament que un jove defensor ha modificat instintivament la seva posició per conservar la pilota i el seu adversari el màxim temps possible al seu camp visual.',
    pregunta: 'Un jove defensor comença a modificar instintivament la seva posició per conservar la pilota i el seu adversari el màxim temps possible al seu camp visual. Quina estratègia visual està assolint espontàniament?',
    opcions: [
      {
        id: 'A',
        text: 'Central iteratiu',
        correcta: false,
        feedback: 'El central iteratiu consisteix en saltar de fixació en fixació amb **llargues pauses**: no permet mantenir simultàniament dos objectius al camp visual. El comportament descrit —modificar posició per conservar dos objectius— és característic d\'un estadi més avançat.'
      },
      {
        id: 'B',
        text: 'Perifèric intermitent',
        correcta: true,
        feedback: 'Exacte! Quan el defensor modifica la seva posició per conservar pilota i adversari **simultàniament** al camp visual, s\'encamina cap al perifèric intermitent: mantenir la mirada a la pilota mentre **controla periféricament** l\'adversari directe. Aquesta és una etapa clau: per primera vegada, el jugador pren consciència del seu comportament oculomotor, ja que ha de mobilitzar l\'atenció cap a un objectiu mentre manté la mirada en un altre.'
      },
      {
        id: 'C',
        text: 'Perifèric iteratiu',
        correcta: false,
        feedback: 'El perifèric iteratiu és el pas **posterior**: requereix fixar la mirada en un punt virtual no informatiu (el "punt de perspectiva") mentre es desplaça l\'atenció pel camp periféric. El comportament descrit —modificar posició per conservar dos objectius— és el pas previ que obre la porta al perifèric intermitent.'
      }
    ],
    seguent: 'scene_13'
  },

  scene_13: {
    id: 'scene_13',
    tipus: 'text_block',
    titol: 'El perifèric iteratiu – L\'estratègia avançada',
    personatge: 'sofia',
    narracio: 'Sofia explica l\'estratègia visual més avançada que es pot treballar sistemàticament. Marc s\'inclina endavant, atent.',
    dialeg: {
      personatge: 'sofia',
      text: '"L\'objectiu del **perifèric iteratiu** és captar informació de tot el camp visual desplaçant l\'atenció sense desplaçar la mirada. L\'ull s\'ancora en un punt estratègic i el cervell 'explora\' la perifèria per atenció pura, sense sacudides.\n\nPer a l\'atacant amb pilota en situació d\'1x1: el punt de fixació és el **centre del travesser**. La porteria és lluny, fixa, sempre al mateix lloc: el punt perfecte per ancorar la mirada.\n\nPer al defensor: el \'**punt de perspectiva**\' és un punt virtual al terra, aproximadament a la bisectriu de l\'angle pilota-defensor-oponent, a una distància de 5 a 10 metres.\n\nEl guany de velocitat és espectacular: desplaçar l\'atenció sense moure els ulls triga **menys de 30 ms**. Una sacudida ocular triga com a mínim **500 ms**. Un defensor exterior amb perifèric iteratiu reacciona en **350 ms** totals. Amb central iteratiu, els mateixos 700 ms. En el moment en que el defensor reacciona, l\'extremer ha recorregut el **doble de camí**."'
    },
    contingut_pedagogic: {
      titol: 'Seqüència de joc amb els tres estats atencionals (defensor extern)',
      text: '**Estat 1 (pilota lluny)** → perifèric difús: atenció distribuïda equitativament per la meitat defensiva del camp.\n**Estat 2 (pilota s\'apropa)** → perifèric iteratiu: mirada al punt de perspectiva (terra), atenció al triangle pilota-adversari directe-porters.\n**Estat 3 (1x1 imminent)** → perifèric intermitent o iteratiu accelerat: mirada fixa, atenció al pivot que pot rebre.\n\nA partir d\'un cert nivell, el defensor exterior pot permetre\'s deixar el seu oponent directe a la semi-lluna monocular (fora del camp binocular) si es manté a 6 metres: la visió monocular periférica és suficient per detectar-ne el moviment.'
    },
    seguent: 'scene_14'
  },

  scene_14: {
    id: 'scene_14',
    tipus: 'checklist',
    titol: 'Identificació de les estratègies visuals',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc revisa els seus apunts sobre les sis estratègies visuals. Sofia li proposa un exercici de verificació.',
    contingut_pedagogic: {
      titol: 'Recorda els criteris de cada estratègia',
      text: '**Central iteratiu** = sacudides amb pauses llargues.\n**Central continu** = fixació única, sacudides inhibides.\n**Central alternat** = vaivé entre dos objectius.\n**Perifèric intermitent** = mirada a un objectiu, atenció a un altre.\n**Perifèric iteratiu** = mirada en punt no informatiu, atenció circula lliurement.\n**Perifèric difús** = atenció igualitària a tot el camp visual.'
    },
    pregunta: 'Marca les afirmacions CORRECTES sobre les estratègies visuals:',
    checklistItems: [
      {
        text: 'Un atacant que mira la porta i és conscient del moviment dels defensors als costats sense desplaçar la mirada utilitza el **perifèric iteratiu**.',
        correcta: true
      },
      {
        text: 'Un jugador que mira l\'adversari contínuament ignorant la resta del camp utilitza el **central continu**.',
        correcta: true
      },
      {
        text: 'El **perifèric difús** és l\'estratègia més fàcil d\'aprendre i és accessible a tots els jugadors.',
        correcta: false
      },
      {
        text: 'En el **perifèric iteratiu**, el punt de perspectiva per al defensor s\'ubica aproximadament a la bisectriu de l\'angle pilota-defensor-oponent.',
        correcta: true
      },
      {
        text: 'Les estratègies visuals s\'aborden primer en **atac** perquè és on hi ha menys pressió temporal.',
        correcta: false
      }
    ],
    feedback: 'Les estratègies visuals segueixen una progressió estricta i s\'introdueixen **primer en defensa** (menys pressió temporal). El perifèric difús espontani és molt rar i propi del talent excepcional.',
    seguent: 'scene_15'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE III – Les Estratègies Visuals (continuació)
  ═══════════════════════════════════════════════════════════ */

  scene_15: {
    id: 'scene_15',
    tipus: 'text_block',
    titol: 'El perifèric difús i la creativitat',
    personatge: 'sofia',
    narracio: 'Sofia explica l\'última estratègia visual i les condicions que fan possible la creativitat tàctica màxima.',
    dialeg: {
      personatge: 'sofia',
      text: '"El **perifèric difús forçat** apareix en situacions de pressió temporal extrema: el contraatac ràpid on el jugador passa la pilota 'sense temps de mirar\'. No és un acte a l\'atzar: el cervell utilitza informació summera sobre els desplaçaments de mòbils no identificats per calcular el passe òptim. El resultat sembla màgic però és ciència.\n\nEl **perifèric difús espontani** és molt rar. El jugador mira 'a cap lloc\' voluntàriament, confiant plenament en el seu cervell. Requereix una confiança absoluta en els processos inconscients. És el gest del jugador de talent excepcional que 'sempre sap on és tothom\'.\n\nUna nota tècnica important: en visió periférica (escala de grisos), dos jugadors que corren en el mateix sentit amb **camisetes del mateix valor lluminós** són indistingibles. Solució: motius molt diferents i de gran superfície en els uniformes."'
    },
    contingut_pedagogic: {
      titol: 'Condicions per a la creativitat tàctica màxima',
      text: '**1. Equilibri entre hemisferis cerebrals** → evitar la dominació de l\'hemisferi esquerre (analític-verbal) durant l\'execució.\n**2. El perifèric iteratiu** = màxima explotació del potencial perceptiu sense sobrecarregar cap canal.\n**3. Evitar instruccions verbals durant l\'acció** → no activar innecessàriament l\'hemisferi esquerre en el moment de la decisió.\n**4. Confiar en les reaccions intuïtives** → el cervell ha après; deixar-lo actuar.\n\nEl jugador creatiu no pensa "he de passar a la dreta". El seu cervell li dona la pilota al company desmarcat sense que ell sàpiga exactament com ha pres la decisió.'
    },
    seguent: 'scene_16'
  },

  scene_16: {
    id: 'scene_16',
    tipus: 'text_block',
    titol: 'Les estratègies visuals de l\'entrenador',
    personatge: 'sofia',
    narracio: 'Sofia recorda a Marc que les estratègies visuals no són únicament per als jugadors. L\'entrenador també necessita cultivar la seva pròpia mirada.',
    dialeg: {
      personatge: 'sofia',
      text: '"Tu, com a entrenador, has de diagnosticar el rendiment tàctic del teu equip. Si segueixes la pilota com un espectador passiu, et perds exactament el que necessites veure: els moviments dels jugadors sense pilota, els espais que s\'obren i es tanquen, les decisions que s\'inicien i s\'abandonen.\n\nHas d\'**ancorar** el teu eix visual en el centre de gravetat virtual de l\'escena observada, no a la pilota. I quan calgui, canviar el punt de perspectiva a mesura que la situació evoluciona.\n\nDespré de cada seqüència, **repassa conscientment** el que has captat inconscientment: tens un màxim de 10 minuts abans que la informació es degradi (com els escaquistes en partides simultànies que repassen les posicions entre jugadors).\n\nUna advertència: el vídeo en càmera lenta és útil per a molts objectius, però **NO beneficia** les capacitats perceptives. Entrenar la percepció en temps real és insubstituïble."'
    },
    contingut_pedagogic: {
      titol: 'Com entrenar l\'observació dels entrenadors',
      text: 'Exercici pràctic per a entrenadors:\n- **Grup observador** posicionat al cercle central, tots mirant al centre del travesser.\n- Mirar la situació de joc "de reüll" (dissociació atenció-mirada).\n- Identificar els moviments dels jugadors sense pilota sense sacudir els ulls.\n- Immediatament après de cada seqüència (màxim 10 min): verbalitzar en veu alta el que s\'ha percebut inconscientment.\n\nAquest exercici, practicat regularment, millora dràsticament la capacitat de diagnòstic tàctic dels entrenadors.'
    },
    seguent: 'scene_17'
  },

  scene_17: {
    id: 'scene_17',
    tipus: 'text_block',
    titol: 'La neuroplasticitat i el calendari biològic',
    personatge: 'narracio',
    narracio: 'Marc llegeix el capítol sobre neuroplasticitat. Per primera vegada, entén per què l\'edat d\'inici és tan crítica en la formació dels jugadors.',
    dialeg: {
      personatge: 'sofia',
      text: '"El cervell del nadó pesa un cinquè del pes adult, però ja conté els 100.000 milions de neurones definitives. El creixement posterior serà gràcies al 'cablatge\' —les connexions sinàptiques— i a la mielinització dels axons, que augmenta dràsticament la velocitat de transmissió de l\'impuls nerviós. Tot el procés pràcticament acaba al final de la pubertat.\n\nGenètica i aprenentatge determinen junts l\'organització final del cervell. La capacitat plàstica és màxima fins al principi de la pubertat: és el que s\'anomena el **Calendari Biològic dels Aprenentatges**.\n\nEl perill contrari és igual de real: la repetició de les mateixes conductes consolida xarxes neuronals rígides que generen el \'jugador sobre rails\': previsible, correcte, però incapaç de solucions creatives quan el joc es desvia del seu guió habitual."'
    },
    contingut_pedagogic: {
      titol: 'Implicació pràctica del calendari biològic',
      text: '**En handbol**: inici recomanat als 6-8 anys. No abans dels 6 (immaduresa del sistema nerviós per a la coordinació específica). Quant més a prop del principi del període crític, més ràpida és l\'adquisició.\n\n**Neuroplasticitat màxima** → fins a la pubertat. La finestra és real i limitada.\n**Calendari biològic** → cada habilitat té el seu moment òptim d\'adquisició. Saltar-se etapes no accelera: retarda.\n**Especialització precoç** → tanca portes. Polivalència en la formació → les manté obertes.\n\nConseqüència directa: el programa de formació perceptiva ha de respectar escrupolosament el calendari biològic, no imposar la lògica dels adults als cervells en formació.'
    },
    seguent: 'scene_18'
  },

  scene_18: {
    id: 'scene_18',
    tipus: 'decisio',
    titol: 'L\'entrenament de la neuroplasticitat',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc dissenya el programa de formació per als prebenjamins (6-8 anys) del Club Handbol Garraf. Ha de decidir l\'ordre d\'introducció de les estratègies visuals.',
    pregunta: 'Quin és l\'enfocament correcte per introduir les estratègies visuals als prebenjamins?',
    opcions: [
      {
        id: 'A',
        punts: 0,
        text: 'Comencem directament amb el perifèric iteratiu (la més avançada) perquè si l\'assoleixen de joves serà més fàcil',
        feedback: 'El **perifèric iteratiu** és una estratègia avançada que requereix el desenvolupament previ de les anteriors. Saltar-se les etapes no accelera el procés; al contrari, genera confusió. Cada estratègia és una fita del **calendari biològic dels aprenentatges** que prepara la següent.',
        seguent: 'scene_19'
      },
      {
        id: 'B',
        punts: 10,
        text: 'Seguim la progressió natural: central iteratiu → central continu → central alternat → perifèric intermitent → perifèric iteratiu, **introduint-ho primer en defensa** per reduir la pressió temporal de la decisió',
        feedback: 'Excel·lent. La progressió segueix el **calendari biològic natural** i respecta un principi fonamental: s\'aborda primer en **defensa** (on la urgència de la decisió és menor) i es trasllada a l\'atac quan ja s\'ha assolit. A cada fase, el jugador "tria inconscientment" les conductes més competitives, sense que l\'entrenador hagi d\'intervenir verbalment en excés.',
        seguent: 'scene_19'
      },
      {
        id: 'C',
        punts: 5,
        text: 'Introduïm el central iteratiu per a prebenjamins i el perifèric intermitent per a benjamins, però en atac, perquè és on es veuen els beneficis',
        feedback: 'La progressió temporal és correcta però **l\'ordre atac-defensa és l\'invers del recomanat**. La raó és clara: en defensa, la urgència temporal és menor i el jugador pot permetre\'s anar assolint la nova estratègia sense l\'estrès de la decisió ofensiva immediata. Introduir-ho en atac primer afegeix una pressió que dificulta l\'aprenentatge.',
        seguent: 'scene_19'
      }
    ],
    seguent: 'scene_19'
  },

  scene_19: {
    id: 'scene_19',
    tipus: 'text_block',
    titol: 'Competir per progressar, divertir-se per aprendre',
    personatge: 'sofia',
    narracio: 'En una conversa al final de l\'entrenament, Sofia i Marc parlen sobre el rol de la competició i el joc en el desenvolupament del jugador jove.',
    dialeg: {
      personatge: 'sofia',
      text: '"Observa nens de 10-12 anys en joc lliure al pati: creativitat motriu rica, solucions imprevistes, joc improvist i variat. Ara observa els mateixos nens en un col·lectiu esportiu organitzat: rigidesa de conductes, aparent pobresa d\'imaginació. Per què?\n\nPerquè els educadors apliquen als nens els mateixos mètodes que als adults: esquemes fixats, rols assignats, posicions definides. I 'fossilitzen\' les aptituds per a la invenció motriu just en el moment en que havien de florir.\n\nLa **competició** és un motor fonamental del progrés tàctic: el joc lliure i la competència autèntica creen problemes que el cervell ha de resoldre de forma creativa. La diversió elimina les inhibicions sobre l\'expressió creativa. I recordes la pregunta correcta? No: 'Com hauries d\'haver-ho fet?\' sinó: 'Què has intentat?\'."'
    },
    contingut_pedagogic: {
      titol: 'Les dues raons per preservar la creativitat',
      text: '**1. Els esports d\'equip com a motor de desenvolupament integral** → el joc esportiu és un dels entorns més rics per al desenvolupament de la intel·ligència tàctica i social del nen.\n**2. La desaparició precoç del jugador creatiu** → l\'excés d\'entrenament rígid extingeix les conductes creatives per condicionament: el jugador aprèn que les solucions "correctes" (les del mànager) reben recompenses, i les solucions creatives (les seves) reben crítiques.\n\nPreguntar **"Què has intentat?"** (objectiu) en lloc de **"Com hauries d\'haver-ho fet?"** (mitjà) preserva l\'autonomia decisional i estimula la reflexió sobre el propòsit de l\'acció.'
    },
    seguent: 'scene_20'
  },

  scene_20: {
    id: 'scene_20',
    tipus: 'text_block',
    titol: 'El treball global i analític',
    personatge: 'sofia',
    narracio: 'Marc pregunta a Sofia com combinar el treball tècnic específic amb el joc global. Sofia explica la trampa de l\'espai protegit.',
    dialeg: {
      personatge: 'sofia',
      text: '"El treball analític i el treball global s\'han de combinar, però **mai en espais aïllats**. El perill dels exercicis analítics en \'espai protegit\' és que generen un hàbit perceptiu invers: el jugador aprèn a focalitzar l\'oponent en un gran espai buit, perquè és l\'única referència disponible. Quan torna al joc real, segueix buscant la silhoueta de l\'oponent en lloc de llegir els espais.\n\nEn joc real, passa el contrari: els espais buits 'tancats\' destaquen millor del fons perquè hi ha molts jugadors. El bon jugador **llegeix la defensa en negatiu**: no veu els jugadors, veu els espais entre els jugadors. Com un escaquista que llegeix el tauler.\n\n**Obstrucció aleatòria del camp visual**: les zones de treball han de **superposar-se**, no separar-se. Els obstacles imprevisibles de les altres parelles obliguen a respostes 'reflex\' per evitar-los, cultivant exactament les capacitats perifériques que busquem."'
    },
    contingut_pedagogic: {
      titol: 'Principis de la sessió d\'entrenament',
      text: '**Ritme i continuïtat** → les sessions no han de tenir llargues pauses; el ritme alt simula la pressió temporal real.\n**Jugadors arbitren** → des de benjamins, els jugadors arbitren els entrenaments. Desenvolupa la comprensió del joc i redueix la dependència de l\'autoritat de l\'entrenador.\n**Zones superposades** → mai separar les zones de treball; les interferències aleatòries són part del procés d\'aprenentatge.\n**Contagi per encadenament** → un exercici "contamina" el següent; les noves habilitats es consoliden per transferència lateral entre activitats.'
    },
    seguent: 'scene_21'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE IV – La Pedagogia Creativa
  ═══════════════════════════════════════════════════════════ */

  scene_21: {
    id: 'scene_21',
    tipus: 'decisio',
    titol: 'El disseny de l\'exercici de transmissions',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc vol dissenyar un exercici d\'1x1 per treballar la lectura de la defensa. Tres opcions es presenten a la seva ment.',
    pregunta: 'Quina és la millor estructura de l\'exercici d\'1x1 per treballar la lectura de la defensa?',
    opcions: [
      {
        id: 'A',
        punts: 0,
        text: 'Dues files d\'un cada una als costats de l\'àrea, actuen una parella cada vegada en l\'espai central lliure i protegit. Marc explica before cada repetició el que ha d\'intentar.',
        feedback: '**Doble error pedagògic**: espai protegit (els impedeix treballar la percepció periférica d\'obstacles) i instruccions verbals prèvies (bloquegen l\'aprenentatge inconscient). En un gran espai buit, els jugadors tendiran a focalitzar l\'oponent (silhoueta aïllada) en lloc de llegir els espais.',
        seguent: 'scene_22'
      },
      {
        id: 'B',
        punts: 10,
        text: 'Diverses parelles treballant simultàniament en zones **superposades**, amb el portador de la pilota obligat a llegir els espais entre els adversaris (els "espais negatius") sense instruccions verbals. Marc modifica les condicions subtilment cada 3-4 repeticions.',
        feedback: 'Perfecte. La simultaneïtat obliga a **l\'atenció periférica**: els obstacles imprevisibles de les altres parelles forcen respostes reflex. Les zones superposades eviten l\'hàbit de focalitzar l\'oponent en gran espai buit. Sense instruccions verbals, l\'aprenentatge inconscient fa la seva feina.',
        seguent: 'scene_22'
      },
      {
        id: 'C',
        punts: 5,
        text: 'Diverses parelles en zones separades, Marc intervé verbalment quan un jugador falla per guiar-lo cap a la decisió correcta',
        feedback: 'La simultaneïtat és bona, però les zones **separades** i les **intervencions verbals** limiten l\'efectivitat. Les zones separades eliminen l\'\'obstrucció aleatòria\' que força respostes perifériques, i les intervencions verbals interrompen el processament inconscient en el moment crític.',
        seguent: 'scene_22'
      }
    ],
    seguent: 'scene_22'
  },

  scene_22: {
    id: 'scene_22',
    tipus: 'text_block',
    titol: 'El rebuig de l\'especialització precoç',
    personatge: 'sofia',
    narracio: 'Sofia explica a Marc per què l\'especialització precoç és un dels errors més greus que pot cometre un formador.',
    dialeg: {
      personatge: 'sofia',
      text: '"L\'especialització precoç no és tan sols contraproduent per crear jugadors polivalents —que ja seria raó suficient—. Sinó perquè **l\'audàcia i la creativitat no es poden expressar** si el jugador no se sent còmode en qualsevol lloc del camp.\n\nSi assignem una funció concreta al jugador, li enviem un missatge implícit: 'Has de fer exactament això, en exactament aquest lloc, exactament d\'aquesta manera.\' I el jugador s\'autocensura quan la situació reclama una solució que surt del seu \'guió\'.\n\nVols que el jugador se senti \'acreditat\' per intentar-ho tot? Llavors no li assignis una funció concreta. La rigidesa de xarxes neuronals es crea per especialització prematura."'
    },
    contingut_pedagogic: {
      titol: 'Cicles de formació i "cursillo inaugural"',
      text: 'La Laia, extremera dreta per costum, és inclosa en exercicis com a central. Primera reacció: inhibida, insegura. Però Sofia diu: "D\'aquí a 10 sessions, miraràs el camp de forma completament diferent."\n\n**Cicles de 6-8 setmanes**: cada nou conjunt d\'exercicis s\'introdueix amb un "cursillo inaugural" on els jugadors descobreixen la nova habilitat per primera vegada. Sense instruccions prèvies detallades: la descoberta autónoma és la base.\n\n**Tres parts de la sessió** amb desfasament de 6-8 setmanes entre elles:\n- Primera part: exercicis en consolidació (assolits fa 6-8 setmanes).\n- Segona part: exercicis en adquisició (en curs).\n- Tercera part: exercicis d\'exploració (nous, presentats per primera vegada).'
    },
    seguent: 'scene_23'
  },

  scene_23: {
    id: 'scene_23',
    tipus: 'quiz',
    titol: 'Les transmissions i l\'anticipació espacial',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc estudia les tècniques de transmissió descrites per Pinaud i Díez. Una en particular li crida l\'atenció: el "passe inductor".',
    pregunta: 'En el context del "passe inductor", quin és el mecanisme fonamental que el portador de la pilota intenta activar?',
    opcions: [
      {
        id: 'A',
        text: 'Forçar una sacudida de la defensa (un desplaçament d\'atenció o de mirada defensiva) per crear un espai lliure per al company',
        correcta: true,
        feedback: 'Exacte. El passe inductor és una acció del portador de la pilota que força una **reacció perceptiva** en la defensa: una sacudida cap a ell, que distreu l\'atenció defensiva del company que s\'ha de demarcar. L\'atacant "indueix" la defensa a reaccionar en el moment i la direcció que ell controla. Aquesta acció requereix que el portador de pilota tingui una **bona percepció periférica** per executar-la mentre observa el company.'
      },
      {
        id: 'B',
        text: 'Passar la pilota el màxim de ràpid possible per sorprendre la defensa',
        correcta: false,
        feedback: 'La velocitat del passe no és el factor clau en el passe inductor. El que importa és el **timing**: el portador ha de crear primer la reacció defensiva amb el seu fint o amenaça, i **llavors** executar el passe cap al company que ha quedat desmarcat. La velocitat pot fins i tot ser contraproduent si no es respecta el ritme del "induïment".'
      },
      {
        id: 'C',
        text: 'Establir contacte visual amb el receptor per confirmar que està preparat per rebre',
        correcta: false,
        feedback: 'Establir contacte visual amb el receptor seria contraproduent: alertaria la defensa sobre la intenció. En el passe inductor, el portador de la pilota ha de **simular** dirigir-se a un costat (induint la sacudida defensiva) mentre periféricament controla el company que es desmarcarà. El contacte visual directe destrueix l\'element sorpresa.'
      }
    ],
    seguent: 'scene_24'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE V – Metodologia Pràctica
  ═══════════════════════════════════════════════════════════ */

  scene_24: {
    id: 'scene_24',
    tipus: 'text_block',
    titol: 'Defensa individual en tot el camp i el bot únic',
    personatge: 'sofia',
    narracio: 'Sofia presenta dos exercicis metodològics fonamentals: la defensa individual en tot el camp per al treball perceptiu del defensor, i el bot únic per al portador de la pilota.',
    dialeg: {
      personatge: 'sofia',
      text: '"El treball perceptiu del defensor és un pas del central iteratiu al perifèric intermitent. Però atenció: la instrucció clau **no fa referència a la \'visió periférica\'**. Definiràs l\'objectiu global al jugador: \'Sense perdre de vista la pilota, has de saber sempre on és el teu adversari.\'\n\nEl jugador, per primera vegada, pren consciència del seu propi comportament oculomotor. Descobreix que havia estat 'alternant\' mecànicament entre pilota i adversari. Ara sap que pot fer-ho millor.\n\nEl **bot únic en tot el camp** sembla un exercici tècnic. En realitat és un exercici perceptiu: elimines el bot com a refugi per forçar la lectura del joc. Quan el jugador no pot fer més d\'un bot, no pot aturar-se a pensar: ha de llegir la defensa en moviment, constantment, sense descans."'
    },
    contingut_pedagogic: {
      titol: 'Progressió de la defensa individual: del central iteratiu al perifèric intermitent',
      text: '**Instrucció correcta** (NO mencionar "visió periférica"): "Sense perdre de vista la pilota, has de saber sempre on és el teu adversari."\n\n**Etapa 1**: el defensor aprèn a mantenir la pilota en el camp central i l\'adversari en la perifèria.\n**Etapa 2**: modifica instintivament la posició per conservar els dos elements el màxim de temps possible simultàniament.\n**Etapa 3**: estabilitza el perifèric intermitent → base per al perifèric iteratiu posterior.\n\n**Bot únic en tot el camp** → el portador ha de llegir la defensa en moviment. Elimina la "muleta" del bot repetit. Força decisions ràpides i aprenentatge de l\'anticipació espacial (passar on ESTARÀ el receptor, no on és).'
    },
    seguent: 'scene_25'
  },

  scene_25: {
    id: 'scene_25',
    tipus: 'decisio',
    titol: 'La resposta al veteran',
    personatge: 'marc',
    punts: 10,
    narracio: 'Jordi, jugador veteran amb 15 anys d\'experiència, interromp l\'explicació de Marc sobre les estratègies visuals.',
    dialeg: {
      personatge: 'jordi',
      text: '"Jo sempre he vist el camp sencer. El secret és seguir la pilota amb els ulls. Sempre m\'ha funcionat."'
    },
    pregunta: 'Com hauria de respondre Marc a Jordi?',
    opcions: [
      {
        id: 'A',
        punts: 0,
        text: 'Tens raó, Jordi. Continua fent el que fas, que et funciona.',
        feedback: 'Deixar Jordi sense qüestionar el seu mètode impedeix el seu progrés. Seguir la pilota amb els ulls = **central iteratiu en el millor cas**. Si Jordi "veu el camp sencer", és perquè ha après a fer sacudides molt ràpides, però segueix **perdent temps** en cada transició. El perifèric iteratiu li donaria un avantatge de 350 ms en cada reacció defensiva.',
        seguent: 'scene_26b'
      },
      {
        id: 'B',
        punts: 10,
        text: 'Entenc que et funciona, Jordi. Però prova un experiment: en el proper 3x3, fixa la mirada en el centre del travesser quan tens la pilota, i intenta que la teva atenció vagi als companys i adversaris sense moure els ulls. Després ens dius si has vist més coses.',
        feedback: 'Perfecte. En lloc de contradir-lo directament, **proposes una experiència pràctica** que li permetrà descobrir per si mateix la diferència. Marc preserva l\'autonomia de Jordi (clau per a la motivació) mentre l\'introdueix al perifèric iteratiu de forma pràctica i no verbal.',
        seguent: 'scene_27'
      },
      {
        id: 'C',
        punts: 5,
        text: 'Jordi, el que fas s\'anomena "central iteratiu". Hi ha 5 estratègies visuals més avançades. T\'explico el perifèric iteratiu...',
        feedback: 'L\'explicació teòrica pot ser útil però en aquest moment és prematura. Jordi aprèn millor a través de l\'**experiència pràctica** que de l\'explicació verbal. Un jugador de 15 anys d\'experiència integra nova informació millor quan la **descobreix ell mateix** en entrenament que quan li és explicada teòricament des de la banda.',
        seguent: 'scene_26b'
      }
    ],
    seguent: 'scene_27'
  },

  scene_26: {
    id: 'scene_26',
    tipus: 'text_block',
    titol: 'El punt de perspectiva',
    personatge: 'sofia',
    narracio: 'Sofia explica a Marc els detalls pràctics del punt de perspectiva per al defensor i els exercicis específics per entrenar-lo.',
    dialeg: {
      personatge: 'sofia',
      text: '"El punt de perspectiva no és un punt fix. Canvia a cada instant en funció de la posició de la pilota i de l\'adversari directe. La ubicació correcta és al terra, aproximadament a la bisectriu de l\'angle format per la pilota, el defensor, i l\'oponent, a una distància de 5 a 10 metres.\n\nA cada desplaçament de la pilota o de l\'oponent, el defensor recalcula mentalment el seu punt de perspectiva i hi dirigeix la mirada. En certa mesura, re-enactua el central iteratiu, però ara les sacudides no van d\'objectiu informatiu a objectiu informatiu: busquen en cada moment el punt de perspectiva que maximitza l\'amplitud del camp visual útil."'
    },
    contingut_pedagogic: {
      titol: 'Exercicis específics per entrenar el punt de perspectiva',
      text: '**Exercici 1: Treball defensiu monocular** → el defensor tanca l\'ull del costat de la pilota. Això l\'obliga a fer la percepció de l\'adversari en visió binocular des del punt de perspectiva correcte, sense la "muleta" de l\'ull directament orientat a la pilota.\n\n**Exercici 2: Focus visual fix (grup observador)** → grup posicionat al cercle central, tots mirant el centre del travesser. Des d\'aquí, practiquen detectar moviments en la perifèria sense sacudir els ulls.\n\n**La recepció sense mirar** → el receptor practica agafar la pilota en visió periférica, sense girar el cap. Força el cervell a calcular la trajectòria balística de la pilota amb informació periférica: cal·libra la visió periférica per a càlculs de profunditat.\n\n**Fixació visual al travesser** → l\'atacant fixa la mirada al centre del travesser i executa el llançament per visió periférica. Sembla impossible; en pocs entrenaments, la precisió millora perquè la visió periférica era ja suficientment precisa per a càlculs balístics en handbol.'
    },
    seguent: 'scene_27'
  },

  scene_26b: {
    id: 'scene_26b',
    tipus: 'text_block',
    titol: 'Per entendre Jordi',
    personatge: 'sofia',
    narracio: 'Sofia s\'acosta a Marc i, en veu baixa, explica per què la resposta que ha donat no és la ideal.',
    dialeg: {
      personatge: 'sofia',
      text: '"Jordi té 15 anys d\'experiència. La seva resistència al canvi és completament normal: ha construït una identitat esportiva al voltant del que sap fer, i qüestionar-ho és qüestionar-lo a ell. La clau no és convèncer-lo amb arguments. La clau és que **ho descobreixi ell sol**.\n\nEl teu paper com a entrenador no és dir als jugadors el que han de fer. És **crear les condicions perquè ho descobreixin**. Organitza experiències, no transmiteixis instruccions. Proposa experiments, no dicti tècniques.\n\nAmb Jordi: proposa-li un experiment pràctic. \'Prova de fixar la mirada al centre del travesser en el proper 3x3 i digues-me si veus més coses.\' Quan l\'experiència li confirmi la diferència, la nova estratègia serà seva, no teva."'
    },
    contingut_pedagogic: {
      titol: 'L\'entrenador com a "organitzador d\'experiències"',
      text: '**Principi pedagògic d\'autonomia**: el jugador que descobreix per si mateix una nova habilitat l\'integra de forma molt més profunda i durable que el jugador al que li és explicada teòricament.\n\n**L\'entrenador no és un transmissor d\'instruccions**. És un dissenyador d\'entorns d\'aprenentatge que ofereix al cervell del jugador les condicions òptimes per a la descoberta autónoma.\n\n**Resistència al canvi en jugadors veterans** → completament normal. No atacar la resistència: crear experiències que la superin orgànicament des de dins.'
    },
    seguent: 'scene_27'
  },

  scene_27: {
    id: 'scene_27',
    tipus: 'checklist',
    titol: 'Bones pràctiques de l\'entrenament perceptiu',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc revisa els principis fonamentals del programa de formació perceptiva que ha après fins ara.',
    contingut_pedagogic: {
      titol: 'Recorda els principis fonamentals del programa de formació perceptiva en handbol',
      text: 'Marca les afirmacions que consideres **correctes** sobre el programa de formació perceptiva:'
    },
    pregunta: 'Marca les afirmacions CORRECTES sobre les bones pràctiques de l\'entrenament perceptiu:',
    checklistItems: [
      {
        text: 'Les instruccions verbals durant l\'acció d\'un jugador s\'han d\'evitar: bloquegen el processament inconscient.',
        correcta: true
      },
      {
        text: 'El vídeo en càmera lenta és el millor recurs per millorar les capacitats perceptives dels jugadors.',
        correcta: false
      },
      {
        text: 'L\'especialització precoç en una posició augmenta la creativitat tàctica del jugador.',
        correcta: false
      },
      {
        text: 'Les zones de treball superposades (en lloc de separades) milloren l\'aprenentatge perceptiu.',
        correcta: true
      },
      {
        text: 'El perifèric iteratiu s\'ha d\'aprendre primer en defensa, quan la pressió temporal de la decisió és menor.',
        correcta: true
      }
    ],
    feedback: 'Les instruccions verbals durant l\'acció interrompen el processament inconscient. El vídeo lent és útil per a d\'altres objectius però NO millora la percepció en temps real. L\'especialització precoç **redueix** la creativitat. Les zones superposades i l\'ordre defensa-atac són fonamentals.',
    seguent: 'scene_28'
  },

  scene_28: {
    id: 'scene_28',
    tipus: 'decisio',
    titol: 'La presentació del programa',
    personatge: 'marc',
    punts: 10,
    narracio: 'Marc ha d\'exposar el seu nou programa de formació al president del club i als pares dels jugadors. La directora tècnica li planteja la pregunta que l\'ha estat preocupant.',
    dialeg: {
      personatge: 'sofia',
      text: 'La directora tècnica del club diu: "Els pares m\'han demanat per què els seus fills no fan exercicis de llançament a porteria en els entrenaments de prebenjamins."'
    },
    pregunta: 'Com hauria de respondre Marc a la directora tècnica i als pares?',
    opcions: [
      {
        id: 'A',
        punts: 10,
        text: 'Perquè en la categoria prebenjamí, el **principal objectiu és el desenvolupament perceptiu i la creativitat tàctica**, que depèn de la neuroplasticitat màxima d\'aquesta franja d\'edat. Els exercicis de llançament aïllats generen hàbits visuals incorrectes que dificulten la creativitat posterior. Ho treballem integrat en el joc.',
        feedback: 'Resposta sòlida i basada en la recerca. Marc posa el **calendari biològic dels aprenentatges** al centre de la decisió pedagògica i explica la raó per la qual els exercicis aïllats (com el llançament a porteria) en espai protegit generen transferència negativa. Els pares i la directora podran entendre la lògica a llarg termini.',
        seguent: 'scene_29'
      },
      {
        id: 'B',
        punts: 0,
        text: 'Perquè el llançament és una habilitat fàcil que ja aprendran sols. El que és difícil és la tàctica.',
        feedback: 'Aquesta resposta és **arrogant i poc convincent** per als pares, i a més és incorrecta des del punt de vista pedagògic: el llançament **ha de treballar-se**, però integrat en el joc, no com a exercici aïllat. La raó per evitar exercicis de llançament aïllats no és que siguin fàcils, sinó que generen hàbits perceptius incorrectes.',
        seguent: 'scene_29'
      },
      {
        id: 'C',
        punts: 5,
        text: 'Perquè seguim el mètode de Pinaud i Díez que és el millor a nivell internacional.',
        feedback: 'Fer referència a l\'autoritat de l\'obra és correcte, però **no explica el perquè** als pares. Una explicació basada en la neurociència i en la pedagogia de la creativitat seria molt més convincent i educativa. Els pares entenen millor els principis que els noms d\'autors.',
        seguent: 'scene_29'
      }
    ],
    seguent: 'scene_29'
  },

  scene_29: {
    id: 'scene_29',
    tipus: 'epilog',
    titol: 'La Nova Mirada – Resultats',
    personatge: 'sofia',
    narracio: 'Han passat sis mesos. El Club Handbol Garraf ha completat el primer cicle complet amb el nou programa de formació perceptiva.',
    mentorMsgs: {
      excellent: '**Excepcional, Marc.** Has assimilat i aplicat els principis de la percepció i la creativitat amb una comprensió profunda. El teu equip tindrà el millor fonament perceptiu possible. Laia ja mira el camp de forma completament diferent, Jordi ha descobert el perifèric iteratiu per experiència pròpia, i els prebenjamins juguen amb una llibertat que sorprèn els visitants. Has entès que **l\'entrenador és un organitzador d\'experiències, no un transmissor d\'instruccions**.',
      good: '**Molt bé, Marc.** Has entès els principis fonamentals i has pres la majoria de decisions correctes. Amb la pràctica i el refinament continu, el teu programa de formació donarà resultats excel·lents. Recorda que els principis de la percepció i la creativitat s\'apliquen millor quan l\'entrenador confies en el procés i deixa que el cervell dels jugadors faci la seva feina.',
      ok: '**Correcte, Marc.** Has captat els conceptes bàsics, però en alguns moments has recaigut en els patrons convencionals: instruccions verbals durant l\'acció, espais aïllats, urgència per veure resultats immediats. El programa de formació perceptiva requereix paciència i confiança en els processos inconscients. Repassa els principis del calendari biològic i la dissociació atenció-mirada.',
      needsWork: '**Necessites aprofundir, Marc.** Els principis de Pinaud i Díez representen un canvi de paradigma profund, i és normal que no s\'integrin tots d\'un cop. Et recomanem repassar el Capítol I sobre bases neurofisiológiques i el Capítol II sobre pedagogia de la creativitat. Posa especial atenció a la diferència entre aprenentatge conscient i inconscient, i al principi fonamental: **l\'entrenador crea condicions, no transmet instruccions**.'
    },
    seguent: 'scene_30'
  },

  /* ═══════════════════════════════════════════════════════════
     EPÍLEG – La Nova Mirada
  ═══════════════════════════════════════════════════════════ */

  scene_30: {
    id: 'scene_30',
    tipus: 'text_block',
    titol: 'Un any després al Garraf',
    personatge: 'narracio',
    narracio: 'Han passat dotze mesos des que Marc Vidal va arribar al Club Handbol Garraf. La diferència és visible a l\'escalfament: els jugadors es mouen amb una lleugeresa diferent, les mirades es creuen sense que ningú no cridi instruccions, i la pilota circula amb una fluïdesa que sorprèn els visitants.\n\nLaia ja no és "l\'extremera dreta". És una jugadora que pot actuar en qualsevol posició i que mira el camp de forma que fa uns mesos era impensable. Jordi, el veteran, ha adoptat el perifèric iteratiu com a propi —perquè el va descobrir ell sol— i ara és qui l\'explica als jugadors més joves.\n\nSofia somriu des de la graderia. Marc s\'acosta.',
    dialeg: {
      personatge: 'sofia',
      text: '"Saps quin és el millor senyal de que el teu programa funciona?"\n\nMarc mira la pista.\n\n"Que no el veig. El veig als jugadors, però no el veig a tu intervenir-hi a cada moment."\n\nMarc somriu. **Ha après a crear les condicions. Ara el cervell dels jugadors fa la resta.**'
    },
    contingut_pedagogic: {
      titol: 'Els tres principis que Marc porta ara a cada entrenament',
      text: '**1. El cervell aprèn millor quan selecciona per si mateix les conductes eficaces** — no instrueixis, dissenya entorns.\n**2. La percepció periférica s\'entrena; la neuroplasticitat és finita** — cada edat té el seu moment biològic.\n**3. El perifèric iteratiu no és un truc: és la màxima explotació del potencial humà** — i comença als 6 anys, en defensa, sense dir-li al jugador el que és.'
    },
    seguent: null
  }

};

/* ============================================================
   ENGINE
   ============================================================ */
const Engine = (function () {
  'use strict';

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
      var done = stage.scenes.every(function (s) { return state.visitedScenes.indexOf(s) !== -1; });
      var active = stage.scenes.indexOf(state.currentScene) !== -1;
      var cls = done ? 'map-done' : active ? 'map-active' : 'map-pending';
      var status = done ? '✓ Completat' : active ? '▶ En curs' : '○ Pendent';
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

    var ch = CHARACTERS[scene.personatge] || CHARACTERS.narracio;
    var r = ch.shape === 'square' ? '4px' : '50%';
    _qs('character-avatar').style.background = ch.color;
    _qs('character-avatar').style.borderRadius = r;
    _qs('character-avatar').textContent = ch.initials;
    _qs('character-name').textContent = ch.name;

    if (scene.imatge) {
      _qs('scene-image').src = scene.imatge;
      _qs('scene-image-container').style.display = 'block';
    }

    _qs('narrative-text').innerHTML = _md(scene.narracio || '');

    if (scene.dialeg) {
      var dc = CHARACTERS[scene.dialeg.personatge] || ch;
      _qs('dialogue-char-name').textContent = dc.name;
      _qs('dialogue-char-name').style.color = dc.color;
      _qs('dialogue-text').innerHTML = _md(scene.dialeg.text);
      _qs('dialogue-box').style.display = 'block';
    }

    if (scene.contingut_pedagogic) {
      var pb = scene.contingut_pedagogic;
      _qs('pedagogic-title').textContent = pb.titol || '';
      _qs('pedagogic-text').innerHTML = _md(pb.text || '');
      _qs('pedagogic-block').style.display = 'block';
    }

    var ia = _qs('interaction-area');
    if (scene.tipus === 'text_block') {
      if (scene.seguent) {
        var btn = document.createElement('button');
        btn.className = 'btn btn-primary btn-enabled';
        btn.textContent = 'Continuar →';
        btn.onclick = function () { Engine.showScene(scene.seguent); };
        ia.appendChild(btn);
      } else {
        var restartBtn = document.createElement('button');
        restartBtn.className = 'btn btn-primary btn-enabled';
        restartBtn.textContent = 'Reiniciar curs';
        restartBtn.onclick = function () {
          state = { currentScene: 'scene_01', score: 0, visitedScenes: [], decisions: [] };
          _save();
          Engine.showScene('scene_01');
        };
        ia.appendChild(restartBtn);
      }
    } else if (scene.tipus === 'quiz') {
      _renderQuiz(scene, ia);
    } else if (scene.tipus === 'decisio') {
      _renderDecisio(scene, ia);
    } else if (scene.tipus === 'checklist') {
      _renderChecklist(scene, ia);
    } else if (scene.tipus === 'epilog') {
      _renderEpilog(scene, ia);
    }
  }

  function _renderQuiz(scene, ia) {
    var label = document.createElement('div');
    label.className = 'decision-label';
    label.textContent = 'Pregunta de comprensió';
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
        if (op.correcta) {
          state.score += pts;
          state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
        } else {
          state.decisions.push({ scene: scene.id, pts: 0, label: scene.titol });
        }
        scene.opcions.forEach(function (o, i) {
          var b = ia.querySelectorAll('.btn-quiz')[i];
          if (o.correcta) b.className = 'btn btn-quiz btn-quiz-correct btn-disabled';
          else if (o.id === op.id) b.className = 'btn btn-quiz btn-quiz-wrong btn-disabled';
          else b.className = 'btn btn-quiz btn-disabled';
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
    label.textContent = 'Decisió';
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
        var pts = op.punts !== undefined ? op.punts : (op.correcta ? 10 : 0);
        state.score += pts;
        state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
        scene.opcions.forEach(function (o, i) {
          ia.querySelectorAll('.btn-option')[i].className = 'btn btn-option btn-disabled';
        });
        btn.className = 'btn btn-option btn-selected btn-disabled';
        _save();
        var icon = pts >= 10 ? 'fb-good' : pts >= 5 ? 'fb-ok' : 'fb-bad';
        var nextScene = (op.seguent) ? op.seguent : scene.seguent;
        _showFeedback(icon, pts, op.feedback, function () {
          Engine.showScene(nextScene);
        });
        _qs('score-display').textContent = state.score + ' pts';
      };
      ia.appendChild(btn);
    });
  }

  function _renderChecklist(scene, ia) {
    var label = document.createElement('div');
    label.className = 'decision-label';
    label.textContent = 'Llista de verificació';
    ia.appendChild(label);

    if (scene.pregunta) {
      var q = document.createElement('div');
      q.className = 'quiz-question';
      q.innerHTML = _md(scene.pregunta);
      ia.appendChild(q);
    }

    var checked = {};
    var submitted = false;
    var itemsDiv = document.createElement('div');
    itemsDiv.className = 'checklist-items';
    ia.appendChild(itemsDiv);

    scene.checklistItems.forEach(function (item, idx) {
      var div = document.createElement('div');
      div.className = 'checklist-item';
      div.innerHTML = '<span class="checkbox-icon">☐</span><span class="checkbox-text">' + _md(item.text) + '</span>';
      div.onclick = function () {
        if (submitted) return;
        checked[idx] = !checked[idx];
        div.classList.toggle('checked', !!checked[idx]);
        div.querySelector('.checkbox-icon').textContent = checked[idx] ? '☑' : '☐';
      };
      itemsDiv.appendChild(div);
    });

    var submitBtn = document.createElement('button');
    submitBtn.className = 'btn btn-primary btn-enabled';
    submitBtn.style.marginTop = '8px';
    submitBtn.textContent = 'Verificar →';
    ia.appendChild(submitBtn);

    submitBtn.onclick = function () {
      if (submitted) return;
      submitted = true;
      submitBtn.style.display = 'none';
      var pts = 0;
      scene.checklistItems.forEach(function (item, idx) {
        var userChecked = !!checked[idx];
        var correct = item.correcta;
        var div = itemsDiv.children[idx];
        div.style.pointerEvents = 'none';
        if (userChecked && correct) { div.classList.add('cl-correct'); pts += 2; }
        else if (!userChecked && !correct) { div.classList.add('cl-correct'); pts += 2; }
        else if (userChecked && !correct) { div.classList.add('cl-incorrect'); }
        else { div.classList.add('cl-missed'); }
      });
      state.score += pts;
      state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
      _save();
      _qs('score-display').textContent = state.score + ' pts';
      var fbDiv = document.createElement('div');
      fbDiv.className = 'checklist-feedback';
      fbDiv.innerHTML = '<strong>' + pts + '/10 pts</strong> — ' + _md(scene.feedback || 'Comprova els ítems ressaltats.');
      ia.appendChild(fbDiv);
      var contBtn = document.createElement('button');
      contBtn.className = 'btn btn-primary btn-enabled';
      contBtn.style.marginTop = '8px';
      contBtn.textContent = 'Continuar →';
      contBtn.onclick = function () { Engine.showScene(scene.seguent); };
      ia.appendChild(contBtn);
    };
  }

  function _renderEpilog(scene, ia) {
    var maxScore = 100;
    var circumference = 2 * Math.PI * 54;
    var offset = circumference * (1 - state.score / maxScore);
    var grade, gradeClass, mentorMsg;
    if (state.score >= 90) {
      grade = 'Excel·lent';
      gradeClass = 'epilogue-excellent';
      mentorMsg = scene.mentorMsgs ? scene.mentorMsgs.excellent : '';
    } else if (state.score >= 70) {
      grade = 'Molt bé';
      gradeClass = 'epilogue-good';
      mentorMsg = scene.mentorMsgs ? scene.mentorMsgs.good : '';
    } else if (state.score >= 50) {
      grade = 'Correcte';
      gradeClass = 'epilogue-ok';
      mentorMsg = scene.mentorMsgs ? scene.mentorMsgs.ok : '';
    } else {
      grade = 'Necessites repassar';
      gradeClass = 'epilogue-needs-work';
      mentorMsg = scene.mentorMsgs ? scene.mentorMsgs.needsWork : '';
    }

    var sofia = CHARACTERS.sofia;
    var html = '<div class="epilogue-container">';
    html += '<div class="epilogue-score-ring">';
    html += '<svg class="score-ring-svg" viewBox="0 0 120 120">';
    html += '<circle cx="60" cy="60" r="54" fill="none" stroke="var(--surface3)" stroke-width="8"/>';
    html += '<circle cx="60" cy="60" r="54" fill="none" stroke="var(--teal)" stroke-width="8"';
    html += ' stroke-dasharray="' + circumference.toFixed(2) + '" stroke-dashoffset="' + offset.toFixed(2) + '"';
    html += ' stroke-linecap="round" transform="rotate(-90 60 60)"/>';
    html += '</svg>';
    html += '<div class="score-ring-text"><span class="score-big">' + state.score + '</span><span class="score-max">/ 100</span></div>';
    html += '</div>';
    html += '<div class="epilogue-grade ' + gradeClass + '">' + grade + '</div>';
    html += '<div class="epilogue-mentor ' + gradeClass + '">';
    html += '<div class="epilogue-mentor-avatar" style="background:' + sofia.color + '">' + sofia.initials + '</div>';
    html += '<p class="epilogue-mentor-msg">' + _md(mentorMsg) + '</p>';
    html += '</div>';
    if (state.decisions.length > 0) {
      html += '<div><div class="epilogue-section-title">Resum de decisions</div>';
      html += '<ul class="decisions-summary">';
      state.decisions.forEach(function (d) {
        var cls = d.pts >= 10 ? 'ds-good' : d.pts >= 5 ? 'ds-ok' : 'ds-bad';
        var icon = d.pts >= 10 ? '✓' : d.pts >= 5 ? '~' : '✗';
        html += '<li class="' + cls + '"><span class="ds-icon">' + icon + '</span>';
        html += '<span>' + (d.label || d.scene) + '</span>';
        html += '<span class="ds-pts">' + d.pts + ' pts</span></li>';
      });
      html += '</ul></div>';
    }
    html += '<div style="margin-top:16px;">';
    html += '<button class="btn btn-primary btn-enabled" onclick="Engine.showScene(\'scene_30\')">';
    html += 'Continuar →</button></div>';
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
      if (!scene) { console.warn('Escena no trobada:', id); return; }
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

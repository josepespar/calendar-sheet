/* ============================================================
   SCENARIO – Control Motor: El Factor Clau Silenciós
   28 escenes · Quiz + Decisions + Checklists · Puntuació màxima: 100 pts
   Idioma: català
   Basat en l'article de Raquel Font-Lladó (UdG)
   VII Seminari Internacional de Tàctica i Tècnica Esportiva
   ============================================================ */

const CHARACTERS = {
  narracio: { name: 'Narrador',           color: '#6B7280', initials: '✦', shape: 'square' },
  raquel:   { name: 'Raquel Font-Lladó',  color: '#F59E0B', initials: 'R', shape: 'circle' },
  joan:     { name: 'Joan (investigador)', color: '#60A5FA', initials: 'J', shape: 'circle' },
  andres:   { name: 'Andrés (entrenador)',color: '#34D399', initials: 'A', shape: 'circle' },
  esther:   { name: 'Esther (atleta)',    color: '#F472B6', initials: 'E', shape: 'circle' },
  pau:      { name: 'Pau Martí',          color: '#8B5CF6', initials: 'P', shape: 'circle' }
};

const JOURNEY_STAGES = [
  {
    id: 'act1',
    label: 'Acte I',
    title: 'El Seminari',
    scenes: ['scene_01','scene_02','scene_03','scene_04','scene_05'],
    color: '#8B5CF6'
  },
  {
    id: 'act2',
    label: 'Acte II',
    title: 'Tres Visions',
    scenes: ['scene_06','scene_07','scene_08','scene_09','scene_10'],
    color: '#60A5FA'
  },
  {
    id: 'act3',
    label: 'Acte III',
    title: 'Biologia, Experiència i Context',
    scenes: ['scene_11','scene_12','scene_13','scene_14','scene_15'],
    color: '#34D399'
  },
  {
    id: 'act4',
    label: 'Acte IV',
    title: 'Aprenentatge i Consciència',
    scenes: ['scene_16','scene_17','scene_18','scene_19','scene_20'],
    color: '#F472B6'
  },
  {
    id: 'act5',
    label: 'Acte V',
    title: 'Mesurar i Planificar',
    scenes: ['scene_21','scene_22','scene_23','scene_24','scene_24b','scene_25'],
    color: '#F59E0B'
  },
  {
    id: 'final',
    label: 'Epíleg',
    title: 'Síntesi i Conclusió',
    scenes: ['scene_26','scene_27','scene_28'],
    color: '#8B5CF6'
  }
];

const CANONICAL_SCENE_COUNT = 28;

const scenes = {

  /* ═══════════════════════════════════════════════════════════
     ACTE I – El Seminari
  ═══════════════════════════════════════════════════════════ */

  scene_01: {
    id: 'scene_01',
    tipus: 'text_block',
    titol: 'Arribar al VII Seminari Internacional',
    personatge: 'pau',
    narracio: 'En Pau Martí, 28 anys, entrenador d\'atletisme júnior, arriba al VII Seminari Internacional de Tàctica i Tècnica Esportiva. Ha recorregut cent cinquanta quilòmetres per assistir a la taula rodona que tancarà la jornada.\n\nLa sala és plena. A la pissarra llegeix el títol de la sessió: «Control motor: ¿El factor clau silenciós del rendiment esportiu?» Tres persones ocupen l\'estrada: una investigadora moderadora, un entrenador de fons i ultrafons de muntanya, i una atleta d\'alt rendiment.\n\nEn Pau obre el seu bloc. Porta anys treballant la tècnica dels seus atletes per intuïció. Avui vol entendre el fonament.',
    dialeg: {
      personatge: 'pau',
      text: '"Per una vegada, vull entendre el perquè, no només el com." En Pau s\'asseu a la primera fila.'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02',
    tipus: 'text_block',
    titol: 'La Definició Provocadora',
    personatge: 'raquel',
    narracio: 'Raquel Font-Lladó, investigadora de la Universitat de Girona, puja a l\'estrada i, en lloc d\'una presentació convencional, llança una definició directa i deliberadament àmplia per generar debat.',
    dialeg: {
      personatge: 'raquel',
      text: '"El control motor és la capacitat del ser humà per produir moviment i mantenir la postura. Es presenta com a potencial intrínsec a l\'individu. Ha de ser desenvolupat a través de l\'experiència, donant resposta a un objectiu motor, integrant la informació rebuda de l\'entorn i del propi cos. Aquesta relació es retroalimenta i s\'interrelaciona per configurar el comportament motor."\n\nRaquel fa una pausa i mira el públic.\n\n"Esteu d\'acord?"'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03',
    tipus: 'text_block',
    titol: 'Tres Grans Perspectives Teòriques',
    personatge: 'raquel',
    narracio: 'Per situar el debat, Raquel presenta les tres grans corrents que la psicologia ha utilitzat per explicar el control motor. Cap d\'elles, assenyala, és la veritat absoluta.',
    contingut_pedagogic: {
      titol: 'Tres grans perspectives sobre el control motor',
      text: '**1. Conductistes i associacionistes** — El moviment és una reacció apresa per estímul-resposta. L\'entrenament repetit consolida les respostes correctes (Lawther, 1968; Rushall i Siedentop, 1972).\n**2. Cognitivistes** — El moviment prové d\'un programa motor emmagatzemat en la memòria, influenciat pels feedbacks. L\'individu construeix representacions internes (Adams, 1971; Schmidt, 1976; Meinel i Schnabel, 1988).\n**3. Sistemes i ecològiques** — El moviment emergeix de la interacció holística entre organisme i entorn. No hi ha director central: el sistema s\'autoorganitza (Bernstein, 1967; Kelso i Tuller, 1984; Thelen, 1987; Gibson).'
    },
    seguent: 'scene_04'
  },

  scene_04: {
    id: 'scene_04',
    tipus: 'quiz',
    titol: 'Comprensió: Definició de Control Motor',
    personatge: 'raquel',
    punts: 10,
    narracio: 'Raquel mira el públic i planteja la primera pregunta de reflexió per verificar la comprensió de la definició.',
    pregunta: 'Quin aspecte és fonamental en la definició de control motor que ha presentat Raquel?',
    opcions: [
      {
        id: 'A',
        text: 'La integració de la informació del cos i l\'entorn per generar moviment orientat a un objectiu',
        correcta: true,
        feedback: 'Exacte. La definició de Raquel posa el focus en la **integració**: informació del cos (propioceptiva) + informació de l\'entorn (perceptiva) + retroalimentació constant, tot al servei d\'un objectiu motor concret. No es tracta de força ni de memòria tècnica aïllada, sinó d\'un sistema circular d\'informació i resposta.'
      },
      {
        id: 'B',
        text: 'La força muscular i la resistència cardiovascular de l\'esportista',
        correcta: false,
        feedback: 'La força i la resistència són condicions físiques necessàries, però la definició de control motor apunta a quelcom diferent: la **capacitat per integrar informació i generar respostes motrius adaptades al context**. Un esportista pot ser molt fort i tenir un control motor deficient si no percep ni integra bé la informació del seu entorn.'
      },
      {
        id: 'C',
        text: 'La memorització de patrons tècnics ideals per reproduir-los en competició',
        correcta: false,
        feedback: 'La memorització de patrons és pròpia de la perspectiva **cognitivista** (programa motor emmagatzemat), però la definició de Raquel és més àmplia: inclou la interacció dinàmica amb l\'entorn i la retroalimentació constant. El control motor no és reproduir un patró fix, sinó adaptar-se contínuament.'
      }
    ],
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05',
    tipus: 'text_block',
    titol: 'La Taula Rodona Comença',
    personatge: 'narracio',
    narracio: 'Raquel presenta els tres ponents que compartiran la taula:\n\n— **Joan**, investigador especialitzat en la perspectiva ecològica del control motor.\n— **Andrés**, entrenador de fons i ultrafons de muntanya amb esportistes d\'alt rendiment.\n— **Esther**, atleta d\'alt rendiment en atletisme de resistència.\n\nTres visions del mateix fenomen: acadèmica, d\'entrenament i pràctica esportiva.',
    dialeg: {
      personatge: 'raquel',
      text: '"I sense més preàmbuls, la primera pregunta que vull que abordem és: **¿Qué es para vosotros el control motor?**"\n\nRaquel mira Joan, convidant-lo a intervenir primer.'
    },
    seguent: 'scene_06'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE II – Tres Visions
  ═══════════════════════════════════════════════════════════ */

  scene_06: {
    id: 'scene_06',
    tipus: 'text_block',
    titol: 'La Perspectiva Ecològica',
    personatge: 'joan',
    narracio: 'Joan pren la paraula des d\'una postura acadèmica. Comença amb una definició però de seguida la matisa des de la seva perspectiva ecològica.',
    dialeg: {
      personatge: 'joan',
      text: '"Per a mi, el control motor és la capacitat de percebre el moviment en, i per a, la generació de respostes motrius."\n\nFa una pausa i afegeix:\n\n"Però, concretament, des de l\'aproximació ecològica, el control motor no existeix inherent a l\'individu, ja que necessita del context. L\'individu i l\'entorn formen un sistema indissociable. El moviment no és la sortida d\'un programa intern: emergeix de la interacció."'
    },
    contingut_pedagogic: {
      titol: 'La perspectiva ecològica (Gibson, Kelso, Thelen)',
      text: 'Des de les teories ecològiques i de sistemes, el moviment **emergeix de la interacció dinàmica** entre l\'organisme i el context. L\'entorn proporciona informació (affordances) que el sistema motor aprofita per generar respostes adaptades.\n\nEl sistema nerviós central deixa de ser un «director» aïllat per passar a formar part d\'un **procés circular**: percepció → resposta → percepció. Bernstein (1967) en fou el primer representant; Thelen, Kelso i Gibson el van consolidar.'
    },
    seguent: 'scene_07'
  },

  scene_07: {
    id: 'scene_07',
    tipus: 'text_block',
    titol: 'La Perspectiva de l\'Entrenador de Muntanya',
    personatge: 'andres',
    narracio: 'Andrés reflexiona sobre el que significa el control motor des de la seva experiència pràctica com a entrenador de corredors i esquiadors de muntanya d\'alt rendiment.',
    dialeg: {
      personatge: 'andres',
      text: '"En els esports de fons i ultrafons de muntanya, el control motor és la percepció i l\'adaptació a allò que estàs percebent. En certa mesura, és la reproducció d\'un mateix en l\'espai i el temps."\n\nReflexiona un instant i continua:\n\n"L\'esportista no controla el terreny: s\'adapta a ell. Cada pas és una decisió motriu nova, dictada per allò que el terreny li ofereix en aquell precís moment."'
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08',
    tipus: 'text_block',
    titol: 'La Perspectiva de l\'Atleta',
    personatge: 'esther',
    narracio: 'Esther introdueix la seva intervenció des de l\'experiència viscuda. La sala escolta amb atenció perquè la seva perspectiva és la més propera a la realitat del dia a dia.',
    dialeg: {
      personatge: 'esther',
      text: '"No m\'agrada treballar-ho. Jo i el meu equip ho associem a modificar la tècnica des d\'un concepte molt ampli: utilització de la força, transferència, gest...\n\nRequereix serietat, concentració, posar el focus en els petits detalls. És complex; exigeix treballar amb tots els sentits posats en el moviment."'
    },
    contingut_pedagogic: {
      titol: 'Tres conceptes clau que emergeixen',
      text: 'De les tres respostes emergeixen tres conceptes que recorreran tota la taula:\n**Percepció** — Joan emfatitza la percepció com a base de la resposta motriu. Sense percebre l\'entorn, no hi ha moviment adaptat.\n**Adaptació** — Andrés posa l\'accent en l\'adaptació permanent a un entorn variable i imprevisible.\n**Focalització tècnica** — Esther destaca la consciència i la concentració en els detalls del propi moviment.'
    },
    seguent: 'scene_09'
  },

  scene_09: {
    id: 'scene_09',
    tipus: 'quiz',
    titol: 'Comprensió: La Perspectiva Ecològica',
    personatge: 'joan',
    punts: 10,
    narracio: 'En Pau subratlla la intervenció de Joan. Intenta reformular-la per assegurar-se que l\'ha entesa bé.',
    pregunta: 'Joan defensa que, des de la perspectiva ecològica, el control motor...',
    opcions: [
      {
        id: 'A',
        text: 'No existeix inherent a l\'individu, perquè necessita del context per existir i es genera en la interacció',
        correcta: true,
        feedback: 'Correcte. Aquesta és la idea central de la perspectiva ecològica: el control motor no és una capacitat que "tenim" independentment de l\'entorn. **Emergeix de la interacció** entre l\'organisme i el context. Per això, entrenant en entorns artificials o estàtics, no es desenvolupa el control motor real.'
      },
      {
        id: 'B',
        text: 'Consisteix en reproduir patrons motors preexistents emmagatzemats en la memòria',
        correcta: false,
        feedback: 'Aquesta és la perspectiva **cognitivista** (Schmidt, Adams): el moviment prové d\'un programa motor intern. La perspectiva ecològica, al contrari, defensa que el moviment **emergeix** de la interacció organisme-entorn en cada moment, sense necessitat d\'un programa previ.'
      },
      {
        id: 'C',
        text: 'És la percepció i l\'adaptació constant al terreny, com explica Andrés en el seu exemple',
        correcta: false,
        feedback: 'La descripció d\'Andrés és molt propera a l\'ecologia, però reflecteix principalment la seva **experiència pràctica** com a entrenador de muntanya. La definició acadèmica de Joan és més radical: el control motor no existeix inherent a l\'individu; necessita el context per existir. No és adaptació al terreny, és emergència en la interacció.'
      }
    ],
    seguent: 'scene_10'
  },

  scene_10: {
    id: 'scene_10',
    tipus: 'quiz',
    titol: 'Comprensió: La Visió d\'Andrés',
    personatge: 'andres',
    punts: 10,
    narracio: 'En Pau pensa en els seus atletes de resistència. La visió d\'Andrés li resulta molt familiar.',
    pregunta: 'Andrés descriu el control motor en l\'esquí i el fons de muntanya com...',
    opcions: [
      {
        id: 'A',
        text: 'La percepció i l\'adaptació a allò que es percep: reproduir-se en l\'espai i el temps',
        correcta: true,
        feedback: 'Exacte. Andrés situa el control motor com una **relació dinàmica** entre el que l\'esportista percep i la seva resposta d\'adaptació. En els esports de muntanya, l\'entorn canvia constantment: el control motor és la capacitat de "ser-hi" a cada moment, adaptat al que ofereix el terreny.'
      },
      {
        id: 'B',
        text: 'La capacitat de repetir el mateix patró tècnic perfecte independentment del terreny',
        correcta: false,
        feedback: 'Aquesta seria la visió conductista o cognitivista: reproduir un patró fix. Andrés defensa justament el contrari: el control motor en la muntanya és **adaptació constant**, no reproducció d\'un patró immutable. El terreny és sempre nou; la resposta motor ha de ser-ho també.'
      },
      {
        id: 'C',
        text: 'La força i la resistència necessàries per superar les proves de llarga distància en muntanya',
        correcta: false,
        feedback: 'La força i la resistència són condicions físiques importants, però Andrés parla d\'una dimensió diferent: la **percepció i adaptació motriu** a l\'entorn. Un esportista pot ser molt resistent i tenir dificultats de control motor si no percep i s\'adapta bé a les exigències canviants del terreny.'
      }
    ],
    seguent: 'scene_11'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE III – Biologia, Experiència i Context
  ═══════════════════════════════════════════════════════════ */

  scene_11: {
    id: 'scene_11',
    tipus: 'text_block',
    titol: 'La Biologia com a Precondició',
    personatge: 'joan',
    narracio: 'Raquel planteja la segona gran pregunta de la taula rodona. Joan pren la paraula sense esperar que li ho demanin.',
    dialeg: {
      personatge: 'raquel',
      text: '"En el desenvolupament del control motor, **quin paper juga la biologia, l\'experiència individual, i la complexitat del context?**"'
    },
    seguent: 'scene_12'
  },

  scene_12: {
    id: 'scene_12',
    tipus: 'text_block',
    titol: 'L\'Exploració com a Clau de l\'Aprenentatge',
    personatge: 'joan',
    narracio: 'Joan és clar i directe en la seva postura. La resposta, per a ell, és inequívoca.',
    dialeg: {
      personatge: 'joan',
      text: '"Des de la meva perspectiva, la biologia és **només una precondició** que exerceix de facilitador per al desenvolupament motor. Són les experiències prèvies del subjecte les determinants clau del desenvolupament, sobretot si han generat aprenentatge.\n\nEn definitiva, l\'aprenentatge del control motor es concreta en l\'exploració i l\'assoliment de noves formes d\'interrelacionar-se amb l\'espai, el temps i els altres individus.\n\nPer tant, com a entrenadors hem de dissenyar contextos que afavoreixin dita exploració."'
    },
    contingut_pedagogic: {
      titol: 'Tres perspectives: el pes de la biologia',
      text: 'Les grans teories discrepan sobre quin factor és determinant:\n**Associacionistes (Lawther, Gesell)**: Creixement + Maduració + Aprenentatge = Desenvolupament. La biologia marca les etapes evolutives del control motor.\n**Cognitivistes (Schmidt, Adams, Piaget)**: El subjecte és agent actiu; construeix representacions a partir de la interacció entre allò biològic i allò contextual.\n**Sistemes ecològics (Bernstein, Gibson, Thelen)**: Biologia, entorn i experiència co-determinen el moviment de manera circular, sense jerarquia fixa entre ells.'
    },
    seguent: 'scene_13'
  },

  scene_13: {
    id: 'scene_13',
    tipus: 'quiz',
    titol: 'Comprensió: El Rol de la Biologia',
    personatge: 'joan',
    punts: 10,
    narracio: 'En Pau para d\'escriure i reflexiona sobre la postura de Joan. Té clar el que ha dit, però vol verificar que ho ha comprès correctament.',
    pregunta: 'Quin paper atribueix Joan a la biologia en el desenvolupament del control motor?',
    opcions: [
      {
        id: 'A',
        text: 'És una precondició necessària, però les experiències prèvies i l\'aprenentatge són els factors realment determinants',
        correcta: true,
        feedback: 'Perfecte. Joan és explícit: la biologia és una **precondició** (és a dir, posa els límits inicials), però no és el determinant del desenvolupament motor. Allò que realment defineix el nivell de control motor és l\'experiència viscuda i, sobretot, si aquesta experiència ha generat **aprenentatge real**: noves formes d\'interactuar amb l\'espai, el temps i els altres.'
      },
      {
        id: 'B',
        text: 'Determina completament el nivell de control motor que pot assolir un esportista',
        correcta: false,
        feedback: 'Aquesta seria la postura madurativa o determinista biològica (Gesell, 1929). Joan pren distància d\'ella: la biologia és una **precondició**, no un determinant final. Si fos determinant, l\'entrenament i l\'experiència no tindrien sentit, i sabem que sí que el tenen.'
      },
      {
        id: 'C',
        text: 'La biologia i l\'experiència tenen exactament el mateix pes en el desenvolupament motor',
        correcta: false,
        feedback: 'Joan no estableix un equilibri igualitari. Afirma que la biologia és **subordinada** a l\'experiència: és una precondició (base necessària), però és l\'experiència i l\'aprenentatge el que veritablement configura el control motor. El pes no és igual: l\'experiència guanya.'
      }
    ],
    seguent: 'scene_14'
  },

  scene_14: {
    id: 'scene_14',
    tipus: 'text_block',
    titol: 'Competició: Quan el Context Ho Canvia Tot',
    personatge: 'esther',
    narracio: 'Esther s\'afegeix a les aportacions de Joan i Andrés amb una reflexió fonamentada en la seva experiència d\'esportista d\'alt rendiment en atletisme.',
    dialeg: {
      personatge: 'esther',
      text: '"La situació de competició presenta alts nivells d\'estrès, magnifica els estímuls externs, i en algun moment la fatiga apareix en la seva màxima expressió. Tot això **modifica el control motor**.\n\nPer tant, entrenar la tècnica sense el context no té sentit. Encara que a vegades treballem de manera descontextualitzada per focalitzar millor l\'atenció, després ens assegurem que hi hagi **transferència**."'
    },
    seguent: 'scene_15'
  },

  scene_15: {
    id: 'scene_15',
    tipus: 'checklist',
    titol: 'Biologia, Experiència, Context: Conceptes Clau',
    personatge: 'raquel',
    narracio: 'Raquel sintetitza les intervencions dels tres ponents. En Pau ha d\'identificar els 5 aspectes que s\'han plantejat com a veritables.',
    pregunta: 'Marca els 5 aspectes veritaders sobre el rol de la biologia, l\'experiència i el context en el control motor:',
    checklistItems: [
      { text: 'La biologia és una precondició necessària, però l\'experiència és el factor determinant clau',           correcta: true  },
      { text: 'L\'aprenentatge motor és explorar noves formes d\'interacció amb l\'espai, el temps i els altres',       correcta: true  },
      { text: 'Cal dissenyar contextos que afavoreixin l\'exploració motriu de l\'esportista',                          correcta: true  },
      { text: 'Entrenar en contextos similars als de competició és imprescindible per garantir la transferència',       correcta: true  },
      { text: 'L\'esportista i l\'entrenador han de treballar conjuntament per millorar el patró motor',               correcta: true  },
      { text: 'La biologia determina definitivament les capacitats motrius al llarg de tota la vida',                  correcta: false },
      { text: 'Reproduir exactament el model tècnic ideal garanteix sempre el màxim rendiment',                        correcta: false }
    ],
    feedback: 'Les dues afirmacions incorrectes representen errors freqüents: la biologia és una precondició, no un determinant permanent; i el "model tècnic ideal" és una referència, no una fórmula universal. La tècnica ha d\'adaptar-se a la biologia de l\'individu i al context de competició.',
    seguent: 'scene_16'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE IV – Aprenentatge i Consciència
  ═══════════════════════════════════════════════════════════ */

  scene_16: {
    id: 'scene_16',
    tipus: 'text_block',
    titol: 'Què Aprenem Quan Treballem el Control Motor?',
    personatge: 'esther',
    narracio: 'Raquel llança la tercera pregunta de la taula rodona. Esther mira el sostre un moment, com si es transportés als seus entrenaments, i respon des de la vivència directa.',
    dialeg: {
      personatge: 'raquel',
      text: '"¿Qué aprendemos cuando trabajamos sobre el control motor? Imatges, sensacions, idees, processos?"'
    },
    seguent: 'scene_17'
  },

  scene_17: {
    id: 'scene_17',
    tipus: 'quiz',
    titol: 'Comprensió: Aprendre sobre el Propi Cos',
    personatge: 'esther',
    punts: 10,
    narracio: 'Esther respon des de la seva pròpia experiència d\'atleta. La sala escolta amb atenció.',
    pregunta: 'L\'Esther assenyala que el treball de control motor li permet, a banda de millorar el rendiment, una cosa que sorprèn la taula. Quina és?',
    opcions: [
      {
        id: 'A',
        text: 'Conèixer millor el propi cos, sentir-se més segura i, indirectament, prevenir lesions corregint patrons lesius',
        correcta: true,
        feedback: 'Exacte. Esther obri una dimensió nova: el control motor **no s\'associa únicament al rendiment directe**. Quan el treball de control motor permet corregir patrons tècnics lesius, contribueix indirectament a la **prevenció de lesions**. Joan aprofita per dir: "Això confirma que el control motor és percepció." Si sentir-se segura és el resultat, vol dir que la percepció del propi moviment ha millorat.'
      },
      {
        id: 'B',
        text: 'Augmentar la força màxima i la potència de sortida en els entrenaments d\'intensitat alta',
        correcta: false,
        feedback: 'La força és una capacitat física independent del control motor, tot i que es relacionen. El que Esther destaca és quelcom de diferent: **l\'autoconeixement corporal**, la confiança en el propi moviment i la capacitat de córrer sense patrons lesius. No és força: és percepció i adaptació.'
      },
      {
        id: 'C',
        text: 'Memoritzar seqüències de moviment que es poden reproduir automàticament en competició',
        correcta: false,
        feedback: 'La memoriztació de seqüències és un mecanisme **cognitivista** (programa motor). Esther, al contrari, parla de **consciència del cos**, de sensació, de seguretat. Aprèn a sentir el seu cos, no a executar una seqüència memoritzada. Joan mateix qüestiona si el que s\'aprèn de manera conscient es transfereix en competició.'
      }
    ],
    seguent: 'scene_18'
  },

  scene_18: {
    id: 'scene_18',
    tipus: 'text_block',
    titol: 'El Repte de Joan: ¿S\'Aprèn Realment?',
    personatge: 'joan',
    narracio: 'Joan intervé amb una pregunta que genera un gran silenci a la sala. En Pau deixa de prendre notes.',
    dialeg: {
      personatge: 'joan',
      text: '"¿Qué aprens? —llança la pregunta a l\'aire—. **¿Realment s\'aprèn el control motor?** La ciència discrepa; té dubtes.\n\nPosar consciència sobre el moviment pot generar una **falsa sensació d\'haver après**. Però és realment així? Faig aquesta pregunta perquè en competició, quan l\'esportista perd el focus sobre el moviment, desapareix el que havia après. O potser, en realitat, no s\'havia après?"'
    },
    contingut_pedagogic: {
      titol: 'Focus intern vs focus extern',
      text: 'El debat sobre la consciència en el control motor planteja preguntes fonamentals:\n**Focus intern** — L\'esportista posa l\'atenció en les sensacions del propi cos. Pot generar autoconeixement, però en competició pot interferir amb la fluïdesa del moviment.\n**Focus extern** — L\'esportista focalitza en l\'efecte del moviment sobre l\'entorn. Molts estudis mostren millors resultats en rendiment real.\n**La paradoxa de la consciència**: allò après amb focus conscient no sempre es transfereix al comportament inconscient de la competició.'
    },
    seguent: 'scene_19'
  },

  scene_19: {
    id: 'scene_19',
    tipus: 'quiz',
    titol: 'El Repte de Joan sobre la Consciència',
    personatge: 'joan',
    punts: 10,
    narracio: 'Esther assenteix. I Andrés reforça la idea de Joan des de la seva experiència amb esportistes en situació de competició.',
    pregunta: 'Joan qüestiona si el control motor s\'aprèn realment. El seu argument central és...',
    opcions: [
      {
        id: 'A',
        text: 'Que en competició, quan l\'esportista perd el focus conscient sobre el moviment, allò après pot desaparèixer, suggerint que potser no s\'havia après realment',
        correcta: true,
        feedback: 'Exacte. Joan identifica una **paradoxa fonamental**: si el canvi depèn del focus conscient per mantenir-se, potser no era un aprenentatge real sinó un canvi temporal. L\'aprenentatge motor profund es manifesta en competició **sense necessitat de focus conscient**. Quan el focus desapareix i el canvi també, cal preguntar-se si hi havia aprenentatge o simplement adaptació superficial.'
      },
      {
        id: 'B',
        text: 'Que l\'esportista no pot mesurar objectivament el seu progrés en control motor',
        correcta: false,
        feedback: 'La mesura és un tema diferent (que tractaran en la propera pregunta). El dubte de Joan és sobre la **naturalesa de l\'aprenentatge**: si allò après és real o si és una il·lusió d\'aprenentatge generada per la consciència temporal del moviment. No és un problema de mesura, és un problema de transferència.'
      },
      {
        id: 'C',
        text: 'Que la biologia limita la capacitat d\'aprendre nous patrons motors en esportistes adults',
        correcta: false,
        feedback: 'Joan ja havia deixat clar que la biologia és una precondició, no un límit definitiu. El seu dubte és sobre la **consciència i la transferència**: allò que s\'aprèn de manera conscient (focus intern), ¿es transfereix al comportament inconscient de la competició? Aquesta és la pregunta que la ciència, diu Joan, no ha respost definitivament.'
      }
    ],
    seguent: 'scene_20'
  },

  scene_20: {
    id: 'scene_20',
    tipus: 'text_block',
    titol: 'Testicles, Feedback i Sensacions',
    personatge: 'andres',
    narracio: 'Andrés reflexiona sobre el que ha dit Joan i hi afegeix la seva pràctica com a entrenador. El debat sobre la consciència el porta a parlar del paper del feedback i dels tests.',
    dialeg: {
      personatge: 'andres',
      text: '"En els entrenaments, testifico molt als esportistes. Sento que els **dóna confiança** sobre les seves possibilitats, i també sobre la meva feina.\n\nTambé pregunto molt als esportistes sobre les sensacions que senten; no tot ho podem observar. A vegades necessito que em diguin, per exemple, \'si estan tirant del gluti o del isquiotibial\'.\n\nEn competició, demano als esportistes que es focalitzin en la tècnica per dos objectius: distreure\'ls del cansament mental, i fer que en situacions molt concretes combinin bé les cames. Sense focus, de manera inconscient, sempre s\'agafa la pujada amb la mateixa cama, la més forta, i això afecta l\'eficiència."'
    },
    seguent: 'scene_21'
  },

  /* ═══════════════════════════════════════════════════════════
     ACTE V – Mesurar i Planificar
  ═══════════════════════════════════════════════════════════ */

  scene_21: {
    id: 'scene_21',
    tipus: 'text_block',
    titol: 'Mesurar el Control Motor',
    personatge: 'esther',
    narracio: 'Raquel planteja la quarta pregunta. Esther obre els ulls de bat a bat. Té una resposta molt clara.',
    dialeg: {
      personatge: 'raquel',
      text: '"¿Es pot mesurar el control motor? ¿Quines evidències utilitzem?"'
    },
    seguent: 'scene_22'
  },

  scene_22: {
    id: 'scene_22',
    tipus: 'quiz',
    titol: 'Comprensió: Mesurar el Control Motor',
    personatge: 'esther',
    punts: 10,
    narracio: 'Esther respon amb convicció: "Clar que el podem mesurar! Cal mesurar-lo!" I explica que utilitzen tests, gravacions de carreres i entrenaments per analitzar-los posteriorment. Però afegeix un matís important que en Pau subratlla immediatament.',
    pregunta: 'Esther assenyala que els **indicadors reals** d\'un canvi de patró motor s\'analitzen quan...',
    opcions: [
      {
        id: 'A',
        text: 'Apareixen de manera repetida al llarg del temps, ja que els canvis efímers (que apareixen i desapareixen) no confirmen l\'aprenentatge',
        correcta: true,
        feedback: 'Exacte. Esther estableix una distinció fonamental: hi ha **canvis efímers** (que apareixen i desapareixen) i **canvis reals** (que es consoliden). Els primers ajuden a entendre els factors que influencien la tècnica (fatiga, estrès, motivació), però no indiquen aprenentatge. Només la **repetició al llarg del temps** confirma que hi ha hagut un canvi real de patró.'
      },
      {
        id: 'B',
        text: 'L\'esportista expressa subjectivament que se sent millor en una sola sessió de tests',
        correcta: false,
        feedback: 'La sensació subjectiva és valuosa (Andrés ho reforçarà), però no és suficient per confirmar un canvi de patró. Esther és clara: a vegades "hi ha dies millors i pitjors". Una sola sessió amb bona sensació no és evidència de canvi real. Cal **repetició temporal** del canvi per confirmar l\'aprenentatge.'
      },
      {
        id: 'C',
        text: 'S\'observa una millora en el cronòmetre en un únic entrenament específic de control motor',
        correcta: false,
        feedback: 'El cronòmetre és una mesura de rendiment global, no necessàriament de control motor específic. A més, Esther deixa clar que l\'evolució "no és lineal": hi ha dies millors i pitjors. Una única millora puntual no confirma un canvi de patró. Cal **analitzar quan els canvis apareixen repetidament**.'
      }
    ],
    seguent: 'scene_23'
  },

  scene_23: {
    id: 'scene_23',
    tipus: 'text_block',
    titol: 'Com Planificar: La Darrera Pregunta',
    personatge: 'raquel',
    narracio: 'Raquel llança la cinquena i darrera pregunta de la taula rodona. Per a en Pau, entrenador en actiu, és la més important de totes.',
    dialeg: {
      personatge: 'raquel',
      text: '"I ara arriba la pregunta que molts de vosaltres porteu als llavis des del principi: **¿Quins serien els aspectes clau per planificar entrenaments centrats en la millora del control motor?**"\n\nEls tres ponents es miren. Ceden la paraula a Andrés.'
    },
    contingut_pedagogic: {
      titol: 'El procés de planificació d\'Andrés: quatre passos',
      text: 'Andrés estructura la seva resposta de manera clara:\n**1. Conèixer l\'esportista** — Tests i preguntes: experiència motriu, morfologia musculo-esquelètica, eficiència metabòlica, lesions prèvies, objectius.\n**2. Planificar el procés** — Elaborar el pla de millora i compartir-lo amb l\'esportista per analitzar-lo i modificar-lo si cal.\n**3. Exposar a contexts** — Aplicar la càrrega i els exercicis per a la millora del patró, en contexts simulats i reals.\n**4. Donar feedbacks significatius** — Feedbacks individualitzats i ajustats al moment de l\'esportista.'
    },
    seguent: 'scene_24'
  },

  scene_24: {
    id: 'scene_24',
    tipus: 'decisio',
    titol: 'La Decisió: Per On Comencem?',
    personatge: 'pau',
    punts: 10,
    narracio: 'En Pau torna a casa en cotxe pensant en un dels seus atletes, un corredor de fons de 24 anys que té un patró de carrera ineficient i lesiu. Ha de fer quelcom. Andrés l\'ha inspirat.',
    dialeg: {
      personatge: 'andres',
      text: '"Davant el full en blanc, jo sempre comença per conèixer el meu esportista: saber des d\'on parteixo."'
    },
    pregunta: 'L\'objectiu de Pau és millorar el control motor del seu corredor. Quin ha de ser el primer pas?',
    opcions: [
      {
        id: 'A',
        punts: 10,
        text: 'Realitzar tests i entrevistes per conèixer l\'esportista: experiència motriu, morfologia, historial de lesions i objectius',
        feedback: 'Perfecte. Andrés ho va explicar amb claredat: "davant el full en blanc, comença per conèixer l\'esportista." Sense conèixer el punt de partida —el patró actual, el marge d\'adaptació, les lesions prèvies— qualsevol pla és una imposició cega. El coneixement de l\'esportista és el fonament de tot el procés.',
        seguent: 'scene_25'
      },
      {
        id: 'B',
        punts: 0,
        text: 'Aplicar directament un programa d\'exercicis de millora del patró motor seleccionats de la literatura científica',
        feedback: 'Andrés és taxatiu: "m\'equivocaria si pretengués generar un patró completament nou sense partir del patró actual." Aplicar exercicis sense conèixer l\'esportista és una imposició que pot generar interferències físiques i psicològiques. Primer cal saber des d\'on es parteix.',
        seguent: 'scene_24b'
      },
      {
        id: 'C',
        punts: 0,
        text: 'Dissenyar un programa basat en el model tècnic ideal del corredor de fons',
        feedback: 'El model tècnic ideal és una referència, no un programa. Andrés adverteix que "modificar un patró necessita moltes hores de dedicació" i que "en moltes situacions de competició el patró original s\'imposa sobre l\'après." Sense conèixer el patró actual de l\'esportista, el model ideal és una meta sense mapa.',
        seguent: 'scene_24b'
      }
    ],
    seguent: 'scene_25'
  },

  scene_24b: {
    id: 'scene_24b',
    tipus: 'text_block',
    titol: 'La Correcció d\'Andrés',
    personatge: 'andres',
    narracio: 'Andrés somriu comprensivament quan en Pau li explica la seva idea al descans. Li ha passat a ell al principi de la seva carrera.',
    dialeg: {
      personatge: 'andres',
      text: '"Al principi jo també pensava que amb bons exercicis seria suficient. Amb el temps vaig aprendre que el motor del canvi no és l\'exercici, sinó **conèixer l\'esportista**.\n\nIdentifica el seu patró actual, el seu marge d\'adaptació, i llavors dissenya. Altrament, qualsevol proposta és una imposició cega que pot generar resistència o, pitjor, una lesió."'
    },
    seguent: 'scene_25'
  },

  scene_25: {
    id: 'scene_25',
    tipus: 'text_block',
    titol: 'El Procés Complet de Planificació',
    personatge: 'andres',
    narracio: 'Andrés detalla, pas a pas, el procés que segueix quan planifica la millora del control motor d\'un esportista.',
    dialeg: {
      personatge: 'andres',
      text: '"Un cop coneixes l\'esportista, **planifiques el procés** de millora considerant tot allò que has recollit: experiència motriu, morfologia, metabolisme, lesions, objectius, circumstàncies personals.\n\nLlavors exposes el teu punt de vista i el teu pla a l\'esportista per analitzar-lo i modificar-lo si cal. **Qui mana és l\'esportista.**\n\nEl darrer pas és l\'aplicació de la càrrega. Amb l\'experiència, he après que si les propostes generen interferències físiques o psicològiques, és important ser capaç de **retrocedir**. En moltes ocasions, aquesta paradoxa és la clau de l\'èxit: saber quan aturar-se és tan important com saber quan avançar."'
    },
    seguent: 'scene_26'
  },

  /* ═══════════════════════════════════════════════════════════
     EPÍLEG – Síntesi i Conclusió
  ═══════════════════════════════════════════════════════════ */

  scene_26: {
    id: 'scene_26',
    tipus: 'checklist',
    titol: 'Claus per Planificar el Control Motor',
    personatge: 'raquel',
    narracio: 'Raquel tanca la taula rodona. En Pau, inspirat, anota els principis que s\'emportarà. Ha d\'identificar els 5 aspectes clau per planificar el treball de control motor en l\'alt rendiment.',
    pregunta: 'Marca els 5 principis veritaders per planificar el treball de control motor en l\'alt rendiment:',
    checklistItems: [
      { text: 'Conèixer l\'esportista: experiència motriu, morfologia, historial de lesions i objectius',             correcta: true  },
      { text: 'Planificar el procés i compartir-lo amb l\'esportista per analitzar-lo i modificar-lo conjuntament',   correcta: true  },
      { text: 'Exposar l\'esportista a contexts simulats i reals de competició',                                       correcta: true  },
      { text: 'Proporcionar feedbacks significatius, individualitzats i ajustats al moment de l\'esportista',          correcta: true  },
      { text: 'Saber retrocedir en l\'avanç si les propostes generen interferències físiques o psicològiques',         correcta: true  },
      { text: 'Aplicar un programa estàndard basat en el model tècnic ideal de l\'esport per a tots els esportistes', correcta: false },
      { text: 'Minimitzar el diàleg amb l\'esportista per mantenir l\'objectivitat de l\'entrenador',                 correcta: false }
    ],
    feedback: 'Els dos ítems incorrectes representen errors habituals: el programa estàndard no s\'adapta a l\'individu (cal personalitzar); i el diàleg constant entre entrenador i esportista és fonamental, no un biaix. Andrés és clar: "quí mana és l\'esportista."',
    seguent: 'scene_27'
  },

  scene_27: {
    id: 'scene_27',
    tipus: 'epilog',
    titol: 'Síntesi Final – Resultats',
    personatge: 'raquel',
    narracio: 'La taula rodona s\'ha tancat. Raquel fa un últim resum i en Pau rep la seva avaluació.',
    mentorMsgs: {
      excellent: '**Excel·lent, Pau.** Has assimilat els principis fonamentals del control motor amb una comprensió profunda. Has entès la distinció entre perspectives teòriques, el rol de la biologia com a precondició, la paradoxa de la consciència en competició, i la importància de conèixer l\'esportista abans de planificar. El teu corredor de fons tindrà un entrenador que pensa, dialoga i s\'adapta. Raquel et diu: **"Ara saps per on has de buscar les respostes."**',
      good: '**Molt bé, Pau.** Has captat els conceptes fonamentals i has pres la majoria de decisions encertades. Recorda que el control motor no té una teoria única: el teu repte com a entrenador és **dialogar entre teoria i pràctica** amb mirada crítica. Cada esportista és un sistema únic; cap programa estàndard et donarà les respostes que necessites.',
      ok: '**Correcte, Pau.** Has entès els conceptes bàsics, però en alguns moments has recaigut en visions reduccionistes: control motor com a força, o com a memorització de patrons. Repassa la distinció entre perspectives teòriques i, sobretot, la idea de Joan: el control motor **emergeix de la interacció** amb l\'entorn. No és un programa intern que cal instal·lar.',
      needsWork: '**Necessites aprofundir, Pau.** Els principis del control motor representen un canvi de perspectiva profund respecte a l\'entrenament convencional. Et recomanem relllegir les aportacions de Joan (perspectiva ecològica), Andrés (conèixer l\'esportista) i Esther (consciència corporal i prevenció de lesions). La clau és acceptar que **no hi ha una teoria única** i que el teu paper és dialogar entre teoria i pràctica amb coherència i ètica.'
    },
    seguent: 'scene_28'
  },

  scene_28: {
    id: 'scene_28',
    tipus: 'text_block',
    titol: 'La Mirada Crítica de l\'Entrenador',
    personatge: 'narracio',
    narracio: 'En Pau surt del seminari amb el bloc ple de notes. Però, sobretot, amb tres preguntes noves que no tenia en entrar: ¿Quin control motor té el meu atleta? ¿Des d\'on partim? ¿Quins contexts li estic dissenyant?\n\nJoan va tancar amb una idea que l\'ha quedat gravada:',
    dialeg: {
      personatge: 'joan',
      text: '"Les consideracions per a la millora del control motor han de ser completament diferents per a l\'elit que per a la formació en les etapes infantils. Però en general, **el control motor ha de proporcionar-te informació sobre la teva forma d\'aprendre per seguir aprenent**.\n\nL\'entrenador hauria de dissenyar experiències properes als límits de l\'esportista, i estar disponible per donar-li el suport necessari en cada moment. Treballar a prop dels límits comporta gestionar moltes emocions, però també moure\'s en l\'espai del que sé fer i el que necessito saber."'
    },
    contingut_pedagogic: {
      titol: 'Tres principis que Pau porta ara a cada entrenament',
      text: '**1. No hi ha una teoria única del control motor** — Dialogar entre la perspectiva conductista, cognitivista i ecològica, de manera crítica i coherent, és la feina de l\'entrenador professional.\n**2. L\'esportista primer, el programa després** — Conèixer el punt de partida, els límits i les circumstàncies de l\'esportista és el pas zero de qualsevol intervenció sobre el control motor.\n**3. Acceptar la incertesa com a part del procés** — Els canvis de patró motor no són lineals, i algunes modificacions seran efímeres. L\'entrenador que accepta la incertesa i el retrocés com a part del procés treballa amb molta més intel·ligència.'
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
        state.score += pts;
        state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
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
        var pts = op.punts !== undefined ? op.punts : 0;
        state.score += pts;
        state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
        scene.opcions.forEach(function (o, i) {
          ia.querySelectorAll('.btn-option')[i].className = 'btn btn-option btn-disabled';
        });
        btn.className = 'btn btn-option btn-selected btn-disabled';
        _save();
        var icon = pts >= 10 ? 'fb-good' : pts >= 5 ? 'fb-ok' : 'fb-bad';
        var nextScene = op.seguent || scene.seguent;
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
        if (userChecked && correct)   { div.classList.add('cl-correct');   pts += 2; }
        else if (!userChecked && !correct) { div.classList.add('cl-correct'); pts += 2; }
        else if (userChecked && !correct)  { div.classList.add('cl-incorrect'); }
        else                               { div.classList.add('cl-missed'); }
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

    var mentor = CHARACTERS.raquel;
    var html = '<div class="epilogue-container">';
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
    html += '<button class="btn btn-primary btn-enabled" onclick="Engine.showScene(\'scene_28\')">';
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

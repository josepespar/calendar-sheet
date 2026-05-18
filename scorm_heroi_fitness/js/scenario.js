/* ============================================================
   SCENARIO – El Viatge de l'Heroi · FitCore v2
   ~30 escenes · Quiz + Decisions · Puntuació màxima: 100 pts
   Imatges a assets/: img_01.jpg…img_05.jpg
   ============================================================ */

const CHARACTERS = {
  harry:    { name: 'Harry',    color: '#4A90D9', initials: 'H',  shape: 'circle' },
  hanna:    { name: 'Hanna',    color: '#27AE60', initials: 'Ha', shape: 'circle' },
  jack:     { name: 'Jack',     color: '#C0392B', initials: 'J',  shape: 'circle' },
  miquel:   { name: 'Miquel',   color: '#8E44AD', initials: 'M',  shape: 'circle' },
  jordi:    { name: 'Jordi',    color: '#E67E22', initials: 'Jo', shape: 'circle' },
  narracio: { name: 'Narrador', color: '#7F8C8D', initials: '✦',  shape: 'square' }
};

const JOURNEY_STAGES = [
  { id: 'act1',  label: 'Acte I',    title: 'El Món Ordinari',           scenes: ['scene_01','scene_02','scene_03','scene_04','scene_05','scene_06','scene_07','scene_08'], color: '#4A90D9' },
  { id: 'act2a', label: 'Acte II-A', title: 'El Primer Llindar',         scenes: ['scene_09','scene_10','scene_11','scene_12','scene_13','scene_14'],                      color: '#FF6B35' },
  { id: 'act2b', label: 'Acte II-B', title: 'La Prova Suprema',          scenes: ['scene_15','scene_16','scene_17','scene_18','scene_19','scene_20'],                      color: '#E67E22' },
  { id: 'act2c', label: 'Acte II-C', title: 'La Recompensa',             scenes: ['scene_21','scene_22','scene_23'],                                                        color: '#D4A017' },
  { id: 'act3',  label: 'Acte III',  title: 'El Retorn',                 scenes: ['scene_24','scene_25','scene_26','scene_26b','scene_27','scene_28','scene_29'],           color: '#27AE60' },
  { id: 'final', label: 'Epíleg',    title: 'El Nou Harry',              scenes: ['scene_30'],                                                                              color: '#8E44AD' }
];

/* Nombre de scenes en el camí principal (sense condicionals) */
const CANONICAL_SCENE_COUNT = 28;

/* ============================================================
   SCENES
   ============================================================ */
const scenes = {

  /* ══════════════════════════════════
     ACTE I – EL MÓN ORDINARI
  ══════════════════════════════════ */

  scene_01: {
    id: 'scene_01', etapa: 'El Món Ordinari', personatge: 'harry',
    tipus: 'text_block', titol: 'Benvingut a FitCore',
    imatge: 'assets/img_01.svg',
    narracio: `Harry té 18 anys, el títol de tècnic esportiu acabat de plastificar i una motxilla plena de blocs de notes sobre periodització. Avui és el seu primer dia com a entrenador a FitCore, el centre de fitness més complet del barri.

La recepció és moderna i acollent. Música suau, aroma de cafè de la màquina del fons, clients que entren i surten amb les bosses al muscle. Harry s'atura a l'entrada i respira fons. Sap perfectament com dissenyar un programa d'entrenament. El que no sap, però, és com parlar amb la gent quan les coses es compliquen.`,
    dialeg: { personatge: 'harry', text: '"Bé. Estic preparat. O almenys això és el que em dic."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Per què les habilitats socials?',
      text: 'Un entrenador no és només un expert en exercici: és un **comunicador**. Les habilitats tècniques obren la porta, però les habilitats comunicatives determinen si el client torna, progresa i confia en tu.'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02', etapa: 'El Món Ordinari', personatge: 'narracio',
    tipus: 'text_block', titol: 'L\'equip: la Hanna i en Jack',
    imatge: 'assets/img_01.svg',
    narracio: `El director Miquel presenta Harry a l'equip. La **Hanna** porta set anys al centre. Quan li estreny la mà, mira Harry als ulls i li diu: "Qualsevol cosa que necessitis, estic aquí." La seva postura és oberta, el somriure és sincer.

En **Jack** porta deu anys. La seva salutació és breu, quasi mecànica. "Benvingut." Dues síl·labes i ja ha donat l'esquena. Els clients el respecten per la seva experiència tècnica, però alguns li diuen que és difícil d'apropar.`,
    dialeg: { personatge: 'narracio', text: '"Dos professionals. Dos estils de comunicació radicalment oposats. Harry haurà de decidir quin model vol seguir."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Estils de comunicació professional',
      text: 'La comunicació professional combina **contingut** (el que saps) i **relació** (com connectes). La Hanna combina totes dues. En Jack domina el contingut però falla en la relació. Els clients necessiten les dues coses per sentir-se ben atesos.'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03', etapa: 'El Món Ordinari', personatge: 'narracio',
    tipus: 'quiz', titol: 'Comprensió: l\'equip de FitCore',
    imatge: 'assets/img_03.svg',
    narracio: `Reflexiona sobre el que acabes de llegir sobre la Hanna i en Jack.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: 'Quina característica diferencia principalment la Hanna d\'en Jack?',
    opcions: [
      { id: 'A', text: 'La Hanna té més experiència tècnica que en Jack', correcta: false, feedback: 'No exactament. En Jack porta deu anys i la Hanna set, i els seus coneixements tècnics no es comparen directament. La diferència clau és una altra.' },
      { id: 'B', text: 'La Hanna combina expertesa tècnica amb connexió relacional; en Jack es queda al contingut', correcta: true, feedback: 'Correcte! La Hanna suma **contingut + relació**: és experta i alhora connecta emocionalment. En Jack domina el contingut però la seva comunicació és distant. Aquesta diferència és la que marca la experiència del client.' },
      { id: 'C', text: 'En Jack comunica millor que la Hanna però és menys empàtic', correcta: false, feedback: 'Al contrari. En Jack té dificultats comunicatives que afecten la relació amb els clients, mentre la Hanna és un exemple de comunicació efectiva.' }
    ],
    seguent: 'scene_04'
  },

  scene_04: {
    id: 'scene_04', etapa: 'La Crida a l\'Aventura', personatge: 'miquel',
    tipus: 'text_block', titol: 'La queixa de la Maria',
    imatge: 'assets/img_02.svg',
    narracio: `Dues hores después del primer dia. El director Miquel crida Harry al despatx. "La Maria Puig ha deixat una queixa formal. Té 55 anys, porta tres mesos al centre i diu que se sent invisible: ningú li pregunta com es troba, li expliquen exercicis sense adaptar-los i no l'escolten."

Miquel mostra la pantalla on apareix el formulari de queixa. La Maria ha escrit: "Em sento com si fos un número, no una persona. He decidit valorar si continuo al centre."`,
    dialeg: { personatge: 'miquel', text: '"Harry, necessito que algú gestioni això ara. Tu, la Hanna, o en Jack podeu parlar amb ella. Decideix tu."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Barreres comunicatives en entorns de fitness',
      text: 'Quan un client diu "ningú m\'escol·ta", sovint no es refereix al contingut sinó al **canal**: la comunicació s\'ha produït, però sense tenir en compte les seves necessitats emocionals. Les barreres més freqüents: terminologia tècnica excessiva, manca de contact visual i absència de feedback empàtic.'
    },
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05', etapa: 'La Crida a l\'Aventura', personatge: 'miquel',
    tipus: 'decision_scenario', titol: 'Com respon Harry a la crida?',
    imatge: 'assets/img_04.svg',
    narracio: `Harry té tres opcions davant. El rellotge avança. La Maria espera al saló.`,
    dialeg: { personatge: 'harry', text: '"Això és real. No és un exercici. Una persona real espera una resposta real."' },
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Acceptar la tasca i demanar consell a la Hanna abans de parlar amb la Maria',
        feedback: '**Decisió òptima.** Buscar orientació d\'una professional experimentada és maduresa professional, no inseguretat. Demostres **escolta activa proactiva**: abans de comunicar, t\'assegures de tenir les eines adequades. La preparació és el primer facilitador de la comunicació efectiva.',
        punts: 10, seguent: 'scene_07'
      },
      {
        id: 'B', text: 'Acceptar la tasca i anar directament a parlar amb la Maria sense preparació',
        feedback: 'Positiu que acceptis la responsabilitat, però anar sense preparació pot agreujar la situació. La Maria ja se sent invisible; si la conversa falla, reforça la seva percepció negativa. Una mínima preparació marca la diferència.',
        punts: 5, seguent: 'scene_07'
      },
      {
        id: 'C', text: 'Derivar el cas a en Jack, que té més experiència',
        feedback: 'Derivar sense acompanyament és una **barrera comunicativa organitzacional**: el client queda sense resposta directa i interpreta que el problema no és prioritari. A més, delegar un conflicte sense supervisió pot agravar-lo. Veuràs el que passa.',
        punts: 0, seguent: 'scene_06'
      }
    ]
  },

  scene_06: {
    id: 'scene_06', etapa: 'Rebuig de la Crida', personatge: 'jack',
    tipus: 'text_block', titol: 'Jack gestiona la Maria',
    imatge: 'assets/img_02.svg',
    narracio: `Harry observa des del passadís. En Jack s'asseu davant de la Maria amb els braços creuats. Sense preguntar res, comença: "Mira, tots els entrenadors aquí sabem el que fem. Si et donava aquells exercicis, era perquè eren els correctes. Potser t'has d'esforçar una mica més."

La Maria encongeix els muscles. En Jack s'aixeca: "Si no n'estàs satisfeta, parla amb direcció." La Maria queda sola, amb els ulls humits.`,
    dialeg: { personatge: 'narracio', text: '"En Jack ha usat tres barreres en menys de dos minuts: postura de domini, invalidació emocional i absència de feedback positiu. El client no ha rebut una resposta: ha rebut una sentència."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Barreres actives: quan la comunicació destrueix la relació',
      text: '**Judicis prematurs** (concloure sense escoltar), **invalidació emocional** (minimitzar el que l\'altre sent) i **absència de feedback** (no confirmar que has entès) fan que el client no se senti respectat, sinó processat.'
    },
    seguent: 'scene_07'
  },

  scene_07: {
    id: 'scene_07', etapa: 'La Trobada amb el Mentor', personatge: 'hanna',
    tipus: 'worked_example', titol: 'La lliçó de la Hanna',
    imatge: 'assets/img_03.svg',
    narracio: `La Hanna porta Harry a la sala d'entrenadors i li prepara un cafè. Parla amb calma: "Cada queixa d'un client és un regal. Ens diu on falla la comunicació. La Maria no es queixa dels exercicis: es queixa de no sentir-se vista."

A continuació modela una conversa amb "la Maria" en role-play. S'asseu al costat de Harry (no enfront), manté contacte visual suau, capeja el cap, i quan la Maria acaba de parlar, reformula: "El que m'estàs dient és que sents que no ens hem pres el temps per entendre el que tu necessites. T'escolto."`,
    dialeg: { personatge: 'hanna', text: '"La comunicació efectiva no és el que dius: és com ho dius, quan ho dius, i sobretot, quant espai dones a l\'altra persona per dir el que necessita."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: 'Com iniciar una conversa difícil: 5 passos',
      text: `**1. Crear l'espai:** lloc tranquil, sense mòbil visible, asseure's al costat.
**2. Obertura empàtica:** "Gràcies per dir-nos el que sents. Vull entendre la teva experiència."
**3. Escolta activa:** deixar parlar sense interrompre, capejar el cap, contacte visual suau.
**4. Reformulació:** "El que m'estàs dient és... ho he entès bé?"
**5. Proposta conjunta:** "Què podríem fer diferent perquè et sentis millor acompanyada?"`
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08', etapa: 'La Trobada amb el Mentor', personatge: 'hanna',
    tipus: 'quiz', titol: 'Comprensió: escolta activa',
    imatge: 'assets/img_03.svg',
    narracio: `La Hanna et fa una pregunta per assegurar-se que has entès el pas 4 del model.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: 'Al pas 4 del model de la Hanna, la "reformulació" serveix per:',
    opcions: [
      { id: 'A', text: 'Repetir literalment el que ha dit el client per demostrar que has escoltat', correcta: false, feedback: 'La repetició literal no és reformulació: és eco. La reformulació implica processar el missatge i tornar-lo amb les teves pròpies paraules per verificar la comprensió.' },
      { id: 'B', text: 'Verificar que has entès el missatge del client i donar-li l\'oportunitat de corregir-te', correcta: true, feedback: 'Exacte! La **reformulació** ("el que m\'estàs dient és...") té dues funcions: confirmar que has entès i demostrar al client que estaves escoltant de veritat. Quan el client et pot corregir, se sent respectat i escoltat.' },
      { id: 'C', text: 'Proposar solucions mentre el client parla per estalviar temps', correcta: false, feedback: 'Proposar solucions mentre l\'altre parla trenca l\'escolta activa: implica que ja has decidit la resposta abans d\'entendre el problema sencer. Cal esperar el pas 5.' }
    ],
    seguent: 'scene_09'
  },

  /* ══════════════════════════════════
     ACTE II-A – EL PRIMER LLINDAR
  ══════════════════════════════════ */

  scene_09: {
    id: 'scene_09', etapa: 'Creuament del Primer Llindar', personatge: 'harry',
    tipus: 'text_block', titol: 'La primera sessió de grup',
    imatge: 'assets/img_04.svg',
    narracio: `Tres dies después. Harry haurà de dirigir la primera sessió de grup com a entrenador principal. El grup és divers: la **Maria** (55 anys, mobilitat limitada, molt sensible), en **Jordi** (35 anys, triatleta avançat, molt exigent) i la **Lluïsa** (65 anys, jubilada, cardiopatia lleu, ve per socialitzar tant com per entrenar).

Harry repassa els fulls de cada client. Tres perfils, tres necessitats, tres estils de comunicació. Exactament el tipus de repte que no surten als manuals.`,
    dialeg: { personatge: 'harry', text: '"Tres persones completament diferents. Haig de trobar una manera de connectar amb cadascuna d\'elles des del primer moment."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Adaptar la comunicació al perfil del client',
      text: 'Un bon entrenador adapta el seu estil comunicatiu a cada client: **directiu** per a qui necessita estructura clara, **empàtic** per a qui necessita suport emocional, **col·laboratiu** per a qui vol participar en les decisions. Usar el mateix estil amb tothom és un error freqüent.'
    },
    seguent: 'scene_10'
  },

  scene_10: {
    id: 'scene_10', etapa: 'Creuament del Primer Llindar', personatge: 'jack',
    tipus: 'decision_scenario', titol: 'El consell de Jack vs. la Hanna',
    imatge: 'assets/img_01.svg',
    narracio: `Minuts abans de la sessió, en Jack entra a la sala d'entrenadors i li dóna un cop a l'espatlla: "Consell d'amic: sigues ferm des del principi. No els deixis que et dominin amb excuses. Aquí venen a entrenar, no a plorar."

Harry recorda el que li ha dit la Hanna: "Pregunta com es troben. Adapta't. Cada persona és un món." Dues filosofies oposades. El grup espera a la porta.`,
    dialeg: { personatge: 'jack', text: '"Que quedi clar qui mana des del primer minut. Si cedeixes un centímetre el primer dia, et mengen sencer."' },
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Seguir el consell de Jack: entrar amb autoritat, sense preguntar com es troben',
        feedback: 'Aplicar l\'enfocament de Jack genera **incoherència comunicativa**: les paraules diuen "benvinguts" però el to i la postura diuen "obeïu". La firmesa sense empatia no és autoritat professional: és distància que genera tensió des del primer minut.',
        punts: 0, seguent: 'scene_11'
      },
      {
        id: 'B', text: 'Aplicar el model de la Hanna: salutació personalitzada i preguntar com es troben',
        feedback: 'Decisió òptima. Iniciar amb una **salutació personalitzada** (dir el nom de cada persona) i preguntar com es troben activa l\'escolta activa des del primer moment. Quan el client sent que l\'entrenador s\'interessa per ell com a persona, la sessió comença des d\'un lloc de confiança.',
        punts: 10, seguent: 'scene_12'
      }
    ]
  },

  scene_11: {
    id: 'scene_11', etapa: 'Creuament del Primer Llindar', personatge: 'narracio',
    tipus: 'text_block', titol: 'Les conseqüències de la firmesa sense empatia',
    imatge: 'assets/img_04.svg',
    narracio: `Harry entra directe als exercicis sense cap salutació personal. En Jordi creua els braços i observa amb escepticisme. La Lluïsa no s'atreveix a preguntar si pot fer l'exercici amb la seva cardiopatia. La Maria mira el terra.

A mitja sessió, en Jordi interromp: "Espera. El meu programa era diferent." Harry no sap de qui parla —en Jordi és client d'un altre entrenador— i la comunicació es converteix en un malentès públic. La sessió acaba amb tensions.`,
    dialeg: { personatge: 'narracio', text: '"La **comunicació no verbal incongruent** i l\'absència d\'escolta activa generen tensió fins i tot quan les paraules semblen correctes. Harry haurà de gestionar les conseqüències."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Comunicació no verbal: el 55% del missatge',
      text: 'Según el model de Mehrabian (en contextos emocionals): **55% postura i expressió**, 38% to de veu, 7% paraules. Una postura oberta, contacte visual suau i somriure natural generen seguretat. Una postura tancada genera por o resistència, independentment del que dius.'
    },
    seguent: 'scene_12'
  },

  scene_12: {
    id: 'scene_12', etapa: 'Proves, Aliats i Enemics', personatge: 'jordi',
    tipus: 'text_block', titol: 'La crítica pública d\'en Jordi',
    imatge: 'assets/img_05.svg',
    narracio: `Una setmana después. Harry ha canviat el programa d'en Jordi: ha reduït les sentadilles i ha afegit treball de core específic per al triatlò. Una millora objectivament justificada. Però no l'ha comunicat al client.

En Jordi arriba a la sala, mira el full nou i es gira cap a Harry en veu alta: "Això és el problema. Canvies el que funciona sense dir res a ningú. No em consultes, no m'expliques res, i esperes que confïi en tu?" Tres clients giren el cap.`,
    dialeg: { personatge: 'jordi', text: '"Potser el problema és que creus que perquè tens el títol ja ho saps tot. Jo porto set anys entrenant per a triatlons."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Assertivitat: expressar-se amb respecte',
      text: 'Hi ha tres estils de comunicació: **Passiu** (cedir sempre, suprimir les pròpies necessitats), **Agressiu** (imposar sense considerar l\'altre) i **Assertiu** (expressar amb claredat i respecte mutu). La resposta assertiva en una situació de crítica pública comença per validar l\'emoció de l\'altre.'
    },
    seguent: 'scene_13'
  },

  scene_13: {
    id: 'scene_13', etapa: 'Proves, Aliats i Enemics', personatge: 'jordi',
    tipus: 'decision_scenario', titol: 'Com respon Harry a la crítica pública?',
    imatge: 'assets/img_05.svg',
    narracio: `Harry té tres opcions. Tres clients observen. El silenci dura un segon que sembla un minut.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Defensar-se públicament: "El canvi era tècnicament millor, i tu hauràs de confiar en el meu criteri"',
        feedback: 'Resposta **agressiva** que escala el conflicte. Tenir raó tècnica no justifica una confrontació pública. En Jordi i els altres clients percebran una batalla d\'egos, no una mostra de professionalitat.',
        punts: 0, seguent: 'scene_14'
      },
      {
        id: 'B', text: 'Disculpar-se excessivament davant de tothom: "Ho sento molt, tens raó en tot, no ho tornaré a fer"',
        feedback: 'Resposta **passiva**. Tot i intentar desactivar el conflicte, envia el missatge que t\'has equivocat greument quan la decisió tècnica era correcta. La passivitat crea un precedent: la crítica pública funciona com a eina de pressió.',
        punts: 5, seguent: 'scene_14'
      },
      {
        id: 'C', text: 'Validar el sentiment d\'en Jordi públicament i proposar parlar en privat: "Entenc la teva frustració. Mereixies saber el canvi. Podem parlar cinc minuts en privat?"',
        feedback: '**Resposta assertiva exemplar.** Validar l\'emoció ("entenc la teva frustració") i proposar un espai privat desactiva l\'escalada pública, protegeix la dignitat de totes dues parts i obre la porta a una solució real. L\'assertivitat comença per reconèixer l\'emoció de l\'altre.',
        punts: 10, seguent: 'scene_14'
      }
    ]
  },

  scene_14: {
    id: 'scene_14', etapa: 'Proves, Aliats i Enemics', personatge: 'hanna',
    tipus: 'quiz', titol: 'Comprensió: assertivitat',
    imatge: 'assets/img_03.svg',
    narracio: `La Hanna et demana que reflexionis sobre el que ha passat.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: 'Per quin motiu la resposta C (assertiva) és millor que la B (passiva), si totes dues eviten el conflicte immediat?',
    opcions: [
      { id: 'A', text: 'Perquè la resposta C és més educada i els clients la valoren més', correcta: false, feedback: 'L\'educació és important, però no és la raó principal. La diferència clau és altra.' },
      { id: 'B', text: 'Perquè la resposta C preserva la dignitat de totes dues parts i estableix un precedent de resolució constructiva, mentre la B crea dependència de la crítica pública', correcta: true, feedback: 'Correcte! La resposta assertiva no només desactiva el conflicte: crea un **precedent positiu**. La resposta passiva, en canvi, ensenya al client que la crítica pública és eficaç per aconseguir el que vol, cosa que pot repetir-se.' },
      { id: 'C', text: 'Perquè la resposta C posa fi al conflicte de manera definitiva', correcta: false, feedback: 'Cap resposta acaba el conflicte "de manera definitiva": la conversió real vindrà en la conversa privada posterior. La resposta C simplement crea les condicions òptimes per a aquella conversa.' }
    ],
    seguent: 'scene_15'
  },

  /* ══════════════════════════════════
     ACTE II-B – LA PROVA SUPREMA
  ══════════════════════════════════ */

  scene_15: {
    id: 'scene_15', etapa: 'Aproximació a la Cova', personatge: 'hanna',
    tipus: 'worked_example', titol: 'Preparació de la conversa difícil',
    imatge: 'assets/img_03.svg',
    narracio: `La Hanna agafa Harry del braç just quan surt de la sala. "Bé fet per no escalar. Però ara ve la part difícil: has de tenir la conversa real amb en Jordi. I per tenir una conversa difícil de manera professional, cal preparar-se."

Explica els cinc elements clau: definir l'objectiu, anticipar les emocions de l'altra part, triar el moment i el lloc, preparar una obertura empàtica i planificar l'escolta activa.`,
    dialeg: { personatge: 'hanna', text: '"Una conversa difícil mal preparada pot destruir una relació que portes mesos construint. Però ben preparada, pot convertir el conflicte en la base d\'una confiança molt més sòlida."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: 'Estructura d\'una conversa difícil',
      text: `**1. Objectiu:** Quin resultat vull? (entendre'ns, no guanyar)
**2. Anticipació:** Com es pot sentir l\'altra persona?
**3. Espai i moment:** Lloc privat, moment tranquil, sense pressa.
**4. Obertura empàtica:** Reconèixer la perspectiva de l\'altre abans d\'explicar la teva.
**5. Escolta activa:** Deixar parlar, reformular, confirmar.`
    },
    seguent: 'scene_16'
  },

  scene_16: {
    id: 'scene_16', etapa: 'Aproximació a la Cova', personatge: 'harry',
    tipus: 'checklist', titol: 'Prepara la conversa amb en Jordi',
    imatge: 'assets/img_04.svg',
    narracio: `Harry s'asseu amb un full en blanc. La Hanna li ha dit que hi ha cinc elements essencials per preparar qualsevol conversa difícil. Marca els que creus que hauries d'incloure en la preparació. Selecciona tots els correctes.`,
    dialeg: { personatge: 'harry', text: '"No és una conversa de cinc minuts. És la conversa que pot canviar com en Jordi em veu com a professional."' },
    contingut_pedagogic: null,
    checklistItems: [
      { id: 'cl1', text: 'Definir quin és el meu objectiu real (entendre\'ns, no demostrar que tinc raó)', correcta: true },
      { id: 'cl2', text: 'Triar un moment tranquil i un lloc privat, fora de la zona d\'entrenament', correcta: true },
      { id: 'cl3', text: 'Preparar una obertura que reconegui la perspectiva d\'en Jordi', correcta: true },
      { id: 'cl4', text: 'Practicar l\'escolta activa: deixar parlar sense interrompre i reformular', correcta: true },
      { id: 'cl5', text: 'Establir un to assertiu: dir el que penso sense atacar ni cedir innecessàriament', correcta: true },
      { id: 'cl6', text: 'Preparar una llista de tots els errors que he comès per disculpar-me per cadascun', correcta: false },
      { id: 'cl7', text: 'Demanar a en Jack que vingui com a suport en cas que en Jordi s\'enfadi', correcta: false }
    ],
    punts_per_item: 2,
    seguent: 'scene_17'
  },

  scene_17: {
    id: 'scene_17', etapa: 'La Prova Suprema', personatge: 'jordi',
    tipus: 'text_block', titol: 'La conversa: el context',
    imatge: 'assets/img_05.svg',
    narracio: `La sala de reunions petita del costat del despatx de Miquel. Dues cadires, una taula estreta, la porta tancada. Harry ha convocat en Jordi amb un missatge clar: "Vull explicar-te per què he canviat el teu programa i escoltar el que penses. Demà a les 10?"

En Jordi arriba puntual però tens. Es creua de braços nada més seure. Harry respira fons. Ara comença la prova suprema.`,
    dialeg: { personatge: 'narracio', text: '"La prova suprema no és física ni tècnica: és emocional. Quan estàs sota pressió i l\'altra persona està enfadada, la teva capacitat de regulació emocional determina si la conversa construeix o destrueix."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Regulació emocional en situacions de tensió',
      text: 'En situacions de tensió, el cervell activa la resposta d\'amenaça, dificultant el pensament racional. Tècniques: **respiració profunda**, pausa conscient ("necessito un moment"), i **reformulació cognitiva** ("el seu enfado no és un atac personal: és una necessitat insatisfeta").'
    },
    seguent: 'scene_18'
  },

  scene_18: {
    id: 'scene_18', etapa: 'La Prova Suprema – Inici', personatge: 'jordi',
    tipus: 'decision_scenario', titol: 'Decisió 1: Com inicia Harry la conversa?',
    imatge: 'assets/img_05.svg',
    narracio: `En Jordi espera. Harry té la paraula. Primer moviment.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"Jordi, entenc la teva frustració. T\'hauré d\'haver consultat el canvi. Vull explicar-te el raonament i escoltar el teu punt de vista."',
        feedback: '**Obertura empàtica perfecta.** Harry reconeix la perspectiva d\'en Jordi, assumeix responsabilitat sense excessos, i estableix l\'objectiu mutu. La comunicació assertiva equilibra drets propis i drets de l\'altre.',
        punts: 10, seguent: 'scene_19'
      },
      {
        id: 'B', text: '"Jordi, vull explicar-te per què he canviat el programa. Era per millorar el teu rendiment al triatlò."',
        feedback: 'Correcte però incomplet. Explicar sense reconèixer primer la perspectiva d\'en Jordi posa la raó pròpia per davant de l\'emoció de l\'altre. La comunicació efectiva en conflictes requereix que primer l\'altra persona se senti escoltada.',
        punts: 5, seguent: 'scene_19'
      },
      {
        id: 'C', text: '"Jordi, sé que vas reaccionar de forma exagerada ahir, però em semblava bé parlar per aclarir les coses."',
        feedback: 'Inici **defensiu i acusatori**. Qualificar la reacció d\'en Jordi com "exagerada" és una invalidació emocional que tancarà la conversa immediatament. L\'altra persona haurà de defensar-se en comptes de dialogar.',
        punts: 0, seguent: 'scene_19'
      }
    ]
  },

  scene_19: {
    id: 'scene_19', etapa: 'La Prova Suprema – Tensió', personatge: 'jordi',
    tipus: 'decision_scenario', titol: 'Decisió 2: En Jordi s\'enfada',
    imatge: 'assets/img_05.svg',
    narracio: `En Jordi s\'incorpora: "Em sap molt greu, però no pots canviar el meu programa sense consultar-me. No és professional." La seva veu ha pujat de to. Harry nota la calor a la cara.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Pausa, respiració, i dir: "Tens raó en el fons. El canvi era correcte però el procés ha fallat. Hauria d\'haver-te informat. Com hauries preferit que ho gestionés?"',
        feedback: '**Regulació emocional excel·lent.** Harry no reacciona a la intensitat emocional: fa una pausa, valida el contingut (no la forma), assumeix la responsabilitat real i redirigeix cap a la solució. Això és **empatia professional**: comprendre sense perdre el rol.',
        punts: 10, seguent: 'scene_20'
      },
      {
        id: 'B', text: 'Dir: "Ho sento, ho sento, tens raó, disculpa, no ho tornaré a fer..."',
        feedback: 'Disculpa excessiva i passiva. Assumir tota la culpa sense matisos no resol el problema i crea un precedent negatiu. La resposta assertiva implica assumir la responsabilitat real, no tota la responsabilitat imaginable.',
        punts: 5, seguent: 'scene_20'
      },
      {
        id: 'C', text: 'Dir: "No és qüestió de consultar-te: jo sóc l\'entrenador i tu has de confiar en el meu criteri professional."',
        feedback: 'Resposta **agressiva** que escala el conflicte. Invocar l\'autoritat quan l\'altre ha expressat una necessitat legítima (ser consultat) tanca la conversa. L\'autoritat professional no s\'imposa: es guanya.',
        punts: 0, seguent: 'scene_20'
      }
    ]
  },

  scene_20: {
    id: 'scene_20', etapa: 'La Prova Suprema – Resolució', personatge: 'jordi',
    tipus: 'decision_scenario', titol: 'Decisió 3: Arribar a un acord',
    imatge: 'assets/img_05.svg',
    narracio: `En Jordi accepta que el canvi tècnic tenia sentit, però vol ser consultat en el futur. Harry ha de proposar com serà la relació d'ara endavant.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"D\'acord. En el futur, quan vulgui modificar el teu programa, t\'enviaré un missatge 48 hores abans explicant el canvi i el motiu. Si tens dubtes, els parlem. Funciona per a tu?"',
        feedback: '**Acord mutu concret.** Un protocol específic (48 hores, missatge, explicació) demostra que has escoltat la necessitat real d\'en Jordi i converteix el conflicte en un acord de treball sostenible. Això és **negociació efectiva**.',
        punts: 10, seguent: 'scene_21'
      },
      {
        id: 'B', text: '"D\'acord, d\'ara endavant tu decideixes el programa i jo l\'executo."',
        feedback: 'Cessió excessiva. Renunciar al rol professional per evitar el conflicte no és un bon acord: és rendició. El client quedarà satisfet a curt termini, però l\'entrenador perdrà l\'autoritat necessària per fer la seva feina correctament.',
        punts: 5, seguent: 'scene_21'
      },
      {
        id: 'C', text: '"Bé, ho intentaré." (sense proposar res concret)',
        feedback: '"Ho intentaré" no és un compromís: és una manera de finalitzar la conversa sense resoldre-la. En Jordi quedarà amb la sensació que res canviarà. Les converses difícils han d\'acabar amb acords clars, no amb bones intencions.',
        punts: 0, seguent: 'scene_21'
      }
    ]
  },

  /* ══════════════════════════════════
     ACTE II-C – LA RECOMPENSA
  ══════════════════════════════════ */

  scene_21: {
    id: 'scene_21', etapa: 'La Recompensa', personatge: 'hanna',
    tipus: 'worked_example', titol: 'La confiança guanyada',
    imatge: 'assets/img_03.svg',
    narracio: `En Jordi surt de la sala i li estreny la mà a Harry. Per a en Jordi, que rarament expressa aprovació, és molt significatiu. "Bé, em sembla bé el que has proposat. Continuem."

La Hanna, que havia escoltat des del passadís, s'apropa i diu simplement: "Has crescut deu anys en una hora." Asseguts a la sala, li explica els cinc estils d'afrontament del conflicte i per quin Harry ha optat.`,
    dialeg: { personatge: 'hanna', text: '"El conflicte no és el problema. El conflicte és l\'oportunitat. Com el gestionis determina si la relació surt enfortida o trencada."' },
    contingut_pedagogic: {
      tipus: 'worked_example', titol: 'Els cinc estils d\'afrontament (Thomas-Kilmann)',
      text: `**Competència:** Imposar la posició. Útil en emergències. Risc: danya relacions.
**Evitació:** Retirar-se. Útil quan el conflicte és trivial. Risc: problemes acumulats.
**Acomodació:** Cedir. Útil per mantenir relació. Risc: pèrdua d\'autoritat.
**Compromís:** Cada part cedeix una mica. Útil per a solucions ràpides.
**Col·laboració:** Solució que satisfà tothom. La millor opció. Requereix temps i confiança.`
    },
    seguent: 'scene_22'
  },

  scene_22: {
    id: 'scene_22', etapa: 'La Recompensa', personatge: 'hanna',
    tipus: 'text_block', titol: 'El conflicte com a oportunitat',
    imatge: 'assets/img_03.svg',
    narracio: `La Hanna li mostra a Harry el que han aconseguit: en Jordi ha passat de desconfiar a proposar millores conjuntes. La queixa de la Maria ha passat a una reunió de revisió del programa on ella participa activament.

"Veus el patró?" li diu la Hanna. "Quan escoltes de veritat, les persones passen de queixar-se a co-crear. El client deixa de ser un problema i es converteix en un aliat."`,
    dialeg: { personatge: 'hanna', text: '"L\'assertivitat no és duresa. És claredat. I la claredat, quan va acompanyada de respecte, crea confiança."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Empatia professional: comprendre sense perdre el rol',
      text: 'L\'empatia professional és la capacitat de comprendre la perspectiva i les emocions del client **sense fusionar-te** amb elles. No és dir "sé exactament com et sents" (ningú ho sap), sinó "entenc per qué et sents així i em sembla legítim".'
    },
    seguent: 'scene_23'
  },

  scene_23: {
    id: 'scene_23', etapa: 'La Recompensa', personatge: 'hanna',
    tipus: 'quiz', titol: 'Comprensió: estils d\'afrontament',
    imatge: 'assets/img_03.svg',
    narracio: `Comprovem que has entès els cinc estils de Thomas-Kilmann.`,
    dialeg: null,
    contingut_pedagogic: null,
    pregunta: 'Quin estil d\'afrontament va usar Harry durant la conversa amb en Jordi (arriba a un protocol concret que satisfà totes dues parts)?',
    opcions: [
      { id: 'A', text: 'Compromís (totes dues parts cedeixen una mica)', correcta: false, feedback: 'El compromís implica que totes dues parts perden alguna cosa. En el cas d\'Harry i Jordi, cap dels dos va perdre res: Harry manté el rol professional i Jordi obté la consulta prèvia que necessitava.' },
      { id: 'B', text: 'Col·laboració (solució que satisfà totes dues parts sense cedir en allò essencial)', correcta: true, feedback: 'Correcte! La **col·laboració** és l\'estil òptim: Harry no renuncia al rol de decidir tècnicament, i en Jordi obté el que necessitava (ser consultat). Cap de les dues parts ha hagut de renunciar a allò essencial.' },
      { id: 'C', text: 'Acomodació (Harry cedeix a les demandes d\'en Jordi)', correcta: false, feedback: 'Si Harry hagués dit "d\'ara endavant tu decideixes el programa", hauria estat acomodació. L\'acord real va ser diferent: Harry va mantenir el rol professional i va afegir un protocol de comunicació.' }
    ],
    seguent: 'scene_24'
  },

  /* ══════════════════════════════════
     ACTE III – EL RETORN
  ══════════════════════════════════ */

  scene_24: {
    id: 'scene_24', etapa: 'El Camí de Retorn', personatge: 'miquel',
    tipus: 'text_block', titol: 'La tensió entre Jack i la Neus',
    imatge: 'assets/img_01.svg',
    narracio: `Dues setmanes después. L'ambient al centre ha millorat notablement gràcies a les noves dinàmiques de Harry. Però a la sala d'entrenadors hi ha una tensió creixent: en Jack i la **Neus** (entrenadora de classes col·lectives) porten tres dies sense parlar-se.

En Jack va criticar públicament les classes de ioga de la Neus davant dels clients. La Neus va respondre amb silenci i evitació total. El director Miquel crida Harry: "Tu has demostrat que saps gestionar persones. Pots intervenir com a mediador informal?"`,
    dialeg: { personatge: 'miquel', text: '"No et demano que soluciones tot sol. Però de vegades un company que intervé amb bona intenció pot desbloquejar el que els protagonistes ja no veuen."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Mediació informal en entorns laborals',
      text: 'La mediació és un procés on una **tercera persona neutral** facilita la comunicació per ajudar les parts a arribar a un acord. Requereix: imparcialitat real, intenció declarada ("vull ajudar, no jutjar"), escoltar les dues parts per separat primer, i no imposar solucions.'
    },
    seguent: 'scene_25'
  },

  scene_25: {
    id: 'scene_25', etapa: 'El Camí de Retorn', personatge: 'miquel',
    tipus: 'decision_scenario', titol: 'Mediar o no mediar?',
    imatge: 'assets/img_01.svg',
    narracio: `Harry té l'oportunitat d'intervenir. Però intervenint s'arrisca a equivocar-se. No intervenint, permet que el conflicte segueixi el seu curs.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: 'Intervenir com a mediador: parlar primer per separat amb Jack i la Neus, i facilitar un diàleg entre ells',
        feedback: '**Decisió madura i professional.** Parlar per separat primer permet que cada part es senti escoltada sense la pressió de l\'altra. Quan les persones se senten escoltades, baixen les defenses i estan més disposades al diàleg. Aquesta és la base de la mediació efectiva.',
        punts: 10, seguent: 'scene_26'
      },
      {
        id: 'B', text: 'No intervenir i deixar que Miquel ho gestioni directament',
        feedback: 'En conflictes interpersonals que afecten l\'entorn laboral, no intervenir quan pots ajudar és una oportunitat perduda. Els protagonistes han demostrat (3 dies de silenci) que no poden resoldre-ho sols. Veuràs el que passa quan el conflicte escala.',
        punts: 0, seguent: 'scene_26b'
      }
    ]
  },

  scene_26: {
    id: 'scene_26', etapa: 'El Camí de Retorn', personatge: 'harry',
    tipus: 'text_block', titol: 'La mediació: el procés',
    imatge: 'assets/img_05.svg',
    narracio: `Harry parla primer a soles amb en Jack. Escolta sense jutjar, reformula, i li pregunta: "Entens per qué la Neus se sent atacada?" En Jack, sorprès, diu que no havia pensat en com l'afectaria.

Llavors parla amb la Neus. Li explica que en Jack no era conscient de l'impacte de les seves paraules. Li demana: "Estairies disposada a tenir una conversa estructurada on tots dos puguin expressar el que necessiten?" La Neus accepta.`,
    dialeg: { personatge: 'harry', text: '"He après que la mediació no és escoltar qui té raó. És ajudar totes dues parts a entendre per qué actuen com actuen."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Presa de decisions amb criteri: quan és millor intervenir',
      text: 'La presa de decisions en situacions de conflicte aliè requereix avaluar: (1) L\'**impacte** del conflicte en l\'entorn, (2) Les **capacitats** pròpies per intervenir de forma neutral, (3) El **moment** oportú, i (4) El **consentiment** de les parts. Intervenir sense consentiment és intrusió; no intervenir quan cal és negligència.'
    },
    seguent: 'scene_27'
  },

  scene_26b: {
    id: 'scene_26b', etapa: 'El Camí de Retorn', personatge: 'narracio',
    tipus: 'text_block', titol: 'El conflicte escala',
    imatge: 'assets/img_01.svg',
    narracio: `Harry decideix no intervenir. Una setmana después, durant una sessió que les dues zones comparteixen, en Jack puja el volum de la música al màxim mentre la Neus explica un exercici de respiració. La Neus perd la compostura i li diu en veu alta que és un maleducat.

Tres clients ho veuen. Dos demanen parlar amb el director. Miquel mira Harry: "Quan et vaig dir que tenies l'oportunitat d'intervenir, ho deia en serio."`,
    dialeg: { personatge: 'miquel', text: '"El silenci davant d\'un conflicte no és innocència. Qui veu un problema i no fa res quan pot fer-ho és part del problema."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'L\'evitació i les seves conseqüències',
      text: 'Els conflictes no resolts no desapareixen: **s\'acumulen i escalen**. Les conseqüències: deteriorament del clima laboral, reducció de la productivitat, efectes sobre clients i necessitat de gestió correctiva molt més costosa.'
    },
    seguent: 'scene_27'
  },

  scene_27: {
    id: 'scene_27', etapa: 'La Resurrecció', personatge: 'miquel',
    tipus: 'text_block', titol: 'El director demana la opinió sobre Jack',
    imatge: 'assets/img_05.svg',
    narracio: `Una setmana después. Les enquestes de satisfacció revelen que tres clients esmenten "un cert entrenador" que els ha fet sentir menyspreats. Miquel sap que és en Jack.

Cita Harry a soles al despatx: "He de prendre una decisió difícil sobre en Jack. Pots donar-me la teva opinió professional i honesta? No la que creus que vull sentir."`,
    dialeg: { personatge: 'miquel', text: '"Et demano l\'opinió d\'algú que ha vist de prop com en Jack treballa. No acusar, no defensar: informació real i respectuosa per prendre una decisió justa."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Donar una opinió professional difícil',
      text: 'Una opinió professional sobre un col·lega ha de: (1) Basar-se en **fets observables**, no en judicis personals. (2) **Separar persona i comportament** ("el comportament X ha generat l\'efecte Y"). (3) Reconèixer el context. (4) Proposar vies constructives quan sigui possible.'
    },
    seguent: 'scene_28'
  },

  scene_28: {
    id: 'scene_28', etapa: 'La Resurrecció', personatge: 'miquel',
    tipus: 'decision_scenario', titol: 'Quin estil comunicatiu usa Harry?',
    imatge: 'assets/img_05.svg',
    narracio: `Harry té la paraula. Miquel espera la seva opinió honesta sobre en Jack.`,
    dialeg: null,
    contingut_pedagogic: null,
    opcions: [
      {
        id: 'A', text: '"En Jack és un mal professional. S\'hauria d\'haver anat fa temps. Els clients no el suporten."',
        feedback: 'Resposta **agressiva** basada en judicis de valor. "Mal professional" és una etiqueta global que no aporta informació útil ni respecta la dignitat d\'en Jack. Una opinió assertiva separa fets de conclusions globals.',
        punts: 0, seguent: 'scene_29'
      },
      {
        id: 'B', text: '"No ho sé, Miquel. Crec que no sóc la persona adequada per opinar sobre un col·lega."',
        feedback: 'Resposta **passiva** que evita la responsabilitat. Miquel ha demanat explícitament la teva opinió com a professional que ha observat en Jack de prop. Evitar opinar quan tens informació rellevant i ets preguntat directament no ajuda la presa de decisions.',
        punts: 5, seguent: 'scene_29'
      },
      {
        id: 'C', text: '"El que he observat és que en Jack usa un estil que tendeix a invalidar els clients quan expressen dificultats. He vist tres situacions concretes. Els seus coneixements tècnics són sòlids. La pregunta és si pot canviar l\'estil comunicatiu amb la formació adequada."',
        feedback: '**Assertivitat avançada exemplar.** Harry es basa en fets observables, reconeix allò positiu, separa persona i comportament, i proposa una via constructiva. Expressar la veritat de forma respectuosa, útil i constructiva és el màxim exponent de l\'assertivitat professional.',
        punts: 10, seguent: 'scene_29'
      }
    ]
  },

  scene_29: {
    id: 'scene_29', etapa: 'El Retorn amb l\'Elixir', personatge: 'harry',
    tipus: 'checklist', titol: 'El decàleg de comunicació de FitCore',
    imatge: 'assets/img_04.svg',
    narracio: `Tres setmanes más tard. En Jack ha acceptat formació en habilitats comunicatives. La Neus i en Jack han tingut una conversa mediada que ha estat difícil però productiva. El director proposa que Harry dirigeixi reunions mensuals de comunicació interna.

Harry prepara el **decàleg de bones pràctiques comunicatives de FitCore**. Marca tots els principis que creus que haurien de formar-ne part.`,
    dialeg: { personatge: 'harry', text: '"Quan vaig arribar pensava que la meva feina era fer programes d\'entrenament. Ara sé que la feina real és crear relacions de confiança."' },
    contingut_pedagogic: {
      tipus: 'idea_clau', titol: 'Les sis competències comunicatives de l\'entrenador professional',
      text: '**Comunicació interpersonal** · **Coherència verbal i no verbal** · **Gestió emocional** · **Empatia professional** · **Assertivitat** · **Resolució constructiva de conflictes**. Aquestes sis competències no s\'ensenyen als clients: s\'ensenyen als entrenadors.'
    },
    checklistItems: [
      { id: 'dc1', text: 'Escoltar activament: deixar parlar, reformular i confirmar que hem entès', correcta: true },
      { id: 'dc2', text: 'Comunicar els canvis de programa als clients amb antelació i explicant el motiu', correcta: true },
      { id: 'dc3', text: 'Gestionar les crítiques públiques amb calma i proposar un espai privat', correcta: true },
      { id: 'dc4', text: 'Expressar opinions professionals de forma assertiva: fets, no judicis', correcta: true },
      { id: 'dc5', text: 'Reconèixer i gestionar les pròpies emocions abans de comunicar sota tensió', correcta: true },
      { id: 'dc6', text: 'Crear espais segurs per als conflictes interns: mediació abans d\'escalada', correcta: true },
      { id: 'dc7', text: 'Mantenir coherència entre el que diem i com ho diem', correcta: true },
      { id: 'dc8', text: 'Adaptar el canal de comunicació (presencial, escrit) a la situació i al client', correcta: true },
      { id: 'dc9', text: 'Donar feedback constructiu als clients: específic, orientat a la millora i respectuós', correcta: true },
      { id: 'dc10', text: 'Demanar i acceptar feedback del client sobre la nostra pràctica professional', correcta: true },
      { id: 'dc11', text: 'Usar la comunicació agressiva quan cal posar límits ferms als clients difícils', correcta: false },
      { id: 'dc12', text: 'Evitar sempre els conflictes per mantenir un ambient positiu al centre', correcta: false },
      { id: 'dc13', text: 'No demanar mai consell a col·legues per no demostrar inseguretat professional', correcta: false }
    ],
    punts_per_item: 1,
    seguent: 'scene_30'
  },

  /* ══════════════════════════════════
     EPÍLEG
  ══════════════════════════════════ */

  scene_30: {
    id: 'scene_30', etapa: 'Epíleg: El Nou Harry', personatge: 'hanna',
    tipus: 'epilogue', titol: 'El Retorn amb l\'Elixir',
    imatge: 'assets/img_01.svg',
    narracio: `Sis setmanes des del primer dia. FitCore, sala de pesos, set del matí.

La Hanna li ha deixat una nota al taquiller: "L'heroi no és el qui no cau. És el qui aprèn cada vegada que es reincorpora."`,
    dialeg: { personatge: 'hanna', text: '"Ja no ets el Harry del primer dia. Ets algú que ha après que la comunicació és una habilitat, no un talent. Les habilitats es practiquen cada dia. Benvingut al club dels professionals de veritat."' },
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
    this.state.checklistData[sceneId] = { selectedIds: selectedIds };
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

    document.getElementById('stage-label').textContent = scene.etapa || '';
    document.getElementById('scene-title').textContent = scene.titol || '';
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
      var hasVisited   = stage.scenes.some(function (s) { return visited.indexOf(s) !== -1; });
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
      status.textContent = hasCompleted ? '✓ Completada' : hasVisited ? '● En curs' : '○ Pendent';

      div.appendChild(badge);
      div.appendChild(title);
      div.appendChild(status);
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
    document.getElementById('pedagogic-text').innerHTML = this._md(scene.contingut_pedagogic.text || '');
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
    var self = this;
    var timerSecs = 15;

    // Timer display
    var timerDiv = document.createElement('div');
    timerDiv.className = 'timer-container';
    var timerTrack = document.createElement('div');
    timerTrack.className = 'timer-bar-track';
    var timerFill = document.createElement('div');
    timerFill.className = 'timer-bar-fill';
    timerTrack.appendChild(timerFill);
    var timerLabel = document.createElement('span');
    timerLabel.id = 'timer-label';
    timerLabel.textContent = timerSecs + 's';
    timerDiv.appendChild(timerTrack);
    timerDiv.appendChild(timerLabel);
    area.appendChild(timerDiv);

    var btn = document.createElement('button');
    btn.className = 'btn btn-primary btn-disabled';
    btn.textContent = 'Continua →';
    btn.addEventListener('click', function () {
      if (!btn.classList.contains('btn-disabled')) Engine.goToScene(scene.seguent);
    });
    area.appendChild(btn);

    // Countdown
    var remaining = timerSecs;
    var handle = setInterval(function () {
      remaining--;
      var pct = (remaining / timerSecs) * 100;
      timerFill.style.width = pct + '%';
      timerFill.style.transition = 'width 1s linear';
      if (timerLabel) timerLabel.textContent = remaining > 0 ? remaining + 's' : '';
      if (remaining <= 0) {
        clearInterval(handle);
        btn.classList.remove('btn-disabled');
        btn.classList.add('btn-enabled');
        if (timerDiv.parentNode) timerDiv.style.display = 'none';
      }
    }, 1000);
    this._timerHandle = handle;
  },

  _renderDecision: function (scene, state, area) {
    var self = this;
    var prev = state.decisions[scene.id];

    var label = document.createElement('p');
    label.className = 'decision-label';
    label.textContent = 'Pren una decisió:';
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
        nextBtn.textContent = 'Continua →';
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

      if (prev) {
        btn.classList.add('btn-disabled');
        if (op.correcta) btn.classList.add('btn-quiz-correct');
        else if (op.id === prev.optionId) btn.classList.add('btn-quiz-wrong');
      }

      btn.addEventListener('click', function () {
        if (prev) return;
        area.querySelectorAll('.btn-quiz').forEach(function (b) { b.classList.add('btn-disabled'); });
        Engine.recordQuiz(scene.id, op.id);

        // Mark correct/wrong
        scene.opcions.forEach(function (o) {
          var b = area.querySelector('[data-qid="' + o.id + '"]');
          if (b) {
            if (o.correcta) b.classList.add('btn-quiz-correct');
            else if (o.id === op.id) b.classList.add('btn-quiz-wrong');
          }
        });

        var fbDiv = document.createElement('div');
        fbDiv.className = 'quiz-feedback ' + (op.correcta ? 'qf-correct' : 'qf-wrong');
        fbDiv.innerHTML = (op.correcta ? '<strong>✓ Correcte!</strong> ' : '<strong>✗ No exactament.</strong> ') +
          self._md(op.feedback);
        area.appendChild(fbDiv);

        var nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-primary btn-enabled';
        nextBtn.style.marginTop = '4px';
        nextBtn.textContent = 'Continua →';
        nextBtn.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
        area.appendChild(nextBtn);
      });

      btn.dataset.qid = op.id;
      area.appendChild(btn);
    });

    if (prev) {
      var prevOp = scene.opcions.find(function (o) { return o.id === prev.optionId; });
      if (prevOp) {
        var fbDiv = document.createElement('div');
        fbDiv.className = 'quiz-feedback ' + (prevOp.correcta ? 'qf-correct' : 'qf-wrong');
        fbDiv.innerHTML = (prevOp.correcta ? '<strong>✓ Correcte!</strong> ' : '<strong>✗ No exactament.</strong> ') +
          self._md(prevOp.feedback);
        area.appendChild(fbDiv);
      }
      var nextBtn = document.createElement('button');
      nextBtn.className = 'btn btn-primary btn-enabled';
      nextBtn.textContent = 'Continua →';
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
    label.textContent = 'Marca els elements importants:';
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

      div.appendChild(icon);
      div.appendChild(txt);

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
        if (item.correcta && was)       d.classList.add('cl-correct');
        else if (!item.correcta && was) d.classList.add('cl-incorrect');
        else if (item.correcta && !was) d.classList.add('cl-missed');
      });
    };

    if (!prev) {
      var confirmBtn = document.createElement('button');
      confirmBtn.className = 'btn btn-primary btn-enabled';
      confirmBtn.textContent = 'Confirmar selecció';
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
        fbDiv.innerHTML = '<strong>' + punts + ' / ' + maxPts + ' punts.</strong> ' +
          '<span style="color:var(--good)">■ Verd = correcte i marcat</span>  ' +
          '<span style="color:var(--bad)">■ Vermell = marcat incorrecte</span>  ' +
          '<span style="color:var(--ok)">■ Taronja = correcte no marcat</span>';
        area.appendChild(fbDiv);

        var nextBtn = document.createElement('button');
        nextBtn.className = 'btn btn-primary btn-enabled';
        nextBtn.textContent = 'Continua →';
        nextBtn.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
        area.appendChild(nextBtn);
      });
      area.appendChild(confirmBtn);
    } else {
      _applyResults();
      var nextBtn2 = document.createElement('button');
      nextBtn2.className = 'btn btn-primary btn-enabled';
      nextBtn2.textContent = 'Continua →';
      nextBtn2.addEventListener('click', function () { Engine.goToScene(scene.seguent); });
      area.appendChild(nextBtn2);
    }
  },

  _renderEpilogue: function (scene, state, area) {
    var score = state.score;
    var pct   = Math.round((score / 100) * 100);

    var hannaMsg, hannaClass;
    if (score >= 85) {
      hannaMsg = '"Has demostrat una comprensió excepcional. FitCore té sort de tenir-te. Continua creixent: cada client és un nou viatge."';
      hannaClass = 'epilogue-excellent';
    } else if (score >= 60) {
      hannaMsg = '"Has après molt en poc temps. La base és sòlida. Les habilitats comunicatives necessiten pràctica diària: cada interacció és una oportunitat d\'aprenentatge."';
      hannaClass = 'epilogue-good';
    } else if (score >= 35) {
      hannaMsg = '"Has donat els primers passos. Les habilitats comunicatives son com l\'entrenament físic: requereixen constància. Revisa les escenes on has tingut dificultats."';
      hannaClass = 'epilogue-ok';
    } else {
      hannaMsg = '"El camí de l\'heroi sovint comença amb errors. El que importa no és on comences, sinó la direcció en què vas. Revisa el viatge i reflexiona sobre cada decisió."';
      hannaClass = 'epilogue-needs-work';
    }

    // Decisions summary
    var dsHtml = '<ul class="decisions-summary">';
    var scored = ['scene_05','scene_10','scene_13','scene_16','scene_18','scene_19','scene_20','scene_25','scene_28','scene_29'];
    scored.forEach(function (sid) {
      var sc = scenes[sid];
      var dec = sid === 'scene_16' || sid === 'scene_29' ? state.checklistData[sid] : state.decisions[sid];
      if (!sc || !dec) return;
      var pts = dec.punts !== undefined ? dec.punts : 0;
      var label = sc.titol;
      var icon = pts >= 10 ? '✓' : pts >= 5 ? '⚠' : '✗';
      var cls  = pts >= 10 ? 'ds-good' : pts >= 5 ? 'ds-ok' : 'ds-bad';
      dsHtml += '<li class="' + cls + '"><span class="ds-icon">' + icon + '</span><span>' + label + '</span><span class="ds-pts">' + pts + ' pts</span></li>';
    });
    dsHtml += '</ul>';

    var skillsHtml = '<ul class="skills-list">' +
      '<li>✓ Comunicació interpersonal i procés comunicatiu</li>' +
      '<li>✓ Comunicació verbal i no verbal: coherència i impacte</li>' +
      '<li>✓ Educació emocional: regulació en situacions de tensió</li>' +
      '<li>✓ Empatia professional en contextos de fitness</li>' +
      '<li>✓ Assertivitat: expressar-se amb respecte i claredat</li>' +
      '<li>✓ Resolució de conflictes i estils d\'afrontament</li>' +
      '</ul>';

    // SVG ring
    var circ = 339.292;
    var dash = circ * pct / 100;

    var container = document.createElement('div');
    container.className = 'epilogue-container';
    container.innerHTML =
      '<div class="epilogue-score-ring">' +
        '<svg viewBox="0 0 120 120" class="score-ring-svg">' +
          '<circle cx="60" cy="60" r="54" fill="none" stroke="#2C2C2C" stroke-width="10"/>' +
          '<circle cx="60" cy="60" r="54" fill="none" stroke="#FF6B35" stroke-width="10" ' +
            'stroke-dasharray="' + dash + ' ' + circ + '" stroke-dashoffset="84.823" stroke-linecap="round"/>' +
        '</svg>' +
        '<div class="score-ring-text">' +
          '<span class="score-big">' + score + '</span>' +
          '<span class="score-max">/100</span>' +
        '</div>' +
      '</div>' +
      '<div class="epilogue-hanna ' + hannaClass + '">' +
        '<div class="epilogue-hanna-avatar" style="background:#27AE60">Ha</div>' +
        '<div class="epilogue-hanna-msg">' + hannaMsg + '</div>' +
      '</div>' +
      '<p class="epilogue-section-title">Resum de les teves decisions clau</p>' +
      dsHtml +
      '<p class="epilogue-section-title">Habilitats treballades</p>' +
      skillsHtml +
      '<div class="epilogue-buttons"><button id="finish-btn" class="btn btn-finish">Finalitzar i enviar puntuació al LMS</button></div>';

    area.appendChild(container);

    // Attach finish listener properly (no onclick attr)
    document.getElementById('finish-btn').addEventListener('click', function () {
      Engine.finish();
      this.disabled = true;
      this.textContent = '✓ Puntuació enviada';
    });
  },

  _showFeedback: function (op, onContinue) {
    var icon = op.punts === 10 ? '✓' : op.punts === 5 ? '⚠' : '✗';
    var cls  = op.punts === 10 ? 'fb-good' : op.punts === 5 ? 'fb-ok' : 'fb-bad';
    document.getElementById('feedback-icon').textContent = icon;
    document.getElementById('feedback-icon').className = 'feedback-icon ' + cls;
    document.getElementById('feedback-pts').textContent = '+' + op.punts + ' punts';
    document.getElementById('feedback-text').innerHTML = this._md(op.feedback);
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
      // Ordered list: "1. " or "**1. " or "**1) " prefixes
      if (/^(\*\*)?[1-9]\d*[.)]\s/.test(first)) {
        return '<ol class="step-list">' +
          lines.filter(Boolean).map(function (l) {
            return '<li>' + inl(l.trim()) + '</li>';
          }).join('') + '</ol>';
      }
      // Unordered bullet: "- " or "• "
      if (/^[-•]\s/.test(first)) {
        return '<ul>' +
          lines.filter(Boolean).map(function (l) {
            return '<li>' + inl(l.trim().replace(/^[-•]\s*/, '')) + '</li>';
          }).join('') + '</ul>';
      }
      // Definition list: multiple lines each starting with **Label**
      if (lines.length > 1 && /^\*\*[^*]/.test(first)) {
        return '<ul class="def-list">' +
          lines.filter(Boolean).map(function (l) {
            return '<li>' + inl(l.trim()) + '</li>';
          }).join('') + '</ul>';
      }
      return '<p>' + inl(para) + '</p>';
    }).join('');
  }
};

// Expose Engine on window so inline handlers and console can access it
window.Engine = Engine;

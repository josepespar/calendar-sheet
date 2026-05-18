/* ============================================================
   SCENARIO DATA – El Viatge de l'Heroi · FitCore
   Sistema de puntuació màxim: 100 punts
   Decisions puntuades: 10 escenes × 10 pts màxim = 100
   ============================================================ */

const CHARACTERS = {
  harry: {
    name: 'Harry',
    role: 'Protagonista',
    color: '#4A90D9',
    initials: 'H',
    shape: 'circle'
  },
  hanna: {
    name: 'Hanna',
    role: 'Mentora',
    color: '#27AE60',
    initials: 'Ha',
    shape: 'circle'
  },
  jack: {
    name: 'Jack',
    role: 'Antagonista',
    color: '#C0392B',
    initials: 'J',
    shape: 'circle'
  },
  miquel: {
    name: 'Miquel',
    role: 'Director',
    color: '#8E44AD',
    initials: 'M',
    shape: 'circle'
  },
  jordi: {
    name: 'Jordi',
    role: 'Client avançat',
    color: '#E67E22',
    initials: 'Jo',
    shape: 'circle'
  },
  narracio: {
    name: 'Narrador',
    role: '',
    color: '#7F8C8D',
    initials: 'N',
    shape: 'square'
  }
};

/* ============================================================
   JOURNEY MAP – etapes del viatge
   ============================================================ */
const JOURNEY_STAGES = [
  { id: 'act1',   label: 'Acte I',   title: 'El Món Ordinari',          scenes: ['scene_01','scene_02','scene_03','scene_04'], color: '#4A90D9' },
  { id: 'act2a',  label: 'Acte II',  title: 'El Món Especial (Entrada)', scenes: ['scene_05','scene_05b','scene_06'],           color: '#FF6B35' },
  { id: 'act2b',  label: 'Acte II',  title: 'El Món Especial (Prova)',   scenes: ['scene_07','scene_08'],                       color: '#FF6B35' },
  { id: 'act2c',  label: 'Acte II',  title: 'El Món Especial (Recomp.)',  scenes: ['scene_09'],                                  color: '#E67E22' },
  { id: 'act3',   label: 'Acte III', title: 'El Retorn',                 scenes: ['scene_10','scene_10b','scene_11','scene_12'], color: '#27AE60' },
  { id: 'final',  label: 'Final',    title: 'Reflexió Final',            scenes: ['scene_13'],                                  color: '#8E44AD' }
];

/* ============================================================
   SCENES DATA
   ============================================================ */
const scenes = {

  /* ─────────────────────────────────────────
     ACTE I – EL MÓN ORDINARI
  ───────────────────────────────────────── */

  scene_01: {
    id: 'scene_01',
    etapa: 'El Món Ordinari',
    personatge: 'harry',
    tipus: 'text_block',
    titol: 'El primer dia a FitCore',
    narracio: `El sol de primera hora de la tarda entra a raig per les finestres del vestíbul de FitCore. La música electrònica ressona suaument entre les màquines cardio i les barres d'halterofília. Harry té 18 anys, porta uns lluentons al nas dels nervis i una motxilla nova on porta el bloc de notes que va omplir de diagrames de periodització durant tota la setmana passada.

Avui és el seu primer dia com a entrenador de condicionament físic. Ha obtingut el títol de tècnic esportiu el mes passat, però el que ningú li ha ensenyat al curs és com parlar amb la gent quan les coses es compliquen.

El centre FitCore té quatre espais: la sala de pesos, la zona cardio, les sales de classes col·lectives i la piscina coberta. Cada racó té el seu propi ecosistema de clients, rutines i expectatives. Harry s'adona que un gimnàs no és tan diferent d'una escola: tothom vol ser vist, escol·tat i respectat.

La Hanna, entrenadora des de fa set anys, li estreny la mà amb un somriure franc. Té una forma de parlar que fa sentir les persones còmodes des del primer segon. A continuació, s'apropa en Jack, veterà de deu anys al sector. La seva salutació és breu, quasi un acte de cortesia obligatòria.`,
    dialeg: {
      personatge: 'harry',
      text: '"Bé, estic aquí. Sé el que he après als llibres. Però la gent real és diferent als exemples dels exàmens. Avui comença la part que de veritat importa."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Les habilitats socials en el fitness professional',
      text: 'Un entrenador de condicionament físic no és només un expert en exercici físic: és un comunicador, un gestor emocional i un mediador en el context del benestar. La comunicació és la competència professional que determina si els clients progressen, si tornen i si confien en tu. Sense habilitats socials sòlides, el millor programa d\'entrenament del món fracassa.'
    },
    seguent: 'scene_02'
  },

  /* ─────────── */

  scene_02: {
    id: 'scene_02',
    etapa: 'La Crida a l\'Aventura',
    personatge: 'miquel',
    tipus: 'decision_scenario',
    titol: 'La queixa de la Maria',
    narracio: `Només dues hores després de l'arribada, el director Miquel crida Harry al seu despatx. El despatx és modest: una taula amb piles de carpetes i una pantalla on s'obren gràfics de renovació de quotes. Miquel és un home de quaranta-cinc anys que parla ràpid i espera respostes encara més ràpides.

La Maria Puig té 55 anys i porta tres mesos al centre. Va venir amb objectius clars: millorar la mobilitat, reduir el dolor de genolls i perdre uns quants quilos sense castigar el cos. Però avui ha deixat una queixa formal a recepció: diu que ningú l'escol·ta, que els entrenadors li expliquen exercicis complicats sense preguntar-li com se sent, i que s'ha sentit invisible.

Miquel mira Harry amb una barreja d'urgència i prova. "Sé que acabes de començar, però necessito que algú gestioni això ara. La Maria és al saló d'espera. Tu, jo, o en Jack podem parlar amb ella." Harry nota la pressió al pit. Aquesta és la primera prova real.`,
    dialeg: {
      personatge: 'miquel',
      text: '"Harry, la Maria és una client valuosa i ara mateix se sent com si no existís per a nosaltres. Necessito que quelcom canviï avui. Tu pots gestionar-ho? Tinc una reunió en deu minuts."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'El procés comunicatiu i les seves barreres',
      text: 'La comunicació no és simplement parlar: és un procés de doble via on l\'emissor codifica un missatge, el receptor el descodifica i es produeix feedback. Les barreres més freqüents en entorns de fitness són: el soroll físic (música alta, màquines), el soroll psicològic (suposicions sobre el client) i les barreres semàntiques (terminologia tècnica incomprensible). Quan un client diu "ningú m\'escol·ta", sovint vol dir que el canal o el codi no és l\'adequat.'
    },
    opcions: [
      {
        id: 'A',
        text: 'Acceptar la tasca i demanar consell a la Hanna abans de parlar amb la Maria',
        feedback: 'Molt bona decisió. Buscar orientació d\'una professional experimentada com la Hanna és un signe de maduresa professional, no de feblesa. Aquesta conducta reflecteix **escolta activa proactiva**: avant de comunicar-te amb el client, t\'assegures de tenir les eines adequades. La preparació és el primer facilitador de la comunicació efectiva.',
        punts: 10,
        seguent: 'scene_04'
      },
      {
        id: 'B',
        text: 'Acceptar la tasca i anar directament a parlar amb la Maria sense preparació prèvia',
        feedback: 'Acceptar la responsabilitat és positiu, però anar sense preparació pot agreujar la situació. La Maria ja se sent invisible; si la conversa no va bé, reforçarà la seva percepció negativa. Qualsevol conversa difícil requereix un mínim de preparació: definir l\'objectiu, anticipar les emocions de l\'altre i tenir clars els missatges clau.',
        punts: 5,
        seguent: 'scene_04'
      },
      {
        id: 'C',
        text: 'Derivar el cas a Jack, que té més experiència',
        feedback: 'Derivar sense acompanyament és una **barrera comunicativa organitzacional**: el client es queda sense resposta directa i pot interpretar que el problema no és prioritari. A més, delegar un conflicte sense supervisió pot agravar-lo si l\'altra persona no té l\'enfocament adequat. Veuràs el que passa a continuació.',
        punts: 0,
        seguent: 'scene_03'
      }
    ]
  },

  /* ─────────── */

  scene_03: {
    id: 'scene_03',
    etapa: 'Rebuig de la Crida',
    personatge: 'jack',
    tipus: 'text_block',
    titol: 'Quan Jack gestiona la Maria',
    narracio: `Harry observa des de l'altre costat del passadís. Jack s'asseu enfront de la Maria al saló d'espera amb els braços creuats i les cames ben obertes, en postura de domini. Comença a parlar sense haver preguntat res.

"Mira, Maria, tots els entrenadors aquí sabem el que fem. Si et donava exercicis és perquè eren els correctes. Potser el problema és que t'has d'esforçar una mica més." La Maria encongeix els muscles i mira cap a terra. Jack continua: "Si vols resultats ràpids has de confiar en el procés. Ara bé, si no n'estàs satisfeta, pots parlar amb direcció." S'aixeca i se'n va.

La Maria queda sola al saló d'espera amb la bossa a la falda i els ulls humits. Quan Harry s'apropa, ella diu en veu baixa: "No importa. Ja entenc que aquí no sóc benvinguda." Harry sent que s'li cau l'estómac.

Aquesta és la conseqüència d'una comunicació sense empatia: el client no se sent escoltat, sinó jutjat i descartat. Jack ha tingut raó en el contingut tècnic, però ha fracassat completament en la forma.`,
    dialeg: {
      personatge: 'narracio',
      text: '"Jack ha demostrat tres barreres comunicatives en menys de dos minuts: postura de domini (comunicació no verbal agressiva), interrupcions implícites (parlar sense escoltar), i invalidació emocional (minimitzar la queixa del client). La Maria no ha rebut resposta: ha rebut una sentència."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Barreres de la comunicació: quan el missatge destrueix el canal',
      text: 'Les barreres comunicatives actives inclouen: (1) Judicis prematurs: concloure abans d\'escoltar. (2) Invalidació emocional: negar o minimitzar el que l\'altra persona sent. (3) Comunicació no verbal incongruent: postura tancada mentre es diuen paraules de "suport". (4) Absència de feedback positiu: no confirmar que s\'ha rebut i comprès el missatge. Quan apareix una o més d\'aquestes barreres, el client no se sent respectat, sinó processat.'
    },
    seguent: 'scene_04'
  },

  /* ─────────── */

  scene_04: {
    id: 'scene_04',
    etapa: 'La Trobada amb el Mentor',
    personatge: 'hanna',
    tipus: 'worked_example',
    titol: 'La lliçó de la Hanna',
    narracio: `La Hanna porta Harry a la sala d'entrenadors, un espai petit però tranquil al fons del passadís. Li prepara un cafè, s'asseu enfront d'ell i comença a parlar amb una veu clara i pausada. No hi ha pressa. La Hanna sap que les coses importants requereixen temps.

"Harry, cada queixa d'un client és un regal. Ens diu exactament on falla la comunicació. La Maria no es queixa dels exercicis: es queixa de no sentir-se vista. I la diferència entre un entrenador mediocre i un de bo no és el programa d'entrenament, és si el client confia en tu."

La Hanna li explica la diferència entre escoltar i sentir. Escoltar és un acte passiu: les paraules entren i surten. L'escolta activa és un acte deliberat: poses l'atenció al cent per cent, confirmes que has entès, i respons al missatge emocional, no només al contingut.

Hanna modela una conversa amb la Maria de manera role-play. Primer s'asseu al costat de Harry (no enfront, per reduir la sensació de confrontació), manté el contacte visual sense pressionar, capeja el cap per confirmar, i quan la Maria acaba de parlar, reformula: "El que m'estàs dient és que sents que no ens hem pres el temps per entendre el que tu necessites. T'escolto. Explica'm."`,
    dialeg: {
      personatge: 'hanna',
      text: '"La comunicació efectiva no és el que dius: és com ho dius, quan ho dius, i sobretot, quant temps i espai dones a l\'altra persona per dir el que necessita. Un entrenador que no sap escoltar no pot fer progressar ningú."'
    },
    contingut_pedagogic: {
      tipus: 'worked_example',
      titol: 'Exemple modelat: com iniciar una conversa difícil amb un client',
      text: `**Pas 1 – Crear l'espai:** Busca un lloc tranquil, asseu-te al costat o en angle (no enfront), treu el mòbil del camp visual.

**Pas 2 – Obertura empàtica:** "Hola [nom], gràcies per dir-nos el que sents. Vull entendre la teva experiència." Evita defensar-te immediatament.

**Pas 3 – Escolta activa:** Deixa parlar sense interrompre. Pren notes mentals. Capeja el cap. Manté contacte visual suau.

**Pas 4 – Reformulació:** "El que m'estàs dient és [resum del que ha dit]... ho he entès bé?" Dona espai per corregir-te.

**Pas 5 – Proposta conjunta:** "Què podríem fer diferent perquè et sentis millor acompanyada?" Fa participar el client en la solució.`
    },
    seguent: 'scene_05'
  },

  /* ─────────────────────────────────────────
     ACTE II – EL MÓN ESPECIAL (ENTRADA)
  ───────────────────────────────────────── */

  scene_05: {
    id: 'scene_05',
    etapa: 'Creuament del Primer Llindar',
    personatge: 'jack',
    tipus: 'decision_scenario',
    titol: 'La primera sessió de grup',
    narracio: `Tres dies después. Harry ha d'iniciar la primera sessió de grup com a entrenador principal. El grup és variat i, francament, intimidant: la Maria (55 anys, mobilitat limitada, molt sensible a la crítica), en Jordi (35 anys, triatleta avançat, molt exigent amb si mateix i amb els altres) i la Lluïsa (65 anys, jubilada, ve per socialitzar tant com per fer exercici, però té una cardiopatia lleu que cal tenir en compte).

Mentre Harry repassa els fulls de cada client a la sala d'entrenadors, en Jack entra i li dóna un cop a l'espatlla. "Consell d'amic, nano: sigues ferm des del principi. No els deixis que et dominin amb excuses. La gent del gimnàs prova fins on pot estirar. Si cedeixes un centímetre el primer dia, et mengen sencer."

Harry recorda la conversa amb la Hanna. Les dues perspectives no podrien ser més diferents. La porta de la sala s'obre: la Maria, en Jordi i la Lluïsa esperen.`,
    dialeg: {
      personatge: 'jack',
      text: '"Mira, jo porto deu anys en aquest sector. Sigues ferm. No escoltis massa excuses. Aquí venen a entrenar, no a plorar. I sobretot: que quedi clar qui mana des del primer minut."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Comunicació verbal i no verbal: la coherència és clau',
      text: 'El 55% de l\'impacte comunicatiu prové del llenguatge corporal, el 38% del to de veu, i només el 7% de les paraules (model Mehrabian aplicat a contextos emocionals). Si el teu missatge verbal és "estic aquí per ajudar-te" però el teu to és distant i la teva postura és tancada, el client percebrà la incoherència. La comunicació no verbal incongruent genera desconfiança, no autoritat.'
    },
    opcions: [
      {
        id: 'A',
        text: 'Seguir el consell de Jack: entrar amb fermesa i autoritat, sense preguntar com es troben',
        feedback: 'Aplicar l\'enfocament de Jack genera **incoherència comunicativa**: les paraules diuen "benvinguts" però la postura i el to diuen "obeïu". En Jordi ho interpretarà com un repte d\'autoritat, la Maria se sentirà novament invisible i la Lluïsa es posarà nerviosa. La fermesa sense empatia és autoritat buida.',
        punts: 0,
        seguent: 'scene_05b'
      },
      {
        id: 'B',
        text: 'Aplicar el que ha après de la Hanna: salutació personalitzada, preguntar com es troben i establir els objectius del dia conjuntament',
        feedback: 'Excel·lent. Iniciar amb una **salutació personalitzada** (dir el nom de cada persona) i preguntar com es troben activa l\'escolta activa des del primer moment. Quan el client sent que l\'entrenador s\'interessa per ell com a persona —no com a usuari d\'una màquina— la sessió comença des d\'un lloc de confiança. La comunicació assertiva comença amb respecte i presència.',
        punts: 10,
        seguent: 'scene_06'
      }
    ]
  },

  /* ─────────── */

  scene_05b: {
    id: 'scene_05b',
    etapa: 'Creuament del Primer Llindar',
    personatge: 'narracio',
    tipus: 'text_block',
    titol: 'Les conseqüències de la fermesa sense empatia',
    narracio: `Harry entra a la sala amb una postura recta, to assertiu i sense pauses. "Bé, avui farem el següent protocol." No pregunta noms, no mira als ulls, comença directament amb la demostració dels exercicis. En Jordi creua els braços i observa amb escepticisme. La Lluïsa fa un somriure nerviós i no s'atreveix a preguntar si pot fer l'exercici amb la seva condició cardíaca. La Maria mira el terra.

A mitja sessió, en Jordi interromp: "Espera. Fa tres setmanes m'havies dit que canviaries el programa de sentadilles. Això és el mateix d'ahir." Harry no sap de qui parla —en Jordi és client d'un altre entrenador— i la comunicació es converteix en un malentès públic davant dels altres.

La sessió acaba amb tensions. La Lluïsa se'n va sense dir res. La Maria surt ràpid. En Jordi queda mirant Harry amb una expressió que diu: "haurem de parlar."

La fermesa sense escolta activa ni empatia no és autoritat professional: és distància que genera malentesos i conflictes evitables.`,
    dialeg: {
      personatge: 'narracio',
      text: '"Harry ha après la primera lliçó dura: la comunicació no verbal incongruent i l\'absència d\'escolta activa generen tensió fins i tot quan les paraules usades semblaven correctes. Ara haurà de gestionar les conseqüències."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Comunicació no verbal i el seu impacte emocional',
      text: 'La comunicació no verbal inclou: postura corporal, expressions facials, contacte visual, to de veu, proximitat física i gestos. En un context de fitness, el client interpreta constantment el que l\'entrenador comunica amb el cos. Una postura oberta (braços relaxats, cos orientat cap al client, somriure natural) genera seguretat. Una postura tancada (braços creuats, mirada distant, to autoritari) genera por o resistència.'
    },
    seguent: 'scene_06'
  },

  /* ─────────── */

  scene_06: {
    id: 'scene_06',
    etapa: 'Proves, Aliats i Enemics',
    personatge: 'jordi',
    tipus: 'decision_scenario',
    titol: 'La crítica pública d\'en Jordi',
    narracio: `Dijous de la setmana següent. Harry ha canviat el programa d'en Jordi: ha reduït el volum de sentadilles i ha introduït treball de core específic per al triatlò, una millora objectivament justificada. Però no ho ha comunicat prèviament al client.

En Jordi arriba a la sala de pesos i mira el full de rutina nou. Primer fa una pausa. Després es gira cap a Harry i, en veu prou alta perquè tres persones properes ho puguin sentir, diu: "Això és exactament el problema. Canvies el que funciona sense dir res a ningú. No em consultes, no m'expliques res, i després esperes que confiï en tu?"

Tres clients més giren el cap. Una noia al aparell de cable atura el moviment. Harry nota la calor a la cara. Té tres opcions davant.`,
    dialeg: {
      personatge: 'jordi',
      text: '"Potser el problema ets tu, que creus que perquè tens el títol ja ho saps tot. Jo porto set anys entrenant per a triatlons. Hauries de consultar-me abans de canviar res."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Assertivitat: el dret a expressar-se amb respecte',
      text: 'L\'assertivitat és la capacitat d\'expressar pensaments, sentiments i necessitats de manera directa, honesta i respectuosa, sense vulnerar els drets dels altres. Hi ha tres estils de comunicació: **passiu** (cedir sempre, suprimir les pròpies necessitats), **agressiu** (imposar sense considerar l\'altre) i **assertiu** (expressar amb claredat i respecte). La resposta assertiva en una situació de crítica pública inclou: no escalar, validar el sentiment del client, i proposar un espai privat per parlar amb calma.'
    },
    opcions: [
      {
        id: 'A',
        text: 'Defensar-se públicament explicant que el canvi era tècnicament millor',
        feedback: 'Defensar-se públicament és una **resposta agressiva** que escala el conflicte. En Jordi interpretarà que l\'estàs humiliant davant dels altres, i la resta de clients presenciaran una disputa que erosiona la teva credibilitat professional. Tenir raó tècnica no justifica una confrontació pública.',
        punts: 0,
        seguent: 'scene_07'
      },
      {
        id: 'B',
        text: 'Disculpar-se immediatament i de manera excessiva davant de tothom',
        feedback: 'Disculpar-se en excés és una **resposta passiva** que, tot i intentar desactivar el conflicte, envia el missatge que et responsabilitzes d\'una mala actuació greu quan en realitat la decisió tècnica era correcta. La passivitat pot solucionar el moment però crea un precedent negatiu: el client aprendrà que la crítica pública funciona com a mecanisme de pressió.',
        punts: 5,
        seguent: 'scene_07'
      },
      {
        id: 'C',
        text: 'Validar el sentiment de Jordi públicament i proposar parlar en privat',
        feedback: 'Resposta assertiva exemplar. Validar l\'emoció d\'en Jordi ("Entenc que et senti frustrat, i tens tot el dret a saber per què he canviat el teu programa") i proposar un espai privat desactiva l\'escalada pública, protegeix la dignitat de totes dues parts i obre la porta a una solució real. La **comunicació assertiva** en situacions de tensió comença sempre per reconèixer l\'emoció de l\'altre, no per defensar la pròpia posició.',
        punts: 10,
        seguent: 'scene_07'
      }
    ]
  },

  /* ─────────── */

  scene_07: {
    id: 'scene_07',
    etapa: 'Aproximació a la Cova',
    personatge: 'harry',
    tipus: 'checklist',
    titol: 'Prepara la conversa difícil amb en Jordi',
    narracio: `La Hanna agafa Harry del braç just quan surt de la sala de pesos. "Bé fet per no escalar. Ara ve la part difícil: has de tenir la conversa de veritat amb en Jordi. I per tenir una conversa difícil de manera professional, cal preparar-se."

Harry s'asseu a la sala d'entrenadors amb un full en blanc. La Hanna li diu que hi ha cinc elements clau per preparar qualsevol conversa difícil amb un client o col·lega. Li demana que els marqui tots, però adverteix: "No és una llista mecànica. Cada element importa perquè canvia la qualitat de la conversa."

**Instrucció:** Marca els cinc elements que hauries de preparar abans de la conversa amb en Jordi. Cada element que marquis correctament et donarà punts. Selecciona tots els que creus que són importants.`,
    dialeg: {
      personatge: 'hanna',
      text: '"Una conversa difícil mal preparada pot destruir una relació que portes mesos construint. Però una conversa difícil ben preparada pot convertir un conflicte en la base d\'una confiança molt més sòlida. La preparació és respecte."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Preparació d\'una conversa difícil: els elements essencials',
      text: 'Les converses difícils comparteixen una estructura universal: (1) aclariment de l\'objectiu propi, (2) anticipació de les emocions de l\'altra part, (3) elecció del moment i l\'espai adequats, (4) preparació d\'una obertura empàtica i no acusatòria, (5) escolta activa planificada. Qui prepara la conversa no l\'imposa: la facilita.'
    },
    checklistItems: [
      { id: 'cl1', text: 'Definir quin és el meu objectiu real de la conversa (no guanyar, sinó entendre\'ns)', correcta: true },
      { id: 'cl2', text: 'Triar un moment tranquil i un lloc privat, fora de la zona d\'entrenament', correcta: true },
      { id: 'cl3', text: 'Preparar una obertura empàtica que reconegui la perspectiva d\'en Jordi', correcta: true },
      { id: 'cl4', text: 'Practicar l\'escolta activa: deixar parlar sense interrompre i reformular', correcta: true },
      { id: 'cl5', text: 'Establir un to assertiu: dir el que penso sense atacar ni cedir innecessàriament', correcta: true },
      { id: 'cl6', text: 'Preparar una llista de tots els errors que he comès per disculpar-me per cadascun', correcta: false },
      { id: 'cl7', text: 'Demanar a en Jack que vingui a la conversa com a suport en cas que en Jordi s\'enfadi', correcta: false }
    ],
    punts_per_item: 2,
    seguent: 'scene_08'
  },

  /* ─────────── */

  scene_08: {
    id: 'scene_08',
    etapa: 'La Prova Suprema',
    personatge: 'jordi',
    tipus: 'ordeal',
    titol: 'La conversa amb en Jordi',
    narracio: `La sala de reunions petita del costat del despatx de Miquel. Dues cadires, una taula estreta, la porta tancada. Harry ha convocat en Jordi amb un missatge clar: "Vull explicar-te per què he canviat el teu programa i escoltar el que penses. Demà a les 10?"

En Jordi arriba puntual però tens. Es creua de braços nada més seure. Harry respira fons.

Aquesta és la prova suprema: el moment on tot el que ha après —sobre comunicació, empatia, assertivitat i regulació emocional— s'ha de posar en pràctica de manera integrada. Tres decisions crítiques determinaran com acaba aquesta conversa.`,
    dialeg: {
      personatge: 'narracio',
      text: '"La prova suprema no és física ni tècnica. És emocional. Quan estàs sota pressió i l\'altra persona està enfadada, la teva capacitat de regular les teves pròpies emocions determina si la conversa construeix o destrueix."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Regulació emocional en situacions de tensió',
      text: 'La regulació emocional és la capacitat de reconèixer, comprendre i gestionar les pròpies emocions de manera que no interfereixin negativament en la comunicació. En situacions de tensió, el cervell activa la resposta d\'amenaça (amígdala), cosa que dificulta el pensament racional. Tècniques útils: respiració profunda, pausa conscient ("necessito un moment"), reformulació cognitiva ("el seu enfado no és un atac personal, és una necessitat insatisfeta").'
    },
    subDecisions: [
      {
        id: 'sd1',
        pregunta: 'Decisió 1: Com inicia Harry la conversa?',
        opcions: [
          {
            id: 'A',
            text: '"Jordi, vull que sàpigues que entenc la teva frustració. T\'hauré consultat el canvi de programa. Per a mi és important que confis en les meves decisions, i per això vull explicar-te el raonament i escoltar el teu punt de vista."',
            feedback: 'Obertura empàtica perfecta. Harry reconeix la perspectiva d\'en Jordi ("entenc la frustració"), assumeix responsabilitat sense excesos ("t\'hauré consultat"), i estableix l\'objectiu mutu ("que confis en les meves decisions"). Aquesta és la base d\'una comunicació assertiva: drets propis i drets de l\'altre, en equilibri.',
            punts: 10
          },
          {
            id: 'B',
            text: '"Jordi, vull explicar-te per què he canviat el programa. Era per millorar el teu rendiment al triatlò."',
            feedback: 'Correcte però incomplet. Harry inicia explicant el seu punt de vista sense haver reconegut primer la perspectiva d\'en Jordi. La comunicació efectiva en situacions de conflicte requereix que primer l\'altra persona se senti escol·tada i validada. Explicar sense reconèixer posa la raó pròpia per davant de l\'emoció de l\'altre.',
            punts: 5
          },
          {
            id: 'C',
            text: '"Jordi, sé que vas reaccionar de forma exagerada ahir, però em semblava bé parlar per aclarir les coses."',
            feedback: 'Inici defensiu i acusatori. Qualificar la reacció d\'en Jordi com "exagerada" és una **invalidació emocional** que immediatament tancarà la conversa. L\'altra persona s\'haurà de defensar en comptes de dialogar. Mai és recomanable iniciar una conversa difícil jutjant l\'emoció de l\'altre.',
            punts: 0
          }
        ]
      },
      {
        id: 'sd2',
        pregunta: 'Decisió 2: En Jordi s\'enfada ("Em sap molt greu, però no pots canviar el meu programa sense consultar-me. No és professional."). Com respon Harry?',
        opcions: [
          {
            id: 'A',
            text: 'Fer una pausa, respirar, i dir: "Tens raó en el fons. El canvi era tècnicament correcte, però el procés ha fallat. Hauria d\'haver-te informat. Pots dir-me com hauries preferit que ho gestionés?"',
            feedback: 'Regulació emocional excel·lent. Harry no reacciona a l\'intensitat emocional d\'en Jordi: fa una pausa, valida el contingut (no la forma), assumeix la part de responsabilitat que li correspon i redirigeix cap a una solució. Això és **empatia professional**: comprendre sense perdre el rol professional ni cedir innecessàriament.',
            punts: 10
          },
          {
            id: 'B',
            text: 'Dir: "Ho sento, ho sento, tens raó, no ho tornaré a fer, disculpa..."',
            feedback: 'Disculpa excessiva i passiva. Assumir tota la culpa sense matisos no resol el problema de fons i crea un precedent negatiu. A més, la repetició de disculpes pot semblar poc sincera. La resposta assertiva implica assumir la responsabilitat real, no tota la responsabilitat imaginable.',
            punts: 5
          },
          {
            id: 'C',
            text: 'Dir: "No és una qüestió de consultar-te: jo sóc l\'entrenador. Tu has de confiar en el meu criteri professional."',
            feedback: 'Resposta agressiva que escala el conflicte. Invocar l\'autoritat professional com a argument quan l\'altre ha expressat una necessitat legítima (ser consultat) genera resistència i tanca la conversa. L\'autoritat professional no s\'imposa: es guanya.',
            punts: 0
          }
        ]
      },
      {
        id: 'sd3',
        pregunta: 'Decisió 3: Arribar a un acord. En Jordi accepta que el canvi tècnic tenia sentit, però vol ser consultat en el futur. Com respon Harry?',
        opcions: [
          {
            id: 'A',
            text: 'Proposar un protocol clar: "D\'acord. En el futur, quan vulgui modificar el teu programa, t\'enviaré un missatge explicant el canvi i el motiu 48 hores abans. Si tens dubtes o preferències, les parlem. Funciona per a tu?"',
            feedback: 'Acord mutu concret i respectuós. Proposar un protocol específic (48 hores, missatge, explicació) demostra que Harry ha escoltat la necessitat real d\'en Jordi (no només la queixa) i ha convertit el conflicte en un acord de treball clar. Això és **negociació efectiva**: arribar a un acord que satisfà totes dues parts de forma sostenible.',
            punts: 10
          },
          {
            id: 'B',
            text: 'Dir: "D\'acord, d\'ara endavant tu decideixes el programa i jo l\'executo."',
            feedback: 'Cessió excessiva. Renunciar al rol professional de presa de decisions tècniques per evitar el conflicte no és un bon acord: és rendició. El client pot quedar satisfet a curt termini, però l\'entrenador haurà perdut l\'autoritat professional necessària per fer la seva feina correctament.',
            punts: 5
          },
          {
            id: 'C',
            text: 'Dir: "Bé, ho intentaré", sense comprometre res concret.',
            feedback: 'Resposta vaga que no tanca el conflicte. "Ho intentaré" no és un compromís: és una manera de finalitzar la conversa sense resoldre-la. En Jordi quedarà amb la sensació que res canviarà. Les converses difícils han d\'acabar amb acords clars i mesurables, no amb bones intencions.',
            punts: 0
          }
        ]
      }
    ],
    seguent: 'scene_09'
  },

  /* ─────────── */

  scene_09: {
    id: 'scene_09',
    etapa: 'La Recompensa',
    personatge: 'hanna',
    tipus: 'worked_example',
    titol: 'La recompensa: la confiança guanyada',
    narracio: `En Jordi surt de la sala i li estreny la mà a Harry. No és un gest exuberant, però per a en Jordi, que rarament expressa aprovació, és molt significatiu. "Bé, em sembla bé el que has proposat. Continuem." Harry treu un llarg sospir.

La Hanna estava al passadís. Ha escoltat tot des de la porta entrebadada —no per espiar, sinó per poder donar feedback immediat. Quan en Jordi marxa, s'apropa a Harry i li diu simplement: "Has crescut deu anys en una hora."

Asseguts a la sala d'entrenadors, la Hanna explica a Harry els cinc estils d'afrontament del conflicte. No tots els estils són dolents en tots els contextos: la clau és saber quin usar en cada situació. "En Jack usa gairebé sempre la competència o l'evitació. Tu, avui, has usat la col·laboració. I ha funcionat."

Harry anota els cinc estils en el bloc de notes. Per primera vegada, no anota tècniques d'entrenament: anota habilitats humanes.`,
    dialeg: {
      personatge: 'hanna',
      text: '"El conflicte no és el problema. El conflicte és l\'oportunitat. Com el gestionis determina si la relació surt enfortida o trencada. Avui l\'has sortit enfortida. Recorda\'t d\'aquest moment quan el proper conflicte sembli insuperable."'
    },
    contingut_pedagogic: {
      tipus: 'concept_definition',
      titol: 'Els cinc estils d\'afrontament del conflicte (Thomas-Kilmann)',
      text: `**1. Competència (Assertiu/No cooperatiu):** Imposar la pròpia posició. Útil en emergències o decisions crítiques. Risc: danya relacions si s\'usa de forma habitual.

**2. Evitació (No assertiu/No cooperatiu):** Retirar-se o ajornar. Útil quan el conflicte és trivial o cal temps. Risc: els problemes no resolts s\'acumulen.

**3. Acomodació (No assertiu/Cooperatiu):** Cedir als interessos de l\'altre. Útil per mantenir la relació a curt termini. Risc: pèrdua de l\'autoritat professional.

**4. Compromís (Moderadament assertiu/cooperatiu):** Cada part cedeix una mica. Útil per a solucions ràpides i equilibrades. Risc: cap de les parts queda del tot satisfeta.

**5. Col·laboració (Assertiu/Cooperatiu):** Buscar solucions que satisfacin totes dues parts. La solució ideal. Requereix temps i confiança mútua.`
    },
    seguent: 'scene_10'
  },

  /* ─────────────────────────────────────────
     ACTE III – EL RETORN
  ───────────────────────────────────────── */

  scene_10: {
    id: 'scene_10',
    etapa: 'El Camí de Retorn',
    personatge: 'miquel',
    tipus: 'decision_scenario',
    titol: 'La tensió entre Jack i la Neus',
    narracio: `Dues setmanes més tard. Les coses a FitCore han canviat subtilment: els clients del grup de Harry cada vegada estan més satisfets, i el boca-orella ha portat tres nous clients a les seves sessions. Però a la sala d'entrenadors hi ha una tensió creixent.

En Jack i la Neus, entrenadora especialitzada en classes col·lectives, porten tres dies sense parlar-se. La situació va escalar quan en Jack va criticar públicament els programes de la Neus davant d'un grup de clients ("les classes de ioga no serveixen per perdre pes"). La Neus ho va prendre com un atac personal i professional, i va respondre amb silenci i evitació.

El director Miquel truca Harry al seu despatx: "Harry, sé que has gestionat bé el tema d'en Jordi. Aquí tenim un altre problema. En Jack i la Neus no es parlen i l'ambient s'ha enrarit. Tu pots intervenir com a mediador informal, o pots quedar-te al marge. La decisió és teva."`,
    dialeg: {
      personatge: 'miquel',
      text: '"No et demano que soluciones tot sol el problema. Però de vegades un company que intervé amb bona intenció pot desbloquejar una situació que els protagonistes ja no poden veure amb claredat. Tens molt a perdre si t\'equivoques, però molt a guanyar si ho fas bé."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Mediació informal: quan i com intervenir en un conflicte aliè',
      text: 'La mediació és un procés en el qual una tercera persona neutral facilita la comunicació entre les parts en conflicte per ajudar-les a arribar a un acord. La mediació informal en entorns laborals requereix: (1) imparcialitat real (no prendre partit), (2) comunicació de la intenció ("vull ajudar, no jutjar"), (3) escoltar les dues parts per separat primer, (4) facilitar un diàleg on cada part pugui expressar les seves necessitats, (5) no imposar solucions, sinó acompanyar el procés.'
    },
    opcions: [
      {
        id: 'A',
        text: 'Intervenir com a mediador: parlar primer per separat amb Jack i amb la Neus, i després facilitar un diàleg entre ells',
        feedback: 'Decisió madura i professional. Intervenir com a mediador —parlant primer per separat i escoltant les dues perspectives— és la millor manera d\'entendre el conflicte real més enllà de les posicions superficials. Quan les persones se senten escoltades per separat, baixen les defenses i estan més disposades al diàleg. Aquesta és la base de la **mediació efectiva**: crear espai de seguretat psicològica per al diàleg.',
        punts: 10,
        seguent: 'scene_11'
      },
      {
        id: 'B',
        text: 'No intervenir i deixar que el problema es resolgui sol o que Miquel ho gestioni directament',
        feedback: 'L\'evitació és de vegades vàlida, però en aquest cas els protagonistes han demostrat que no poden resoldre el conflicte per ells mateixos (tres dies de silenci i evitació). Quedar-se al marge quan hi ha l\'oportunitat d\'ajudar és una oportunitat perduda de creixement professional i de millora de l\'ambient de treball. Veuràs el que passa quan el conflicte escala.',
        punts: 0,
        seguent: 'scene_10b'
      }
    ]
  },

  /* ─────────── */

  scene_10b: {
    id: 'scene_10b',
    etapa: 'El Camí de Retorn',
    personatge: 'narracio',
    tipus: 'text_block',
    titol: 'Quan el conflicte escala',
    narracio: `Harry decideix no intervenir. "No és el meu problema", es diu. Passa la setmana. En Jack i la Neus continuen sense parlar-se. Un dia, durant una sessió de grup que totes dues zones comparteixen, en Jack puja el volum de la música de la zona de pesos al màxim en el moment en què la Neus intenta explicar un exercici de respiració als seus clients. La Neus perd la compostura i li diu en veu alta que és un maleducat.

Tres clients ho veuen. Dos demanen parlar amb el director. Un terç diu que "si l'ambient al centre és aquest" es planteja canviar de gimnàs.

Miquel truca tothom al seu despatx. La conversa és tensa. En acabar, mira Harry i diu: "Quan et vaig dir que tenies l'oportunitat d'intervenir, ho deia en serio." Harry entén que l'evitació no és neutralitat: és una elecció que té conseqüències.`,
    dialeg: {
      personatge: 'miquel',
      text: '"En els entorns professionals, el silenci davant d\'un conflicte no és innocència. Qui veu un problema i no fa res quan pot fer-ho és part del problema. Ara hem de gestionar algo molt més gran del que era fa una setmana."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'L\'evitació del conflicte i les seves conseqüències',
      text: 'L\'estil d\'evitació pot semblar una opció neutral, però en conflictes interpersonals que afecten l\'entorn laboral, no intervenir implica permetre que el conflicte segueixi el seu curs natural —que tendeix a escalar. Els conflictes no resolts generen: (1) deteriorament del clima laboral, (2) reducció de la productivitat, (3) efectes sobre tercers (clients, col·legues), (4) necessitat de gestió correctiva més costosa.'
    },
    seguent: 'scene_11'
  },

  /* ─────────── */

  scene_11: {
    id: 'scene_11',
    etapa: 'La Resurrecció',
    personatge: 'miquel',
    tipus: 'decision_scenario',
    titol: 'L\'opinió sobre en Jack',
    narracio: `La setmana siguiente. La situació amb en Jack ha tornat a escalfar-se: un informe intern revela que tres clients han dit en les enquestes de satisfacció que "un cert entrenador" els ha fet sentir menyspreats o incompetents. Miquel sap que es refereixen a en Jack.

Miquel demana una reunió a soles amb Harry. "Harry, sé que coneixes en Jack millor que jo ara. He de prendre una decisió difícil: si li done una última oportunitat amb formació i seguiment, o si els nostres camins s'han de separar. Necessito la teva opinió professional i honesta. No la que creus que vull sentir."

Harry nota el pes de la situació. En Jack no li cau bé, però és una persona real amb una vida real. I al mateix temps, els clients mereixen un entorn segur i respectuós. La manera com Harry expressi la seva opinió dirà molt sobre qui és com a professional.`,
    dialeg: {
      personatge: 'miquel',
      text: '"Harry, et demano l\'opinió d\'un professional que ha vist de prop com en Jack treballa. No et demano que l\'acusis ni que el defenses. Et demano que siguis honest, respectuós i que em donis informació real per prendre una decisió justa."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Assertivitat avançada: com donar una opinió professional difícil',
      text: 'Donar una opinió professional difícil sobre un col·lega requereix: (1) basar-se en fets observables, no en judicis de valor personals, (2) separar la persona del comportament ("el comportament X ha generat l\'efecte Y"), (3) reconèixer el context i les possibles causes, (4) proposar vies de millora quan siguin possibles, (5) ser clar en allò que has observat directament i honest sobre allò que desconeixem. L\'assertivitat no és duresa ni suavitat: és precisió i respecte simultanis.'
    },
    opcions: [
      {
        id: 'A',
        text: '"Miquel, en Jack és un mal professional i s\'hauria d\'haver anat fa temps. Els clients no el suporten i l\'equip tampoc."',
        feedback: 'Resposta agressiva basada en judicis de valor. "Mal professional" és una etiqueta global que no aporta informació útil ni respecta la dignitat d\'en Jack. Aquesta opinió pot reflectir frustració acumulada, però no és una opinió professional: és un judici personal. Una persona assertiva separa els fets observables de les conclusions globals.',
        punts: 0,
        seguent: 'scene_12'
      },
      {
        id: 'B',
        text: '"No ho sé, Miquel. Crec que no sóc la persona adequada per opinar sobre un col·lega."',
        feedback: 'Resposta passiva que evita la responsabilitat. Miquel ha demanat explícitament l\'opinió d\'Harry com a professional que ha observat en Jack de prop. Evitar opinar quan tens informació rellevant i ets preguntat directament és una forma de passivitat que no ajuda la presa de decisions. La correcció d\'un entorn laboral és responsabilitat col·lectiva.',
        punts: 5,
        seguent: 'scene_12'
      },
      {
        id: 'C',
        text: '"Miquel, el que he observat directament és que en Jack utilitza un estil de comunicació que tendeix a invalidar els clients quan expressen dificultats. He vist tres situacions concretes on un client s\'ha sentit menyspreuat. Al mateix temps, els seus coneixements tècnics són sòlids. La pregunta és si pot canviar el seu estil comunicatiu amb la formació adequada."',
        feedback: 'Resposta assertiva exemplar. Harry es basa en fets observables ("he vist tres situacions concretes"), reconeix allò positiu ("els seus coneixements tècnics"), separa la persona del comportament i proposa una via constructiva. Això és **assertivitat avançada**: expressar la veritat de forma respectuosa, útil i constructiva, sense atacar ni evitar.',
        punts: 10,
        seguent: 'scene_12'
      }
    ]
  },

  /* ─────────── */

  scene_12: {
    id: 'scene_12',
    etapa: 'El Retorn amb l\'Elixir',
    personatge: 'harry',
    tipus: 'checklist',
    titol: 'La reunió d\'equip: establint les bases',
    narracio: `Tres setmanes més tard. En Jack ha acceptat fer un programa de formació en habilitats comunicatives i ha quedat al centre amb supervisió. La Neus i en Jack han tingut una conversa mediada per Harry que ha estat difícil però productiva. Les enquestes de satisfacció han pujat un 18%.

El director Miquel proposa que Harry dirigeixi una reunió d'equip mensual de comunicació interna. "Tu has demostrat que saps gestionar persones. Ara has d'ajudar a crear una cultura de comunicació al centre."

Harry prepara la reunió. Sap que no es tracta de donar un curs, sinó d'establir uns principis comuns que tothom pugui aplicar. Revisa tot el que ha après en les últimes setmanes i elabora un decàleg de bones pràctiques comunicatives per a FitCore.

**Instrucció:** Marca tots els principis que creus que haurien de formar part del decàleg de comunicació de FitCore. Cada principi correcte suma punts.`,
    dialeg: {
      personatge: 'harry',
      text: '"Quan vaig arribar a FitCore pensava que la feina d\'entrenador era fer programes d\'entrenament. Ara sé que la feina real és crear relacions de confiança. Els exercicis canvien el cos. La comunicació canvia les persones."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'Integració: les sis competències comunicatives de l\'entrenador professional',
      text: 'Un entrenador de condicionament físic complet domina: (1) la comunicació interpersonal clara i empàtica, (2) la coherència entre comunicació verbal i no verbal, (3) la gestió emocional pròpia i la comprensió de les emocions dels clients, (4) l\'empatia professional (comprendre sense perdre el rol), (5) l\'assertivitat per expressar opinions i límits amb respecte, (6) la resolució constructiva dels conflictes. Aquestes sis competències no s\'ensenyen als clients: s\'ensenyen als entrenadors.'
    },
    checklistItems: [
      { id: 'dc1', text: 'Escoltar activament els clients: deixar parlar, reformular i confirmar que hem entès', correcta: true },
      { id: 'dc2', text: 'Comunicar els canvis de programa als clients amb antelació i explicant el motiu', correcta: true },
      { id: 'dc3', text: 'Gestionar les crítiques públiques amb calma i proposar un espai privat per parlar', correcta: true },
      { id: 'dc4', text: 'Expressar opinions professionals de forma assertiva: basant-se en fets, no en judicis', correcta: true },
      { id: 'dc5', text: 'Reconèixer i gestionar les pròpies emocions abans de comunicar en situacions de tensió', correcta: true },
      { id: 'dc6', text: 'Crear espais segurs per als conflictes interns: mediació abans d\'escalada', correcta: true },
      { id: 'dc7', text: 'Mantenir coherència entre el que diem i com ho diem (verbal i no verbal)', correcta: true },
      { id: 'dc8', text: 'Adaptar el canal de comunicació (presencial, escrit, telèfon) a la situació i al client', correcta: true },
      { id: 'dc9', text: 'Donar feedback constructiu als clients: específic, orientat a la millora i respectuós', correcta: true },
      { id: 'dc10', text: 'Demanar i acceptar feedback del client sobre la nostra pràctica professional', correcta: true },
      { id: 'dc11', text: 'Usar la comunicació agressiva quan cal posar límits ferms als clients difícils', correcta: false },
      { id: 'dc12', text: 'Evitar sempre els conflictes per mantenir un ambient positiu al centre', correcta: false },
      { id: 'dc13', text: 'No demanar mai consell a col·legues per no demostrar inseguretat professional', correcta: false }
    ],
    punts_per_item: 1,
    seguent: 'scene_13'
  },

  /* ─────────── */

  scene_13: {
    id: 'scene_13',
    etapa: 'Epíleg: El Nou Harry',
    personatge: 'hanna',
    tipus: 'epilogue',
    titol: 'El Retorn amb l\'Elixir',
    narracio: `Sis setmanes des del primer dia. La sala de pesos de FitCore a les set del matí. Harry es prepara per a la primera sessió del dia. Però ara alguna cosa ha canviat: no és el nerviosisme d'aquell primer dia. Ara hi ha una calma diferent, una seguretat que no prové de saber-ho tot, sinó de saber com aprendre de cada interacció.

La Hanna s'acosta i li deixa una nota enganxada a la porta del taquiller. A la nota hi ha una sola frase: "L'heroi no és el qui no cau. És el qui aprèn cada vegada que es reincorpora."

Harry mira el grup de clients que esperen a l'entrada. Hi ha la Maria, que ara arriba amb un somriure. En Jordi, que li fa un gest de cap. La Lluïsa, que ha portat una amiga nova. I alguns rostres nous que encara no coneix, però que aviat coneixerà.

El viatge ha canviat Harry. No li ha ensenyat tècniques noves d'entrenament. Li ha ensenyat a ser un professional complet: algú que sap que la millor tècnica del món no serveix de res si no saps connectar amb les persones que vols ajudar.`,
    dialeg: {
      personatge: 'hanna',
      text: '"Ja no ets el Harry del primer dia. Ets algú que ha après que la comunicació és una habilitat, no un talent. I les habilitats es practiquen cada dia. Benvingut al club dels professionals de veritat."'
    },
    contingut_pedagogic: {
      tipus: 'idea_clau',
      titol: 'El Viatge de l\'Heroi com a metàfora de l\'aprenentatge professional',
      text: 'El Viatge de l\'Heroi (Campbell / Vogler) descriu el procés universal de creixement: sortir del món conegut, enfrontar proves, trobar aliats i mentors, superar la prova suprema, i tornar transformat amb un "elixir" —el coneixement que pot compartir amb els altres. En el context de la formació professional, l\'elixir no és una tècnica: és la comprensió profunda de que les relacions humanes son la base de qualsevol professió que treballa amb persones.'
    }
  }
};

/* ============================================================
   ENGINE – Lògica de navegació i puntuació
   ============================================================ */

const Engine = {
  state: {
    currentScene: 'scene_01',
    score: 0,
    decisions: {},        // sceneId: { optionId, punts }
    checklistData: {},    // sceneId: { selectedIds: [] }
    ordealData: {},       // sceneId: { sd1: optId, sd2: optId, sd3: optId }
    visitedScenes: [],
    completedScenes: [],
    finished: false
  },

  init: function () {
    SCORM.init();
    var saved = SCORM.loadSuspendData();
    if (saved) {
      this.state = saved;
      console.info('[Engine] Estat recuperat:', this.state.currentScene);
    }
    this._renderScene(this.state.currentScene);
    this._renderProgressBar();
    this._renderJourneyMap();
  },

  _save: function () {
    SCORM.saveSuspendData(this.state);
  },

  _allSceneIds: function () {
    return Object.keys(scenes);
  },

  _progressPercent: function () {
    var all = this._allSceneIds();
    var done = this.state.completedScenes.length;
    return Math.round((done / all.length) * 100);
  },

  _markVisited: function (sceneId) {
    if (this.state.visitedScenes.indexOf(sceneId) === -1) {
      this.state.visitedScenes.push(sceneId);
    }
  },

  _markCompleted: function (sceneId) {
    if (this.state.completedScenes.indexOf(sceneId) === -1) {
      this.state.completedScenes.push(sceneId);
    }
  },

  goToScene: function (sceneId) {
    if (!scenes[sceneId]) {
      console.warn('[Engine] Escena no trobada:', sceneId);
      return;
    }
    this._markCompleted(this.state.currentScene);
    this.state.currentScene = sceneId;
    this._markVisited(sceneId);
    this._save();
    this._renderScene(sceneId);
    this._renderProgressBar();
    this._renderJourneyMap();
    window.scrollTo(0, 0);
  },

  recordDecision: function (sceneId, optionId, punts) {
    this.state.decisions[sceneId] = { optionId: optionId, punts: punts };
    this.state.score += punts;
    this._save();
  },

  recordChecklist: function (sceneId, selectedIds, puntsTotals) {
    this.state.checklistData[sceneId] = { selectedIds: selectedIds };
    this.state.score += puntsTotals;
    this._save();
  },

  recordOrdeal: function (sceneId, subDecisions) {
    // subDecisions = [{ id, optionId, punts }, ...]
    var total = 0;
    subDecisions.forEach(function (sd) { total += sd.punts; });
    this.state.ordealData[sceneId] = subDecisions;
    this.state.score += total;
    this._save();
  },

  finish: function () {
    this.state.finished = true;
    this._markCompleted(this.state.currentScene);
    this._save();
    SCORM.finish(this.state.score);
  },

  _renderScene: function (sceneId) {
    var scene = scenes[sceneId];
    if (!scene) return;
    UI.render(scene, this.state);
  },

  _renderProgressBar: function () {
    UI.updateProgressBar(this._progressPercent(), this.state.completedScenes.length, this._allSceneIds().length);
  },

  _renderJourneyMap: function () {
    UI.updateJourneyMap(JOURNEY_STAGES, this.state.visitedScenes, this.state.completedScenes);
  }
};

/* ============================================================
   UI – Renderitzat de l'interfície
   ============================================================ */

const UI = {
  _currentOrdealStep: 0,
  _ordealAnswers: [],
  _continueTimer: null,
  _readyToAdvance: false,

  render: function (scene, state) {
    this._currentOrdealStep = 0;
    this._ordealAnswers = [];
    this._readyToAdvance = false;

    document.getElementById('stage-label').textContent = scene.etapa || '';
    document.getElementById('scene-title').textContent = scene.titol || '';
    document.getElementById('score-display').textContent = 'Puntuació: ' + state.score;

    this._renderCharacter(scene);
    this._renderNarrative(scene);
    this._renderDialogue(scene);
    this._renderPedagogic(scene);
    this._renderInteraction(scene, state);
  },

  _renderCharacter: function (scene) {
    var char = CHARACTERS[scene.personatge] || CHARACTERS.narracio;
    var avatarEl = document.getElementById('character-avatar');
    var nameEl = document.getElementById('character-name');
    avatarEl.style.background = char.color;
    avatarEl.textContent = char.initials;
    avatarEl.style.borderRadius = char.shape === 'square' ? '8px' : '50%';
    nameEl.textContent = char.name;
    nameEl.style.color = char.color;
  },

  _renderNarrative: function (scene) {
    var el = document.getElementById('narrative-text');
    el.innerHTML = this._md(scene.narracio || '');
  },

  _renderDialogue: function (scene) {
    if (!scene.dialeg) {
      document.getElementById('dialogue-box').style.display = 'none';
      return;
    }
    var char = CHARACTERS[scene.dialeg.personatge] || CHARACTERS.narracio;
    var box = document.getElementById('dialogue-box');
    box.style.display = 'block';
    box.style.borderLeftColor = char.color;
    document.getElementById('dialogue-char-name').textContent = char.name;
    document.getElementById('dialogue-char-name').style.color = char.color;
    document.getElementById('dialogue-text').innerHTML = this._md(scene.dialeg.text || '');
  },

  _renderPedagogic: function (scene) {
    if (!scene.contingut_pedagogic) {
      document.getElementById('pedagogic-block').style.display = 'none';
      return;
    }
    var block = document.getElementById('pedagogic-block');
    block.style.display = 'block';
    document.getElementById('pedagogic-title').textContent = scene.contingut_pedagogic.titol || '';
    document.getElementById('pedagogic-text').innerHTML = this._md(scene.contingut_pedagogic.text || '');
  },

  _renderInteraction: function (scene, state) {
    var interactionEl = document.getElementById('interaction-area');
    interactionEl.innerHTML = '';

    if (scene.tipus === 'text_block' || scene.tipus === 'worked_example' || scene.tipus === 'concept_definition') {
      this._renderTextBlock(scene, state, interactionEl);
    } else if (scene.tipus === 'decision_scenario') {
      this._renderDecision(scene, state, interactionEl);
    } else if (scene.tipus === 'checklist') {
      this._renderChecklist(scene, state, interactionEl);
    } else if (scene.tipus === 'ordeal') {
      this._renderOrdeal(scene, state, interactionEl);
    } else if (scene.tipus === 'epilogue') {
      this._renderEpilogue(scene, state, interactionEl);
    }
  },

  _renderTextBlock: function (scene, state, container) {
    var self = this;
    var btn = this._makeButton('Continua →', 'btn-primary btn-disabled', function () {
      if (self._readyToAdvance) {
        Engine.goToScene(scene.seguent);
      }
    });
    btn.id = 'continue-btn';
    container.appendChild(btn);

    clearTimeout(this._continueTimer);
    this._readyToAdvance = false;
    this._continueTimer = setTimeout(function () {
      self._readyToAdvance = true;
      var b = document.getElementById('continue-btn');
      if (b) {
        b.classList.remove('btn-disabled');
        b.classList.add('btn-enabled');
      }
    }, 3000);
  },

  _renderDecision: function (scene, state, container) {
    // Si ja s'ha pres la decisió en aquesta escena (per si es torna a renderitzar)
    var prev = state.decisions[scene.id];

    var label = document.createElement('p');
    label.className = 'decision-label';
    label.textContent = 'Pren una decisió:';
    container.appendChild(label);

    var self = this;
    scene.opcions.forEach(function (opcio) {
      var btn = self._makeButton(opcio.id + ') ' + opcio.text, 'btn-option' + (prev && prev.optionId === opcio.id ? ' btn-selected' : ''), function () {
        if (prev) return; // ja decidit
        Engine.recordDecision(scene.id, opcio.id, opcio.punts);
        self._showFeedback(opcio, function () {
          Engine.goToScene(opcio.seguent);
        });
        // Disable all options
        container.querySelectorAll('.btn-option').forEach(function (b) {
          b.classList.add('btn-disabled');
        });
      });
      container.appendChild(btn);
    });

    if (prev) {
      // Mostra feedback de l'opció triada
      var triedOpcio = scene.opcions.find(function (o) { return o.id === prev.optionId; });
      if (triedOpcio) {
        var continueBtn = self._makeButton('Continua →', 'btn-primary btn-enabled', function () {
          Engine.goToScene(triedOpcio.seguent);
        });
        container.appendChild(continueBtn);
      }
    }
  },

  _renderChecklist: function (scene, state, container) {
    var self = this;
    var prev = state.checklistData[scene.id];

    var label = document.createElement('p');
    label.className = 'decision-label';
    label.textContent = 'Marca els elements que creus importants:';
    container.appendChild(label);

    var selected = prev ? prev.selectedIds.slice() : [];
    var itemsDiv = document.createElement('div');
    itemsDiv.className = 'checklist-items';

    scene.checklistItems.forEach(function (item) {
      var itemDiv = document.createElement('div');
      itemDiv.className = 'checklist-item' + (selected.indexOf(item.id) !== -1 ? ' checked' : '');
      itemDiv.dataset.id = item.id;

      var checkbox = document.createElement('span');
      checkbox.className = 'checkbox-icon';
      checkbox.textContent = selected.indexOf(item.id) !== -1 ? '☑' : '☐';

      var text = document.createElement('span');
      text.className = 'checkbox-text';
      text.textContent = item.text;

      itemDiv.appendChild(checkbox);
      itemDiv.appendChild(text);

      if (!prev) {
        itemDiv.addEventListener('click', function () {
          var idx = selected.indexOf(item.id);
          if (idx === -1) {
            selected.push(item.id);
            itemDiv.classList.add('checked');
            checkbox.textContent = '☑';
          } else {
            selected.splice(idx, 1);
            itemDiv.classList.remove('checked');
            checkbox.textContent = '☐';
          }
        });
      }

      itemsDiv.appendChild(itemDiv);
    });

    container.appendChild(itemsDiv);

    if (!prev) {
      var confirmBtn = self._makeButton('Confirmar selecció', 'btn-primary btn-enabled', function () {
        // Calcular punts
        var punts = 0;
        selected.forEach(function (id) {
          var item = scene.checklistItems.find(function (i) { return i.id === id; });
          if (item && item.correcta) {
            punts += scene.punts_per_item;
          }
        });

        Engine.recordChecklist(scene.id, selected, punts);

        // Mostrar resultats
        scene.checklistItems.forEach(function (item) {
          var itemDiv = itemsDiv.querySelector('[data-id="' + item.id + '"]');
          if (!itemDiv) return;
          itemDiv.style.pointerEvents = 'none';
          var wasSelected = selected.indexOf(item.id) !== -1;
          if (item.correcta && wasSelected) {
            itemDiv.classList.add('cl-correct');
          } else if (!item.correcta && wasSelected) {
            itemDiv.classList.add('cl-incorrect');
          } else if (item.correcta && !wasSelected) {
            itemDiv.classList.add('cl-missed');
          }
        });

        var feedback = document.createElement('div');
        feedback.className = 'checklist-feedback';
        var maxPunts = scene.checklistItems.filter(function (i) { return i.correcta; }).length * scene.punts_per_item;
        feedback.innerHTML = '<strong>Has obtingut ' + punts + ' de ' + maxPunts + ' punts possibles.</strong><br>' +
          '<span style="color:#27AE60">☑ Verd = correcte i seleccionat</span> &nbsp; ' +
          '<span style="color:#C0392B">☒ Vermell = incorrecte seleccionat</span> &nbsp; ' +
          '<span style="color:#E67E22">○ Taronja = correcte no seleccionat</span>';
        container.appendChild(feedback);
        confirmBtn.remove();

        var nextBtn = self._makeButton('Continua →', 'btn-primary btn-enabled', function () {
          Engine.goToScene(scene.seguent);
        });
        container.appendChild(nextBtn);
      });
      container.appendChild(confirmBtn);
    } else {
      // Mostrar resultats previs
      scene.checklistItems.forEach(function (item) {
        var itemDiv = itemsDiv.querySelector('[data-id="' + item.id + '"]');
        if (!itemDiv) return;
        itemDiv.style.pointerEvents = 'none';
        var wasSelected = selected.indexOf(item.id) !== -1;
        if (item.correcta && wasSelected) itemDiv.classList.add('cl-correct');
        else if (!item.correcta && wasSelected) itemDiv.classList.add('cl-incorrect');
        else if (item.correcta && !wasSelected) itemDiv.classList.add('cl-missed');
      });

      var nextBtn2 = self._makeButton('Continua →', 'btn-primary btn-enabled', function () {
        Engine.goToScene(scene.seguent);
      });
      container.appendChild(nextBtn2);
    }
  },

  _renderOrdeal: function (scene, state, container) {
    var self = this;
    var step = self._currentOrdealStep;
    var prevOrdeal = state.ordealData[scene.id];

    if (prevOrdeal) {
      // Ja resolt
      var doneDiv = document.createElement('div');
      doneDiv.className = 'ordeal-done';
      var totalPunts = 0;
      prevOrdeal.forEach(function (sd) { totalPunts += sd.punts; });
      doneDiv.innerHTML = '<p>Has completat la prova suprema amb <strong>' + totalPunts + ' punts</strong>.</p>';
      container.appendChild(doneDiv);
      var nextBtn = self._makeButton('Continua →', 'btn-primary btn-enabled', function () {
        Engine.goToScene(scene.seguent);
      });
      container.appendChild(nextBtn);
      return;
    }

    var sd = scene.subDecisions[step];
    var preguntaEl = document.createElement('p');
    preguntaEl.className = 'ordeal-question';
    preguntaEl.textContent = sd.pregunta;
    container.appendChild(preguntaEl);

    sd.opcions.forEach(function (opcio) {
      var btn = self._makeButton(opcio.id + ') ' + opcio.text, 'btn-option', function () {
        container.querySelectorAll('.btn-option').forEach(function (b) { b.classList.add('btn-disabled'); });

        self._ordealAnswers.push({ id: sd.id, optionId: opcio.id, punts: opcio.punts });

        var fbDiv = document.createElement('div');
        fbDiv.className = 'feedback-inline ' + (opcio.punts === 10 ? 'fb-good' : opcio.punts === 5 ? 'fb-ok' : 'fb-bad');
        var icon = opcio.punts === 10 ? '✓' : opcio.punts === 5 ? '⚠' : '✗';
        fbDiv.innerHTML = '<span class="fb-icon">' + icon + '</span>' +
          '<span class="fb-pts">+' + opcio.punts + ' pts</span>' +
          '<p class="fb-text">' + self._md(opcio.feedback) + '</p>';
        container.appendChild(fbDiv);

        var nextStepBtn = document.createElement('button');
        nextStepBtn.className = 'btn btn-primary btn-enabled';

        if (self._currentOrdealStep < scene.subDecisions.length - 1) {
          nextStepBtn.textContent = 'Següent decisió →';
          nextStepBtn.addEventListener('click', function () {
            self._currentOrdealStep++;
            container.innerHTML = '';
            self._renderOrdeal(scene, state, container);
          });
        } else {
          nextStepBtn.textContent = 'Continua →';
          nextStepBtn.addEventListener('click', function () {
            Engine.recordOrdeal(scene.id, self._ordealAnswers);
            Engine.goToScene(scene.seguent);
          });
        }
        container.appendChild(nextStepBtn);
      });
      container.appendChild(btn);
    });
  },

  _renderEpilogue: function (scene, state, container) {
    var score = state.score;
    var maxScore = 100;
    var pct = Math.round((score / maxScore) * 100);

    // Missatge de la Hanna segons puntuació
    var hannaMsg = '';
    var hannaClass = '';
    if (score >= 85) {
      hannaMsg = '"Has demostrat una comprensió excepcional de les habilitats comunicatives. FitCore té sort de tenir-te. Continua creixent: cada client és un nou viatge."';
      hannaClass = 'epilogue-excellent';
    } else if (score >= 60) {
      hannaMsg = '"Has après molt en poc temps. Tens una base sòlida. Les habilitats que has treballat necessiten pràctica diària. Continua observant, escoltant i aprenent de cada interacció."';
      hannaClass = 'epilogue-good';
    } else if (score >= 35) {
      hannaMsg = '"Has donat els primers passos. Les habilitats comunicatives son com l\'entrenament físic: requereixen constància i repetició. Torna a revisar les escenes on has tingut dificultats i reflexiona sobre les decisions alternatives."';
      hannaClass = 'epilogue-ok';
    } else {
      hannaMsg = '"El camí de l\'heroi sovint comença amb errors. El que importa no és on comences, sinó la direcció en què vas. Revisa el viatge, reflexiona sobre cada decisió i torna-hi. La comunicació efectiva s\'aprèn, no és innata."';
      hannaClass = 'epilogue-needs-work';
    }

    // Resum de decisions
    var decisionsHtml = '<ul class="decisions-summary">';
    Object.keys(state.decisions).forEach(function (sid) {
      var sc = scenes[sid];
      var dec = state.decisions[sid];
      if (!sc || !dec) return;
      var opcio = sc.opcions.find(function (o) { return o.id === dec.optionId; });
      if (!opcio) return;
      var icon = dec.punts === 10 ? '✓' : dec.punts === 5 ? '⚠' : '✗';
      var cls = dec.punts === 10 ? 'ds-good' : dec.punts === 5 ? 'ds-ok' : 'ds-bad';
      decisionsHtml += '<li class="' + cls + '"><span class="ds-icon">' + icon + '</span> <strong>' + sc.titol + ':</strong> ' + opcio.text + ' <span class="ds-pts">(' + dec.punts + ' pts)</span></li>';
    });
    decisionsHtml += '</ul>';

    // Habilitats treballades
    var skillsHtml = '<ul class="skills-list">' +
      '<li>✓ Comunicació interpersonal i procés comunicatiu</li>' +
      '<li>✓ Comunicació verbal i no verbal: coherència i impacte</li>' +
      '<li>✓ Educació emocional: regulació en situacions de tensió</li>' +
      '<li>✓ Empatia professional en contextos de fitness</li>' +
      '<li>✓ Assertivitat: expressar-se amb respecte i claredat</li>' +
      '<li>✓ Resolució de conflictes i estilos d\'afrontament</li>' +
      '</ul>';

    container.innerHTML =
      '<div class="epilogue-container">' +
        '<div class="epilogue-score-ring">' +
          '<svg viewBox="0 0 120 120" class="score-ring-svg">' +
            '<circle cx="60" cy="60" r="54" fill="none" stroke="#2C2C2C" stroke-width="10"/>' +
            '<circle cx="60" cy="60" r="54" fill="none" stroke="#FF6B35" stroke-width="10" ' +
              'stroke-dasharray="' + (339.292 * pct / 100) + ' 339.292" ' +
              'stroke-dashoffset="84.823" stroke-linecap="round"/>' +
          '</svg>' +
          '<div class="score-ring-text"><span class="score-big">' + score + '</span><span class="score-max">/100</span></div>' +
        '</div>' +

        '<div class="epilogue-hanna ' + hannaClass + '">' +
          '<div class="epilogue-hanna-avatar" style="background:#27AE60">Ha</div>' +
          '<div class="epilogue-hanna-msg">' + hannaMsg + '</div>' +
        '</div>' +

        '<h3 class="epilogue-section-title">Resum de les teves decisions</h3>' +
        decisionsHtml +

        '<h3 class="epilogue-section-title">Habilitats treballades en aquest viatge</h3>' +
        skillsHtml +

        '<div class="epilogue-buttons">' +
          '<button class="btn btn-finish" onclick="Engine.finish()">Finalitzar i enviar puntuació</button>' +
        '</div>' +
      '</div>';
  },

  _showFeedback: function (opcio, onContinue) {
    var overlay = document.getElementById('feedback-overlay');
    var icon = opcio.punts === 10 ? '✓' : opcio.punts === 5 ? '⚠' : '✗';
    var cls = opcio.punts === 10 ? 'fb-good' : opcio.punts === 5 ? 'fb-ok' : 'fb-bad';
    document.getElementById('feedback-icon').textContent = icon;
    document.getElementById('feedback-icon').className = 'feedback-icon ' + cls;
    document.getElementById('feedback-pts').textContent = '+' + opcio.punts + ' punts';
    document.getElementById('feedback-text').innerHTML = this._md(opcio.feedback);
    overlay.style.display = 'flex';

    document.getElementById('feedback-continue').onclick = function () {
      overlay.style.display = 'none';
      onContinue();
    };
  },

  updateProgressBar: function (pct, done, total) {
    document.getElementById('progress-bar-fill').style.width = pct + '%';
    document.getElementById('progress-label').textContent = done + '/' + total + ' escenes';
  },

  updateJourneyMap: function (stages, visited, completed) {
    var mapEl = document.getElementById('journey-map-content');
    if (!mapEl) return;
    mapEl.innerHTML = '';
    stages.forEach(function (stage) {
      var stageDiv = document.createElement('div');
      stageDiv.className = 'map-stage';

      var hasVisited = stage.scenes.some(function (s) { return visited.indexOf(s) !== -1; });
      var hasCompleted = stage.scenes.every(function (s) { return completed.indexOf(s) !== -1; });

      stageDiv.classList.add(hasCompleted ? 'map-done' : hasVisited ? 'map-active' : 'map-pending');
      stageDiv.style.borderLeftColor = stage.color;

      var badge = document.createElement('span');
      badge.className = 'map-badge';
      badge.textContent = stage.label;
      badge.style.background = stage.color;

      var title = document.createElement('span');
      title.className = 'map-title';
      title.textContent = stage.title;

      var status = document.createElement('span');
      status.className = 'map-status';
      status.textContent = hasCompleted ? '✓ Completada' : hasVisited ? '● En curs' : '○ Pendent';

      stageDiv.appendChild(badge);
      stageDiv.appendChild(title);
      stageDiv.appendChild(status);
      mapEl.appendChild(stageDiv);
    });
  },

  _makeButton: function (text, classes, onClick) {
    var btn = document.createElement('button');
    btn.className = 'btn ' + classes;
    btn.textContent = text;
    if (onClick) btn.addEventListener('click', onClick);
    return btn;
  },

  /* Simple Markdown-like: **bold**, *italic*, \n\n paragraphs, lists */
  _md: function (text) {
    if (!text) return '';
    return text
      .split('\n\n')
      .map(function (para) {
        para = para.trim();
        if (!para) return '';
        // Convert **text** to <strong>
        para = para.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        // Convert *text* to <em>
        para = para.replace(/\*([^*]+)\*/g, '<em>$1</em>');
        // Numbered list items
        if (/^\d+\.\s/.test(para)) {
          var items = para.split('\n').filter(function (l) { return l.trim(); });
          return '<ol>' + items.map(function (li) {
            return '<li>' + li.replace(/^\d+\.\s*/, '').replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>') + '</li>';
          }).join('') + '</ol>';
        }
        // Unordered list (lines starting with - or *)
        if (/^[-•]\s/.test(para)) {
          var lis = para.split('\n').filter(function (l) { return l.trim(); });
          return '<ul>' + lis.map(function (li) {
            return '<li>' + li.replace(/^[-•]\s*/, '').replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>') + '</li>';
          }).join('') + '</ul>';
        }
        return '<p>' + para + '</p>';
      })
      .join('');
  }
};

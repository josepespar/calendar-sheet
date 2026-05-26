/* ============================================================
   SCENARIO – Motor Control: The Silent Key Factor in Sports Performance
   28 scenes · Quizzes + Decisions + Checklists · Max score: 100 pts
   Language: English
   Based on the article by Raquel Font-Lladó (University of Girona)
   VII International Seminar on Tactics and Sports Technique
   ============================================================ */

const CHARACTERS = {
  narracio: { name: 'Narrator',              color: '#6B7280', initials: '✦', shape: 'square' },
  raquel:   { name: 'Raquel Font-Lladó',     color: '#F59E0B', initials: 'R', shape: 'circle' },
  joan:     { name: 'Viktor (researcher)',    color: '#60A5FA', initials: 'V', shape: 'circle' },
  andres:   { name: 'Andrés (coach)',         color: '#34D399', initials: 'A', shape: 'circle' },
  esther:   { name: 'Esther (athlete)',       color: '#F472B6', initials: 'E', shape: 'circle' },
  pau:      { name: 'Pau Martí',              color: '#8B5CF6', initials: 'P', shape: 'circle' }
};

const JOURNEY_STAGES = [
  {
    id: 'act1',
    label: 'Act I',
    title: 'The Seminar',
    scenes: ['scene_01','scene_02','scene_03','scene_04','scene_05'],
    color: '#8B5CF6'
  },
  {
    id: 'act2',
    label: 'Act II',
    title: 'Three Perspectives',
    scenes: ['scene_06','scene_07','scene_08','scene_09','scene_10'],
    color: '#60A5FA'
  },
  {
    id: 'act3',
    label: 'Act III',
    title: 'Biology, Experience & Context',
    scenes: ['scene_11','scene_12','scene_13','scene_14','scene_15'],
    color: '#34D399'
  },
  {
    id: 'act4',
    label: 'Act IV',
    title: 'Learning & Awareness',
    scenes: ['scene_16','scene_17','scene_18','scene_19','scene_20'],
    color: '#F472B6'
  },
  {
    id: 'act5',
    label: 'Act V',
    title: 'Measuring & Planning',
    scenes: ['scene_21','scene_22','scene_23','scene_24','scene_24b','scene_25'],
    color: '#F59E0B'
  },
  {
    id: 'final',
    label: 'Epilogue',
    title: 'Synthesis & Conclusion',
    scenes: ['scene_26','scene_27','scene_28'],
    color: '#8B5CF6'
  }
];

const CANONICAL_SCENE_COUNT = 28;

const scenes = {

  /* ═══════════════════════════════════════════════════════════
     ACT I – The Seminar
  ═══════════════════════════════════════════════════════════ */

  scene_01: {
    id: 'scene_01',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_01.webp',
    tipus: 'text_block',
    titol: 'Arriving at the VII International Seminar',
    personatge: 'pau',
    narracio: 'Pau Martí, 28, a junior athletics coach, arrives at the VII International Seminar on Tactics and Sports Technique. He has driven 150 kilometers to attend the round table that will close the day.\n\nThe room is packed. On the board he reads the title of the session: "Motor Control: The Silent Key Factor in Sports Performance?" Three people take the stage: a researcher-moderator, a mountain trail and ultra-trail running coach, and an elite athlete.\n\nPau opens his notebook. He has spent years working on his athletes\' technique by intuition. Today he wants to understand the foundations.',
    dialeg: {
      personatge: 'pau',
      text: '"For once, I want to understand the why, not just the how." Pau takes a seat in the front row.'
    },
    seguent: 'scene_02'
  },

  scene_02: {
    id: 'scene_02',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_02.webp',
    tipus: 'text_block',
    titol: 'The Provocative Definition',
    personatge: 'raquel',
    narracio: 'Raquel Font-Lladó, researcher at the University of Girona, takes the stage and, instead of a conventional presentation, delivers a direct and deliberately broad definition to spark debate.',
    dialeg: {
      personatge: 'raquel',
      text: '"Motor control is the human capacity to produce movement and maintain posture. It is presented as an intrinsic potential of the individual. It must be developed through experience, responding to a motor objective, integrating information received from the environment and from one\'s own body. This relationship feeds back on itself and is interrelated to shape motor behavior."\n\nRaquel pauses and looks at the audience.\n\n"Do you agree?"'
    },
    seguent: 'scene_03'
  },

  scene_03: {
    id: 'scene_03',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_03.webp',
    tipus: 'text_block',
    titol: 'Three Major Theoretical Perspectives',
    personatge: 'raquel',
    narracio: 'To set the stage for debate, Raquel presents the three major schools of thought that psychology has used to explain motor control. None of them, she notes, is the absolute truth.',
    contingut_pedagogic: {
      titol: 'Three Major Perspectives on Motor Control',
      text: '**1. Behaviorists and Associationists** — Movement is a learned reaction through stimulus-response. Repeated training consolidates correct responses (Lawther, 1968; Rushall & Siedentop, 1972).\n**2. Cognitivists** — Movement comes from a motor program stored in memory, influenced by feedback. The individual builds internal representations (Adams, 1971; Schmidt, 1976; Meinel & Schnabel, 1988).\n**3. Ecological Systems** — Movement emerges from the holistic interaction between organism and environment. There is no central director: the system self-organizes (Bernstein, 1967; Kelso & Tuller, 1984; Thelen, 1987; Gibson).'
    },
    seguent: 'scene_04'
  },

  scene_04: {
    id: 'scene_04',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_04.webp',
    tipus: 'quiz',
    titol: 'Comprehension: Definition of Motor Control',
    personatge: 'raquel',
    punts: 10,
    narracio: 'Raquel looks at the audience and poses the first reflection question to verify understanding of the definition.',
    pregunta: 'Which aspect is fundamental to the definition of motor control that Raquel has presented?',
    opcions: [
      {
        id: 'A',
        text: 'The integration of information from the body and the environment to generate goal-directed movement',
        correcta: true,
        feedback: 'Exactly. Raquel\'s definition focuses on **integration**: information from the body (proprioceptive) + information from the environment (perceptual) + constant feedback, all in service of a specific motor objective. It is not about isolated strength or technical memory, but a circular system of information and response.'
      },
      {
        id: 'B',
        text: 'The muscular strength and cardiovascular endurance of the athlete',
        correcta: false,
        feedback: 'Strength and endurance are necessary physical conditions, but the definition of motor control points to something different: the **capacity to integrate information and generate motor responses adapted to the context**. An athlete can be very strong and have poor motor control if they do not perceive and integrate information from their environment well.'
      },
      {
        id: 'C',
        text: 'The memorization of ideal technical patterns to reproduce in competition',
        correcta: false,
        feedback: 'The memorization of patterns is characteristic of the **cognitivist** perspective (stored motor program), but Raquel\'s definition is broader: it includes dynamic interaction with the environment and constant feedback. Motor control is not reproducing a fixed pattern, but continuously adapting.'
      }
    ],
    seguent: 'scene_05'
  },

  scene_05: {
    id: 'scene_05',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_05.webp',
    tipus: 'text_block',
    titol: 'The Round Table Begins',
    personatge: 'narracio',
    narracio: 'Raquel introduces the three panelists who will share the table:\n\n— **Viktor**, a researcher specializing in the ecological perspective of motor control.\n— **Andrés**, a mountain trail and ultra-trail running coach working with elite athletes.\n— **Esther**, an elite endurance runner.\n\nThree perspectives on the same phenomenon: academic, coaching, and practical sports.',
    dialeg: {
      personatge: 'raquel',
      text: '"And without further ado, the first question I want us to address is: **What does motor control mean to you?**"\n\nRaquel looks at Viktor, inviting him to speak first.'
    },
    seguent: 'scene_06'
  },

  /* ═══════════════════════════════════════════════════════════
     ACT II – Three Perspectives
  ═══════════════════════════════════════════════════════════ */

  scene_06: {
    id: 'scene_06',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_06.webp',
    tipus: 'text_block',
    titol: 'The Ecological Perspective',
    personatge: 'joan',
    narracio: 'Viktor takes the floor from an academic standpoint. He opens with a definition but immediately qualifies it from his ecological perspective.',
    dialeg: {
      personatge: 'joan',
      text: '"For me, motor control is the capacity to perceive movement in, and for, the generation of motor responses."\n\nHe pauses and adds:\n\n"But specifically, from the ecological approach, motor control does not exist inherent to the individual, as it requires context. The individual and the environment form an inseparable system. Movement is not the output of an internal program: it emerges from the interaction."'
    },
    contingut_pedagogic: {
      titol: 'The Ecological Perspective (Gibson, Kelso, Thelen)',
      text: 'From ecological and systems theories, movement **emerges from the dynamic interaction** between the organism and the context. The environment provides information (affordances) that the motor system uses to generate adapted responses.\n\nThe central nervous system is no longer an isolated "director" but becomes part of a **circular process**: perception → response → perception. Bernstein (1967) was the first pioneer; Thelen, Kelso, and Gibson later consolidated it.'
    },
    seguent: 'scene_07'
  },

  scene_07: {
    id: 'scene_07',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_07.webp',
    tipus: 'text_block',
    titol: 'The Mountain Coach\'s Perspective',
    personatge: 'andres',
    narracio: 'Andrés reflects on what motor control means from his practical experience as a coach of mountain runners and skiers at the elite level.',
    dialeg: {
      personatge: 'andres',
      text: '"In mountain trail and ultra-trail sports, motor control is the perception and adaptation to what you are perceiving. In a sense, it is the reproduction of oneself in space and time."\n\nHe reflects for a moment and continues:\n\n"The athlete doesn\'t control the terrain: they adapt to it. Each step is a new motor decision, dictated by what the terrain offers at that precise moment."'
    },
    seguent: 'scene_08'
  },

  scene_08: {
    id: 'scene_08',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_08.webp',
    tipus: 'text_block',
    titol: 'The Athlete\'s Perspective',
    personatge: 'esther',
    narracio: 'Esther opens her contribution with lived experience. The room listens attentively because her perspective is closest to everyday reality.',
    dialeg: {
      personatge: 'esther',
      text: '"I don\'t enjoy working on it. My team and I associate it with modifying technique from a very broad concept: use of force, transfer, movement...\n\nIt requires seriousness, concentration, focusing on small details. It is complex; it demands working with all the senses fully engaged in the movement."'
    },
    contingut_pedagogic: {
      titol: 'Three Key Concepts That Emerge',
      text: 'From the three responses, three concepts emerge that will run through the entire round table:\n**Perception** — Viktor emphasizes perception as the basis of motor response. Without perceiving the environment, there is no adapted movement.\n**Adaptation** — Andrés focuses on permanent adaptation to a variable and unpredictable environment.\n**Technical Focus** — Esther highlights awareness and concentration on the details of one\'s own movement.'
    },
    seguent: 'scene_09'
  },

  scene_09: {
    id: 'scene_09',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_09.webp',
    tipus: 'quiz',
    titol: 'Comprehension: The Ecological Perspective',
    personatge: 'joan',
    punts: 10,
    narracio: 'Pau underlines Viktor\'s contribution. He tries to rephrase it to make sure he has understood it correctly.',
    pregunta: 'Viktor argues that, from the ecological perspective, motor control...',
    opcions: [
      {
        id: 'A',
        text: 'Does not exist inherent to the individual, because it requires context to exist and is generated in the interaction',
        correcta: true,
        feedback: 'Correct. This is the central idea of the ecological perspective: motor control is not a capacity we "have" independently of the environment. It **emerges from the interaction** between the organism and the context. This is why training in artificial or static environments does not develop real motor control.'
      },
      {
        id: 'B',
        text: 'Consists of reproducing pre-existing motor patterns stored in memory',
        correcta: false,
        feedback: 'This is the **cognitivist** perspective (Schmidt, Adams): movement comes from an internal motor program. The ecological perspective, on the contrary, holds that movement **emerges** from the organism-environment interaction at each moment, without the need for a prior program.'
      },
      {
        id: 'C',
        text: 'Is the perception and constant adaptation to the terrain, as Andrés explains in his example',
        correcta: false,
        feedback: 'Andrés\'s description is very close to ecology, but it mainly reflects his **practical experience** as a mountain coach. Viktor\'s academic definition is more radical: motor control does not exist inherent to the individual; it requires context to exist. It is not adaptation to terrain, it is emergence in the interaction.'
      }
    ],
    seguent: 'scene_10'
  },

  scene_10: {
    id: 'scene_10',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_10.webp',
    tipus: 'quiz',
    titol: 'Comprehension: Andrés\' View',
    personatge: 'andres',
    punts: 10,
    narracio: 'Pau thinks about his endurance athletes. Andrés\' view feels very familiar.',
    pregunta: 'Andrés describes motor control in skiing and mountain running as...',
    opcions: [
      {
        id: 'A',
        text: 'The perception and adaptation to what is perceived: reproducing oneself in space and time',
        correcta: true,
        feedback: 'Exactly. Andrés frames motor control as a **dynamic relationship** between what the athlete perceives and their adaptive response. In mountain sports, the environment changes constantly: motor control is the capacity to "be there" at each moment, adapted to what the terrain offers.'
      },
      {
        id: 'B',
        text: 'The ability to repeat the same perfect technical pattern regardless of the terrain',
        correcta: false,
        feedback: 'This would be the behaviorist or cognitivist view: reproducing a fixed pattern. Andrés argues precisely the opposite: motor control in the mountains is **constant adaptation**, not the reproduction of an unchanging pattern. The terrain is always new; the motor response must be too.'
      },
      {
        id: 'C',
        text: 'The strength and endurance needed to complete long-distance mountain events',
        correcta: false,
        feedback: 'Strength and endurance are important physical conditions, but Andrés is talking about a different dimension: the **motor perception and adaptation** to the environment. An athlete can be very fit and still struggle with motor control if they don\'t perceive and adapt well to the changing demands of the terrain.'
      }
    ],
    seguent: 'scene_11'
  },

  /* ═══════════════════════════════════════════════════════════
     ACT III – Biology, Experience & Context
  ═══════════════════════════════════════════════════════════ */

  scene_11: {
    id: 'scene_11',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_11.webp',
    tipus: 'text_block',
    titol: 'Biology as a Precondition',
    personatge: 'joan',
    narracio: 'Raquel poses the second major question of the round table. Viktor takes the floor without waiting to be asked.',
    dialeg: {
      personatge: 'raquel',
      text: '"In the development of motor control, **what role does biology, individual experience, and the complexity of context play?**"'
    },
    seguent: 'scene_12'
  },

  scene_12: {
    id: 'scene_12',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_12.webp',
    tipus: 'text_block',
    titol: 'Exploration as the Key to Learning',
    personatge: 'joan',
    narracio: 'Viktor is clear and direct in his stance. The answer, for him, is unambiguous.',
    dialeg: {
      personatge: 'joan',
      text: '"From my perspective, biology is **only a precondition** that acts as a facilitator for motor development. It is the individual\'s prior experiences that are the truly determining key factors of development, especially when they have generated learning.\n\nUltimately, the learning of motor control comes down to the exploration and achievement of new ways of interacting with space, time, and other individuals.\n\nTherefore, as coaches we must design contexts that favor this exploration."'
    },
    contingut_pedagogic: {
      titol: 'Three Perspectives: The Weight of Biology',
      text: 'The major theories disagree on which factor is determinant:\n**Associationists (Lawther, Gesell)**: Growth + Maturation + Learning = Development. Biology marks the evolutionary stages of motor control.\n**Cognitivists (Schmidt, Adams, Piaget)**: The subject is an active agent; they build representations from the interaction between the biological and the contextual.\n**Ecological Systems (Bernstein, Gibson, Thelen)**: Biology, environment, and experience co-determine movement in a circular manner, without a fixed hierarchy between them.'
    },
    seguent: 'scene_13'
  },

  scene_13: {
    id: 'scene_13',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_13.webp',
    tipus: 'quiz',
    titol: 'Comprehension: The Role of Biology',
    personatge: 'joan',
    punts: 10,
    narracio: 'Pau stops writing and reflects on Viktor\'s stance. He is clear on what was said, but wants to verify he has understood it correctly.',
    pregunta: 'What role does Viktor attribute to biology in the development of motor control?',
    opcions: [
      {
        id: 'A',
        text: 'It is a necessary precondition, but prior experiences and learning are the truly determining factors',
        correcta: true,
        feedback: 'Perfect. Viktor is explicit: biology is a **precondition** (it sets the initial limits), but it is not the determinant of motor development. What truly defines the level of motor control is lived experience and, above all, whether that experience has generated **real learning**: new ways of interacting with space, time, and others.'
      },
      {
        id: 'B',
        text: 'It completely determines the level of motor control an athlete can achieve',
        correcta: false,
        feedback: 'This would be the maturational or biological determinist position (Gesell, 1929). Viktor distances himself from it: biology is a **precondition**, not a final determinant. If it were determinant, training and experience would make no sense — and we know they do.'
      },
      {
        id: 'C',
        text: 'Biology and experience carry exactly the same weight in motor development',
        correcta: false,
        feedback: 'Viktor does not establish an equal balance. He states that biology is **subordinate** to experience: it is a precondition (a necessary base), but it is experience and learning that truly shapes motor control. The weight is not equal: experience prevails.'
      }
    ],
    seguent: 'scene_14'
  },

  scene_14: {
    id: 'scene_14',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_14.webp',
    tipus: 'text_block',
    titol: 'Competition: When Context Changes Everything',
    personatge: 'esther',
    narracio: 'Esther adds to Viktor and Andrés\'s contributions with a reflection grounded in her experience as an elite endurance runner.',
    dialeg: {
      personatge: 'esther',
      text: '"The competition situation presents high stress levels, magnifies external stimuli, and at some point fatigue appears at its most intense. All of this **modifies motor control**.\n\nTherefore, training technique without context makes no sense. Even though we sometimes work in a decontextualized way to focus attention better, we then make sure there is **transfer**."'
    },
    seguent: 'scene_15'
  },

  scene_15: {
    id: 'scene_15',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_15.webp',
    tipus: 'checklist',
    titol: 'Biology, Experience, Context: Key Concepts',
    personatge: 'raquel',
    narracio: 'Raquel summarizes the contributions of the three panelists. Pau must identify the 5 aspects that have been presented as true.',
    pregunta: 'Check the 5 true statements about the role of biology, experience, and context in motor control:',
    checklistItems: [
      { text: 'Biology is a necessary precondition, but experience is the truly determining factor',                     correcta: true  },
      { text: 'Motor learning is the exploration of new ways of interacting with space, time, and others',               correcta: true  },
      { text: 'We must design contexts that favor the athlete\'s motor exploration',                                      correcta: true  },
      { text: 'Training in contexts similar to competition is essential to ensure transfer',                              correcta: true  },
      { text: 'The athlete and coach must work together to improve the motor pattern',                                    correcta: true  },
      { text: 'Biology definitively determines motor capabilities throughout an athlete\'s entire life',                  correcta: false },
      { text: 'Reproducing the ideal technical model exactly always guarantees maximum performance',                      correcta: false }
    ],
    feedback: 'The two incorrect statements represent common errors: biology is a precondition, not a permanent determinant; and the "ideal technical model" is a reference, not a universal formula. Technique must adapt to the individual\'s biology and the competition context.',
    seguent: 'scene_16'
  },

  /* ═══════════════════════════════════════════════════════════
     ACT IV – Learning & Awareness
  ═══════════════════════════════════════════════════════════ */

  scene_16: {
    id: 'scene_16',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_16.webp',
    tipus: 'text_block',
    titol: 'What Do We Learn When Working on Motor Control?',
    personatge: 'esther',
    narracio: 'Raquel poses the third round table question. Esther looks up at the ceiling for a moment, as if transporting herself back to her training sessions, and responds from direct experience.',
    dialeg: {
      personatge: 'raquel',
      text: '"What do we learn when we work on motor control? Images, sensations, ideas, processes?"'
    },
    seguent: 'scene_17'
  },

  scene_17: {
    id: 'scene_17',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_17.webp',
    tipus: 'quiz',
    titol: 'Comprehension: Learning About One\'s Own Body',
    personatge: 'esther',
    punts: 10,
    narracio: 'Esther responds from her own experience as an athlete. The room listens attentively.',
    pregunta: 'Esther points out that working on motor control allows her, beyond improving performance, something that surprises the panel. What is it?',
    opcions: [
      {
        id: 'A',
        text: 'To know her body better, feel more secure, and indirectly prevent injuries by correcting harmful patterns',
        correcta: true,
        feedback: 'Exactly. Esther opens a new dimension: motor control work **is not solely associated with direct performance**. When motor control work allows the correction of injury-causing technical patterns, it contributes indirectly to **injury prevention**. Viktor adds: "This confirms that motor control is perception." If feeling secure is the outcome, it means perception of one\'s own movement has improved.'
      },
      {
        id: 'B',
        text: 'To increase maximum strength and power output in high-intensity training sessions',
        correcta: false,
        feedback: 'Strength is an independent physical capacity from motor control, though the two are related. What Esther highlights is something different: **bodily self-knowledge**, confidence in one\'s own movement, and the ability to run without harmful patterns. It is not strength: it is perception and adaptation.'
      },
      {
        id: 'C',
        text: 'To memorize movement sequences that can be reproduced automatically in competition',
        correcta: false,
        feedback: 'Memorizing sequences is a **cognitivist** mechanism (motor program). Esther, on the contrary, talks about **body awareness**, sensation, security. She learns to sense her body, not to execute a memorized sequence. Viktor himself questions whether what is learned consciously transfers to the unconscious behavior of competition.'
      }
    ],
    seguent: 'scene_18'
  },

  scene_18: {
    id: 'scene_18',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_18.webp',
    tipus: 'text_block',
    titol: 'Viktor\'s Challenge: Is It Really Learned?',
    personatge: 'joan',
    narracio: 'Viktor intervenes with a question that generates a great silence in the room. Pau stops taking notes.',
    dialeg: {
      personatge: 'joan',
      text: '"What do you learn? —he throws the question into the air—. **Is motor control really learned?** Science disagrees; it has doubts.\n\nPlacing awareness on movement can generate a **false sense of having learned**. But is it really like that? I ask this because in competition, when the athlete loses focus on the movement, what had been learned disappears. Or perhaps, in reality, it had never been learned at all?"'
    },
    contingut_pedagogic: {
      titol: 'Internal Focus vs External Focus',
      text: 'The debate about awareness in motor control raises fundamental questions:\n**Internal focus** — The athlete places attention on the sensations of their own body. It can generate self-knowledge, but in competition it can interfere with the fluidity of movement.\n**External focus** — The athlete focuses on the effect of the movement on the environment. Many studies show better results in actual performance.\n**The paradox of awareness**: what is learned with conscious focus does not always transfer to the unconscious behavior of competition.'
    },
    seguent: 'scene_19'
  },

  scene_19: {
    id: 'scene_19',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_19.webp',
    tipus: 'quiz',
    titol: 'Viktor\'s Challenge on Awareness',
    personatge: 'joan',
    punts: 10,
    narracio: 'Esther nods. And Andrés reinforces Viktor\'s idea from his experience with athletes in competition.',
    pregunta: 'Viktor questions whether motor control is really learned. His central argument is...',
    opcions: [
      {
        id: 'A',
        text: 'That in competition, when the athlete loses conscious focus on the movement, what was learned can disappear, suggesting it may never have been truly learned at all',
        correcta: true,
        feedback: 'Exactly. Viktor identifies a **fundamental paradox**: if the change depends on conscious focus to be maintained, perhaps it was not real learning but a temporary change. Deep motor learning manifests in competition **without the need for conscious focus**. When the focus disappears and the change does too, we must ask whether there was real learning or merely superficial adaptation.'
      },
      {
        id: 'B',
        text: 'That the athlete cannot objectively measure their progress in motor control',
        correcta: false,
        feedback: 'Measurement is a different topic (which they will address in the next scene). Viktor\'s doubt is about the **nature of learning**: whether what was learned is real or an illusion of learning generated by temporary conscious awareness of movement. It is not a measurement problem, it is a transfer problem.'
      },
      {
        id: 'C',
        text: 'That biology limits the capacity to learn new motor patterns in adult athletes',
        correcta: false,
        feedback: 'Viktor had already made clear that biology is a precondition, not a definitive limit. His doubt is about **awareness and transfer**: does what is learned consciously (internal focus) transfer to the unconscious behavior of competition? This is the question that science, Viktor says, has not definitively answered.'
      }
    ],
    seguent: 'scene_20'
  },

  scene_20: {
    id: 'scene_20',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_20.webp',
    tipus: 'text_block',
    titol: 'Testing, Feedback, and Sensations',
    personatge: 'andres',
    narracio: 'Andrés reflects on what Viktor said and adds his own experience as a coach. The debate about awareness leads him to talk about the role of feedback and testing.',
    dialeg: {
      personatge: 'andres',
      text: '"During training, I test my athletes a lot. I feel it **gives them confidence** about their own potential, and also about my work.\n\nI also ask my athletes a lot about the sensations they feel; not everything can be observed. Sometimes I need them to tell me, for example, \'whether they\'re pulling from the glute or the hamstring\'.\n\nIn competition, I ask athletes to focus on technique for two objectives: to distract them from mental fatigue, and to make them combine their legs well in very specific situations. Without focus, unconsciously, they always approach the climb with the same leg, the stronger one, and that affects efficiency."'
    },
    seguent: 'scene_21'
  },

  /* ═══════════════════════════════════════════════════════════
     ACT V – Measuring & Planning
  ═══════════════════════════════════════════════════════════ */

  scene_21: {
    id: 'scene_21',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_21.webp',
    tipus: 'text_block',
    titol: 'Measuring Motor Control',
    personatge: 'esther',
    narracio: 'Raquel poses the fourth question. Esther opens her eyes wide. She has a very clear answer.',
    dialeg: {
      personatge: 'raquel',
      text: '"Can motor control be measured? What evidence do we use?"'
    },
    seguent: 'scene_22'
  },

  scene_22: {
    id: 'scene_22',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_22.webp',
    tipus: 'quiz',
    titol: 'Comprehension: Measuring Motor Control',
    personatge: 'esther',
    punts: 10,
    narracio: 'Esther responds with conviction: "Of course we can measure it! We need to measure it!" She explains that they use tests, race and training recordings to analyze later. But she adds an important nuance that Pau immediately underlines.',
    pregunta: 'Esther points out that the **real indicators** of a change in motor pattern are analyzed when...',
    opcions: [
      {
        id: 'A',
        text: 'They appear repeatedly over time, as fleeting changes (that appear and disappear) do not confirm learning',
        correcta: true,
        feedback: 'Exactly. Esther establishes a fundamental distinction: there are **fleeting changes** (that appear and disappear) and **real changes** (that consolidate). The former help understand the factors that influence technique (fatigue, stress, motivation), but they do not indicate learning. Only **repetition over time** confirms that a real pattern change has occurred.'
      },
      {
        id: 'B',
        text: 'The athlete subjectively expresses feeling better in a single test session',
        correcta: false,
        feedback: 'Subjective sensation is valuable (Andrés will reinforce this), but it is not sufficient to confirm a pattern change. Esther is clear: sometimes "there are better and worse days." A single session with a good feeling is not evidence of real change. **Temporal repetition** of the change is needed to confirm learning.'
      },
      {
        id: 'C',
        text: 'An improvement in the stopwatch is observed in a single specific motor control training session',
        correcta: false,
        feedback: 'The stopwatch is a measure of overall performance, not necessarily of specific motor control. Furthermore, Esther makes clear that progress "is not linear": there are better and worse days. A single point-in-time improvement does not confirm a pattern change. We must **analyze when changes appear repeatedly**.'
      }
    ],
    seguent: 'scene_23'
  },

  scene_23: {
    id: 'scene_23',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_23.webp',
    tipus: 'text_block',
    titol: 'How to Plan: The Final Question',
    personatge: 'raquel',
    narracio: 'Raquel poses the fifth and final question of the round table. For Pau, a working coach, this is the most important one of all.',
    dialeg: {
      personatge: 'raquel',
      text: '"And now comes the question that many of you have had on your lips from the start: **What would be the key aspects for planning training sessions focused on improving motor control?**"\n\nThe three panelists look at each other. They yield the floor to Andrés.'
    },
    contingut_pedagogic: {
      titol: 'Andrés\' Planning Process: Four Steps',
      text: 'Andrés structures his response clearly:\n**1. Know the athlete** — Tests and interviews: motor experience, musculoskeletal morphology, metabolic efficiency, previous injuries, objectives.\n**2. Plan the process** — Develop the improvement plan and share it with the athlete to analyze and modify it if necessary.\n**3. Expose to contexts** — Apply the load and exercises for pattern improvement, in simulated and real contexts.\n**4. Give meaningful feedback** — Individualized feedback, adjusted to the athlete\'s current moment.'
    },
    seguent: 'scene_24'
  },

  scene_24: {
    id: 'scene_24',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_24.webp',
    tipus: 'decisio',
    titol: 'The Decision: Where Do We Start?',
    personatge: 'pau',
    punts: 10,
    narracio: 'Pau drives home thinking about one of his athletes, a 24-year-old long-distance runner who has an inefficient and injury-causing running pattern. He must do something. Andrés has inspired him.',
    dialeg: {
      personatge: 'andres',
      text: '"Faced with a blank page, I always start by getting to know my athlete: knowing where we\'re starting from."'
    },
    pregunta: 'Pau\'s goal is to improve his runner\'s motor control. What should the first step be?',
    opcions: [
      {
        id: 'A',
        punts: 10,
        text: 'Conduct tests and interviews to know the athlete: motor experience, morphology, injury history, and objectives',
        feedback: 'Perfect. Andrés explained it clearly: "faced with a blank page, start by knowing your athlete." Without knowing the starting point — the current pattern, the margin of adaptation, previous injuries — any plan is a blind imposition. Knowledge of the athlete is the foundation of the entire process.',
        seguent: 'scene_25'
      },
      {
        id: 'B',
        punts: 0,
        text: 'Directly apply a motor pattern improvement exercise program selected from the scientific literature',
        feedback: 'Andrés is categorical: "I would be wrong if I tried to generate a completely new pattern without starting from the current pattern." Applying exercises without knowing the athlete is an imposition that can generate physical and psychological interference. First you need to know where you\'re starting from.',
        seguent: 'scene_24b'
      },
      {
        id: 'C',
        punts: 0,
        text: 'Design a program based on the ideal technical model of the long-distance runner',
        feedback: 'The ideal technical model is a reference, not a program. Andrés warns that "modifying a pattern requires many hours of dedication" and that "in many competition situations the original pattern overrides the learned one." Without knowing the athlete\'s current pattern, the ideal model is a destination without a map.',
        seguent: 'scene_24b'
      }
    ],
    seguent: 'scene_25'
  },

  scene_24b: {
    id: 'scene_24b',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_25.webp',
    tipus: 'text_block',
    titol: 'Andrés\' Correction',
    personatge: 'andres',
    narracio: 'Andrés smiles understandingly when Pau explains his idea during the break. It happened to him too at the start of his career.',
    dialeg: {
      personatge: 'andres',
      text: '"At first I also thought that good exercises would be enough. Over time I learned that the engine of change is not the exercise, but **knowing the athlete**.\n\nIdentify their current pattern, their margin of adaptation, and then design. Otherwise, any proposal is a blind imposition that can generate resistance or, worse, an injury."'
    },
    seguent: 'scene_25'
  },

  scene_25: {
    id: 'scene_25',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_26.webp',
    tipus: 'text_block',
    titol: 'The Complete Planning Process',
    personatge: 'andres',
    narracio: 'Andrés details, step by step, the process he follows when planning the improvement of an athlete\'s motor control.',
    dialeg: {
      personatge: 'andres',
      text: '"Once you know the athlete, you **plan the improvement process** considering everything you\'ve gathered: motor experience, morphology, metabolism, injuries, objectives, personal circumstances.\n\nThen you present your point of view and your plan to the athlete to analyze and modify it if necessary. **The athlete is in charge.**\n\nThe last step is applying the load. With experience, I\'ve learned that if the proposals generate physical or psychological interference, it\'s important to be able to **step back**. On many occasions, this paradox is the key to success: knowing when to stop is as important as knowing when to move forward."'
    },
    seguent: 'scene_26'
  },

  /* ═══════════════════════════════════════════════════════════
     EPILOGUE – Synthesis & Conclusion
  ═══════════════════════════════════════════════════════════ */

  scene_26: {
    id: 'scene_26',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_27.webp',
    tipus: 'checklist',
    titol: 'Keys to Planning Motor Control Work',
    personatge: 'raquel',
    narracio: 'Raquel closes the round table. Pau, inspired, writes down the principles he\'ll take away. He must identify the 5 key aspects for planning motor control work in elite sports.',
    pregunta: 'Check the 5 true principles for planning motor control work in elite sports:',
    checklistItems: [
      { text: 'Know the athlete: motor experience, morphology, injury history, and objectives',                                correcta: true  },
      { text: 'Plan the process and share it with the athlete to analyze and modify it jointly',                               correcta: true  },
      { text: 'Expose the athlete to simulated and real competition contexts',                                                  correcta: true  },
      { text: 'Provide meaningful, individualized feedback adjusted to the athlete\'s current moment',                         correcta: true  },
      { text: 'Know when to step back if the proposals generate physical or psychological interference',                       correcta: true  },
      { text: 'Apply a standard program based on the ideal technical model of the sport for all athletes',                     correcta: false },
      { text: 'Minimize dialogue with the athlete to maintain the coach\'s objectivity',                                       correcta: false }
    ],
    feedback: 'The two incorrect items represent common mistakes: a standard program does not adapt to the individual (personalization is needed); and constant dialogue between coach and athlete is fundamental, not a bias. Andrés is clear: "the athlete is in charge."',
    seguent: 'scene_27'
  },

  scene_27: {
    id: 'scene_27',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_28.webp',
    tipus: 'epilog',
    titol: 'Final Synthesis – Results',
    personatge: 'raquel',
    narracio: 'The round table has closed. Raquel gives one final summary and Pau receives his assessment.',
    mentorMsgs: {
      excellent: '**Excellent, Pau.** You have assimilated the fundamental principles of motor control with deep understanding. You have grasped the distinction between theoretical perspectives, the role of biology as a precondition, the paradox of awareness in competition, and the importance of knowing your athlete before planning. Your long-distance runner will have a coach who thinks, listens, and adapts. Raquel tells you: **"Now you know where to look for the answers."**',
      good: '**Well done, Pau.** You have grasped the core concepts and made most of the right decisions. Remember that motor control has no single theory: your challenge as a coach is to **dialogue between theory and practice** with a critical eye. Each athlete is a unique system; no standard program will give you the answers you need.',
      ok: '**Good effort, Pau.** You understood the basic concepts, but at times you fell back into reductionist views: motor control as strength, or as memorization of patterns. Review the distinction between theoretical perspectives and, above all, Viktor\'s idea: motor control **emerges from the interaction** with the environment. It is not an internal program to be installed.',
      needsWork: '**You need to go deeper, Pau.** The principles of motor control represent a profound shift in perspective from conventional training. We recommend re-reading Viktor\'s contributions (ecological perspective), Andrés\'s (know the athlete), and Esther\'s (body awareness and injury prevention). The key is accepting that **there is no single theory** and that your role is to dialogue between theory and practice with consistency and ethics.'
    },
    seguent: 'scene_28'
  },

  scene_28: {
    id: 'scene_28',
    imatge: 'https://raw.githubusercontent.com/josepespar/scorm_control_motor-img-/main/raquel_escena_29.webp',
    tipus: 'text_block',
    titol: 'The Coach\'s Critical Perspective',
    personatge: 'narracio',
    narracio: 'Pau leaves the seminar with his notebook full of notes. But above all, with three new questions he didn\'t have when he walked in: What is my athlete\'s motor control like? Where are we starting from? What contexts am I designing for them?\n\nViktor closed with an idea that has stayed with Pau:',
    dialeg: {
      personatge: 'joan',
      text: '"The considerations for improving motor control must be completely different for elite athletes versus young athletes in developmental stages. But in general, **motor control should give you information about how you learn so you can keep learning**.\n\nThe coach should design experiences close to the athlete\'s limits, and be available to provide the necessary support at each moment. Working close to the limits involves managing many emotions, but also moving in the space between what I know how to do and what I still need to learn."'
    },
    contingut_pedagogic: {
      titol: 'Three Principles Pau Now Brings to Every Training Session',
      text: '**1. There is no single theory of motor control** — Critically and coherently dialoguing between the behaviorist, cognitivist, and ecological perspectives is the professional coach\'s job.\n**2. Athlete first, program second** — Knowing the starting point, limits, and circumstances of the athlete is step zero of any motor control intervention.\n**3. Accept uncertainty as part of the process** — Motor pattern changes are not linear, and some modifications will be fleeting. The coach who accepts uncertainty and regression as part of the process works with much greater intelligence.'
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
      var done   = stage.scenes.every(function (s) { return state.visitedScenes.indexOf(s) !== -1; });
      var active = stage.scenes.indexOf(state.currentScene) !== -1;
      var cls    = done ? 'map-done' : active ? 'map-active' : 'map-pending';
      var status = done ? '✓ Completed' : active ? '▶ In progress' : '○ Pending';
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
      _qs('pedagogic-title').textContent   = pb.titol || '';
      _qs('pedagogic-text').innerHTML      = _md(pb.text || '');
      _qs('pedagogic-block').style.display = 'block';
    }

    var ia = _qs('interaction-area');
    if (scene.tipus === 'text_block') {
      if (scene.seguent) {
        var btn = document.createElement('button');
        btn.className = 'btn btn-primary btn-enabled';
        btn.textContent = 'Continue →';
        btn.onclick = function () { Engine.showScene(scene.seguent); };
        ia.appendChild(btn);
      } else {
        var restartBtn = document.createElement('button');
        restartBtn.className = 'btn btn-primary btn-enabled';
        restartBtn.textContent = 'Restart course';
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
    label.textContent = 'Comprehension question';
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
          if (o.correcta)        b.className = 'btn btn-quiz btn-quiz-correct btn-disabled';
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
        contBtn.textContent = 'Continue →';
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
    label.textContent = 'Decision';
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
    label.textContent = 'Checklist';
    ia.appendChild(label);

    if (scene.pregunta) {
      var q = document.createElement('div');
      q.className = 'quiz-question';
      q.innerHTML = _md(scene.pregunta);
      ia.appendChild(q);
    }

    var checked   = {};
    var submitted = false;
    var itemsDiv  = document.createElement('div');
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
    submitBtn.textContent = 'Check →';
    ia.appendChild(submitBtn);

    submitBtn.onclick = function () {
      if (submitted) return;
      submitted = true;
      submitBtn.style.display = 'none';
      var pts = 0;
      scene.checklistItems.forEach(function (item, idx) {
        var userChecked = !!checked[idx];
        var correct     = item.correcta;
        var div         = itemsDiv.children[idx];
        div.style.pointerEvents = 'none';
        if (userChecked && correct)        { div.classList.add('cl-correct');   pts += 2; }
        else if (!userChecked && !correct) { div.classList.add('cl-correct');   pts += 2; }
        else if (userChecked && !correct)  { div.classList.add('cl-incorrect'); }
        else                               { div.classList.add('cl-missed'); }
      });
      state.score += pts;
      state.decisions.push({ scene: scene.id, pts: pts, label: scene.titol });
      _save();
      _qs('score-display').textContent = state.score + ' pts';
      var fbDiv = document.createElement('div');
      fbDiv.className = 'checklist-feedback';
      fbDiv.innerHTML = '<strong>' + pts + '/10 pts</strong> — ' + _md(scene.feedback || 'Review the highlighted items.');
      ia.appendChild(fbDiv);
      var contBtn = document.createElement('button');
      contBtn.className = 'btn btn-primary btn-enabled';
      contBtn.style.marginTop = '8px';
      contBtn.textContent = 'Continue →';
      contBtn.onclick = function () { Engine.showScene(scene.seguent); };
      ia.appendChild(contBtn);
    };
  }

  function _renderEpilog(scene, ia) {
    var maxScore      = 100;
    var circumference = 2 * Math.PI * 54;
    var offset        = circumference * (1 - state.score / maxScore);
    var grade, gradeClass, mentorMsg;

    if (state.score >= 90) {
      grade      = 'Excellent';
      gradeClass = 'epilogue-excellent';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.excellent  : '';
    } else if (state.score >= 70) {
      grade      = 'Well done';
      gradeClass = 'epilogue-good';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.good       : '';
    } else if (state.score >= 50) {
      grade      = 'Good effort';
      gradeClass = 'epilogue-ok';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.ok         : '';
    } else {
      grade      = 'Needs revision';
      gradeClass = 'epilogue-needs-work';
      mentorMsg  = scene.mentorMsgs ? scene.mentorMsgs.needsWork  : '';
    }

    var mentor = CHARACTERS.raquel;
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
    html += '<p class="epilogue-mentor-msg">' + _md(mentorMsg) + '</p>';
    html += '</div>';

    if (state.decisions.length > 0) {
      html += '<div><div class="epilogue-section-title">Decision summary</div>';
      html += '<ul class="decisions-summary">';
      state.decisions.forEach(function (d) {
        var cls  = d.pts >= 10 ? 'ds-good' : d.pts >= 5 ? 'ds-ok' : 'ds-bad';
        var icon = d.pts >= 10 ? '✓' : d.pts >= 5 ? '~' : '✗';
        html += '<li class="' + cls + '"><span class="ds-icon">' + icon + '</span>';
        html += '<span>' + (d.label || d.scene) + '</span>';
        html += '<span class="ds-pts">' + d.pts + ' pts</span></li>';
      });
      html += '</ul></div>';
    }

    html += '<div style="margin-top:16px;">';
    html += '<button class="btn btn-primary btn-enabled" onclick="Engine.showScene(\'scene_28\')">';
    html += 'Continue →</button></div>';
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
      if (!scene) { console.warn('[Engine] Scene not found:', id); return; }
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

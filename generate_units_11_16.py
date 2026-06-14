"""
ELT Grammar Course Designer — Units 11-16 Generator
Generates Level 1 Units 11-16 (DOCX + PPTX) covering Past Simple Regular and
Irregular Verbs, Past Simple Negative and Questions, Past Continuous,
Past Continuous vs Past Simple, Present Perfect Introduction, and
Present Perfect just/already/yet.
"""

import sys
import os

sys.path.insert(0, '/projects/sandbox/APP-GENERATION-PPT-FILE-ETC')

from generate_docx import build_docx
from generate_pptx import build_pptx

# ─────────────────────────────────────────────────────────────────────────────
# UNIT DATA — Units 11 through 16
# ─────────────────────────────────────────────────────────────────────────────

UNITS_11_16 = []

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 11 — Past Simple: Regular and Irregular Verbs (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "11",
    "cefr": "A1",
    "topic": "Past Simple — Regular and Irregular Verbs",
    "sub_focus": "Regular verbs + -ed (worked, played, studied); irregular verbs (went, had, saw, came, took, made, got, gave); pronunciation of -ed (/t/, /d/, /ɪd/); time markers (yesterday, last week, ago)",
    "objectives": [
        "Form the Past Simple affirmative of regular verbs by adding -ed.",
        "Apply spelling rules for -ed: double consonant (stopped), -e verbs (lived), consonant + y (studied).",
        "Use common irregular past forms correctly (went, had, saw, came, took, made, got, gave).",
        "Identify time markers that signal Past Simple (yesterday, last week, two days ago).",
        "Produce sentences about completed past actions and events.",
    ],
    "warmup": {
        "name": "🕰️ What Did You Do Yesterday?",
        "time": "5 minutes",
        "steps": [
            "Think of five things you did yesterday. Write them quickly on a piece of paper.",
            "Tell your partner: 'I woke up early. I walked to school. I studied English.'",
            "Your partner listens and counts: how many regular verbs? How many irregular?",
        ],
    },
    "lead_in": [
        "What did you do last weekend? Did you stay at home or go out?",
        "Think about yesterday evening. What happened? Tell your partner.",
    ],
    "vocab": [
        ("walked",      "verb",      "A1", "past of walk",                     "I walked to school yesterday."),
        ("played",      "verb",      "A1", "past of play",                     "They played football last week."),
        ("worked",      "verb",      "A1", "past of work",                     "She worked late on Monday."),
        ("studied",     "verb",      "A1", "past of study",                    "He studied all night."),
        ("stopped",     "verb",      "A1", "past of stop",                     "The bus stopped suddenly."),
        ("lived",       "verb",      "A1", "past of live",                     "We lived in Paris for two years."),
        ("wanted",      "verb",      "A1", "past of want",                     "I wanted a new phone."),
        ("needed",      "verb",      "A1", "past of need",                     "She needed more time."),
        ("watched",     "verb",      "A1", "past of watch",                    "They watched a film last night."),
        ("listened",    "verb",      "A1", "past of listen",                   "He listened to music."),
        ("went",        "verb",      "A1", "past of go (irregular)",           "I went to the park."),
        ("had",         "verb",      "A1", "past of have (irregular)",         "She had breakfast at seven."),
        ("saw",         "verb",      "A1", "past of see (irregular)",          "We saw a great film."),
        ("came",        "verb",      "A1", "past of come (irregular)",         "He came home late."),
        ("took",        "verb",      "A1", "past of take (irregular)",         "She took the bus."),
        ("made",        "verb",      "A1", "past of make (irregular)",         "They made a cake together."),
        ("got",         "verb",      "A1", "past of get (irregular)",          "I got a new job last month."),
        ("gave",        "verb",      "A1", "past of give (irregular)",         "He gave me a present."),
        ("ate",         "verb",      "A1", "past of eat (irregular)",          "We ate pizza for dinner."),
        ("drank",       "verb",      "A1", "past of drink (irregular)",        "She drank coffee this morning."),
        ("bought",      "verb",      "A1", "past of buy (irregular)",          "He bought a new car."),
        ("thought",     "verb",      "A1", "past of think (irregular)",        "I thought it was good."),
        ("told",        "verb",      "A1", "past of tell (irregular)",         "She told me a story."),
        ("found",       "verb",      "A1", "past of find (irregular)",         "They found the keys."),
        ("left",        "verb",      "A1", "past of leave (irregular)",        "He left early yesterday."),
        ("yesterday",   "adverb",    "A1", "the day before today",            "I walked to work yesterday."),
        ("ago",         "adverb",    "A1", "in the past; before now",         "That was three days ago."),
        ("last",        "adjective", "A1", "the most recent before now",      "We met last Friday."),
        ("suddenly",    "adverb",    "A1", "quickly and unexpectedly",        "The car stopped suddenly."),
        ("finally",     "adverb",    "A1", "at last; after a long time",      "We finally arrived home."),
    ],
    "grammar": {
        "rule": "For regular verbs, add -ed to the base form. Irregular verbs have unique past forms that must be memorised.",
        "form_headers": ["Type", "Rule", "Examples"],
        "form_rows": [
            ["Regular: most verbs", "base + -ed",             "worked, played, listened"],
            ["Regular: ends in -e", "base + -d",              "lived, liked, used"],
            ["Regular: consonant + y", "y to -ied",           "studied, tried, worried"],
            ["Regular: CVC (short vowel)", "double + -ed",    "stopped, planned, dropped"],
            ["Irregular", "unique form (memorise)",            "go-went, have-had, see-saw"],
        ],
        "form_col_widths": [2400, 2400, 3400],
        "note": "There is NO rule for irregular verbs. You must learn them by heart. Common ones: go-went, have-had, see-saw, come-came, take-took, make-made, get-got, give-gave, eat-ate, drink-drank, buy-bought.",
        "usage": [
            "Completed actions: I walked to school yesterday.",
            "Past events: She saw a great film last night.",
            "Sequences: He woke up, had breakfast, and went to work.",
            "Time markers: two days ago, last Monday, yesterday morning.",
            "Life events: They got married in 2019. She lived in Japan for a year.",
        ],
    },
    "guided": {
        "instruction": "Write the Past Simple form of the verb in brackets. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ (walk) to school yesterday.", "I WALKED to school yesterday."],
            ["2 ✅", "She ___ (go) to the cinema last night.", "She WENT to the cinema last night."],
            ["3",   "He ___ (study) English for two hours.", "________________________"],
            ["4",   "They ___ (have) dinner at eight.", "________________________"],
            ["5",   "We ___ (watch) a film on Saturday.", "________________________"],
            ["6",   "I ___ (buy) a new phone last week.", "________________________"],
            ["7",   "She ___ (stop) the car suddenly.", "________________________"],
            ["8",   "He ___ (give) me his number.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete sentences in the Past Simple using the prompts.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / walk / to the park / yesterday)",     "________________________"],
            ["2", "(she / see / her friend / last week)",     "________________________"],
            ["3", "(they / eat / pizza / on Friday)",         "________________________"],
            ["4", "(he / come / home / late / last night)",   "________________________"],
            ["5", "(we / make / a cake / for her birthday)",  "________________________"],
            ["6", "(I / take / the bus / two days ago)",      "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "He goed to school yesterday.", "________________________"],
            ["2", "She studyed all night.", "________________________"],
            ["3", "They taked the bus.", "________________________"],
            ["4", "I haved breakfast at seven.", "________________________"],
            ["5", "We stoped at the traffic light.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I go to work by bus. (change to yesterday)", "________________________"],
            ["2", "She watches TV every evening. (change to last night)", "________________________"],
            ["3", "They eat lunch at 1. (change to yesterday)", "________________________"],
            ["4", "He makes breakfast. (change to this morning)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Weekend Story Chain",
        "instructions": [
            "In groups of 4, take turns adding a sentence about last weekend.",
            "Use Past Simple: 'I woke up late. I had breakfast. I went shopping.'",
            "Each person adds one sentence. Keep the story going for 2 rounds.",
            "Report the funniest or most interesting story to the class.",
        ],
    },
    "reading": {
        "text": (
            "Last Saturday, Maria woke up early. She had breakfast and then walked to the "
            "market. She bought some fruit and vegetables. In the afternoon, she made a "
            "big lunch for her family. Her brother came at two o'clock. They ate together "
            "and talked for hours. It was a lovely day."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Maria woke up late on Saturday.",     "___"],
            ["2", "She walked to the market.",           "___"],
            ["3", "She bought some meat.",               "___"],
            ["4", "Her brother came at two o'clock.",    "___"],
            ["5", "They ate dinner together.",           "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about what you did last weekend using the Past Simple.",
        "prompts": [
            "What time did you wake up? I woke up at ___.",
            "Where did you go? I went to ___.",
            "What did you eat? I ate / had ___.",
            "Who did you see? I saw ___.",
            "What did you buy? I bought ___.",
            "What did you do in the evening? I ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 10 sentences about a memorable day in the past. Use at least 5 regular "
            "verbs and 5 irregular verbs in the Past Simple. Underline the irregular verbs."
        ),
        "item2": (
            "Learn the past forms of these 15 irregular verbs: go, have, see, come, take, "
            "make, get, give, eat, drink, buy, think, tell, find, leave. Write each one in a sentence."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 12 — Past Simple: Negative and Questions (A1-A2)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "12",
    "cefr": "A1-A2",
    "topic": "Past Simple — Negative and Questions",
    "sub_focus": "didn't + base form; Did + subject + base form?; Wh- questions in the past; short answers (Yes, I did / No, she didn't)",
    "objectives": [
        "Form negative sentences with didn't (did not) + base form of the verb.",
        "Understand that the main verb returns to the base form after didn't.",
        "Form Yes/No questions with Did + subject + base verb?",
        "Construct Wh- questions in the Past Simple (What did you do? Where did he go?).",
        "Give short answers: Yes, I did. / No, she didn't.",
    ],
    "warmup": {
        "name": "🔍 Did You or Didn't You?",
        "time": "5 minutes",
        "steps": [
            "Teacher reads statements: 'You watched TV last night. You ate breakfast today.'",
            "If TRUE for you, stand up and say: 'Yes, I did!' If FALSE, stay seated and say: 'No, I didn't!'",
            "Ask a standing classmate a follow-up Wh- question: 'What did you watch?'",
        ],
    },
    "lead_in": [
        "Did you sleep well last night? What time did you go to bed?",
        "What didn't you do yesterday that you usually do every day?",
    ],
    "vocab": [
        ("forget",     "verb",      "A1", "to fail to remember",              "I didn't forget your birthday."),
        ("remember",   "verb",      "A1", "to keep in your memory",           "Did you remember to call?"),
        ("invite",     "verb",      "A1", "to ask someone to come",           "She didn't invite him."),
        ("arrive",     "verb",      "A1", "to reach a place",                 "Did they arrive on time?"),
        ("decide",     "verb",      "A1", "to make a choice",                 "He didn't decide yet."),
        ("enjoy",      "verb",      "A1", "to get pleasure from something",   "Did you enjoy the party?"),
        ("happen",     "verb",      "A1", "to take place; occur",             "What happened yesterday?"),
        ("stay",       "verb",      "A1", "to remain in a place",             "I didn't stay long."),
        ("leave",      "verb",      "A1", "to go away from a place",          "When did she leave?"),
        ("pay",        "verb",      "A1", "to give money for something",      "He didn't pay the bill."),
        ("answer",     "verb",      "A1", "to give a reply",                  "She didn't answer my call."),
        ("finish",     "verb",      "A1", "to complete something",            "Did you finish your homework?"),
        ("pass",       "verb",      "A1", "to succeed in a test/exam",        "He didn't pass the exam."),
        ("fail",       "verb",      "A1", "to not succeed",                   "Did she fail the test?"),
        ("rain",       "verb",      "A1", "water falling from clouds",        "It didn't rain yesterday."),
        ("cost",       "verb",      "A1", "to have a price",                  "How much did it cost?"),
        ("break",      "verb",      "A1", "to damage so it's in pieces",     "Who broke the window?"),
        ("lose",       "verb",      "A1", "to not be able to find",           "Did you lose your keys?"),
        ("win",        "verb",      "A1", "to be the best in a competition",  "They didn't win the game."),
        ("spend",      "verb",      "A1", "to use money or time",             "How much did you spend?"),
        ("abroad",     "adverb",    "A1", "in or to a foreign country",       "Did you go abroad?"),
        ("instead",    "adverb",    "A2", "in place of something else",       "I stayed home instead."),
        ("anyway",     "adverb",    "A2", "in spite of something",            "I went anyway."),
        ("perhaps",    "adverb",    "A1", "maybe; possibly",                  "Perhaps she didn't hear."),
        ("recently",   "adverb",    "A2", "not long ago",                     "Did you see him recently?"),
        ("whole",      "adjective", "A1", "all of; complete",                 "I didn't sleep the whole night."),
        ("late",       "adjective", "A1", "after the expected time",          "Did he arrive late?"),
        ("early",      "adjective", "A1", "before the expected time",         "She didn't come early."),
        ("ready",      "adjective", "A1", "prepared",                         "The food wasn't ready yet."),
        ("angry",      "adjective", "A1", "feeling strong displeasure",       "Was she angry? Did she shout?"),
    ],
    "grammar": {
        "rule": "For Past Simple negatives use didn't + base form. For questions use Did + subject + base form?",
        "form_headers": ["Form", "Structure", "Example"],
        "form_rows": [
            ["Negative",          "Subject + didn't + base verb",      "I didn't go to school."],
            ["Yes/No question",   "Did + subject + base verb?",        "Did you see the film?"],
            ["Short answer (Yes)","Yes, + subject + did.",              "Yes, I did."],
            ["Short answer (No)", "No, + subject + didn't.",            "No, she didn't."],
            ["Wh- question",      "Wh + did + subject + base verb?",   "Where did he go?"],
        ],
        "form_col_widths": [2000, 2800, 3400],
        "note": "After didn't and Did, ALWAYS use the base form. NEVER say 'I didn't went' or 'Did she saw?'. The past tense is carried by 'did', not the main verb.",
        "usage": [
            "Negatives: I didn't eat breakfast. She didn't come to class.",
            "Yes/No questions: Did you sleep well? Did he finish the homework?",
            "Wh- questions: What did you do? Where did they go? When did it happen?",
            "Short answers: Did you enjoy it? Yes, I did. / No, I didn't.",
            "Common patterns: How much did it cost? How long did it take?",
        ],
    },
    "guided": {
        "instruction": "Make negatives or questions as indicated. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I went to school. (negative)", "I DIDN'T GO to school."],
            ["2 ✅", "She saw the film. (question)", "DID SHE SEE the film?"],
            ["3",   "He enjoyed the party. (negative)", "________________________"],
            ["4",   "They arrived on time. (question)", "________________________"],
            ["5",   "We stayed long. (negative)", "________________________"],
            ["6",   "You finished the homework. (question)", "________________________"],
            ["7",   "She paid the bill. (negative)", "________________________"],
            ["8",   "Where / he / go / yesterday (?)", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete Past Simple negative sentences or questions.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / not / sleep / well / last night)",       "________________________"],
            ["2", "(she / not / invite / me / to the party)",    "________________________"],
            ["3", "(Did / you / enjoy / the concert?)",          "________________________"],
            ["4", "(What / you / do / last weekend?)",           "________________________"],
            ["5", "(he / not / pass / the exam)",                "________________________"],
            ["6", "(Where / they / go / on holiday?)",           "________________________"],
            ["7", "(I / not / forget / your birthday)",          "________________________"],
            ["8", "(How much / it / cost?)",                     "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "I didn't went to school.", "________________________"],
            ["2", "Did she saw the film?", "________________________"],
            ["3", "He didn't arrived on time.", "________________________"],
            ["4", "Where did you went?", "________________________"],
            ["5", "She didn't enjoyed the party.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I went shopping. (make negative)", "________________________"],
            ["2", "She cooked dinner. (make a Did question)", "________________________"],
            ["3", "They stayed at home. (make a Where question)", "________________________"],
            ["4", "He spent a lot. (make a How much question)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Last Weekend Interview",
        "instructions": [
            "Write 5 Did you...? questions about last weekend.",
            "Interview your partner and note the answers.",
            "Ask follow-up Wh- questions: What did you eat? Where did you go?",
            "Report to the class: 'Ali didn't stay at home. He went to the park.'",
        ],
    },
    "reading": {
        "text": (
            "Tom didn't go to work last Monday. He was sick. He didn't eat much, but he "
            "drank a lot of water. His friend called and asked: 'Did you go to the doctor?' "
            "Tom said: 'No, I didn't. I just stayed in bed.' He didn't feel better until "
            "Wednesday. He went back to work on Thursday."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Tom went to work on Monday.",         "___"],
            ["2", "He didn't eat much.",                 "___"],
            ["3", "His friend visited him at home.",     "___"],
            ["4", "He went to the doctor.",              "___"],
            ["5", "He felt better on Wednesday.",        "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences: 3 negative and 3 questions about the past.",
        "prompts": [
            "Something you didn't do yesterday: I didn't ___.",
            "Something a friend didn't do: He/She didn't ___.",
            "Ask about last weekend: Did you ___?",
            "Ask a where question: Where did ___?",
            "Ask a what question: What did ___?",
            "Something that didn't happen: It didn't ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 10 sentences about last week: 5 negatives (didn't + base form) and "
            "5 questions (Did...? or Wh- + did...?). Make sure the main verb is ALWAYS base form."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each one, write a negative sentence "
            "and a question. Example: forget - 'I didn't forget.' / 'Did you forget?'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 13 — Past Continuous (A2)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "13",
    "cefr": "A2",
    "topic": "Past Continuous",
    "sub_focus": "was/were + verb-ing; actions in progress at a specific past time; setting the scene in stories; negative and question forms",
    "objectives": [
        "Form the Past Continuous: was/were + verb-ing.",
        "Use the Past Continuous to describe actions in progress at a specific time in the past.",
        "Form negatives: wasn't/weren't + verb-ing.",
        "Form questions: Was/Were + subject + verb-ing?",
        "Use the Past Continuous to set the scene in a story or narrative.",
    ],
    "warmup": {
        "name": "🕗 At 8 PM Last Night...",
        "time": "5 minutes",
        "steps": [
            "Think: What were you doing at 8 PM last night?",
            "Tell your partner: 'I was watching TV.' / 'I was cooking dinner.'",
            "Find someone in the class who was doing the SAME thing as you.",
        ],
    },
    "lead_in": [
        "What were you doing at 7 AM this morning? Were you sleeping or getting ready?",
        "Think about yesterday at lunchtime. What was happening around you?",
    ],
    "vocab": [
        ("while",      "conjunction","A2", "during the time that",             "I was reading while he was cooking."),
        ("during",     "preposition","A1", "throughout a period of time",      "She was sleeping during the film."),
        ("scene",      "noun",      "A2", "a part of a story or event",        "Set the scene: it was raining."),
        ("background", "noun",      "A2", "what is behind the main action",    "Music was playing in the background."),
        ("situation",  "noun",      "A1", "the conditions at a time",          "The situation was getting worse."),
        ("progress",   "noun",      "A2", "moving forward; development",       "The work was in progress."),
        ("noise",      "noun",      "A1", "a sound, especially unwanted",      "A strange noise was coming from outside."),
        ("conversation","noun",     "A2", "a talk between people",             "They were having a conversation."),
        ("direction",  "noun",      "A2", "the way something moves",           "He was walking in my direction."),
        ("dream",      "noun",      "A1", "images during sleep",               "I was having a strange dream."),
        ("ring",       "verb",      "A1", "to make a phone sound",             "The phone was ringing."),
        ("shine",      "verb",      "A1", "to give off light",                 "The sun was shining brightly."),
        ("blow",       "verb",      "A1", "wind moving air",                   "The wind was blowing hard."),
        ("pour",       "verb",      "A2", "to rain very heavily",              "It was pouring with rain."),
        ("wave",       "verb",      "A1", "to move your hand in greeting",     "She was waving at me."),
        ("chat",       "verb",      "A2", "to talk in a friendly way",         "They were chatting happily."),
        ("argue",      "verb",      "A2", "to disagree strongly",              "The couple were arguing."),
        ("whisper",    "verb",      "A2", "to speak very quietly",             "He was whispering something."),
        ("rush",       "verb",      "A2", "to move very quickly",              "People were rushing to work."),
        ("search",     "verb",      "A2", "to look carefully for something",   "She was searching for her keys."),
        ("nervous",    "adjective", "A1", "worried or slightly afraid",        "I was feeling nervous."),
        ("loud",       "adjective", "A1", "making a lot of noise",             "The music was getting loud."),
        ("dark",       "adjective", "A1", "with no light",                     "It was getting dark outside."),
        ("wet",        "adjective", "A1", "covered in water",                  "The ground was getting wet."),
        ("busy",       "adjective", "A1", "doing many things",                 "Everyone was busy working."),
        ("asleep",     "adjective", "A1", "sleeping",                          "The baby was still asleep."),
        ("awake",      "adjective", "A1", "not sleeping",                      "Were you awake at midnight?"),
        ("outside",    "adverb",    "A1", "not inside a building",             "Children were playing outside."),
        ("quietly",    "adverb",    "A2", "in a quiet way",                    "She was reading quietly."),
        ("heavily",    "adverb",    "A2", "in a heavy way; a lot",             "It was raining heavily."),
    ],
    "grammar": {
        "rule": "The Past Continuous is formed with was/were + verb-ing. It describes an action in progress at a specific time in the past.",
        "form_headers": ["Form", "I / He / She / It", "You / We / They"],
        "form_rows": [
            ["Affirmative",   "was + verb-ing",       "were + verb-ing"],
            ["Negative",      "wasn't + verb-ing",    "weren't + verb-ing"],
            ["Question",      "Was I/he/she/it + verb-ing?", "Were you/we/they + verb-ing?"],
            ["Short answer +","Yes, I was.",           "Yes, they were."],
            ["Short answer -","No, she wasn't.",       "No, we weren't."],
        ],
        "form_col_widths": [1800, 3200, 3200],
        "note": "Use Past Continuous for actions IN PROGRESS at a past moment. It sets the scene: 'The sun was shining. Birds were singing.' Do NOT use it for completed actions; use Past Simple for those.",
        "usage": [
            "Action in progress: At 8 PM, I was watching TV.",
            "Setting the scene: It was raining. People were rushing home.",
            "Parallel actions: She was cooking while he was cleaning.",
            "Background: The phone was ringing, but nobody answered.",
            "Interrupted action: I was sleeping when the alarm went off.",
        ],
    },
    "guided": {
        "instruction": "Complete with the Past Continuous (was/were + -ing). Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "At 9 PM, I ___ (watch) a film.", "At 9 PM, I WAS WATCHING a film."],
            ["2 ✅", "They ___ (play) in the garden.", "They WERE PLAYING in the garden."],
            ["3",   "She ___ (read) a book at that time.", "________________________"],
            ["4",   "We ___ (not / sleep) at midnight.", "________________________"],
            ["5",   "___ he ___ (work) yesterday evening?", "________________________"],
            ["6",   "The children ___ (run) in the park.", "________________________"],
            ["7",   "It ___ (rain) heavily all morning.", "________________________"],
            ["8",   "I ___ (not / listen) to music.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete sentences in the Past Continuous.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / study / at 10 PM last night)",         "________________________"],
            ["2", "(she / not / sleep / at midnight)",          "________________________"],
            ["3", "(they / wait / for the bus)",                "________________________"],
            ["4", "(he / cook / dinner / at 7 o'clock)",       "________________________"],
            ["5", "(we / not / watch / TV / at that time)",     "________________________"],
            ["6", "(the sun / shine / brightly)",               "________________________"],
            ["7", "(you / talk / on the phone?)",               "________________________"],
            ["8", "(people / rush / to work)",                  "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She were reading a book.", "________________________"],
            ["2", "They was playing football.", "________________________"],
            ["3", "I was watch TV at 8 PM.", "________________________"],
            ["4", "He were sleeping at midnight.", "________________________"],
            ["5", "Was they waiting for us?", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I watch TV. (at 9 PM last night, Past Continuous)", "________________________"],
            ["2", "She was cooking. (make negative)", "________________________"],
            ["3", "They were sleeping. (make a question)", "________________________"],
            ["4", "He reads a book. (at that moment, Past Continuous)", "________________________"],
        ],
        "transform_col_widths": [500, 3800, 3900],
    },
    "speaking": {
        "name": "💬 What Was Everyone Doing?",
        "instructions": [
            "Think about 8 PM last night. Describe what was happening in your home.",
            "Use Past Continuous: 'My mother was cooking. My brother was doing homework.'",
            "Ask your partner: 'What were you doing at 8 PM? Was anyone else at home?'",
            "Report to the class: 'At 8 PM, Carlos was watching TV and his sister was studying.'",
        ],
    },
    "reading": {
        "text": (
            "It was a cold evening. The wind was blowing and it was raining heavily. Inside "
            "the house, the family was having dinner. The children were talking and laughing. "
            "The dog was sleeping under the table. Outside, a cat was sitting on the wall, "
            "watching the rain. Everything was quiet and peaceful."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "It was a warm evening.",                "___"],
            ["2", "The family was having dinner.",          "___"],
            ["3", "The children were sleeping.",            "___"],
            ["4", "The dog was sleeping under the table.",  "___"],
            ["5", "A cat was sitting inside the house.",    "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences describing what was happening at a specific time in the past.",
        "prompts": [
            "At 7 AM this morning, I was ___.",
            "At lunchtime yesterday, my friend was ___.",
            "Last Sunday at 3 PM, we were ___.",
            "During the lesson, the teacher was ___.",
            "While I was ___, my brother was ___.",
            "At midnight, everyone was ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write a short paragraph (8 sentences) describing a scene from last weekend. "
            "Use Past Continuous to set the scene. What was happening? Who was doing what?"
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. Write a Past Continuous sentence "
            "for each one. Example: 'The wind was blowing hard all night.'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 14 — Past Continuous vs Past Simple (A2)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "14",
    "cefr": "A2",
    "topic": "Past Continuous vs Past Simple",
    "sub_focus": "when + Past Simple (short action) interrupts Past Continuous (long action); while + Past Continuous; two parallel past actions; narrative sequences",
    "objectives": [
        "Understand when to use Past Continuous vs Past Simple in the same sentence.",
        "Use 'when' with Past Simple for the interrupting (short) action.",
        "Use 'while' with Past Continuous for the background (long) action.",
        "Describe two parallel actions using Past Continuous + while.",
        "Tell stories combining both tenses for background and main events.",
    ],
    "warmup": {
        "name": "⚡ When the Phone Rang...",
        "time": "5 minutes",
        "steps": [
            "Teacher says: 'When the phone rang...' and students complete: 'I was taking a shower.'",
            "Take turns: one student says a Past Simple interruption, another says the background action.",
            "Make funny combinations: 'When the teacher arrived, I was dancing on the desk!'",
        ],
    },
    "lead_in": [
        "What were you doing when you heard some surprising news recently?",
        "Think of a time something unexpected happened. What was happening at that moment?",
    ],
    "vocab": [
        ("suddenly",   "adverb",    "A1", "quickly and unexpectedly",         "Suddenly, the lights went out."),
        ("while",      "conjunction","A2", "during the time that",             "While I was walking, it started raining."),
        ("when",       "conjunction","A1", "at the time that",                 "When I arrived, she was cooking."),
        ("interrupt",  "verb",      "A2", "to stop someone while speaking",   "Don't interrupt me!"),
        ("accident",   "noun",      "A1", "something bad that happens by chance","There was an accident on the road."),
        ("crash",      "noun",      "A2", "a violent vehicle collision",       "We heard a loud crash."),
        ("scream",     "verb",      "A2", "to shout very loudly",              "Someone screamed suddenly."),
        ("slip",       "verb",      "A2", "to lose your balance",              "He slipped on the wet floor."),
        ("fall",       "verb",      "A1", "to drop down",                      "She fell while she was running."),
        ("notice",     "verb",      "A1", "to see or become aware of",         "I noticed something strange."),
        ("appear",     "verb",      "A2", "to become visible; to show up",     "A cat appeared from nowhere."),
        ("disappear",  "verb",      "A2", "to become impossible to see",       "The sun disappeared behind clouds."),
        ("knock",      "verb",      "A1", "to hit a door to get attention",    "Someone knocked on the door."),
        ("shout",      "verb",      "A1", "to say something very loudly",      "He shouted my name."),
        ("drop",       "verb",      "A1", "to let something fall",             "She dropped her phone."),
        ("break",      "verb",      "A1", "to separate into pieces",           "The glass broke on the floor."),
        ("steal",      "verb",      "A2", "to take without permission",        "Someone stole my bag."),
        ("bite",       "verb",      "A2", "to use teeth to cut into",          "A dog bit him."),
        ("hurt",       "verb",      "A1", "to cause pain",                     "I hurt my leg."),
        ("escape",     "verb",      "A2", "to get away from danger",           "The bird escaped from the cage."),
        ("surprised",  "adjective", "A1", "feeling unexpected wonder",         "I was surprised."),
        ("frightened", "adjective", "A2", "feeling very afraid",               "She was frightened."),
        ("strange",    "adjective", "A1", "unusual or unexpected",             "Something strange happened."),
        ("unlucky",    "adjective", "A2", "having bad luck",                   "He was very unlucky that day."),
        ("immediately","adverb",    "A2", "at once; without delay",            "I called the police immediately."),
        ("fortunately","adverb",    "A2", "luckily",                           "Fortunately, nobody was hurt."),
        ("unfortunately","adverb",  "A2", "unluckily",                         "Unfortunately, it started raining."),
        ("exactly",    "adverb",    "A2", "precisely; no more, no less",       "What exactly were you doing?"),
        ("already",    "adverb",    "A2", "before the expected time",          "When I arrived, they were already eating."),
        ("just",       "adverb",    "A1", "a very short time ago",             "I was just leaving when he called."),
    ],
    "grammar": {
        "rule": "Use Past Continuous for the longer background action and Past Simple for the shorter interrupting action. Use 'when' with the short action and 'while' with the long action.",
        "form_headers": ["Pattern", "Structure", "Example"],
        "form_rows": [
            ["When + short action", "Past Cont. + when + Past Simple", "I was sleeping when the phone rang."],
            ["While + long action", "While + Past Cont., Past Simple", "While I was walking, it started to rain."],
            ["Two parallel actions", "Past Cont. + while + Past Cont.", "She was reading while he was cooking."],
            ["Sequence of events", "Past Simple + Past Simple",        "He woke up, got dressed, and left."],
        ],
        "form_col_widths": [2200, 2800, 3200],
        "note": "WHEN + Past Simple (the interruption). WHILE + Past Continuous (the background). If both actions are short and completed, use Past Simple + Past Simple: 'I opened the door and saw him.'",
        "usage": [
            "Interrupted action: I was reading when the doorbell rang.",
            "Background + event: While they were eating, the power went out.",
            "Parallel actions: She was studying while her brother was playing games.",
            "Story telling: The sun was shining. Birds were singing. Suddenly, it started to rain.",
            "Explanation: I was running to catch the bus when I slipped and fell.",
        ],
    },
    "guided": {
        "instruction": "Complete with Past Simple or Past Continuous. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ (sleep) when the alarm ___ (go) off.", "I WAS SLEEPING when the alarm WENT off."],
            ["2 ✅", "While she ___ (cook), he ___ (arrive).", "While she WAS COOKING, he ARRIVED."],
            ["3",   "They ___ (walk) home when it ___ (start) to rain.", "________________________"],
            ["4",   "I ___ (read) a book when someone ___ (knock).", "________________________"],
            ["5",   "While we ___ (have) dinner, the phone ___ (ring).", "________________________"],
            ["6",   "She ___ (drive) when she ___ (see) the accident.", "________________________"],
            ["7",   "He ___ (study) while his sister ___ (watch) TV.", "________________________"],
            ["8",   "The children ___ (play) when it ___ (get) dark.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete sentences using Past Continuous and Past Simple together.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / walk / to work / when / I / see / an accident)", "________________________"],
            ["2", "(while / she / sleep / the baby / start / crying)",    "________________________"],
            ["3", "(he / cook / when / the fire alarm / go off)",         "________________________"],
            ["4", "(they / have / lunch / when / the boss / call)",       "________________________"],
            ["5", "(while / I / study / my phone / ring)",                "________________________"],
            ["6", "(she / run / in the park / when / she / fall)",        "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "I was sleeping when the phone was ringing.", "________________________"],
            ["2", "While I walked home, it started to rain.", "________________________"],
            ["3", "She was cooking when he was arriving.", "________________________"],
            ["4", "When I was reading, the doorbell was ringing.", "________________________"],
            ["5", "While they had dinner, the lights went out.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I slept. The phone rang. (combine with when)", "________________________"],
            ["2", "She walked. It rained. (combine with while)", "________________________"],
            ["3", "He cooked. She cleaned. (two parallel actions with while)", "________________________"],
            ["4", "They watched TV. The power went out. (combine with when)", "________________________"],
        ],
        "transform_col_widths": [500, 3800, 3900],
    },
    "speaking": {
        "name": "💬 Tell Me Your Story",
        "instructions": [
            "Think of a time something unexpected happened to you.",
            "Set the scene with Past Continuous: 'I was walking home. It was getting dark.'",
            "Add the event with Past Simple: 'Suddenly, I heard a loud noise.'",
            "Your partner asks questions: 'What happened next? What were you doing exactly?'",
        ],
    },
    "reading": {
        "text": (
            "Last Tuesday, I was walking to work when I saw an accident. A man was crossing "
            "the road while he was talking on his phone. He didn't see the car. Fortunately, "
            "the driver was driving slowly and stopped in time. Nobody was hurt, but the man "
            "was very frightened. He dropped his phone and it broke."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "The writer was driving to work.",          "___"],
            ["2", "The man was talking on his phone.",        "___"],
            ["3", "The driver was driving fast.",             "___"],
            ["4", "Nobody was hurt.",                         "___"],
            ["5", "The man dropped his phone.",               "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write a short story (6 sentences) using Past Continuous and Past Simple together.",
        "prompts": [
            "Set the scene (Past Continuous): The sun was ___. I was ___.",
            "Add an interruption (Past Simple): Suddenly, ___.",
            "Describe a reaction (Past Simple): I ___.",
            "Add parallel actions (Past Continuous): While I was ___, someone was ___.",
            "Add another event (Past Simple): Then, ___.",
            "End the story: Finally / Fortunately / Unfortunately, ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write a story (8-10 sentences) about an unexpected event. Use Past Continuous "
            "to set the scene and Past Simple for main events. Use 'when' and 'while' at least twice each."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each one, write a sentence "
            "combining Past Continuous and Past Simple. Example: 'I was walking when I suddenly slipped.'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 15 — Present Perfect: Introduction (A2)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "15",
    "cefr": "A2",
    "topic": "Present Perfect — Introduction",
    "sub_focus": "have/has + past participle; life experiences (I have been to...); unfinished time (today, this week); ever/never in questions and answers",
    "objectives": [
        "Form the Present Perfect with have/has + past participle.",
        "Use regular past participles (worked, played, studied) and common irregular ones (been, gone, seen, done, had).",
        "Talk about life experiences: I have been to Paris. She has never seen snow.",
        "Use ever in questions: Have you ever...?",
        "Use never in negative statements: I have never...",
    ],
    "warmup": {
        "name": "🌍 Have You Ever...?",
        "time": "5 minutes",
        "steps": [
            "Stand up. Teacher asks: 'Have you ever been to another country?'",
            "If YES, stay standing. If NO, sit down.",
            "Those standing share: 'I have been to Turkey.' Then ask a classmate: 'Have you ever...?'",
        ],
    },
    "lead_in": [
        "Have you ever eaten sushi? Have you ever been to a concert? Tell your partner.",
        "What interesting things have you done in your life? Make a quick list.",
    ],
    "vocab": [
        ("ever",       "adverb",    "A2", "at any time (in questions)",       "Have you ever been to London?"),
        ("never",      "adverb",    "A1", "not at any time; not ever",        "I have never seen snow."),
        ("been",       "verb",      "A1", "past participle of be/go",         "She has been to Japan."),
        ("gone",       "verb",      "A1", "past participle of go (away now)", "He has gone to work."),
        ("seen",       "verb",      "A1", "past participle of see",           "Have you seen this film?"),
        ("done",       "verb",      "A1", "past participle of do",            "I have done my homework."),
        ("eaten",      "verb",      "A1", "past participle of eat",           "She has never eaten sushi."),
        ("visited",    "verb",      "A1", "past participle of visit",         "We have visited many cities."),
        ("tried",      "verb",      "A1", "past participle of try",           "Have you ever tried Thai food?"),
        ("met",        "verb",      "A1", "past participle of meet",          "I have met the new teacher."),
        ("read",       "verb",      "A1", "past participle of read",          "She has read that book."),
        ("written",    "verb",      "A2", "past participle of write",         "He has written three emails."),
        ("spoken",     "verb",      "A2", "past participle of speak",         "Have you spoken to her?"),
        ("driven",     "verb",      "A2", "past participle of drive",         "She has never driven a car."),
        ("flown",      "verb",      "A2", "past participle of fly",           "I have flown in a helicopter."),
        ("broken",     "verb",      "A1", "past participle of break",         "He has broken his arm."),
        ("forgotten",  "verb",      "A1", "past participle of forget",        "I have forgotten his name."),
        ("chosen",     "verb",      "A2", "past participle of choose",        "She has chosen a dress."),
        ("experience", "noun",      "A2", "something that happens to you",    "It was an amazing experience."),
        ("abroad",     "adverb",    "A1", "in a foreign country",             "Have you ever lived abroad?"),
        ("once",       "adverb",    "A1", "one time",                         "I have been there once."),
        ("twice",      "adverb",    "A1", "two times",                        "She has visited twice."),
        ("several",    "adjective", "A2", "more than two but not many",       "He has been there several times."),
        ("amazing",    "adjective", "A1", "very surprising and good",         "It was an amazing trip."),
        ("wonderful",  "adjective", "A1", "extremely good",                   "We have had a wonderful time."),
        ("horrible",   "adjective", "A2", "very bad or unpleasant",           "I have had a horrible day."),
        ("lucky",      "adjective", "A1", "having good things by chance",     "I have been very lucky."),
        ("famous",     "adjective", "A1", "known by many people",             "Have you met anyone famous?"),
        ("different",  "adjective", "A1", "not the same",                     "I have tried many different foods."),
        ("exciting",   "adjective", "A1", "causing strong enthusiasm",        "It has been an exciting week."),
    ],
    "grammar": {
        "rule": "The Present Perfect is formed with have/has + past participle. Use it for life experiences, unfinished time periods, and actions with a result in the present.",
        "form_headers": ["Form", "I / You / We / They", "He / She / It"],
        "form_rows": [
            ["Affirmative",   "have + past participle (I've done)", "has + past participle (She's done)"],
            ["Negative",      "haven't + past participle",          "hasn't + past participle"],
            ["Question",      "Have you/they + past participle?",   "Has he/she + past participle?"],
            ["Short answer +","Yes, I have.",                        "Yes, she has."],
            ["Short answer -","No, we haven't.",                     "No, he hasn't."],
        ],
        "form_col_widths": [1800, 3200, 3200],
        "note": "Use 'been to' for visiting and returning: 'I have been to Paris.' (= I visited and came back.) Use 'gone to' for someone who is still away: 'She has gone to work.' (= She is not here now.)",
        "usage": [
            "Life experience: I have visited ten countries. She has never eaten sushi.",
            "Ever/never: Have you ever been to Japan? No, I have never been there.",
            "Unfinished time: I have read two books this week. He has worked hard today.",
            "Result now: She has broken her leg. (= It is broken now.)",
            "How many times: I have seen that film three times.",
        ],
    },
    "guided": {
        "instruction": "Complete with the Present Perfect (have/has + past participle). Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ (visit) London twice.", "I HAVE VISITED London twice."],
            ["2 ✅", "She ___ never ___ (eat) sushi.", "She HAS never EATEN sushi."],
            ["3",   "They ___ (be) to Japan.", "________________________"],
            ["4",   "___ you ever ___ (see) a whale?", "________________________"],
            ["5",   "He ___ (not / do) his homework.", "________________________"],
            ["6",   "We ___ (meet) the new teacher.", "________________________"],
            ["7",   "She ___ (write) three emails today.", "________________________"],
            ["8",   "I ___ never ___ (fly) in a helicopter.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete sentences using the Present Perfect.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / be / to / three different countries)", "________________________"],
            ["2", "(she / never / drive / a car)",             "________________________"],
            ["3", "(they / see / that film / twice)",          "________________________"],
            ["4", "(he / not / finish / his project)",         "________________________"],
            ["5", "(Have / you / ever / try / Thai food?)",    "________________________"],
            ["6", "(we / visit / many interesting places)",    "________________________"],
            ["7", "(she / read / five books / this month)",    "________________________"],
            ["8", "(I / not / speak / to her / today)",        "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "I have went to London.", "________________________"],
            ["2", "She has never ate sushi.", "________________________"],
            ["3", "Have you ever been went to Japan?", "________________________"],
            ["4", "He have visited Paris twice.", "________________________"],
            ["5", "They has seen that film.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I visited France. (use Present Perfect + ever)", "________________________"],
            ["2", "She didn't eat sushi. (use Present Perfect + never)", "________________________"],
            ["3", "He saw that film. (use Present Perfect + twice)", "________________________"],
            ["4", "They went to Japan. (use Present Perfect + been to)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Life Experiences Bingo",
        "instructions": [
            "Make a 3x3 grid with experiences: 'been to Europe', 'eaten sushi', 'seen a famous person'...",
            "Walk around and ask classmates: 'Have you ever been to Europe?'",
            "If they say YES, write their name. First person to complete a line wins!",
            "Report: 'Maria has been to three countries. Ali has never flown in a plane.'",
        ],
    },
    "reading": {
        "text": (
            "My friend Sara has travelled a lot. She has been to fifteen countries. She has "
            "visited Japan, Brazil, and Egypt. She has tried many different foods and has "
            "met interesting people everywhere. But she has never been to Australia. It is "
            "her dream. She says: 'I have always wanted to see kangaroos in the wild!'"
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Sara has been to fifteen countries.",     "___"],
            ["2", "She has visited Australia.",              "___"],
            ["3", "She has tried many different foods.",     "___"],
            ["4", "She has never met interesting people.",   "___"],
            ["5", "She has always wanted to visit Australia.","___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about your life experiences using the Present Perfect.",
        "prompts": [
            "A country you have visited: I have been to ___.",
            "Something you have never done: I have never ___.",
            "A food you have tried: I have tried ___.",
            "A film you have seen: I have seen ___.",
            "Something you have done this week: I have ___ this week.",
            "Ask someone a question: Have you ever ___?",
        ],
    },
    "homework": {
        "item1": (
            "Write 10 sentences about your life experiences using the Present Perfect. "
            "Include at least 3 with 'never', 2 with 'ever' (questions), and 2 with time expressions (twice, three times)."
        ),
        "item2": (
            "Learn these irregular past participles: be-been, go-gone, see-seen, do-done, "
            "eat-eaten, write-written, speak-spoken, drive-driven, fly-flown, break-broken, "
            "forget-forgotten, choose-chosen, meet-met, read-read. Write each in a Present Perfect sentence."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 16 — Present Perfect: just / already / yet (A2)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_11_16.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "16",
    "cefr": "A2",
    "topic": "Present Perfect — just / already / yet",
    "sub_focus": "just (a moment ago); already (sooner than expected); yet (until now, in negatives and questions); word order with just/already/yet",
    "objectives": [
        "Use 'just' with Present Perfect to talk about very recent actions.",
        "Use 'already' to say something happened sooner than expected.",
        "Use 'yet' in negative sentences and questions to talk about expected actions.",
        "Place just and already between have/has and the past participle.",
        "Place yet at the end of the sentence in negatives and questions.",
    ],
    "warmup": {
        "name": "✅ Have You Done It Yet?",
        "time": "5 minutes",
        "steps": [
            "Teacher gives a list: 'eat breakfast, do homework, check your phone, drink water'.",
            "Students say which they have already done and which they haven't done yet.",
            "Example: 'I have already eaten breakfast, but I haven't done my homework yet.'",
        ],
    },
    "lead_in": [
        "What have you already done today? What haven't you done yet?",
        "Has anyone just arrived? What has your teacher just said?",
    ],
    "vocab": [
        ("just",       "adverb",    "A2", "a very short time ago",            "I have just finished lunch."),
        ("already",    "adverb",    "A2", "before now; sooner than expected", "She has already left."),
        ("yet",        "adverb",    "A2", "until now (negatives/questions)",  "He hasn't arrived yet."),
        ("still",      "adverb",    "A1", "continuing until now",             "She still hasn't called."),
        ("finished",   "adjective", "A1", "completed; done",                  "I have just finished."),
        ("arrived",    "verb",      "A1", "past participle of arrive",        "Has she arrived yet?"),
        ("started",    "verb",      "A1", "past participle of start",         "The film has already started."),
        ("cleaned",    "verb",      "A1", "past participle of clean",         "I have just cleaned the house."),
        ("paid",       "verb",      "A1", "past participle of pay",           "Have you paid the bill yet?"),
        ("sent",       "verb",      "A1", "past participle of send",          "She has already sent the email."),
        ("received",   "verb",      "A2", "past participle of receive",       "I haven't received it yet."),
        ("replied",    "verb",      "A2", "past participle of reply",         "He hasn't replied yet."),
        ("booked",     "verb",      "A2", "past participle of book",          "Have you booked the tickets yet?"),
        ("packed",     "verb",      "A2", "past participle of pack",          "I have already packed my bag."),
        ("prepared",   "verb",      "A2", "past participle of prepare",       "She has just prepared dinner."),
        ("completed",  "verb",      "A2", "past participle of complete",      "They have already completed the task."),
        ("downloaded", "verb",      "A2", "past participle of download",      "I have just downloaded the app."),
        ("charged",    "verb",      "A2", "past participle of charge",        "Has your phone charged yet?"),
        ("woken",      "verb",      "A1", "past participle of wake",          "He has just woken up."),
        ("eaten",      "verb",      "A1", "past participle of eat",           "Have you eaten yet?"),
        ("recently",   "adverb",    "A2", "not long ago",                     "I have recently changed jobs."),
        ("completely", "adverb",    "A2", "totally; in every way",            "I have completely finished."),
        ("finally",    "adverb",    "A1", "at last; after a long time",       "She has finally replied!"),
        ("almost",     "adverb",    "A1", "nearly; not quite",                "I have almost finished."),
        ("probably",   "adverb",    "A2", "most likely",                      "She has probably left already."),
        ("actually",   "adverb",    "A2", "in fact; really",                  "I have actually done it!"),
        ("definitely", "adverb",    "A2", "certainly; for sure",              "He has definitely arrived."),
        ("task",       "noun",      "A2", "a piece of work to do",            "Have you finished the task yet?"),
        ("message",    "noun",      "A1", "a piece of information sent",      "I have just sent a message."),
        ("laundry",    "noun",      "A2", "clothes that need washing",        "She hasn't done the laundry yet."),
    ],
    "grammar": {
        "rule": "Use 'just' for very recent actions, 'already' for actions completed sooner than expected, and 'yet' in negatives/questions for actions expected but not done.",
        "form_headers": ["Word", "Meaning", "Position", "Example"],
        "form_rows": [
            ["just",    "a moment ago",          "have/has + JUST + past participle",    "I have just eaten."],
            ["already", "sooner than expected",   "have/has + ALREADY + past participle", "She has already left."],
            ["yet",     "until now (neg/question)","at the END of the sentence",          "He hasn't arrived yet."],
            ["yet",     "until now (question)",    "at the END of the sentence",          "Have you finished yet?"],
        ],
        "form_col_widths": [1200, 2000, 2800, 2200],
        "note": "JUST and ALREADY go between have/has and the past participle. YET goes at the END of the sentence and is used ONLY in negatives and questions. Do NOT say 'I have yet finished.' Say 'I haven't finished yet.'",
        "usage": [
            "Just (very recent): I have just arrived. She has just called me.",
            "Already (done sooner): The film has already started. I have already eaten.",
            "Yet (negative): I haven't finished yet. He hasn't called yet.",
            "Yet (question): Have you done your homework yet? Has she arrived yet?",
            "Contrast: I have already cleaned my room, but I haven't done the laundry yet.",
        ],
    },
    "guided": {
        "instruction": "Complete with just, already, or yet in the correct position. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ finished my homework. (just)", "I have JUST finished my homework."],
            ["2 ✅", "She hasn't called ___. (yet)", "She hasn't called YET."],
            ["3",   "They have ___ left the house. (already)", "________________________"],
            ["4",   "Have you eaten ___? (yet)", "________________________"],
            ["5",   "He has ___ woken up. (just)", "________________________"],
            ["6",   "We haven't booked the tickets ___. (yet)", "________________________"],
            ["7",   "I have ___ sent the email. (already)", "________________________"],
            ["8",   "Has the film started ___? (yet)", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete Present Perfect sentences using just, already, or yet.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / just / finish / lunch)",                "________________________"],
            ["2", "(she / already / leave / the office)",       "________________________"],
            ["3", "(they / not / arrive / yet)",                "________________________"],
            ["4", "(Have / you / do / your homework / yet?)",   "________________________"],
            ["5", "(he / just / send / the message)",           "________________________"],
            ["6", "(we / already / book / the hotel)",          "________________________"],
            ["7", "(I / not / reply / to her email / yet)",     "________________________"],
            ["8", "(she / just / wake up)",                     "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "I have finished just my homework.", "________________________"],
            ["2", "She hasn't already called.", "________________________"],
            ["3", "Have you yet eaten?", "________________________"],
            ["4", "He has yet not arrived.", "________________________"],
            ["5", "They have left already the building.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I finished my homework. (add 'just', Present Perfect)", "________________________"],
            ["2", "She left. (add 'already', Present Perfect)", "________________________"],
            ["3", "He didn't arrive. (use 'yet', Present Perfect)", "________________________"],
            ["4", "Did you eat? (use 'yet', Present Perfect question)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 The To-Do List Game",
        "instructions": [
            "Write a to-do list of 6 things for today (eat breakfast, do homework, call a friend...).",
            "Tell your partner which ones you have already done and which you haven't done yet.",
            "Your partner asks: 'Have you done your homework yet?' You reply: 'Yes, I have already done it.' or 'No, I haven't done it yet.'",
            "Report: 'Maria has already eaten lunch, but she hasn't called her friend yet.'",
        ],
    },
    "reading": {
        "text": (
            "It is Saturday morning. Tom has a list of things to do. He has already eaten "
            "breakfast and cleaned his room. He has just finished washing the dishes. But he "
            "hasn't done the laundry yet. He hasn't been to the supermarket yet either. His "
            "mother asks: 'Have you called your grandmother yet?' 'No, not yet!' he says."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Tom has already eaten breakfast.",          "___"],
            ["2", "He hasn't cleaned his room yet.",           "___"],
            ["3", "He has just finished the dishes.",          "___"],
            ["4", "He has already done the laundry.",          "___"],
            ["5", "He has called his grandmother.",            "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences using just, already, and yet about your day today.",
        "prompts": [
            "Something you have just done: I have just ___.",
            "Something you have already done: I have already ___.",
            "Something you haven't done yet: I haven't ___ yet.",
            "Ask a friend a question with yet: Have you ___ yet?",
            "Something that has already happened today: The ___ has already ___.",
            "Something very recent: My teacher has just ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 10 sentences about your day today. Use 'just' in 3 sentences, 'already' in 4 sentences, "
            "and 'yet' in 3 sentences (include both negative and question forms with 'yet')."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. Write a sentence for each using "
            "just, already, or yet. Example: 'I have just received your message.' / 'She hasn't replied yet.'"
        ),
    },
})


# ─────────────────────────────────────────────────────────────────────────────
# Generate all DOCX + PPTX files
# ─────────────────────────────────────────────────────────────────────────────

OUT_DIR = "/projects/sandbox/APP-GENERATION-PPT-FILE-ETC/generated"
os.makedirs(OUT_DIR, exist_ok=True)

topic_slugs = [
    "PastSimple_Regular_Irregular",
    "PastSimple_Negative_Questions",
    "PastContinuous",
    "PastContinuous_vs_PastSimple",
    "PresentPerfect_Introduction",
    "PresentPerfect_Just_Already_Yet",
]

if __name__ == "__main__":
    for i, (unit, slug) in enumerate(zip(UNITS_11_16, topic_slugs)):
        unit_num = i + 11

        # Generate DOCX
        docx_fname = f"Level1_Unit{unit_num}_{slug}.docx"
        docx_path = os.path.join(OUT_DIR, docx_fname)
        docx_data = build_docx(unit)
        with open(docx_path, "wb") as f:
            f.write(docx_data)
        print(f"  Generated: {docx_fname}  ({len(docx_data):,} bytes)")

        # Generate PPTX
        pptx_fname = f"Level1_Unit{unit_num}_{slug}.pptx"
        pptx_path = os.path.join(OUT_DIR, pptx_fname)
        pptx_data = build_pptx(unit)
        with open(pptx_path, "wb") as f:
            f.write(pptx_data)
        print(f"  Generated: {pptx_fname}  ({len(pptx_data):,} bytes)")

    print(f"\nAll {len(UNITS_11_16) * 2} files generated successfully (6 DOCX + 6 PPTX).")

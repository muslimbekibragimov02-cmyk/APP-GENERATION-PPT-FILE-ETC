"""
ELT Grammar Course Designer — Units 5-10 Generator
Generates Level 1 Units 5-10 (DOCX + PPTX) covering Present Simple, Present
Simple vs Present Continuous, Have Got/Has Got, and Verb To Be Past Simple.
"""

import sys
import os

sys.path.insert(0, '/projects/sandbox/APP-GENERATION-PPT-FILE-ETC')

from generate_docx import build_docx
from generate_pptx import build_pptx

# ─────────────────────────────────────────────────────────────────────────────
# UNIT DATA — Units 5 through 10
# ─────────────────────────────────────────────────────────────────────────────

UNITS_5_10 = []

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 5 — Present Simple: Affirmative (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "5",
    "cefr": "A1",
    "topic": "Present Simple — Affirmative",
    "sub_focus": "I/we/you/they + base form; he/she/it + -s/-es; spelling rules (-es after s/sh/ch, -y to -ies); adverbs of frequency (always/never/often/sometimes/usually)",
    "objectives": [
        "Form affirmative Present Simple sentences with all subject pronouns.",
        "Add -s or -es correctly to verbs after he, she, and it.",
        "Apply spelling rules: -es after s/sh/ch/x/o and -y to -ies.",
        "Use adverbs of frequency (always, usually, often, sometimes, never) in the correct position.",
        "Produce sentences about daily routines and habits.",
    ],
    "warmup": {
        "name": "🕐 My Morning Routine",
        "time": "5 minutes",
        "steps": [
            "Think of five things you do every morning.",
            "Tell your partner: 'I wake up at 7. I brush my teeth. I drink coffee...'",
            "Your partner listens and remembers. Then they repeat YOUR routine to the class.",
        ],
    },
    "lead_in": [
        "What time do you wake up every day? What do you do first?",
        "Does your mother or father cook breakfast? What do they usually make?",
    ],
    "vocab": [
        ("morning",    "noun",      "A1", "the early part of the day",         "I wake up every morning."),
        ("evening",    "noun",      "A1", "the later part of the day",         "She reads in the evening."),
        ("breakfast",  "noun",      "A1", "the first meal of the day",         "He eats breakfast at seven."),
        ("lunch",      "noun",      "A1", "the meal in the middle of the day", "We have lunch at one."),
        ("dinner",     "noun",      "A1", "the main evening meal",             "They cook dinner together."),
        ("school",     "noun",      "A1", "a place for learning",              "She goes to school daily."),
        ("office",     "noun",      "A1", "a place where people work",         "He works in an office."),
        ("every",      "determiner","A1", "each one without exception",        "I drink water every day."),
        ("always",     "adverb",    "A1", "at all times; on every occasion",   "She always smiles."),
        ("usually",    "adverb",    "A1", "most of the time",                  "He usually walks to work."),
        ("often",      "adverb",    "A1", "many times; frequently",            "We often visit our parents."),
        ("sometimes",  "adverb",    "A1", "on some occasions",                 "I sometimes take the bus."),
        ("never",      "adverb",    "A1", "not at any time",                   "She never drinks coffee."),
        ("early",      "adverb",    "A1", "before the expected time",          "He arrives early today."),
        ("finish",     "verb",      "A1", "to bring to an end",                "She finishes work at five."),
        ("start",      "verb",      "A1", "to begin",                          "Classes start at nine."),
        ("live",       "verb",      "A1", "to have your home somewhere",       "They live in London."),
        ("drink",      "verb",      "A1", "to take liquid into the body",      "He drinks tea every day."),
        ("teach",      "verb",      "A1", "to give lessons",                   "She teaches English."),
        ("wash",       "verb",      "A1", "to clean with water",               "He washes his car weekly."),
        ("catch",      "verb",      "A1", "to get hold of something moving",   "She catches the bus at 8."),
        ("fix",        "verb",      "A1", "to repair something",               "He fixes computers."),
        ("enjoy",      "verb",      "A1", "to get pleasure from",              "They enjoy music a lot."),
        ("brush",      "verb",      "A1", "to clean with a brush",             "She brushes her hair."),
        ("fly",        "verb",      "A1", "to travel through the air",         "He flies to Paris often."),
        ("worry",      "verb",      "A1", "to feel anxious",                   "She worries about exams."),
        ("copy",       "verb",      "A1", "to make something the same",        "He copies the notes daily."),
        ("bus",        "noun",      "A1", "a large vehicle for passengers",    "I take the bus to school."),
        ("tooth",      "noun",      "A1", "one of the hard white objects in your mouth", "Brush your teeth daily."),
        ("weekend",    "noun",      "A1", "Saturday and Sunday",               "He relaxes at the weekend."),
    ],
    "grammar": {
        "rule": "Use the base form of the verb for I/you/we/they. Add -s or -es for he/she/it.",
        "form_headers": ["Subject", "Verb Form", "Example"],
        "form_rows": [
            ["I / You / We / They", "base form (no change)", "I work every day."],
            ["He / She / It",       "verb + -s",             "She works every day."],
            ["He / She / It",       "verb + -es (s/sh/ch/x/o)", "He watches TV at night."],
            ["He / She / It",       "consonant + y = -ies",  "She studies English."],
        ],
        "form_col_widths": [2200, 2800, 3200],
        "spelling_headers": ["Rule", "Base Form", "He/She/It Form"],
        "spelling_rows": [
            ["Most verbs: + s",           "work, like, eat",    "works, likes, eats"],
            ["After s, sh, ch, x, o: + es","wash, teach, go",   "washes, teaches, goes"],
            ["Consonant + y: y to ies",    "study, fly, worry", "studies, flies, worries"],
            ["Vowel + y: + s",             "play, enjoy, stay", "plays, enjoys, stays"],
        ],
        "spelling_col_widths": [2600, 2600, 3000],
        "note": "The verb have is irregular: he/she/it HAS (not haves). Do NOT add -s to modal verbs (can, must, will).",
        "usage": [
            "Daily routines: I wake up at 7. He brushes his teeth.",
            "Habits: She always drinks coffee in the morning.",
            "Facts: The sun rises in the east.",
            "Frequency: We usually eat lunch at 1 o'clock.",
            "Jobs: My father works in a factory. He teaches maths.",
        ],
    },
    "guided": {
        "instruction": "Complete the sentences with the correct Present Simple form. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "She ___ (teach) English at a school.", "She TEACHES English at a school."],
            ["2 ✅", "They ___ (live) in a big city.", "They LIVE in a big city."],
            ["3",   "He ___ (wash) his car every Saturday.", "________________________"],
            ["4",   "I ___ (drink) tea every morning.", "________________________"],
            ["5",   "She ___ (finish) work at 6 o'clock.", "________________________"],
            ["6",   "We ___ (catch) the bus at 8:15.", "________________________"],
            ["7",   "My brother ___ (fly) to Berlin often.", "________________________"],
            ["8",   "They ___ (enjoy) cooking together.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete affirmative sentences using the Present Simple.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(Maria / teach / science)",        "________________________"],
            ["2", "(my parents / live / in Rome)",    "________________________"],
            ["3", "(he / catch / the 7:30 train)",    "________________________"],
            ["4", "(she / always / brush / her teeth)","________________________"],
            ["5", "(we / usually / have / lunch / at 1)", "________________________"],
            ["6", "(Tom / fix / cars / every day)",   "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She teachs English.", "________________________"],
            ["2", "He washs his car.", "________________________"],
            ["3", "They goes to school every day.", "________________________"],
            ["4", "My sister flys to London.", "________________________"],
            ["5", "He studys at night.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I drink coffee. (change subject to He)", "________________________"],
            ["2", "They watch TV. (change subject to She)", "________________________"],
            ["3", "We study English. (change subject to He)", "________________________"],
            ["4", "I go to work by bus. (change subject to My mother)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Daily Habits Survey",
        "instructions": [
            "Ask three classmates about their daily routine using Present Simple questions.",
            "Use prompts: wake up / have breakfast / go to work or school / finish / relax.",
            "Report to the class: 'Maria wakes up at 6. She always has breakfast.'",
            "Use adverbs of frequency: always, usually, often, sometimes, never.",
        ],
    },
    "reading": {
        "text": (
            "Tom lives in Manchester. He works in an office. He usually starts at nine "
            "and finishes at five. He always catches the bus to work. In the evening, "
            "he sometimes cooks dinner, but he often buys food from a shop. He never "
            "goes to bed late because he wakes up early every morning."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Tom lives in London.",                   "___"],
            ["2", "He starts work at nine.",                "___"],
            ["3", "He always drives to work.",              "___"],
            ["4", "He sometimes cooks dinner.",             "___"],
            ["5", "He goes to bed late every night.",       "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about your daily routine using the Present Simple.",
        "prompts": [
            "When do you wake up? I wake up at ___.",
            "How do you get to work/school? I ___ to ___.",
            "What does your best friend do every day? He/She ___.",
            "What do you always do in the evening? I always ___.",
            "What do you never do? I never ___.",
            "What does a family member usually do at the weekend? He/She usually ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about a family member's daily routine using the Present Simple. "
            "Include at least 3 different adverbs of frequency. Underline every verb + -s/-es form."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. Write them in a sentence of your own "
            "using the Present Simple affirmative. Check your spelling of the he/she/it forms."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 6 — Present Simple: Negative (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "6",
    "cefr": "A1",
    "topic": "Present Simple — Negative",
    "sub_focus": "don't / doesn't + base form; I don't like / He doesn't like; common errors (NOT doesn't speaks)",
    "objectives": [
        "Form negative sentences with don't (do not) for I/you/we/they.",
        "Form negative sentences with doesn't (does not) for he/she/it.",
        "Understand that the main verb stays in the base form after don't/doesn't.",
        "Avoid the common error: He doesn't likes (incorrect) vs He doesn't like (correct).",
        "Express likes, dislikes, and habits in the negative form.",
    ],
    "warmup": {
        "name": "🚫 I Don't Do That!",
        "time": "5 minutes",
        "steps": [
            "Teacher says a sentence: 'I eat pizza every day.' Students who don't say: 'I DON'T eat pizza every day.'",
            "Students take turns making positive statements; classmates who disagree stand up and say the negative.",
            "Keep it fast! If you are too slow, you sit down.",
        ],
    },
    "lead_in": [
        "What food don't you like? Tell your partner: 'I don't like ...'",
        "Does your best friend like the same things as you? What doesn't he/she like?",
    ],
    "vocab": [
        ("hate",       "verb",      "A1", "to dislike very strongly",          "I don't hate anybody."),
        ("agree",      "verb",      "A1", "to have the same opinion",          "He doesn't agree with me."),
        ("smoke",      "verb",      "A1", "to breathe in cigarette fumes",     "She doesn't smoke."),
        ("spend",      "verb",      "A1", "to use money to buy things",        "They don't spend much."),
        ("travel",     "verb",      "A1", "to go from one place to another",   "He doesn't travel often."),
        ("understand", "verb",      "A1", "to know the meaning of something",  "I don't understand this."),
        ("believe",    "verb",      "A1", "to think something is true",        "She doesn't believe him."),
        ("belong",     "verb",      "A1", "to be the property of",             "This doesn't belong to me."),
        ("matter",     "verb",      "A1", "to be important",                   "It doesn't matter now."),
        ("cost",       "verb",      "A1", "to have a particular price",        "It doesn't cost much."),
        ("forget",     "verb",      "A1", "to fail to remember",               "I don't forget names."),
        ("mind",       "verb",      "A1", "to be upset or annoyed by",         "She doesn't mind noise."),
        ("sugar",      "noun",      "A1", "a sweet substance for food/drinks", "He doesn't take sugar."),
        ("meat",       "noun",      "A1", "food from an animal",               "She doesn't eat meat."),
        ("noise",      "noun",      "A1", "loud or unwanted sound",            "I don't like noise."),
        ("mistake",    "noun",      "A1", "something done wrongly",            "He doesn't make mistakes."),
        ("habit",      "noun",      "A1", "something you do regularly",        "She doesn't have bad habits."),
        ("reason",     "noun",      "A1", "why something happens",             "I don't know the reason."),
        ("truth",      "noun",      "A1", "what is true or real",              "He doesn't tell the truth."),
        ("problem",    "noun",      "A1", "a difficult situation",             "We don't have a problem."),
        ("either",     "adverb",    "A1", "also not; as well (negatives)",     "I don't like it either."),
        ("anymore",    "adverb",    "A1", "no longer; not now",                "She doesn't live here anymore."),
        ("really",     "adverb",    "A1", "truly; very much",                  "I don't really like fish."),
        ("foreign",    "adjective", "A1", "from another country",              "He doesn't speak foreign languages."),
        ("wrong",      "adjective", "A1", "not correct",                       "That doesn't sound wrong."),
        ("lazy",       "adjective", "A1", "not wanting to work",               "She isn't lazy at all."),
        ("healthy",    "adjective", "A1", "in good physical condition",        "He doesn't eat healthy food."),
        ("spicy",      "adjective", "A1", "having a strong hot flavor",        "She doesn't like spicy food."),
        ("boring",     "adjective", "A1", "not interesting",                   "This film isn't boring."),
        ("alone",      "adjective", "A1", "without other people",              "He doesn't live alone."),
    ],
    "grammar": {
        "rule": "To make a Present Simple negative, use do not (don't) or does not (doesn't) + base form of the verb.",
        "form_headers": ["Subject", "Negative Form", "Example"],
        "form_rows": [
            ["I / You / We / They", "don't + base verb",    "I don't like coffee."],
            ["He / She / It",       "doesn't + base verb",  "She doesn't eat meat."],
        ],
        "form_col_widths": [2200, 2800, 3200],
        "note": "After doesn't, ALWAYS use the base form. NEVER say 'He doesn't likes' or 'She doesn't goes'. The -s is already in 'does'.",
        "usage": [
            "Likes/dislikes: I don't like spicy food. He doesn't enjoy sports.",
            "Habits: She doesn't smoke. We don't eat meat.",
            "Knowledge: They don't understand the question.",
            "Possession: He doesn't have a car.",
            "Frequency: I don't usually wake up early. She doesn't often travel.",
        ],
    },
    "guided": {
        "instruction": "Make the sentences negative. Items 1-2 are completed for you.",
        "headers": ["#", "Positive Sentence", "Negative Sentence"],
        "items": [
            ["1 ✅", "She likes coffee.", "She DOESN'T LIKE coffee."],
            ["2 ✅", "They eat meat.", "They DON'T EAT meat."],
            ["3",   "He smokes.", "________________________"],
            ["4",   "I understand the question.", "________________________"],
            ["5",   "She spends a lot of money.", "________________________"],
            ["6",   "We agree with the teacher.", "________________________"],
            ["7",   "It costs a lot.", "________________________"],
            ["8",   "He believes the story.", "________________________"],
        ],
        "col_widths": [500, 3500, 4200],
    },
    "controlled": {
        "instruction": "Write complete negative sentences in the Present Simple.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(she / not / eat / meat)",          "________________________"],
            ["2", "(I / not / like / noise)",          "________________________"],
            ["3", "(he / not / travel / often)",       "________________________"],
            ["4", "(we / not / understand / this)",    "________________________"],
            ["5", "(it / not / matter)",               "________________________"],
            ["6", "(they / not / agree / with us)",    "________________________"],
            ["7", "(my father / not / smoke)",         "________________________"],
            ["8", "(this / not / belong / to me)",     "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "He doesn't likes football.", "________________________"],
            ["2", "She don't eat meat.", "________________________"],
            ["3", "They doesn't understand.", "________________________"],
            ["4", "I doesn't have a car.", "________________________"],
            ["5", "He don't speaks English.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I like fish. (make negative)",              "________________________"],
            ["2", "She eats cheese. (make negative)",          "________________________"],
            ["3", "They travel often. (make negative)",         "________________________"],
            ["4", "He has a big car. (make negative)",          "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Things We Don't Do",
        "instructions": [
            "Work in pairs. Tell your partner 5 things you don't do.",
            "Example: 'I don't smoke. I don't eat meat. I don't watch horror films.'",
            "Your partner reports to the class: 'Ana doesn't smoke. She doesn't eat meat.'",
            "Find one thing that BOTH of you don't do: 'We don't like...'",
        ],
    },
    "reading": {
        "text": (
            "My friend Sara is vegetarian. She doesn't eat meat or fish. She doesn't "
            "drink milk either. Her brother is different. He doesn't like vegetables, "
            "but he eats everything else. They don't agree about food, but they don't "
            "argue. They cook separately and respect each other."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Sara eats fish.",                        "___"],
            ["2", "Sara doesn't drink milk.",               "___"],
            ["3", "Her brother likes vegetables.",          "___"],
            ["4", "They argue about food.",                 "___"],
            ["5", "They cook separately.",                  "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 negative sentences about yourself or people you know.",
        "prompts": [
            "Write something you don't eat: I don't eat ___.",
            "Write something you don't do in the morning: I don't ___.",
            "Write something your friend doesn't like: He/She doesn't like ___.",
            "Write something your parents don't do: They don't ___.",
            "Write something that doesn't cost much: ___ doesn't cost much.",
            "Write something you don't understand: I don't understand ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 negative sentences about things your best friend doesn't do. "
            "Use doesn't at least 5 times. Check that the base form follows doesn't."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each word, write one negative sentence "
            "using don't or doesn't. Example: 'She doesn't smoke.'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 7 — Present Simple: Questions and Short Answers (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "7",
    "cefr": "A1",
    "topic": "Present Simple — Questions & Short Answers",
    "sub_focus": "Do/Does + subject + base form?; Wh- questions; short answers (Yes, I do / No, she doesn't)",
    "objectives": [
        "Form Yes/No questions using Do you/we/they...? and Does he/she/it...?",
        "Give correct short answers: Yes, I do. / No, she doesn't.",
        "Construct Wh- questions with correct word order (What do you do?).",
        "Understand that 'What do you do?' asks about someone's job.",
        "Use question forms in everyday conversation about habits and routines.",
    ],
    "warmup": {
        "name": "🎲 Question Dice",
        "time": "5 minutes",
        "steps": [
            "Roll a die: 1-2 = Do you...? 3-4 = Does he/she...? 5-6 = What/Where/When...?",
            "Make a question using the category and ask a classmate.",
            "Your classmate gives a short answer or a full answer.",
        ],
    },
    "lead_in": [
        "Do you like English? Do you study every day? Ask your partner these questions.",
        "What does your father do? Where does your mother work? Tell the class.",
    ],
    "vocab": [
        ("job",        "noun",      "A1", "work that you do for money",        "What does he do for a job?"),
        ("hobby",      "noun",      "A1", "an activity you enjoy in free time","Do you have a hobby?"),
        ("country",    "noun",      "A1", "a nation with its own government",  "Which country do you live in?"),
        ("language",   "noun",      "A1", "a system of words for communication","Does she speak a language?"),
        ("city",       "noun",      "A1", "a large town",                      "Do they live in a big city?"),
        ("holiday",    "noun",      "A1", "a time of rest from work",          "Where do you go on holiday?"),
        ("neighbour",  "noun",      "A1", "a person living next to you",       "Does your neighbour work?"),
        ("subject",    "noun",      "A1", "a topic studied in school",         "What subjects do you study?"),
        ("brother",    "noun",      "A1", "a male sibling",                    "Does your brother work?"),
        ("sister",     "noun",      "A1", "a female sibling",                  "Does your sister drive?"),
        ("prefer",     "verb",      "A1", "to like one thing more than another","Do you prefer tea or coffee?"),
        ("need",       "verb",      "A1", "to require something",              "Does he need any help?"),
        ("know",       "verb",      "A1", "to have information in your mind",  "Do you know the answer?"),
        ("want",       "verb",      "A1", "to desire or wish for something",   "Does she want a new phone?"),
        ("mean",       "verb",      "A1", "to have a particular meaning",      "What does this word mean?"),
        ("happen",     "verb",      "A1", "to take place; to occur",           "What happens next?"),
        ("use",        "verb",      "A1", "to do something with a tool/thing", "Do you use a dictionary?"),
        ("begin",      "verb",      "A1", "to start",                          "What time does class begin?"),
        ("belong",     "verb",      "A1", "to be owned by",                    "Does this bag belong to you?"),
        ("seem",       "verb",      "A1", "to appear to be",                   "Does she seem happy?"),
        ("how",        "adverb",    "A1", "in what way or manner",             "How do you get to school?"),
        ("when",       "adverb",    "A1", "at what time",                      "When does the film start?"),
        ("where",      "adverb",    "A1", "in or to what place",               "Where do you live?"),
        ("why",        "adverb",    "A1", "for what reason",                   "Why does he study English?"),
        ("what",       "pronoun",   "A1", "asking for information",            "What do you do?"),
        ("which",      "determiner","A1", "asking about a choice",             "Which bus do you take?"),
        ("usually",    "adverb",    "A1", "most of the time",                  "Do you usually walk?"),
        ("often",      "adverb",    "A1", "many times; frequently",            "How often do you exercise?"),
        ("together",   "adverb",    "A1", "with each other",                   "Do you work together?"),
        ("still",      "adverb",    "A1", "continuing until now",              "Does she still live there?"),
    ],
    "grammar": {
        "rule": "For Yes/No questions: Do/Does + subject + base verb? For Wh- questions: Wh-word + do/does + subject + base verb?",
        "form_headers": ["Type", "Structure", "Example"],
        "form_rows": [
            ["Yes/No (I/you/we/they)", "Do + subject + verb?",          "Do you like coffee?"],
            ["Yes/No (he/she/it)",     "Does + subject + verb?",        "Does she work here?"],
            ["Short answer (Yes)",     "Yes, + subject + do/does.",      "Yes, I do. / Yes, she does."],
            ["Short answer (No)",      "No, + subject + don't/doesn't.","No, I don't. / No, she doesn't."],
            ["Wh- question",           "Wh + do/does + subject + verb?","Where do you live?"],
        ],
        "form_col_widths": [2400, 2800, 3000],
        "note": "After Do or Does, ALWAYS use the base form. NEVER say 'Does she likes...?' The question word 'What do you do?' asks about jobs.",
        "usage": [
            "Asking about habits: Do you drink coffee? Does he smoke?",
            "Asking about jobs: What do you do? What does she do?",
            "Asking about place: Where do they live? Where does he work?",
            "Asking about time: When does the class start? What time do you wake up?",
            "Asking about reason: Why do you study English? Why does she travel?",
        ],
    },
    "guided": {
        "instruction": "Write the correct question form. Items 1-2 are completed for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "you / like / English (?)", "DO YOU LIKE English?"],
            ["2 ✅", "she / work / here (?)", "DOES SHE WORK here?"],
            ["3",   "they / live / in this city (?)", "________________________"],
            ["4",   "he / speak / French (?)", "________________________"],
            ["5",   "Where / you / work (?)", "________________________"],
            ["6",   "What time / she / start (?)", "________________________"],
            ["7",   "Why / they / study / English (?)", "________________________"],
            ["8",   "your brother / have / a car (?)", "________________________"],
        ],
        "col_widths": [500, 3500, 4200],
    },
    "controlled": {
        "instruction": "Write Yes/No questions and give short answers.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(you / like / your job?) - Yes",     "________________________"],
            ["2", "(she / live / near here?) - No",     "________________________"],
            ["3", "(they / know / the answer?) - Yes",  "________________________"],
            ["4", "(he / need / help?) - No",           "________________________"],
            ["5", "(your parents / travel / often?) - Yes", "________________________"],
            ["6", "(the bus / stop / here?) - No",      "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "Does she likes coffee?", "________________________"],
            ["2", "Do he live here?", "________________________"],
            ["3", "Where does they work?", "________________________"],
            ["4", "What you do?", "________________________"],
            ["5", "Does she has a car?", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "She lives in Paris. (make a Yes/No question)", "________________________"],
            ["2", "He works at a bank. (make a Where question)", "________________________"],
            ["3", "They study English. (make a Why question)", "________________________"],
            ["4", "I wake up at 7. (make a What time question with 'she')", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Find Out About Your Partner",
        "instructions": [
            "Write 5 questions to ask your partner (use Do you / Does your...).",
            "Ask and answer in pairs. Give short answers first, then full answers.",
            "Report interesting answers to the class: 'Maria doesn't like coffee!'",
            "Try at least 2 Wh- questions: Where / What / Why / When.",
        ],
    },
    "reading": {
        "text": (
            "Anna works in a bank. She starts at 8:30 and finishes at 5. She doesn't "
            "drive to work. She takes the train. Does she like her job? Yes, she does. "
            "She meets new people every day. Her colleagues are friendly. She doesn't "
            "work at the weekend, so she visits her parents on Saturdays."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Anna works in a hospital.",          "___"],
            ["2", "She drives to work.",                "___"],
            ["3", "She likes her job.",                 "___"],
            ["4", "Her colleagues are unfriendly.",     "___"],
            ["5", "She works on Saturdays.",            "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 5 questions and answers about a friend or family member.",
        "prompts": [
            "Write a Do question: Do you ___? Yes, I do. / No, I don't.",
            "Write a Does question: Does he/she ___? Yes, he/she does.",
            "Write a Where question: Where does ___?",
            "Write a What question: What do you ___?",
            "Write a When question: When does ___?",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 questions about a classmate's daily routine (use Do/Does). "
            "Then write the answers. Include at least 3 Wh- questions."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each word, write a question "
            "using Do or Does. Example: 'Do you have a hobby?'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 8 — Present Simple vs Present Continuous (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "8",
    "cefr": "A1",
    "topic": "Present Simple vs Present Continuous",
    "sub_focus": "I do (habits/general) vs I am doing (now); state verbs (like, want, know, understand, need, believe, remember, forget, depend, prefer, love, hate, mean) never use -ing",
    "objectives": [
        "Distinguish between Present Simple (habits/facts) and Present Continuous (happening now).",
        "Choose the correct tense based on time markers (every day vs right now).",
        "Identify state verbs that cannot be used in the continuous form.",
        "Produce sentences comparing what people generally do vs what they are doing now.",
        "Recognize signal words: always/usually/every vs now/at the moment/today.",
    ],
    "warmup": {
        "name": "📸 What's Different?",
        "time": "5 minutes",
        "steps": [
            "Teacher shows two pictures: one of a routine (She walks to work) and one of a current action (She is running to work).",
            "Students describe the difference: 'She WALKS every day, but today she IS RUNNING.'",
            "Pairs make similar contrasts about themselves: 'I usually sit here, but today I am standing.'",
        ],
    },
    "lead_in": [
        "What do you usually do on Saturday? And what are you doing THIS Saturday?",
        "Think of a friend. What does he/she do every day? What is he/she doing right now?",
    ],
    "vocab": [
        ("routine",     "noun",      "A1", "a regular series of actions",       "My routine is the same daily."),
        ("moment",      "noun",      "A1", "a very short period of time",       "Wait a moment, please."),
        ("difference",  "noun",      "A1", "a way things are not the same",     "I see no difference."),
        ("habit",       "noun",      "A1", "something you do regularly",        "Smoking is a bad habit."),
        ("action",      "noun",      "A1", "the process of doing something",    "The action is happening now."),
        ("fact",        "noun",      "A1", "something known to be true",        "That is a well-known fact."),
        ("situation",   "noun",      "A1", "the conditions at a particular time","The situation is changing."),
        ("example",     "noun",      "A1", "something used to explain",         "Give me an example."),
        ("today",       "adverb",    "A1", "on this present day",               "I am working from home today."),
        ("right now",   "adverb",    "A1", "at this exact moment",              "She is sleeping right now."),
        ("at the moment","adverb",   "A1", "now; currently",                    "He is busy at the moment."),
        ("generally",   "adverb",    "A1", "in most cases; usually",            "I generally walk to work."),
        ("currently",   "adverb",    "A2", "at the present time",               "She is currently studying."),
        ("temporary",   "adjective", "A2", "lasting for a short time only",     "This is a temporary job."),
        ("permanent",   "adjective", "A2", "lasting forever or a long time",    "Is this a permanent change?"),
        ("normal",      "adjective", "A1", "usual; typical",                    "This is not normal for me."),
        ("different",   "adjective", "A1", "not the same",                      "Today is different."),
        ("same",        "adjective", "A1", "not different; identical",           "It is always the same."),
        ("remember",    "verb",      "A1", "to keep in your memory",            "I remember his name."),
        ("depend",      "verb",      "A1", "to be influenced by something",     "It depends on the weather."),
        ("change",      "verb",      "A1", "to become different",               "Things are changing fast."),
        ("happen",      "verb",      "A1", "to take place; to occur",           "What is happening outside?"),
        ("seem",        "verb",      "A1", "to appear to be",                   "You seem tired today."),
        ("notice",      "verb",      "A1", "to become aware of",                "I notice a difference."),
        ("compare",     "verb",      "A2", "to look at similarities/differences","Compare the two sentences."),
        ("describe",    "verb",      "A1", "to say what something is like",     "Describe what you see."),
        ("choose",      "verb",      "A1", "to pick from options",              "Choose the correct answer."),
        ("guess",       "verb",      "A1", "to give an answer without knowing", "Guess what I am doing."),
        ("complete",    "verb",      "A1", "to finish; to make whole",          "Complete the sentence."),
        ("correct",     "adjective", "A1", "right; without mistakes",           "Is this answer correct?"),
    ],
    "grammar": {
        "rule": "Use Present Simple for habits, routines, and general facts. Use Present Continuous for actions happening right now or temporary situations.",
        "form_headers": ["Tense", "Use", "Time Words", "Example"],
        "form_rows": [
            ["Present Simple",     "habits / routines / facts",   "every day, always, usually, often, never", "She drinks coffee every morning."],
            ["Present Continuous", "happening NOW / temporary",    "now, right now, at the moment, today",     "She is drinking tea right now."],
            ["State verbs",        "NEVER continuous",             "always use Simple",                        "I know the answer. (NOT knowing)"],
        ],
        "form_col_widths": [1800, 2000, 2400, 2800],
        "note": "State verbs (like, want, know, understand, need, believe, remember, forget, depend, prefer, love, hate, mean) NEVER use -ing. Say 'I like it' NOT 'I am liking it'.",
        "usage": [
            "Habit vs now: I usually walk to work, but today I am taking the bus.",
            "Fact vs temporary: She works in London. This week she is working in Paris.",
            "State verb: I know the answer. (NOT I am knowing.)",
            "State verb: She wants a coffee. (NOT She is wanting.)",
            "Signal words: He ALWAYS eats lunch at 1 (Simple). He IS EATING lunch NOW (Continuous).",
        ],
    },
    "guided": {
        "instruction": "Choose Present Simple or Present Continuous. Items 1-2 are completed for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "She ___ (drink) coffee every morning.", "She DRINKS coffee every morning. [habit]"],
            ["2 ✅", "Look! He ___ (run) in the park now.", "He IS RUNNING in the park now. [now]"],
            ["3",   "I ___ (know) the answer. [state verb]", "________________________"],
            ["4",   "They ___ (watch) TV at the moment.", "________________________"],
            ["5",   "We ___ (go) to the gym every Friday.", "________________________"],
            ["6",   "She ___ (want) a new phone. [state verb]", "________________________"],
            ["7",   "Today he ___ (work) from home.", "________________________"],
            ["8",   "My mother ___ (cook) dinner every evening.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Complete with the correct tense: Present Simple or Present Continuous.",
        "headers": ["#", "Sentence", "Write your answer"],
        "items": [
            ["1", "She usually (walk) ___ but today she (take) ___ the bus.", "________________________"],
            ["2", "I (not understand) ___ this word. [state]", "________________________"],
            ["3", "He (read) ___ a book right now.", "________________________"],
            ["4", "They (play) ___ tennis every Saturday.", "________________________"],
            ["5", "Listen! Someone (sing) ___ outside.", "________________________"],
            ["6", "I (prefer) ___ tea to coffee. [state]", "________________________"],
            ["7", "We (learn) ___ Spanish this term. [temporary]", "________________________"],
            ["8", "The sun (rise) ___ in the east. [fact]", "________________________"],
        ],
        "col_widths": [500, 4200, 3500],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "I am knowing the answer.", "________________________"],
            ["2", "She is wanting a coffee.", "________________________"],
            ["3", "He drinks tea right now.", "________________________"],
            ["4", "They are usually going to the gym.", "________________________"],
            ["5", "I am believing you.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I walk to work. (happening now, use today)", "________________________"],
            ["2", "She is cooking dinner. (make it a routine, every day)", "________________________"],
            ["3", "He watches TV. (happening at this moment)", "________________________"],
            ["4", "I like this music. (explain why NOT 'I am liking')", "________________________"],
        ],
        "transform_col_widths": [500, 3800, 3900],
    },
    "speaking": {
        "name": "💬 What I Do vs What I'm Doing",
        "instructions": [
            "Make two lists: 5 things you USUALLY do + 5 things you are doing THIS WEEK differently.",
            "Tell your partner: 'I usually eat at home, but this week I am eating at a restaurant.'",
            "Your partner asks follow-up questions: 'Why are you eating at a restaurant?'",
            "Report to the class: 'Carlos usually walks, but this week he is taking the bus.'",
        ],
    },
    "reading": {
        "text": (
            "Marta lives in Madrid. She usually works in an office, but this month she "
            "is working from home. She normally takes the train, but now she doesn't need "
            "to travel. She likes the change because she has more time with her family. "
            "Right now she is sitting in her garden and drinking coffee."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Marta usually works from home.",         "___"],
            ["2", "This month she is working from home.",   "___"],
            ["3", "She normally drives to work.",           "___"],
            ["4", "She likes working from home.",           "___"],
            ["5", "Right now she is in her office.",        "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write sentences comparing your habits with what you are doing today.",
        "prompts": [
            "I usually ___ but today I am ___.",
            "My friend normally ___ but this week he/she is ___.",
            "I always ___ every day, but right now I am ___.",
            "I ___ (state verb: like/know/want/need) because ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences: 4 about your usual routine (Present Simple) and 4 about "
            "what is different today or this week (Present Continuous). Underline the verbs."
        ),
        "item2": (
            "Write the 13 state verbs from the lesson in your notebook. For each one, write "
            "a correct example sentence in Present Simple. Do NOT use -ing with any of them."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 9 — Have Got / Has Got (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "9",
    "cefr": "A1",
    "topic": "Have Got / Has Got",
    "sub_focus": "I've got / He's got; haven't got / hasn't got; Have you got?; also don't have / doesn't have",
    "objectives": [
        "Use have got / has got to talk about possession and relationships.",
        "Form negatives: haven't got / hasn't got.",
        "Form questions: Have you got? / Has she got?",
        "Give short answers: Yes, I have. / No, she hasn't.",
        "Understand the difference between have got (British) and have (American).",
    ],
    "warmup": {
        "name": "🎒 What's in Your Bag?",
        "time": "5 minutes",
        "steps": [
            "Without looking, guess 3 things your partner has got in their bag.",
            "Ask: 'Have you got a pen? Have you got a phone? Have you got keys?'",
            "Your partner checks and answers: 'Yes, I have.' or 'No, I haven't.'",
        ],
    },
    "lead_in": [
        "Have you got a pet? Have you got brothers or sisters? Tell the class.",
        "What has your teacher got on the desk? Describe what you can see.",
    ],
    "vocab": [
        ("pet",        "noun",      "A1", "an animal kept at home",            "I've got a pet cat."),
        ("garden",     "noun",      "A1", "a piece of land for growing plants","She's got a big garden."),
        ("phone",      "noun",      "A1", "a device for communication",        "Have you got a phone?"),
        ("wallet",     "noun",      "A1", "a small case for money/cards",      "He hasn't got his wallet."),
        ("umbrella",   "noun",      "A1", "something to keep rain off",        "I've got an umbrella."),
        ("bicycle",    "noun",      "A1", "a two-wheeled vehicle",             "She's got a new bicycle."),
        ("headache",   "noun",      "A1", "a pain in the head",                "I've got a headache."),
        ("fever",      "noun",      "A1", "a high body temperature",           "He's got a fever today."),
        ("toothache",  "noun",      "A1", "a pain in a tooth",                 "She's got a toothache."),
        ("appointment","noun",      "A1", "an arranged meeting time",          "I've got an appointment."),
        ("enough",     "adjective", "A1", "as much as needed",                 "Have you got enough time?"),
        ("spare",      "adjective", "A1", "extra; not being used",             "I've got a spare pen."),
        ("own",        "adjective", "A1", "belonging to that person",          "She's got her own room."),
        ("free",       "adjective", "A1", "not busy; available",               "Have you got free time?"),
        ("ready",      "adjective", "A1", "prepared for something",            "I haven't got everything ready."),
        ("dark",       "adjective", "A1", "with little or no light",           "She's got dark hair."),
        ("fair",       "adjective", "A1", "light in colour (hair/skin)",       "He's got fair skin."),
        ("curly",      "adjective", "A1", "having curls or curves",            "She's got curly hair."),
        ("straight",   "adjective", "A1", "not bent or curving",               "He's got straight hair."),
        ("long",       "adjective", "A1", "measuring a great distance",        "She's got long legs."),
        ("short",      "adjective", "A1", "not long; not tall",                "He's got short hair."),
        ("blue",       "adjective", "A1", "having a blue colour",              "She's got blue eyes."),
        ("brown",      "adjective", "A1", "having a brown colour",             "He's got brown eyes."),
        ("younger",    "adjective", "A1", "less old (comparative)",            "I've got a younger sister."),
        ("older",      "adjective", "A1", "more old (comparative)",            "She's got an older brother."),
        ("idea",       "noun",      "A1", "a thought or plan",                 "I've got an idea!"),
        ("chance",     "noun",      "A1", "a possibility of something",        "Have you got a chance?"),
        ("ticket",     "noun",      "A1", "a piece of paper for entry",        "He's got two tickets."),
        ("license",    "noun",      "A1", "an official document of permission","Has she got a license?"),
        ("cold",       "noun",      "A1", "a common illness",                  "I've got a cold."),
    ],
    "grammar": {
        "rule": "Use have got (I/you/we/they) or has got (he/she/it) to talk about possession, family, illness, and appearance.",
        "form_headers": ["Form", "I / You / We / They", "He / She / It"],
        "form_rows": [
            ["Affirmative",   "I have got (I've got)",         "He has got (He's got)"],
            ["Negative",      "I have not got (haven't got)",  "She has not got (hasn't got)"],
            ["Question",      "Have you got...?",              "Has he got...?"],
            ["Short answer +","Yes, I have.",                  "Yes, she has."],
            ["Short answer -","No, we haven't.",               "No, he hasn't."],
        ],
        "form_col_widths": [1800, 3200, 3200],
        "note": "Have got = have (same meaning). 'I've got a car' = 'I have a car'. In questions: 'Have you got?' = 'Do you have?'. Do NOT say 'I've got to' for obligation here; that is a different use.",
        "usage": [
            "Possession: I've got a new phone. She hasn't got a car.",
            "Family: He's got two brothers. Have you got any sisters?",
            "Appearance: She's got blue eyes and dark hair.",
            "Illness: I've got a headache. He's got a cold.",
            "Available items: Have you got a pen? I haven't got enough money.",
        ],
    },
    "guided": {
        "instruction": "Complete using have got / has got (or the negative/question form). Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ (have got) a new bicycle.", "I'VE GOT a new bicycle."],
            ["2 ✅", "___ she ___ (have got) blue eyes?", "HAS she GOT blue eyes?"],
            ["3",   "He ___ (not / have got) a wallet.", "________________________"],
            ["4",   "They ___ (have got) two children.", "________________________"],
            ["5",   "___ you ___ (have got) a pet?", "________________________"],
            ["6",   "She ___ (have got) long, curly hair.", "________________________"],
            ["7",   "We ___ (not / have got) enough time.", "________________________"],
            ["8",   "___ he ___ (have got) a headache?", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write sentences using have got / has got based on the prompts.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / a younger brother)",            "________________________"],
            ["2", "(she / not / a car)",                "________________________"],
            ["3", "(they / a big garden)",              "________________________"],
            ["4", "(he / not / any pets)",              "________________________"],
            ["5", "(you / a spare pen?)",               "________________________"],
            ["6", "(she / dark hair and blue eyes)",    "________________________"],
            ["7", "(we / not / enough money)",          "________________________"],
            ["8", "(he / a cold?)",                     "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She have got a new phone.", "________________________"],
            ["2", "Has they got any children?", "________________________"],
            ["3", "I hasn't got a ticket.", "________________________"],
            ["4", "He's not got got a car.", "________________________"],
            ["5", "Do she has got blue eyes?", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I have a car. (rewrite with have got)",         "________________________"],
            ["2", "She's got a cat. (make negative)",              "________________________"],
            ["3", "They've got three children. (make a question)","________________________"],
            ["4", "He doesn't have a garden. (use hasn't got)",   "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 Describe Your Partner",
        "instructions": [
            "Look at your partner carefully for 30 seconds.",
            "Turn away and describe them: 'She's got brown hair. She's got a blue bag.'",
            "Your partner confirms: 'Yes, I have.' or corrects: 'No, I haven't. I've got a black bag.'",
            "Ask 3 extra questions: 'Have you got a pet? Have you got any brothers?'",
        ],
    },
    "reading": {
        "text": (
            "My name is Leo. I've got a big family. I've got two brothers and one sister. "
            "My older brother has got a car, but my younger brother hasn't got one yet. "
            "My sister has got curly hair and brown eyes. We've got a dog called Max. "
            "He's got short legs and a long tail. Have you got any pets?"
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Leo has got one brother.",              "___"],
            ["2", "His older brother has got a car.",      "___"],
            ["3", "His sister has got straight hair.",     "___"],
            ["4", "They've got a cat.",                    "___"],
            ["5", "The dog has got short legs.",           "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences using have got / has got about yourself and people you know.",
        "prompts": [
            "Describe your appearance: I've got ___ hair and ___ eyes.",
            "Talk about your family: I've got ___ (number of siblings).",
            "Talk about possession: I've got a ___. I haven't got a ___.",
            "Describe a friend: He/She's got ___.",
            "Ask a question: Have you got ___?",
            "Talk about health: I've got a ___ / I haven't got a ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences describing your family members using have got / has got. "
            "Include appearance (hair, eyes), possessions, and at least 2 negatives (haven't/hasn't got)."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each word, write a sentence using "
            "have got or has got. Example: 'My sister has got curly hair.'"
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 10 — Verb To Be: Past Simple (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "10",
    "cefr": "A1",
    "topic": "Verb To Be — Past Simple",
    "sub_focus": "was (I/he/she/it); were (you/we/they); wasn't / weren't; questions and short answers",
    "objectives": [
        "Use was for I/he/she/it and were for you/we/they in past sentences.",
        "Form negatives with wasn't and weren't.",
        "Ask questions: Was he...? Were they...?",
        "Give short answers: Yes, I was. / No, they weren't.",
        "Talk about past states, locations, and descriptions using was/were.",
    ],
    "warmup": {
        "name": "🕰️ Yesterday vs Today",
        "time": "5 minutes",
        "steps": [
            "Think about yesterday. Where were you at 10 AM? How were you feeling?",
            "Tell your partner: 'Yesterday I was at home. I was tired.'",
            "Your partner asks: 'Were you alone? Was it a good day?'",
        ],
    },
    "lead_in": [
        "Where were you last weekend? Were you at home or somewhere else?",
        "Was the weather good yesterday? Was it hot or cold?",
    ],
    "vocab": [
        ("yesterday",  "adverb",    "A1", "the day before today",             "I was at home yesterday."),
        ("last",       "adjective", "A1", "most recent; the one before now",  "She was here last week."),
        ("ago",        "adverb",    "A1", "in the past; before now",          "That was two days ago."),
        ("born",       "adjective", "A1", "brought into life",                "He was born in 1995."),
        ("absent",     "adjective", "A1", "not present; away",                "She was absent yesterday."),
        ("present",    "adjective", "A1", "in a particular place",            "Were you present in class?"),
        ("closed",     "adjective", "A1", "not open",                         "The shop was closed."),
        ("empty",      "adjective", "A1", "containing nothing",               "The room was empty."),
        ("full",       "adjective", "A1", "containing as much as possible",   "The bus was full."),
        ("crowded",    "adjective", "A1", "full of people",                   "The park was crowded."),
        ("famous",     "adjective", "A1", "known by many people",             "He was a famous writer."),
        ("quiet",      "adjective", "A1", "making little noise",              "The house was very quiet."),
        ("loud",       "adjective", "A1", "making a lot of noise",            "The music was too loud."),
        ("scared",     "adjective", "A1", "feeling afraid",                   "I was scared of the dark."),
        ("surprised",  "adjective", "A1", "feeling unexpected wonder",        "We were surprised."),
        ("excited",    "adjective", "A1", "feeling happy and enthusiastic",   "They were very excited."),
        ("upset",      "adjective", "A1", "unhappy or disappointed",          "She was upset about it."),
        ("sick",       "adjective", "A1", "not well; ill",                    "He was sick last Monday."),
        ("lucky",      "adjective", "A1", "having good fortune",              "We were lucky that day."),
        ("difficult",  "adjective", "A1", "hard to do or understand",         "The test was difficult."),
        ("easy",       "adjective", "A1", "not hard; simple",                 "The homework was easy."),
        ("terrible",   "adjective", "A1", "very bad",                         "The weather was terrible."),
        ("wonderful",  "adjective", "A1", "extremely good",                   "The trip was wonderful."),
        ("right",      "adjective", "A1", "correct; true",                    "Your answer was right."),
        ("safe",       "adjective", "A1", "free from danger",                 "We were safe at home."),
        ("busy",       "adjective", "A1", "having a lot to do",               "She was busy all day."),
        ("available",  "adjective", "A2", "free to use or see",               "The room wasn't available."),
        ("century",    "noun",      "A1", "a period of 100 years",            "It was the 20th century."),
        ("war",        "noun",      "A1", "a fight between countries",         "The war was terrible."),
        ("childhood",  "noun",      "A1", "the time of being a child",        "My childhood was happy."),
    ],
    "grammar": {
        "rule": "The past of am/is is was. The past of are is were. Use was/were to talk about past states, feelings, and locations.",
        "form_headers": ["Form", "I / He / She / It", "You / We / They"],
        "form_rows": [
            ["Affirmative",   "was",              "were"],
            ["Negative",      "was not (wasn't)", "were not (weren't)"],
            ["Question",      "Was I/he/she/it...?","Were you/we/they...?"],
            ["Short answer +","Yes, I was.",       "Yes, they were."],
            ["Short answer -","No, he wasn't.",    "No, we weren't."],
        ],
        "form_col_widths": [1800, 3200, 3200],
        "note": "Was/were is the ONLY past form of to be. Do NOT say 'I were' or 'He were'. I/he/she/it = WAS. You/we/they = WERE.",
        "usage": [
            "Past location: I was at home yesterday. They were at school.",
            "Past feeling: She was happy. We were tired after the trip.",
            "Past description: The film was boring. The questions were easy.",
            "Past time: It was Monday. It was 3 o'clock.",
            "Birth: She was born in 2001. They were born in London.",
        ],
    },
    "guided": {
        "instruction": "Complete with was, were, wasn't, or weren't. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ at home yesterday.", "I WAS at home yesterday."],
            ["2 ✅", "They ___ not at school.", "They WERE not at school. (WEREN'T)"],
            ["3",   "She ___ born in 2003.", "________________________"],
            ["4",   "The children ___ very excited.", "________________________"],
            ["5",   "___ you at the party last night?", "________________________"],
            ["6",   "The shops ___ (not) open yesterday.", "________________________"],
            ["7",   "It ___ a beautiful day last Sunday.", "________________________"],
            ["8",   "___ he sick last week? Yes, he ___.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write complete sentences in the past using was or were.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "(I / tired / yesterday)",              "________________________"],
            ["2", "(the weather / terrible / last week)", "________________________"],
            ["3", "(they / not / at home / last night)",  "________________________"],
            ["4", "(the test / easy?)",                   "________________________"],
            ["5", "(we / lucky / on holiday)",            "________________________"],
            ["6", "(she / not / happy / about it)",       "________________________"],
            ["7", "(you / present / in class?)",          "________________________"],
            ["8", "(the room / empty / when I arrived)",  "________________________"],
        ],
        "col_widths": [500, 3600, 4100],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "They was at school yesterday.", "________________________"],
            ["2", "He were born in London.", "________________________"],
            ["3", "I were tired after the trip.", "________________________"],
            ["4", "Was they surprised?", "________________________"],
            ["5", "She weren't at work.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I am tired today. (change to yesterday)", "________________________"],
            ["2", "They are at school. (change to last week)", "________________________"],
            ["3", "She was happy. (make negative)", "________________________"],
            ["4", "It was cold. (make a question)", "________________________"],
        ],
        "transform_col_widths": [500, 3600, 4100],
    },
    "speaking": {
        "name": "💬 My Last Weekend",
        "instructions": [
            "Tell your partner about last weekend. Use: I was... / I wasn't... / It was...",
            "Ask your partner questions: Were you at home? Was it fun? Was the weather good?",
            "Give short answers: Yes, I was. / No, it wasn't.",
            "Report to the class: 'Maria was at the beach. The weather was wonderful.'",
        ],
    },
    "reading": {
        "text": (
            "Last Saturday was my birthday. I was at a restaurant with my family. The food "
            "was wonderful and the music was loud. My parents were very happy. My little "
            "brother was excited because there was a big cake. The restaurant was crowded, "
            "but our table was ready at seven. It was a perfect evening."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "The birthday was on Sunday.",            "___"],
            ["2", "The food was wonderful.",                "___"],
            ["3", "The music was quiet.",                   "___"],
            ["4", "The restaurant was empty.",              "___"],
            ["5", "The table was ready at seven.",          "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about a past event using was and were.",
        "prompts": [
            "Where were you last weekend? I was at ___.",
            "How was the weather? The weather was ___.",
            "How were you feeling? I was ___.",
            "Was it a good day? Yes, it was ___ / No, it wasn't ___.",
            "Were your friends with you? They were ___ / They weren't ___.",
            "Describe something: The ___ was ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about a trip or event in the past. Use was/were at least 6 times. "
            "Include at least 2 negatives (wasn't/weren't) and 2 questions (Was/Were...?)."
        ),
        "item2": (
            "Choose 10 vocabulary words from Section 4. For each word, write a sentence "
            "using was or were. Example: 'The park was crowded last Sunday.'"
        ),
    },
})


# ─────────────────────────────────────────────────────────────────────────────
# Generate all DOCX + PPTX files
# ─────────────────────────────────────────────────────────────────────────────

OUT_DIR = "/projects/sandbox/APP-GENERATION-PPT-FILE-ETC/generated"
os.makedirs(OUT_DIR, exist_ok=True)

topic_slugs = [
    "PresentSimple_Affirmative",
    "PresentSimple_Negative",
    "PresentSimple_Questions",
    "PresentSimple_vs_PresentContinuous",
    "HaveGot_HasGot",
    "VerbToBe_PastSimple",
]

if __name__ == "__main__":
    for i, (unit, slug) in enumerate(zip(UNITS_5_10, topic_slugs)):
        unit_num = i + 5

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

    print(f"\nAll {len(UNITS_5_10) * 2} files generated successfully (6 DOCX + 6 PPTX).")

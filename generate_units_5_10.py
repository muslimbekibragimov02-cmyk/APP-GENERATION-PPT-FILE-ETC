"""
ELT Grammar Course Designer — Units 5-10 Generator
Generates Level 1 Units 5-10 (DOCX + PPTX) using build_docx and build_pptx.
"""

import os
import sys

sys.path.insert(0, '/projects/sandbox/APP-GENERATION-PPT-FILE-ETC')
from generate_docx import build_docx
from generate_pptx import build_pptx

# ─────────────────────────────────────────────────────────────────────────────
# UNIT DATA — Units 5-10
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
    "sub_focus": "he/she/it + -s/-es; I/you/we/they + base form; daily routines",
    "objectives": [
        "Form affirmative present simple sentences with I/you/we/they + base verb.",
        "Add -s or -es correctly with he/she/it subjects.",
        "Apply spelling rules for -es (goes, watches) and consonant+y (studies).",
        "Use the present simple to describe daily routines and habits.",
        "Produce sentences about other people's habits using third person -s.",
    ],
    "warmup": {
        "name": "🕐 My Morning Routine",
        "time": "5 minutes",
        "steps": [
            "Write three things you do every morning using 'I + verb'.",
            "Share with a partner. Your partner reports to the class using 'He/She + verb-s'.",
            "Example: 'I drink coffee.' -> Partner says: 'She drinks coffee.'",
        ],
    },
    "lead_in": [
        "What time do you usually wake up? What do you do first?",
        "Tell me about your best friend. What does he or she do every day?",
    ],
    "vocab": [
        ("wake up",    "verb",      "A1", "to stop sleeping",                  "I wake up at seven."),
        ("get up",     "verb",      "A1", "to leave your bed",                 "She gets up early."),
        ("brush",      "verb",      "A1", "to clean with a brush",             "He brushes his teeth."),
        ("wash",       "verb",      "A1", "to clean with water",               "She washes her face."),
        ("have",       "verb",      "A1", "to eat or drink",                   "We have breakfast at eight."),
        ("go",         "verb",      "A1", "to move to another place",          "He goes to work by bus."),
        ("start",      "verb",      "A1", "to begin doing something",          "School starts at nine."),
        ("finish",     "verb",      "A1", "to complete or end",                "He finishes work at five."),
        ("cook",       "verb",      "A1", "to prepare food using heat",        "My mother cooks dinner."),
        ("watch",      "verb",      "A1", "to look at something for a time",   "He watches TV in the evening."),
        ("listen",     "verb",      "A1", "to pay attention to sounds",        "She listens to music."),
        ("read",       "verb",      "A1", "to look at words and understand",   "They read books at night."),
        ("work",       "verb",      "A1", "to do a job or task",               "My father works in a factory."),
        ("live",       "verb",      "A1", "to have your home somewhere",       "We live in London."),
        ("study",      "verb",      "A1", "to learn about a subject",          "She studies English every day."),
        ("play",       "verb",      "A1", "to take part in a game",            "They play football on Saturdays."),
        ("drive",      "verb",      "A1", "to operate a car",                  "He drives to the office."),
        ("walk",       "verb",      "A1", "to move on foot",                   "I walk to school."),
        ("sleep",      "verb",      "A1", "to rest with eyes closed",          "The baby sleeps a lot."),
        ("drink",      "verb",      "A1", "to take liquid into the mouth",     "She drinks tea every morning."),
        ("eat",        "verb",      "A1", "to put food in the mouth",          "We eat lunch at noon."),
        ("teach",      "verb",      "A1", "to give lessons to students",       "She teaches maths."),
        ("speak",      "verb",      "A1", "to say words; talk",                "He speaks three languages."),
        ("leave",      "verb",      "A1", "to go away from a place",           "She leaves home at eight."),
        ("arrive",     "verb",      "A1", "to reach a place",                  "The bus arrives at nine."),
        ("usually",    "adverb",    "A1", "most of the time",                  "I usually walk to work."),
        ("always",     "adverb",    "A1", "every time; all the time",          "She always gets up early."),
        ("often",      "adverb",    "A1", "many times; frequently",            "We often eat together."),
        ("every day",  "adverb",    "A1", "each day",                          "He goes to the gym every day."),
        ("sometimes",  "adverb",    "A1", "on some occasions",                 "I sometimes watch films."),
    ],
    "grammar": {
        "rule": "Use the present simple to talk about habits, routines, and facts. Add -s/-es to the verb for he/she/it.",
        "form_headers": ["Subject", "Verb form", "Example"],
        "form_rows": [
            ["I / You / We / They", "base form (no change)", "I work in a bank."],
            ["He / She / It",       "verb + -s",              "She works in a hospital."],
            ["He / She / It",       "verb + -es (go, do, watch, wash)", "He watches TV."],
            ["He / She / It",       "consonant + y -> -ies",  "She studies English."],
        ],
        "form_col_widths": [2200, 3000, 3000],
        "spelling_headers": ["Rule", "Base form", "He/She/It form"],
        "spelling_rows": [
            ["Most verbs: add -s",         "work, play, live",    "works, plays, lives"],
            ["-ch, -sh, -ss, -x, -o: add -es", "watch, go, wash", "watches, goes, washes"],
            ["Consonant + y: drop y, add -ies", "study, carry, fly", "studies, carries, flies"],
            ["Vowel + y: just add -s",     "play, say, enjoy",    "plays, says, enjoys"],
            ["Irregular: have -> has",     "have",                 "has"],
        ],
        "spelling_col_widths": [2800, 2400, 3000],
        "note": "Remember: I have breakfast. BUT She has breakfast. The verb 'have' is irregular in third person.",
        "usage": [
            "Daily routines: I wake up at seven. She gets up at six.",
            "Habits: He drinks coffee every morning.",
            "Facts: The sun rises in the east.",
            "Jobs: My father works in a hospital.",
            "Frequency: We always eat dinner together.",
        ],
    },
    "guided": {
        "instruction": "Complete the sentences with the correct present simple form. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "She ___ (get up) at 6 a.m.", "She GETS UP at 6 a.m."],
            ["2 ✅", "They ___ (play) football on Saturdays.", "They PLAY football on Saturdays."],
            ["3",   "He ___ (watch) TV every evening.", "________________________"],
            ["4",   "I ___ (walk) to school.", "________________________"],
            ["5",   "She ___ (study) English after dinner.", "________________________"],
            ["6",   "We ___ (have) lunch at noon.", "________________________"],
            ["7",   "My father ___ (drive) to work.", "________________________"],
            ["8",   "The train ___ (leave) at 8 o'clock.", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write the verb in brackets in the correct present simple form.",
        "headers": ["#", "Sentence", "Write your answer"],
        "items": [
            ["1", "Maria ___ (teach) at a school.",        "________________________"],
            ["2", "My brother ___ (go) to the gym.",       "________________________"],
            ["3", "They ___ (live) in a small town.",      "________________________"],
            ["4", "She ___ (wash) her hair every day.",    "________________________"],
            ["5", "He ___ (finish) work at six.",          "________________________"],
            ["6", "The shop ___ (close) at nine p.m.",     "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She go to work by bus.",           "________________________"],
            ["2", "He watchs TV every night.",        "________________________"],
            ["3", "They plays football on Sundays.",  "________________________"],
            ["4", "She studys English after school.", "________________________"],
            ["5", "My mother haves breakfast early.", "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I get up at seven. (change to She)",           "________________________"],
            ["2", "We have lunch at noon. (change to He)",        "________________________"],
            ["3", "They watch TV in the evening. (change to My sister)", "________________________"],
            ["4", "I study French. (change to Tom)",              "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Daily Routine Interview",
        "instructions": [
            "Work in pairs. Ask your partner about their daily routine.",
            "Use questions like: 'What time do you wake up? What do you do after school?'",
            "Listen and take notes. Then tell the class about your partner.",
            "Use he/she + -s: 'Maria gets up at seven. She walks to school.'",
        ],
    },
    "reading": {
        "text": (
            "Tom is a bus driver. He wakes up at 5 a.m. every day. He has a quick breakfast and "
            "leaves home at 5:30. He drives a big red bus in the city centre. He finishes work at "
            "2 p.m. After work, he goes to the gym. In the evening, he watches TV or reads a book. "
            "He goes to bed at 10 p.m. On Saturdays, he plays football with his friends."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Tom wakes up at 6 a.m.",               "___"],
            ["2", "He drives a bus.",                      "___"],
            ["3", "He finishes work at 5 p.m.",           "___"],
            ["4", "He goes to the gym after work.",       "___"],
            ["5", "He plays football on Sundays.",        "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about your daily routine using the present simple.",
        "prompts": [
            "What time do you wake up? I wake up at ___.",
            "How do you go to school/work? I ___ to school/work.",
            "What do you do in the afternoon? I ___.",
            "What does your mother/father do every day? My ___ ___.",
            "What do you do in the evening? In the evening, I ___.",
            "What time do you go to bed? I go to bed at ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about a family member's daily routine. "
            "Use he/she + present simple verbs with correct -s/-es spelling. "
            "Underline all third-person verb endings."
        ),
        "item2": (
            "Choose 10 verbs from the vocabulary list. Write the base form and the "
            "he/she/it form for each verb. Example: go -> goes, study -> studies."
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
    "sub_focus": "don't / doesn't + base form; I/you/we/they vs he/she/it",
    "objectives": [
        "Form negative present simple sentences using don't + base verb.",
        "Form negative present simple sentences using doesn't + base verb.",
        "Remember that the main verb stays in the base form after doesn't.",
        "Contrast affirmative and negative forms in the same context.",
        "Produce negative sentences about habits and preferences.",
    ],
    "warmup": {
        "name": "🚫 Things I Don't Do",
        "time": "5 minutes",
        "steps": [
            "Write three things you DON'T do in the morning (e.g., 'I don't eat breakfast.').",
            "Share with a partner. Find one thing you both don't do.",
            "Report to the class: 'We don't watch TV in the morning.'",
        ],
    },
    "lead_in": [
        "Do you like getting up early? Why or why not?",
        "Think of your best friend. What doesn't he or she do?",
    ],
    "vocab": [
        ("like",       "verb",      "A1", "to enjoy; to find pleasant",        "I don't like mornings."),
        ("want",       "verb",      "A1", "to wish to have or do",             "She doesn't want coffee."),
        ("need",       "verb",      "A1", "to require something",              "We don't need help."),
        ("know",       "verb",      "A1", "to have information about",         "He doesn't know the answer."),
        ("understand", "verb",      "A1", "to comprehend the meaning",         "I don't understand this word."),
        ("remember",   "verb",      "A1", "to keep in mind",                   "She doesn't remember my name."),
        ("forget",     "verb",      "A1", "to fail to remember",               "They don't forget homework."),
        ("believe",    "verb",      "A1", "to accept as true",                 "I don't believe that story."),
        ("agree",      "verb",      "A1", "to have the same opinion",          "He doesn't agree with me."),
        ("hate",       "verb",      "A1", "to dislike very strongly",          "She doesn't hate her job."),
        ("enjoy",      "verb",      "A1", "to get pleasure from",              "They don't enjoy cooking."),
        ("prefer",     "verb",      "A2", "to like one thing more than another","He doesn't prefer tea."),
        ("smoke",      "verb",      "A1", "to breathe in cigarette smoke",     "I don't smoke."),
        ("run",        "verb",      "A1", "to move fast on foot",              "She doesn't run in the morning."),
        ("swim",       "verb",      "A1", "to move through water",             "We don't swim in winter."),
        ("dance",      "verb",      "A1", "to move the body to music",         "He doesn't dance well."),
        ("sing",       "verb",      "A1", "to make music with the voice",      "I don't sing in public."),
        ("travel",     "verb",      "A1", "to go on a journey",                "They don't travel often."),
        ("spend",      "verb",      "A1", "to use money or time",              "She doesn't spend a lot."),
        ("exercise",   "noun",      "A1", "physical activity for health",      "He doesn't do exercise."),
        ("breakfast",  "noun",      "A1", "the first meal of the day",         "I don't eat breakfast."),
        ("lunch",      "noun",      "A1", "a meal in the middle of the day",   "He doesn't have lunch at home."),
        ("dinner",     "noun",      "A1", "the main meal of the day",          "We don't cook dinner on Fridays."),
        ("coffee",     "noun",      "A1", "a hot dark drink",                  "She doesn't drink coffee."),
        ("meat",       "noun",      "A1", "food from animals",                 "They don't eat meat."),
        ("early",      "adverb",    "A1", "before the usual time",             "He doesn't wake up early."),
        ("never",      "adverb",    "A1", "not at any time",                   "I never eat fast food."),
        ("together",   "adverb",    "A1", "with each other",                   "We don't live together."),
        ("alone",      "adverb",    "A1", "without other people",              "She doesn't like being alone."),
        ("really",     "adverb",    "A1", "truly; very much",                  "I don't really like sport."),
    ],
    "grammar": {
        "rule": "To make a negative in present simple, use do not (don't) or does not (doesn't) + base verb. The main verb has NO -s.",
        "form_headers": ["Subject", "Negative (full)", "Negative (contraction)", "Example"],
        "form_rows": [
            ["I / You / We / They", "do not + base verb",   "don't + base verb",   "I don't like coffee."],
            ["He / She / It",       "does not + base verb", "doesn't + base verb", "She doesn't like tea."],
        ],
        "form_col_widths": [2000, 2000, 2000, 2200],
        "note": "After doesn't, the verb is ALWAYS in the base form. She doesn't LIKES is WRONG. She doesn't LIKE is correct.",
        "usage": [
            "Habits: I don't eat breakfast. He doesn't drink coffee.",
            "Preferences: She doesn't like sport. We don't prefer tea.",
            "Facts: Penguins don't fly. A cat doesn't swim.",
            "Routines: They don't work on Sundays. He doesn't drive to work.",
            "Abilities: I don't speak French. She doesn't understand German.",
        ],
    },
    "guided": {
        "instruction": "Make negative sentences using don't or doesn't. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I / like / mornings",            "I don't like mornings."],
            ["2 ✅", "She / drink / coffee",           "She doesn't drink coffee."],
            ["3",   "He / eat / meat",                "________________________"],
            ["4",   "We / live / in London",          "________________________"],
            ["5",   "They / work / on Saturdays",     "________________________"],
            ["6",   "My brother / smoke",             "________________________"],
            ["7",   "I / understand / this exercise", "________________________"],
            ["8",   "She / want / to go out",         "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Rewrite each affirmative sentence as a negative sentence.",
        "headers": ["#", "Affirmative sentence", "Write your answer"],
        "items": [
            ["1", "She likes football.",              "________________________"],
            ["2", "I eat meat.",                      "________________________"],
            ["3", "They live in Paris.",              "________________________"],
            ["4", "He smokes.",                       "________________________"],
            ["5", "We travel a lot.",                 "________________________"],
            ["6", "My sister drinks coffee.",         "________________________"],
            ["7", "Tom watches TV every day.",        "________________________"],
            ["8", "I remember her name.",             "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She don't like coffee.",            "________________________"],
            ["2", "He doesn't likes football.",        "________________________"],
            ["3", "They doesn't eat meat.",            "________________________"],
            ["4", "I doesn't want to go.",             "________________________"],
            ["5", "We not live in London.",            "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I like sport. (make negative)",                    "________________________"],
            ["2", "She watches TV. (make negative)",                  "________________________"],
            ["3", "He does not eat breakfast. (use contraction)",     "________________________"],
            ["4", "They do not travel often. (use contraction)",      "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Find the Difference",
        "instructions": [
            "Work in pairs. Tell your partner 4 things you do every day.",
            "Your partner says what they DON'T do: 'I don't get up early.'",
            "Find 3 differences between you.",
            "Report to the class: 'I eat breakfast, but Maria doesn't eat breakfast.'",
        ],
    },
    "reading": {
        "text": (
            "Lisa is a vegetarian. She doesn't eat meat or fish. She doesn't drink milk either. "
            "She eats a lot of vegetables and fruit. Her husband, Mark, likes meat but he doesn't "
            "cook it at home. They don't argue about food. They both enjoy cooking together. "
            "On weekends, they don't eat at home. They go to their favourite restaurant."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Lisa eats meat.",                       "___"],
            ["2", "She doesn't drink milk.",               "___"],
            ["3", "Mark doesn't like meat.",               "___"],
            ["4", "They argue about food.",                "___"],
            ["5", "They eat at a restaurant on weekends.", "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 negative sentences about yourself or people you know.",
        "prompts": [
            "Something you don't eat: I don't eat ___.",
            "Something you don't do in the morning: I don't ___.",
            "Something your friend doesn't like: My friend doesn't like ___.",
            "Something your family doesn't do on weekdays: We don't ___.",
            "Something a family member doesn't do: My ___ doesn't ___.",
            "Something you don't understand: I don't understand ___.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about things people in your family DON'T do. "
            "Use don't for I/we/they and doesn't for he/she/it. "
            "Example: My mother doesn't eat chocolate. We don't watch TV at night."
        ),
        "item2": (
            "Choose 10 words from the vocabulary list. Write one affirmative sentence "
            "and one negative sentence for each word. Example: like -> I like pizza. / I don't like fish."
        ),
    },
})

# ══════════════════════════════════════════════════════════════════════════════
# UNIT 7 — Present Simple: Questions (A1)
# ══════════════════════════════════════════════════════════════════════════════
UNITS_5_10.append({
    "level": "1",
    "course": "English Grammar in Action",
    "unit_num": "7",
    "cefr": "A1",
    "topic": "Present Simple — Questions",
    "sub_focus": "Do/Does + subject + base form? Short answers; Wh- questions",
    "objectives": [
        "Form Yes/No questions using Do/Does + subject + base verb.",
        "Give correct short answers: Yes, I do. / No, she doesn't.",
        "Form Wh- questions with What/Where/When/How often + do/does.",
        "Apply correct word order in present simple questions.",
        "Use questions to ask about routines, habits, and preferences.",
    ],
    "warmup": {
        "name": "🎲 Question Dice",
        "time": "5 minutes",
        "steps": [
            "Teacher gives each pair a list of verbs: eat, go, play, like, work, live.",
            "Student A picks a verb and makes a question: 'Do you play tennis?'",
            "Student B answers with a short answer: 'Yes, I do.' or 'No, I don't.'",
            "Switch roles and continue until all verbs are used.",
        ],
    },
    "lead_in": [
        "What questions do you ask when you meet someone new?",
        "How do you find out about a person's daily life?",
    ],
    "vocab": [
        ("where",      "adverb",    "A1", "asking about a place",              "Where do you live?"),
        ("when",       "adverb",    "A1", "asking about a time",               "When does the class start?"),
        ("how often",  "adverb",    "A1", "asking about frequency",            "How often do you exercise?"),
        ("what time",  "adverb",    "A1", "asking about a specific time",      "What time do you wake up?"),
        ("why",        "adverb",    "A1", "asking about a reason",             "Why do you study English?"),
        ("who",        "pronoun",   "A1", "asking about a person",             "Who do you live with?"),
        ("what",       "pronoun",   "A1", "asking about a thing or action",    "What does she eat for lunch?"),
        ("live",       "verb",      "A1", "to have your home somewhere",       "Where do you live?"),
        ("work",       "verb",      "A1", "to have a job",                     "Where does your father work?"),
        ("start",      "verb",      "A1", "to begin",                          "What time does school start?"),
        ("come",       "verb",      "A1", "to move toward the speaker",        "Where do you come from?"),
        ("mean",       "verb",      "A1", "to have a particular meaning",      "What does this word mean?"),
        ("cost",       "verb",      "A1", "to have a price of",                "How much does this cost?"),
        ("happen",     "verb",      "A1", "to take place; to occur",           "What happens next?"),
        ("think",      "verb",      "A1", "to use the mind; to believe",       "What do you think?"),
        ("meet",       "verb",      "A1", "to see someone for the first time", "Where do you meet friends?"),
        ("help",       "verb",      "A1", "to make something easier",          "Do you help at home?"),
        ("exercise",   "verb",      "A1", "to do physical activity",           "Do you exercise every day?"),
        ("relax",      "verb",      "A1", "to rest and feel calm",             "How do you relax?"),
        ("weekend",    "noun",      "A1", "Saturday and Sunday",               "What do you do at the weekend?"),
        ("hobby",      "noun",      "A1", "an activity you enjoy",             "Do you have a hobby?"),
        ("language",   "noun",      "A1", "a system of communication",         "How many languages do you speak?"),
        ("country",    "noun",      "A1", "a nation with its own government",  "What country does she come from?"),
        ("job",        "noun",      "A1", "work that you do for money",        "Does he like his job?"),
        ("time",       "noun",      "A1", "hours and minutes",                 "What time do you finish?"),
        ("answer",     "noun",      "A1", "a response to a question",          "Do you know the answer?"),
        ("often",      "adverb",    "A1", "many times; frequently",            "Do you often go out?"),
        ("usually",    "adverb",    "A1", "most of the time",                  "Do you usually walk to work?"),
        ("about",      "preposition","A1","concerning; regarding",             "What do you think about it?"),
        ("together",   "adverb",    "A1", "with each other",                   "Do you eat together?"),
    ],
    "grammar": {
        "rule": "To form a present simple question, use Do/Does + subject + base verb. For Wh- questions, put the question word first.",
        "form_headers": ["Question Type", "Structure", "Example"],
        "form_rows": [
            ["Yes/No (I/you/we/they)", "Do + subject + base verb?",   "Do you like pizza?"],
            ["Yes/No (he/she/it)",     "Does + subject + base verb?", "Does she work here?"],
            ["Short answer (Yes)",     "Yes, + subject + do/does.",    "Yes, I do. / Yes, she does."],
            ["Short answer (No)",      "No, + subject + don't/doesn't.", "No, I don't. / No, she doesn't."],
            ["Wh- question",           "Wh- + do/does + subject + base verb?", "Where does he live?"],
        ],
        "form_col_widths": [2400, 2800, 3000],
        "note": "After does, the verb is ALWAYS in the base form. 'Does she likes?' is WRONG. 'Does she like?' is correct.",
        "usage": [
            "Yes/No: Do you speak English? - Yes, I do.",
            "Where: Where does your sister live? - She lives in Rome.",
            "What time: What time do you get up? - At seven.",
            "How often: How often does he exercise? - Twice a week.",
            "Why: Why do they study English? - Because they need it for work.",
        ],
    },
    "guided": {
        "instruction": "Write the correct question for each answer. Items 1-2 are done for you.",
        "headers": ["#", "Answer given", "Question"],
        "items": [
            ["1 ✅", "Yes, I do. (like football?)",       "Do you like football?"],
            ["2 ✅", "She lives in Madrid. (where?)",     "Where does she live?"],
            ["3",   "No, he doesn't. (speak French?)",   "________________________"],
            ["4",   "At 8 o'clock. (what time / start?)", "________________________"],
            ["5",   "Yes, they do. (work on Saturdays?)", "________________________"],
            ["6",   "In a hospital. (where / he work?)", "________________________"],
            ["7",   "Three languages. (how many?)",      "________________________"],
            ["8",   "Because I enjoy it. (why / study?)", "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Put the words in order to make a present simple question.",
        "headers": ["#", "Words", "Write your answer"],
        "items": [
            ["1", "you / do / where / live",           "________________________"],
            ["2", "does / what / she / eat / lunch / for", "________________________"],
            ["3", "start / does / what time / school", "________________________"],
            ["4", "do / you / how often / exercise",   "________________________"],
            ["5", "he / does / like / his job",        "________________________"],
            ["6", "do / what / they / at the weekend / do", "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "Does you like pizza?",              "________________________"],
            ["2", "Where does she lives?",             "________________________"],
            ["3", "Do he work here?",                  "________________________"],
            ["4", "What time starts the class?",       "________________________"],
            ["5", "How often you go to the gym?",      "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "She works in a bank. (make a question)",            "________________________"],
            ["2", "They play tennis. (make a Yes/No question)",        "________________________"],
            ["3", "He lives in Paris. (make a Where question)",        "________________________"],
            ["4", "I get up at seven. (make a What time question)",    "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Find Someone Who...",
        "instructions": [
            "Walk around the class. Ask questions to find someone who...",
            "...likes coffee, plays a sport, speaks three languages, gets up before 6.",
            "Use: 'Do you like coffee?' When someone says 'Yes, I do,' write their name.",
            "Report back: 'Maria likes coffee. Tom plays football.'",
        ],
    },
    "reading": {
        "text": (
            "Anna is a journalist. She works for a newspaper in Berlin. She starts work at 9 a.m. "
            "and finishes at 6 p.m. She speaks three languages: German, English, and Spanish. "
            "She doesn't work on weekends. On Saturdays, she plays tennis with her friends. "
            "On Sundays, she relaxes at home and reads books. She loves her job because she meets "
            "interesting people every day."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Anna works in a hospital.",            "___"],
            ["2", "She starts work at 9 a.m.",           "___"],
            ["3", "She speaks two languages.",            "___"],
            ["4", "She works on weekends.",               "___"],
            ["5", "She plays tennis on Saturdays.",       "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 questions you would ask a new friend about their life.",
        "prompts": [
            "Ask about where they live: ___?",
            "Ask about their job: ___?",
            "Ask about what they do at the weekend: ___?",
            "Ask about their hobbies: ___?",
            "Ask about languages they speak: ___?",
            "Ask about what time they get up: ___?",
        ],
    },
    "homework": {
        "item1": (
            "Interview a family member. Write 6 questions using Do/Does and write their answers. "
            "Include at least 2 Wh- questions (Where/What/When/How often)."
        ),
        "item2": (
            "Write the question word (What, Where, When, How often, Why) that matches each situation: "
            "asking about a place, asking about time, asking about frequency, asking about a reason, "
            "asking about a thing."
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
    "sub_focus": "when to use each tense; signal words (always, now, every day, at the moment)",
    "objectives": [
        "Distinguish between present simple (habits) and present continuous (now).",
        "Identify signal words for each tense (always, usually vs now, at the moment).",
        "Choose the correct tense based on context in gap-fill exercises.",
        "Produce sentences using both tenses correctly in the same paragraph.",
        "Explain why a particular tense is used in a given sentence.",
    ],
    "warmup": {
        "name": "📸 Two Pictures",
        "time": "5 minutes",
        "steps": [
            "Look at Picture A (a daily routine scene) and Picture B (a 'right now' action).",
            "Describe Picture A: 'She usually drinks coffee in the morning.'",
            "Describe Picture B: 'Right now, she is drinking tea.'",
            "Discuss: What is the difference between the two sentences?",
        ],
    },
    "lead_in": [
        "What do you usually do in the evening? What are you doing right now?",
        "Is there a difference between 'I work' and 'I am working'? When do we use each?",
    ],
    "vocab": [
        ("always",       "adverb",    "A1", "every time; all the time",          "I always eat breakfast."),
        ("usually",      "adverb",    "A1", "most of the time",                  "She usually walks to work."),
        ("often",        "adverb",    "A1", "many times; frequently",            "We often go to the park."),
        ("sometimes",    "adverb",    "A1", "on some occasions; not always",     "He sometimes reads at night."),
        ("never",        "adverb",    "A1", "not at any time",                   "They never eat fast food."),
        ("every day",    "adverb",    "A1", "each day; daily",                   "She exercises every day."),
        ("now",          "adverb",    "A1", "at this moment",                    "I am reading now."),
        ("right now",    "adverb",    "A1", "at this exact moment",              "She is sleeping right now."),
        ("at the moment","adverb",    "A1", "at this time; now",                 "He is working at the moment."),
        ("today",        "adverb",    "A1", "on this day",                       "I am staying home today."),
        ("this week",    "adverb",    "A1", "during the current week",           "We are studying hard this week."),
        ("routine",      "noun",      "A1", "a regular way of doing things",     "My morning routine is simple."),
        ("habit",        "noun",      "A1", "something you do regularly",        "Reading is a good habit."),
        ("action",       "noun",      "A1", "the process of doing something",    "This action is happening now."),
        ("moment",       "noun",      "A1", "a very short period of time",       "Wait a moment, please."),
        ("difference",   "noun",      "A1", "the way things are not the same",   "What is the difference?"),
        ("still",        "adverb",    "A1", "continuing until now",              "He is still working."),
        ("temporary",    "adjective", "A2", "lasting for a short time only",     "This is temporary."),
        ("permanent",    "adjective", "A2", "lasting forever or a long time",    "It is a permanent job."),
        ("general",      "adjective", "A1", "relating to all; not specific",     "It is a general fact."),
        ("specific",     "adjective", "A2", "particular; not general",           "This is a specific situation."),
        ("look",         "verb",      "A1", "to use your eyes to see",           "Look! She is dancing."),
        ("listen",       "verb",      "A1", "to pay attention to sounds",        "Listen! The baby is crying."),
        ("rain",         "verb",      "A1", "water falls from the sky",          "It is raining now."),
        ("wear",         "verb",      "A1", "to have clothes on your body",      "She is wearing a red dress."),
        ("sit",          "verb",      "A1", "to be in a seated position",        "He is sitting on the sofa."),
        ("stand",        "verb",      "A1", "to be on your feet",                "They are standing outside."),
        ("happen",       "verb",      "A1", "to take place",                     "What is happening?"),
        ("change",       "verb",      "A1", "to become different",               "Things are changing fast."),
        ("learn",        "verb",      "A1", "to get knowledge or skill",         "She is learning to drive."),
    ],
    "grammar": {
        "rule": "Use Present Simple for habits/routines/facts. Use Present Continuous for actions happening NOW or temporary situations.",
        "form_headers": ["Tense", "Use", "Signal words", "Example"],
        "form_rows": [
            ["Present Simple",      "habits, routines, facts",    "always, usually, often, every day, never", "She drinks coffee every morning."],
            ["Present Continuous",   "actions happening now",      "now, right now, at the moment, today",      "She is drinking tea right now."],
            ["Present Simple",      "permanent situations",       "---",                                        "He works in a bank."],
            ["Present Continuous",   "temporary situations",       "this week, today, currently",               "He is working from home this week."],
        ],
        "form_col_widths": [1800, 2000, 2200, 2200],
        "note": "If you see NOW, RIGHT NOW, LOOK!, LISTEN! -> use Present Continuous. If you see ALWAYS, USUALLY, EVERY DAY -> use Present Simple.",
        "usage": [
            "Habit vs Now: I drink tea every morning. (habit) / I am drinking tea now. (this moment)",
            "Permanent vs Temporary: She lives in Rome. (permanent) / She is staying in a hotel this week. (temporary)",
            "Fact: Water boils at 100 degrees. (always true - Present Simple)",
            "Right now: Look! It is raining. (happening at this moment - Present Continuous)",
            "Signal word: He usually drives, but today he is walking.",
        ],
    },
    "guided": {
        "instruction": "Choose Present Simple or Present Continuous. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "She ___ (drink) coffee every morning.",    "She DRINKS coffee every morning. (habit)"],
            ["2 ✅", "Look! It ___ (rain).",                     "Look! It IS RAINING. (now)"],
            ["3",   "He usually ___ (walk) to work.",           "________________________"],
            ["4",   "Right now, she ___ (read) a book.",        "________________________"],
            ["5",   "We ___ (go) to the park every Sunday.",    "________________________"],
            ["6",   "Listen! The baby ___ (cry).",              "________________________"],
            ["7",   "I ___ (study) English this week.",         "________________________"],
            ["8",   "My father ___ (work) in a factory.",       "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write P.S. (Present Simple) or P.C. (Present Continuous) and explain why.",
        "headers": ["#", "Sentence", "Write your answer"],
        "items": [
            ["1", "She always gets up at 6 a.m.",          "________________________"],
            ["2", "They are watching TV at the moment.",   "________________________"],
            ["3", "He works in a hospital.",               "________________________"],
            ["4", "I am wearing a blue shirt today.",      "________________________"],
            ["5", "We often eat lunch together.",          "________________________"],
            ["6", "Look! The cat is sleeping on the sofa.","________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She is drinking coffee every morning.", "________________________"],
            ["2", "Look! He walks in the rain.",           "________________________"],
            ["3", "I am usually going to bed at 10.",      "________________________"],
            ["4", "They play football right now.",         "________________________"],
            ["5", "He is working in a bank. (permanent)",  "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I read a book. (add: right now)",                    "________________________"],
            ["2", "She is cooking dinner. (change to: every evening)",  "________________________"],
            ["3", "He drives to work. (add: today / make temporary)",   "________________________"],
            ["4", "We eat lunch at 1 p.m. (add: at the moment)",       "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 What Are They Doing?",
        "instructions": [
            "Look at the pictures on the board. Describe what the people are doing NOW.",
            "Then say what they USUALLY do (use your imagination).",
            "Example: 'Right now, he is eating pizza. He usually eats salad.'",
            "Work in pairs. Take turns describing and contrasting.",
        ],
    },
    "reading": {
        "text": (
            "Sarah is a teacher. She usually gets up at 6:30 and goes to school at 7:45. "
            "She teaches English to teenagers. She always eats lunch at the school cafeteria. "
            "But today is different. Today, Sarah is not at school. She is sitting in a cafe "
            "and reading a book. She is drinking hot chocolate. She is not working today because "
            "it is a holiday. She is enjoying her day off."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Sarah usually gets up at 7:30.",          "___"],
            ["2", "She teaches English.",                    "___"],
            ["3", "Today she is at school.",                 "___"],
            ["4", "She is drinking coffee right now.",       "___"],
            ["5", "It is a holiday today.",                  "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write a short paragraph (6 sentences) about your normal routine and what you are doing differently today.",
        "prompts": [
            "Write 3 sentences about what you USUALLY do (Present Simple).",
            "Write 3 sentences about what you are doing NOW/TODAY (Present Continuous).",
            "Use signal words: usually, always, every day, right now, at the moment, today.",
            "Example: 'I usually study at the library, but today I am studying at home.'",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences: 4 in Present Simple about your habits and 4 in Present Continuous "
            "about what is happening right now (or imagine a situation). Underline the signal words."
        ),
        "item2": (
            "Sort the vocabulary signal words into two groups: Present Simple words and "
            "Present Continuous words. Write an example sentence for each word."
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
    "sub_focus": "affirmative, negative, and questions with have got / has got; possessions and descriptions",
    "objectives": [
        "Form affirmative sentences using have got (I/you/we/they) and has got (he/she/it).",
        "Form negative sentences using haven't got and hasn't got.",
        "Form questions: Have you got...? Has she got...? and give short answers.",
        "Use have got to talk about possessions, family, and physical descriptions.",
        "Distinguish between have got (British) and have (American) usage.",
    ],
    "warmup": {
        "name": "🎒 What's in Your Bag?",
        "time": "5 minutes",
        "steps": [
            "Look in your bag or pockets. Write 3 things you have got.",
            "Tell your partner: 'I've got a phone, a pen, and a notebook.'",
            "Ask your partner: 'Have you got a dictionary? Have you got any keys?'",
        ],
    },
    "lead_in": [
        "What have you got in your classroom? Look around and describe.",
        "Have you got any brothers or sisters? Tell your partner about your family.",
    ],
    "vocab": [
        ("brother",    "noun",      "A1", "a male sibling",                    "I've got two brothers."),
        ("sister",     "noun",      "A1", "a female sibling",                  "She's got a little sister."),
        ("pet",        "noun",      "A1", "an animal kept at home",            "Have you got a pet?"),
        ("dog",        "noun",      "A1", "a common domestic animal",          "He's got a big dog."),
        ("cat",        "noun",      "A1", "a small furry pet",                 "We've got two cats."),
        ("car",        "noun",      "A1", "a vehicle with four wheels",        "They haven't got a car."),
        ("phone",      "noun",      "A1", "a device for communication",        "I've got a new phone."),
        ("computer",   "noun",      "A1", "an electronic machine",             "She hasn't got a computer."),
        ("garden",     "noun",      "A1", "land next to a house with plants",  "They've got a big garden."),
        ("bicycle",    "noun",      "A1", "a two-wheeled vehicle",             "He's got a red bicycle."),
        ("money",      "noun",      "A1", "coins and notes used to buy things","I haven't got any money."),
        ("idea",       "noun",      "A1", "a thought or plan",                 "I've got an idea!"),
        ("problem",    "noun",      "A1", "a difficulty or trouble",           "She's got a problem."),
        ("headache",   "noun",      "A1", "pain in the head",                  "He's got a headache."),
        ("cold",       "noun",      "A1", "a common illness",                  "I've got a cold."),
        ("ticket",     "noun",      "A1", "paper allowing entry or travel",    "Have you got your ticket?"),
        ("passport",   "noun",      "A1", "official travel document",          "She hasn't got her passport."),
        ("umbrella",   "noun",      "A1", "device for protection from rain",   "I haven't got an umbrella."),
        ("key",        "noun",      "A1", "a metal device for locks",          "Have you got your keys?"),
        ("notebook",   "noun",      "A1", "a small book for writing",          "He's got a new notebook."),
        ("camera",     "noun",      "A1", "a device for taking photos",        "They've got a good camera."),
        ("long",       "adjective", "A1", "not short in length",               "She's got long hair."),
        ("short",      "adjective", "A1", "not tall; not long",                "He's got short brown hair."),
        ("curly",      "adjective", "A1", "forming curves or spirals",         "She's got curly hair."),
        ("straight",   "adjective", "A1", "not curved or bent",                "He's got straight hair."),
        ("blue",       "adjective", "A1", "the colour of the sky",             "She's got blue eyes."),
        ("brown",      "adjective", "A1", "a dark warm colour",                "He's got brown eyes."),
        ("big",        "adjective", "A1", "large in size",                     "They've got a big house."),
        ("small",      "adjective", "A1", "little in size",                    "She's got a small bag."),
        ("any",        "determiner","A1", "some (in questions/negatives)",     "Have you got any brothers?"),
    ],
    "grammar": {
        "rule": "Use have got / has got to talk about possessions, family, physical features, and illnesses. Has got is for he/she/it.",
        "form_headers": ["Subject", "Affirmative", "Negative", "Question"],
        "form_rows": [
            ["I / You / We / They", "have got (I've got)",    "haven't got",  "Have you got...?"],
            ["He / She / It",       "has got (she's got)",    "hasn't got",   "Has she got...?"],
        ],
        "form_col_widths": [2000, 2200, 1800, 2200],
        "spelling_headers": ["Full form", "Contraction", "Example"],
        "spelling_rows": [
            ["I have got",     "I've got",      "I've got a cat."],
            ["She has got",    "She's got",     "She's got blue eyes."],
            ["have not got",   "haven't got",   "I haven't got a car."],
            ["has not got",    "hasn't got",    "He hasn't got a pet."],
        ],
        "spelling_col_widths": [2200, 2200, 3800],
        "note": "Short answers: Have you got a car? Yes, I have. / No, I haven't. (NOT: Yes, I've got. / No, I haven't got.)",
        "usage": [
            "Possessions: I've got a new phone. She hasn't got a computer.",
            "Family: He's got two sisters. They haven't got any children.",
            "Physical: She's got long brown hair. He's got blue eyes.",
            "Illness: I've got a headache. She's got a cold.",
            "Questions: Have you got any pets? Has he got a car?",
        ],
    },
    "guided": {
        "instruction": "Complete the sentences with have got, has got, haven't got, or hasn't got. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ a new phone. (+)",           "I HAVE GOT a new phone."],
            ["2 ✅", "She ___ a car. (-)",               "She HASN'T GOT a car."],
            ["3",   "They ___ two children. (+)",       "________________________"],
            ["4",   "He ___ any brothers. (-)",         "________________________"],
            ["5",   "___ you ___ a pet? (?)",           "________________________"],
            ["6",   "We ___ a big garden. (+)",         "________________________"],
            ["7",   "She ___ blue eyes. (+)",           "________________________"],
            ["8",   "I ___ any money. (-)",             "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write sentences using have got / has got about the pictures (described below).",
        "headers": ["#", "Description", "Write your answer"],
        "items": [
            ["1", "Tom / a red bicycle (+)",                 "________________________"],
            ["2", "Maria / long curly hair (+)",             "________________________"],
            ["3", "They / any pets (-)",                     "________________________"],
            ["4", "She / a headache (+)",                    "________________________"],
            ["5", "We / a big house (-)",                    "________________________"],
            ["6", "He / brown eyes and short hair (+)",      "________________________"],
            ["7", "I / my passport (-)",                     "________________________"],
            ["8", "You / any brothers or sisters (?)",       "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She have got a cat.",               "________________________"],
            ["2", "He's has got blue eyes.",           "________________________"],
            ["3", "Have he got a car?",                "________________________"],
            ["4", "I hasn't got any money.",           "________________________"],
            ["5", "They hasn't got a garden.",         "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "I have got a dog. (make negative)",          "________________________"],
            ["2", "She has got a new phone. (make a question)", "________________________"],
            ["3", "They have got a big garden. (use contraction)", "________________________"],
            ["4", "He has not got a car. (use contraction)",    "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Describe Your Partner",
        "instructions": [
            "Work in pairs. Look at your partner and describe them using has got.",
            "Talk about: hair (colour, length, style), eyes (colour), and accessories.",
            "Example: 'Maria has got long brown hair. She's got brown eyes.'",
            "Ask questions: 'Have you got any brothers? Have you got a pet?'",
        ],
    },
    "reading": {
        "text": (
            "My name is Sophie. I've got a big family. I've got two brothers and one sister. "
            "My brother Jack has got red hair and green eyes. My sister Emma has got long blonde hair. "
            "We've got a dog called Max and two cats. We haven't got a garden, but we've got a "
            "balcony. My parents have got an old car. It hasn't got air conditioning, but we love it."
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "Sophie has got one brother.",             "___"],
            ["2", "Jack has got red hair.",                  "___"],
            ["3", "Emma has got short hair.",                "___"],
            ["4", "They haven't got any pets.",              "___"],
            ["5", "The family car hasn't got air conditioning.", "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences describing yourself and your family using have got / has got.",
        "prompts": [
            "Describe your hair: I've got ___ hair.",
            "Describe your eyes: I've got ___ eyes.",
            "Write about siblings: I've got ___ / I haven't got ___.",
            "Write about a pet: I've got / I haven't got ___.",
            "Describe a family member: My ___ has got ___.",
            "Write about possessions: I've got ___ but I haven't got ___.",
        ],
    },
    "homework": {
        "item1": (
            "Draw or find a picture of a person. Write 8 sentences describing them using "
            "has got and hasn't got. Include hair, eyes, accessories, and possessions."
        ),
        "item2": (
            "Write a list of 5 things you have got and 5 things you haven't got. "
            "Example: I've got a bike. I haven't got a car."
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
    "sub_focus": "was/were: affirmative, negative, and questions; talking about the past",
    "objectives": [
        "Form affirmative past simple sentences with was (I/he/she/it) and were (you/we/they).",
        "Form negative sentences using wasn't and weren't.",
        "Form Yes/No questions: Was he...? Were they...? and give short answers.",
        "Form Wh- questions with was/were (Where were you? How was it?).",
        "Use was/were to describe past situations, locations, and states.",
    ],
    "warmup": {
        "name": "🕰️ Yesterday's Story",
        "time": "5 minutes",
        "steps": [
            "Think about yesterday. Where were you at 3 p.m.?",
            "Tell your partner: 'I was at school / at home / in the park.'",
            "Ask your partner: 'Were you tired yesterday? Was it a good day?'",
        ],
    },
    "lead_in": [
        "Where were you last weekend? Were you at home or outside?",
        "Think about your last birthday. Was it fun? Who was there?",
    ],
    "vocab": [
        ("yesterday",  "adverb",    "A1", "the day before today",              "I was tired yesterday."),
        ("last week",  "adverb",    "A1", "the week before this one",          "She was ill last week."),
        ("last night", "adverb",    "A1", "the previous evening/night",        "Were you at home last night?"),
        ("ago",        "adverb",    "A1", "in the past (two days ago)",        "He was here an hour ago."),
        ("born",       "adjective", "A1", "brought into life",                 "I was born in 1999."),
        ("late",       "adjective", "A1", "not on time",                       "We were late for class."),
        ("early",      "adjective", "A1", "before the expected time",          "She was early today."),
        ("absent",     "adjective", "A1", "not present; away",                 "Tom was absent yesterday."),
        ("present",    "adjective", "A1", "in a place; not absent",            "All students were present."),
        ("famous",     "adjective", "A1", "known by many people",              "He was a famous actor."),
        ("ill",        "adjective", "A1", "not in good health; sick",          "She was ill last week."),
        ("angry",      "adjective", "A1", "feeling strong displeasure",        "My father was angry."),
        ("worried",    "adjective", "A1", "feeling anxious or troubled",       "We were worried about you."),
        ("surprised",  "adjective", "A1", "feeling unexpected wonder",         "I was surprised."),
        ("scared",     "adjective", "A1", "feeling afraid",                    "The children were scared."),
        ("excited",    "adjective", "A1", "feeling very happy and eager",      "She was excited about the trip."),
        ("bored",      "adjective", "A1", "feeling uninterested",              "They were bored at home."),
        ("crowded",    "adjective", "A1", "full of people",                    "The bus was crowded."),
        ("empty",      "adjective", "A1", "having nothing inside",             "The room was empty."),
        ("closed",     "adjective", "A1", "not open",                          "The shop was closed."),
        ("open",       "adjective", "A1", "not shut",                          "Was the library open?"),
        ("difficult",  "adjective", "A1", "hard to do or understand",          "The test was difficult."),
        ("easy",       "adjective", "A1", "not hard; simple",                  "The homework was easy."),
        ("sunny",      "adjective", "A1", "with a lot of sunshine",            "It was sunny yesterday."),
        ("rainy",      "adjective", "A1", "with a lot of rain",               "It was rainy last week."),
        ("birthday",   "noun",      "A1", "the anniversary of birth",          "My birthday was last month."),
        ("party",      "noun",      "A1", "a social gathering",                "The party was great."),
        ("meeting",    "noun",      "A1", "a gathering of people to discuss",  "The meeting was at 3 p.m."),
        ("holiday",    "noun",      "A1", "a day of rest or celebration",      "It was a national holiday."),
        ("childhood",  "noun",      "A2", "the time when you are a child",     "My childhood was happy."),
    ],
    "grammar": {
        "rule": "Use was (I/he/she/it) and were (you/we/they) to talk about past states, descriptions, and locations.",
        "form_headers": ["Subject", "Affirmative", "Negative", "Question"],
        "form_rows": [
            ["I / He / She / It",   "was",   "was not (wasn't)",   "Was I/he/she/it...?"],
            ["You / We / They",     "were",  "were not (weren't)", "Were you/we/they...?"],
        ],
        "form_col_widths": [2000, 1800, 2200, 2200],
        "spelling_headers": ["Form", "Contraction", "Example"],
        "spelling_rows": [
            ["was not",    "wasn't",    "It wasn't cold yesterday."],
            ["were not",   "weren't",   "They weren't at school."],
        ],
        "spelling_col_widths": [2200, 2200, 3800],
        "note": "Short answers: Was she tired? Yes, she was. / No, she wasn't. (NOT: Yes, she was tired. for a short answer)",
        "usage": [
            "Past state: I was tired yesterday. She wasn't happy.",
            "Past location: We were at the park. He was at work.",
            "Time: The meeting was at 3 p.m. It was Monday.",
            "Description: The weather was sunny. The film was boring.",
            "Questions: Were you at home last night? Where was she born?",
        ],
    },
    "guided": {
        "instruction": "Complete the sentences with was, were, wasn't, or weren't. Items 1-2 are done for you.",
        "headers": ["#", "Prompt", "Answer"],
        "items": [
            ["1 ✅", "I ___ at home yesterday. (+)",          "I WAS at home yesterday."],
            ["2 ✅", "They ___ at school. (-)",               "They WEREN'T at school."],
            ["3",   "She ___ ill last week. (+)",            "________________________"],
            ["4",   "We ___ late for class. (-)",            "________________________"],
            ["5",   "___ you at the party? (?)",             "________________________"],
            ["6",   "The shop ___ closed yesterday. (+)",    "________________________"],
            ["7",   "The children ___ scared. (-)",          "________________________"],
            ["8",   "It ___ sunny last weekend. (+)",        "________________________"],
        ],
        "col_widths": [500, 3800, 3900],
    },
    "controlled": {
        "instruction": "Write sentences about the past using was/were based on the prompts.",
        "headers": ["#", "Prompt", "Write your answer"],
        "items": [
            ["1", "you / at home / last night? (?)",         "________________________"],
            ["2", "the film / boring (-)",                    "________________________"],
            ["3", "my parents / at the meeting (+)",          "________________________"],
            ["4", "I / born / in 2005 (+)",                   "________________________"],
            ["5", "the weather / cold / yesterday? (?)",      "________________________"],
            ["6", "she / absent / from class (+)",            "________________________"],
            ["7", "the shops / open / on Sunday? (?)",        "________________________"],
            ["8", "we / excited / about the holiday (+)",     "________________________"],
        ],
        "col_widths": [500, 3400, 4300],
    },
    "semi_controlled": {
        "error_headers": ["#", "Incorrect Sentence", "Corrected Sentence"],
        "error_items": [
            ["1", "She were at home yesterday.",        "________________________"],
            ["2", "They was late for school.",          "________________________"],
            ["3", "I were born in London.",             "________________________"],
            ["4", "Was you at the party?",             "________________________"],
            ["5", "The weather were nice last week.",   "________________________"],
        ],
        "error_col_widths": [500, 3500, 4200],
        "transform_headers": ["#", "Original", "Rewrite as instructed"],
        "transform_items": [
            ["1", "She was at school. (make negative)",          "________________________"],
            ["2", "They were happy. (make a question)",          "________________________"],
            ["3", "I was not tired. (use contraction)",          "________________________"],
            ["4", "The film was good. (make a Wh- question: How)", "________________________"],
        ],
        "transform_col_widths": [500, 3400, 4300],
    },
    "speaking": {
        "name": "💬 Last Weekend Interview",
        "instructions": [
            "Work in pairs. Ask your partner about last weekend.",
            "Use questions with was/were: 'Where were you? Was it fun? Were you tired?'",
            "Take notes and report to the class.",
            "Example: 'Maria was at the cinema. The film was exciting. She wasn't bored.'",
        ],
    },
    "reading": {
        "text": (
            "Last Saturday was my birthday. I was very excited. The weather was sunny and warm. "
            "My friends were all at my house for a party. The food was delicious and the music was "
            "great. My best friend Tom was there too. He was late, but it was OK. We were all very "
            "happy. The party wasn't boring at all. It was the best birthday ever!"
        ),
        "tf_headers": ["#", "Statement", "T / F"],
        "tf_items": [
            ["1", "The birthday was on Sunday.",             "___"],
            ["2", "The weather was sunny.",                  "___"],
            ["3", "Tom was early to the party.",             "___"],
            ["4", "The food was delicious.",                 "___"],
            ["5", "The party was boring.",                   "___"],
        ],
        "tf_col_widths": [500, 6500, 1200],
    },
    "writing": {
        "instruction": "Write 6 sentences about a day in the past using was/were.",
        "prompts": [
            "Where were you last Saturday? I was at ___.",
            "How was the weather? It was ___.",
            "Who was with you? ___ was/were with me.",
            "How was the day? It was ___.",
            "Was it fun or boring? It was/wasn't ___.",
            "What time was it when you got home? It was ___ o'clock.",
        ],
    },
    "homework": {
        "item1": (
            "Write 8 sentences about your last holiday or a special day. "
            "Use was/were in affirmative and negative forms. Include at least "
            "2 questions with was/were and their short answers."
        ),
        "item2": (
            "Write 10 sentences about famous people from history. "
            "Example: Einstein was a scientist. He wasn't a musician. Was he German? Yes, he was."
        ),
    },
})


# ─────────────────────────────────────────────────────────────────────────────
# Main execution — generate DOCX + PPTX for each unit
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
        fname_docx = f"Level1_Unit{unit_num}_{slug}.docx"
        path_docx = os.path.join(OUT_DIR, fname_docx)
        data_docx = build_docx(unit)
        with open(path_docx, "wb") as f:
            f.write(data_docx)
        print(f"  Generated: {fname_docx}  ({len(data_docx):,} bytes)")

        # Generate PPTX
        fname_pptx = f"Level1_Unit{unit_num}_{slug}.pptx"
        path_pptx = os.path.join(OUT_DIR, fname_pptx)
        data_pptx = build_pptx(unit)
        with open(path_pptx, "wb") as f:
            f.write(data_pptx)
        print(f"  Generated: {fname_pptx}  ({len(data_pptx):,} bytes)")

    print(f"\nAll {len(UNITS_5_10)} units generated successfully (12 files total).")

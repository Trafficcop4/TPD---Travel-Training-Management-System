"""Seed data for Lists, Agencies, Instructors, ExamMaster, SkillsMaster,
holidays, PT events, and the course-critique (evaluation) form.

Sources: current V5.4 workbook (agencies, exam/skills masters, dropdowns),
BPOC 7 hourly-calendar workbook (holidays, instructor roster), Academy policy
manual May 2026 (PT standards), existing "Class Evaluation and Assessment"
critique form (eval questions, lightly modernized).
"""

# ---------------------------------------------------------------- dropdowns
LISTS = {
    "Yes/No": ["Yes", "No"],
    "Cadet Status": ["Active", "Separated", "Graduated", "On Leave"],
    # "Spelling" was offered here but NOTHING consumes an exam typed
    # Spelling: every category count, average and weight on sysGrades
    # keys on Major/Minor/Final only, and the 10% spelling category
    # comes from the Spelling SHEET (sysGrades G/K read Spelling!Q/P).
    # Picking it printed and emailed a score that moved no grade.
    "Exam Type": ["Major", "Minor", "Final"],
    "Attendance Event": ["Tardy", "Early Departure", "Late Arrival",
                         "Removed From Class", "Absent Full Day",
                         "Absent Partial Day", "PT Missed", "PT Modified",
                         "PT Refused"],
    "Reason": ["Agency Recall", "LE Work", "Personal", "Medical", "Military",
               "Disciplinary", "Other"],
    "Documentation": ["Not Required", "Pending", "Received"],
    "Excused?": ["Counted", "Excused"],
    # Only the two types the attendance caps actually credit. The list used
    # to offer "Skills" and "Admin Approved" as well, and the Makeup Row
    # Check then rejected both unconditionally as TYPE NOT CREDITED - the
    # dropdown invited two answers the sheet immediately painted red.
    # Classroom skills time cannot be made up and there is no administrative
    # waiver, so those two answers were never valid to begin with.
    "Makeup Type": ["Classroom", "PT"],
    "Scoring Mode": ["Score", "Pass/Fail"],
    "Skill Result": ["Pass", "Fail", "Pending"],
    "Incident Direction": ["Positive", "Negative", "Informational"],
    "Severity": ["Informational", "Minor", "Moderate", "Major", "Critical"],
    "Resolution": ["Open", "In Review", "Resolved", "Closed"],
    "Assignment Type": ["Report", "Memo", "Essay", "Affidavit", "Other"],
    "Submission": ["Not Submitted", "Submitted", "Late", "Resubmitted"],
    # Skills genuinely allows up to 5 (SkillsMaster "Max Attempts":
    # Driving 5, Firearms/MOA/SFST/Medical 3), so this list stays 1-5
    # for Skills only.
    "Attempt #": ["1", "2", "3", "4", "5"],
    # ...but policy 300.5 allows ONE retest per exam, and every exam
    # rule in the workbook is written for attempts 1 and 2. A third
    # attempt used to be an ordinary dropdown pick that silently
    # replaced a passed retest with the failed first-attempt score.
    "Exam Attempt #": ["1", "2"],
    "IssueType": ["Absence", "Failure to Train", "Deficiency", "Conduct",
                  "Injury", "Equipment", "Other"],
    # -- new in V6 --
    "Counseling Type": ["Verbal Counseling", "Written Counseling", "Tutoring",
                        "Remedial Training", "Agency Notification",
                        "Performance Plan", "Commendation", "Other"],
    "Medical Status": ["Cleared - Full Duty", "Restricted", "Pending Eval",
                       "Expired"],
    "PT Session": ["Baseline", "Midpoint", "Final"],
    "Award": ["Valedictorian", "Physical Fitness", "Top Gun", "Grit"],
    "Instructor Type": ["TCOLE Instructor", "SME (Letter on File)",
                        "SME (Letter Needed)", "Guest/Outside"],
    "Materials Status": ["On File", "Requested", "Missing", "N/A"],
    "Requirement Source": ["TCOLE Rule", "Academy Policy"],
    "Review Type": ["Dismissal Review", "Academic Review", "Conduct Review"],
    "Review Outcome": ["Pending", "Retained", "Separated", "Retained w/ Plan"],
    # which engine trigger this review closes. The engine matches these
    # strings EXACTLY (sysChecks DismissReview), so they must not be
    # reworded without updating sheets_engine.build_syschecks.
    "Dismissal Trigger": ["Failed retest (exam)", "Skills failed out",
                          "Chain-of-command incident review",
                          "Unexcused missed exams",
                          "Other (not engine-tracked)"],
}

# ------------------------------------------------------------------ holidays
# (name, formula template over year cell {Y}) — observed-date logic from the
# coordinator's calendar workbook (Sat->Fri, Sun->Mon shifts where marked).
# observed Christmas Day: the federal weekend-shift rule (Sat -> Fri,
# Sun -> Mon). Referenced twice below, so it lives in one place.
_XMAS_OBS = ("IF(WEEKDAY(DATE({Y},12,25),2)=6,DATE({Y},12,25)-1,"
             "IF(WEEKDAY(DATE({Y},12,25),2)=7,DATE({Y},12,25)+1,"
             "DATE({Y},12,25)))")

HOLIDAYS = [
    ("New Year's Day",
     "IF(WEEKDAY(DATE({Y},1,1),2)=6,DATE({Y},1,1)-1,"
     "IF(WEEKDAY(DATE({Y},1,1),2)=7,DATE({Y},1,1)+1,DATE({Y},1,1)))"),
    ("Martin Luther King Jr. Day",
     "DATE({Y},1,1)+MOD(8-WEEKDAY(DATE({Y},1,1),2),7)+14"),
    ("Good Friday",
     "LET(Y,{Y},A,MOD(Y,19),B,INT(Y/100),D,INT(B/4),E,MOD(B,4),"
     "F,INT((B+8)/25),G,INT((B-F+1)/3),H,MOD(19*A+B-D-G+15,30),"
     "I,INT(MOD(Y,100)/4),K,MOD(MOD(Y,100),4),"
     "L,MOD(32+2*E+2*I-H-K,7),M,INT((A+11*H+22*L)/451),"
     "N,INT((H+L-7*M+114)/31),O,MOD(H+L-7*M+114,31),"
     "DATE(Y,N,O+1)-2)"),
    ("Memorial Day",
     "DATE({Y},6,0)-MOD(WEEKDAY(DATE({Y},6,0),2)-1,7)"),
    ("Juneteenth",
     "IF(WEEKDAY(DATE({Y},6,19),2)=6,DATE({Y},6,19)-1,"
     "IF(WEEKDAY(DATE({Y},6,19),2)=7,DATE({Y},6,19)+1,DATE({Y},6,19)))"),
    ("Independence Day",
     "IF(WEEKDAY(DATE({Y},7,4),2)=6,DATE({Y},7,4)-1,"
     "IF(WEEKDAY(DATE({Y},7,4),2)=7,DATE({Y},7,4)+1,DATE({Y},7,4)))"),
    ("Labor Day",
     "DATE({Y},9,1)+MOD(8-WEEKDAY(DATE({Y},9,1),2),7)"),
    ("Thanksgiving",
     "DATE({Y},11,1)+MOD(11-WEEKDAY(DATE({Y},11,1),2),7)+21"),
    ("Black Friday",
     "DATE({Y},11,1)+MOD(11-WEEKDAY(DATE({Y},11,1),2),7)+22"),
    # Christmas Eve and Christmas Day may never resolve to the SAME observed
    # date. Applying the weekend-shift rule to each date independently
    # collided in two of the seven calendars: Dec 25 on a Saturday shifted
    # Christmas Day back onto Friday the 24th (already Christmas Eve), and
    # Dec 25 on a Monday shifted Christmas Eve (Sunday the 24th) forward onto
    # Monday the 25th. Either collision silently deletes one closure day from
    # nrAllClosures, which pulls every later class day - and every Day #,
    # retest deadline, memo due date and writing date computed from it - one
    # day earlier with no error anywhere.
    #
    # Christmas Day keeps the ordinary federal weekend-shift rule; Christmas
    # Eve is defined as the WEEKDAY BEFORE the observed Christmas Day, which
    # is always distinct and always a Mon-Fri class day:
    #   Fri 12/25 -> Eve Thu 12/24      Sat 12/25 -> Day Fri 12/24, Eve Thu 12/23
    #   Mon 12/25 -> Day Mon 12/25, Eve Fri 12/22
    #   Sun 12/25 -> Day Mon 12/26, Eve Fri 12/23
    ("Christmas Eve",
     "(" + _XMAS_OBS + ")-IF(WEEKDAY(" + _XMAS_OBS + ",2)=1,3,1)"),
    ("Christmas Day", _XMAS_OBS),
]

# ---------------------------------------------------------------- instructors
# Non-individual schedule entries (proctors, venues, outside orgs) — seeded
# as Guest/Outside so schedule references resolve without instructor-file
# requirements (guest speakers not teaching LOs need no bio per the IRG).
GUEST_ENTITIES = [
    "Training Staff", "Cadre", "TJC", "TFD", "Kilgore College",
    "The Mayfair Building",
]

# Roster seed (names only; PID/bio/certs entered by coordinator).
INSTRUCTORS = [
    "Tyler Pride", "Trey Edwards", "Judson Moore", "Keven Fite",
    "Alejandra Flowers", "Reggie Johnson", "Mike Malone", "Jessica Doughten",
    "Caleb Westbrook", "April Molina", "James Goodman", "Rebekah Hutson",
    "Ethan Johnson", "Garrett Martin", "Magnolia Custer", "Stephen Thomas",
    "Tommy Guerrero", "Israel Camarena", "Andy Erbaugh", "JH Burge",
    "Brandon Crawley", "Jim Holt", "Blake Lockhart", "Luke Shafer",
    "Luis Aparicio", "David Alexander", "James Reeves", "Jon Phillips",
    "Matthew Dahl", "Elliott Patterson", "Bianca Smedley", "James Freeman",
    "Kyle Liggitt", "Ryan Caldwell", "Justin Lambert", "Chuck Boyce",
    "Adam Riggle", "Adam Tarrant", "Spencer McGregor", "Josh Allen",
    "Gavin Kirkhart", "Jimmy Turner", "Ken Gardner", "Will Sinclair",
    "Eddie Zapata", "Jordan Hill", "Chris Mackey", "Braden Barns",
    "John Hebert", "DJ Schick", "Brandon Lott", "Ryan Tack",
    "Amanda Cook", "Teressa Dell", "Steve Black", "Chad Homer",
    "Mike Saxion", "Justin Utley", "Jason Burton", "Nathan Elliott",
    "Lyndsay Rogers", "Pedro Maya",
]

# ------------------------------------------------------------------ agencies
# (AgencyID, AgencyName, Contact, Email, Phone, Address, ActiveYN)
AGENCIES = [
    ("TPD", "Tyler PD", "Eddie Sheffield",
     "esheffield@tylertexas.com; dshafer@tylertexas.com; "
     "blockhart@tylertexas.com; jmoore@tylertexas.com",
     "903-531-1090", "711 W Ferguson St, Tyler TX", "Yes"),
    ("TFD", "Tyler Fire Marshal", "Joey Hooten", "jhooton@tylertexas.com",
     "", "227 N Spring Ave, Tyler TX", "No"),
    ("APD", "Athens PD", "B. Lee", "blee@athenstx.gov",
     "", "110 Wildcat Dr, Whitehouse TX", "Yes"),
    ("JPD", "Jacksonville PD", "Steven Markasky",
     "steven.markasky@jacksonvilletx.org", "", "", "Yes"),
    ("LPD", "Longview PD", "Shannon Purdon", "Jpurdon@longviewtexas.com",
     "", "", "Yes"),
    ("TJCPD", "TJC PD", "Michael Seale", "michael.seale@tjc.edu", "", "", "No"),
    ("WPD", "Whitehouse PD", "Brian Tomlin", "btomlin@whitehousetx.org",
     "", "", "Yes"),
    ("ArpPD", "Arp PD", "", "", "", "105 Cannery Row, Lindale TX", "No"),
]

# ---------------------------------------------------------------- exam master
# (code, name, type, passing, seq)
EXAMS = [
    ("E01", "Professionalism & Ethics", "Minor", 70, 1),
    ("E02", "Constitutional Law", "Major", 70, 2),
    ("E03", "Penal Code", "Major", 70, 3),
    ("E04", "Code of Criminal Procedure", "Major", 70, 4),
    ("E05", "Arrest, Search & Seizure", "Major", 70, 5),
    ("E06", "Health & Safety Code", "Minor", 70, 6),
    ("E07", "Family Violence", "Minor", 70, 7),
    ("E08", "Victims of Crime & Human Trafficking", "Minor", 70, 8),
    ("E09", "Transportation Code", "Minor", 70, 9),
    ("E10", "Crash, TIM & SFST", "Major", 70, 10),
    ("E11", "Communications", "Minor", 70, 11),
    ("E12", "Spanish & De-Escalation", "Minor", 70, 12),
    ("E13", "Force Options", "Major", 70, 13),
    ("E14", "Crisis Intervention", "Major", 70, 14),
    ("E15", "Criminal Investigations", "Minor", 70, 15),
    ("E16", "Professional Police Driving", "Major", 70, 16),
    ("E17", "Emergency Medical, Radio Comms, K9, Deaf & Hard of Hearing",
     "Minor", 70, 17),
    ("FIN", "Final Comprehensive Exam", "Final", 70, 20),
]

# --------------------------------------------------------------- skills master
# (category, max_attempts, mode, passing)
SKILLS = [
    ("Firearms", 3, "Score", 70),
    ("Driving", 5, "Pass/Fail", None),
    ("MOA", 3, "Pass/Fail", None),
    ("OC", 1, "Pass/Fail", None),
    ("SFST", 3, "Pass/Fail", None),
    ("Medical", 3, "Pass/Fail", None),
]

# -------------------------------------------------------- per-cadet certs
# IRG-mandatory per-student completions: (short name, linked chapter/source)
CERTS = [
    ("TIM", "Ch 22 - National TIM Responder (4 hr)"),
    ("SFST", "Ch 23 - NHTSA SFST course"),
    ("TCIC", "Ch 36 - DPS TCIC/TLETS #4800"),
    ("CPR/AED", "Ch 40 - accredited CPR/AED (AHA/ARC)"),
    ("ALERRT", "Ch 43 - ALERRT Level 1 (SB 1852)"),
    ("ICS 100", "Writing #29 - FEMA IS-100"),
    ("ICS 200", "Writing #29 - FEMA IS-200"),
    ("ICS 700", "Writing #29 - FEMA IS-700"),
]

# ------------------------------------------------------------------ PT events
# (event, baseline standard text, unit, higher_is_better)
# ---------------------------------------------- final PT assessment rubric
# Approved "PT Test Score Chart" (v1, 09/03/2026). FIVE events - bench press
# and vertical jump are BASELINE standards only and are not scored here.
# Each entry is (event, PT-sheet source column, higher_is_better, measure,
# five tier thresholds = the value a cadet must REACH for tiers 1..5).
#
# Tier points are 12 / 14 / 16 / 18 / 20, so 60 = every minimum met and
# 100 = tier 5 across the board. Passing is cfgPTFinalMinPoints (70).
#
# Thresholds are the tier's ENTRY value, not the printed ranges. That
# matters: the chart's 300 m tier 1 ends at 61.5 s while tier 2 begins at
# 61.40 s, leaving 61.41-61.49 s uncovered. Scoring on "the best tier whose
# threshold is met" closes that gap conservatively (61.45 s scores tier 1)
# instead of returning nothing.
#
# The 1.5-mile row is in SECONDS while the PT sheet takes decimal minutes -
# the scoring formula multiplies by 60. Whole seconds keep the thresholds
# exact; 12:35 in decimal minutes is 12.58333... and rounding that would
# fail a cadet who ran the qualifying time exactly.
PT_FINAL_BANDS = [
    ("Push-Ups", "O", True, "reps (no time limit)",
     [23, 33, 41, 51, 79]),
    ("Sit-Ups", "P", True, "reps in 1:00",
     [23, 33, 41, 51, 79]),
    ("Agility Run", "Q", False, "seconds (lower is better)",
     [20.2, 18.0, 15.8, 13.6, 12.0]),
    ("1.5 Mile Run", "R", False,
     "SECONDS (15:42=942, 14:08=848, 12:35=755, 11:02=662, 9:30=570); "
     "the PT sheet takes decimal minutes and the engine converts",
     [942, 848, 755, 662, 570]),
    ("300 M Sprint", "S", False, "seconds (lower is better)",
     [66.99, 61.40, 55.80, 50.20, 45.99]),
]
PT_TIER_POINTS = [12, 14, 16, 18, 20]

# The BASELINE standards in PT_EVENTS below and the FINAL chart above are two
# DIFFERENT tests and are SUPPOSED to disagree. Baseline is the entry standard
# from the policy manual; the chart is the exit standard, and the whole point
# of the academy is that a cadet improves, so every final minimum is stricter
# than its baseline. Do not "reconcile" them - they are not a contradiction.
#
# Note the push-up protocol genuinely differs too: baseline is 18 reps IN ONE
# MINUTE, the final is 23 reps with NO TIME LIMIT. Different tests, not a
# typo.
#
# Baseline values restated numerically, in the SAME units the final bands use
# (seconds for the 1.5 mile), purely so verify_build can assert that no final
# standard is ever EASIER than the baseline it is meant to exceed. Bench press
# and vertical jump are baseline-only and have no final equivalent.
PT_BASELINE_NUMERIC = {
    "Push-Ups":     (18, True),      # 18 reps in 1:00
    "Sit-Ups":      (16, True),      # 16 reps in 1:00
    "Agility Run":  (22.0, False),   # Illinois agility, under 22.0 s
    "1.5 Mile Run": (18 * 60 + 48, False),   # 18:48 = 1128 s
    "300 M Sprint": (76, False),     # 76 s
}

PT_EVENTS = [
    ("Bench Press", "62% of body weight (1 rep)", "lbs", True),
    ("Vertical Jump", "14.5 inches", "in", True),
    ("Push-ups (1 min)", "18 reps", "reps", True),
    ("Sit-ups (1 min)", "16 reps", "reps", True),
    ("Illinois Agility", "under 22.0 seconds", "sec", False),
    ("1.5 Mile Run", "18:48 or less", "mm:ss", False),
    ("300 Meter Run", "76 seconds or less", "sec", False),
]

# ------------------------------------------------------- course critique form
# Existing "Class Evaluation and Assessment" questions, lightly modernized.
# VERBATIM from the academy's approved "Class Evaluation and Assessment"
# form (Chapter 01 packet), verified against the PDF. Five of these had been
# silently copy-edited here - "my job needs" had become "your job needs",
# which changes what is being asked - and the coordinator's instruction is to
# match the approved form exactly. The source's own grammar in Q6 and Q8 is
# intentional and must NOT be "fixed": this is an audit artifact and it has
# to read the way the form the cadets sign reads.
EVAL_QUESTIONS = [
    "Were the goals and objectives for the course provided at the beginning "
    "of the class?",
    "Was the material presented in a logical order?",
    "Was the course related to my job needs?",
    "Were real life examples or situations used to relate the materials?",
    "Will this course improve your job performance?",
    "Were the teaching aids (demonstrations) handout materials of adequate "
    "quality?",
    "Was the instructor knowledgeable about the course material?",
    "Did the instructor present the information a clear and understandable "
    "way?",
    "Did the instructor give you an opportunity to interact with the class "
    "or ask questions?",
    "Did you get satisfactory answers to your questions about the materials?",
    "Would you recommend this course to others?",
]
# also verbatim from the approved form
EVAL_SCALE = ("CHECK APPROPRIATE BOX.  1 LOWEST TO 5 HIGHEST RATING FOR EACH "
              "OF THE BELOW LISTED QUESTIONS.")
EVAL_FOOTER = ("Please use the below listed space to elaborate on any of the "
               "above points or any personal comments you may have. Your "
               "responses will be used to shape and refine the quality and "
               "content of future schools.")
ACADEMY_ADDRESS = "711 W. Ferguson St., Tyler, TX 75702 — (903) 531-1018  Fax (903) 535-0102"

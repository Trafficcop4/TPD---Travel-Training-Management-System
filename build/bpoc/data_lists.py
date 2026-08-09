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
    "Exam Type": ["Major", "Minor", "Spelling", "Final"],
    "Attendance Event": ["Tardy", "Early Departure", "Late Arrival",
                         "Removed From Class", "Absent Full Day",
                         "Absent Partial Day", "PT Missed", "PT Modified",
                         "PT Refused"],
    "Reason": ["LE Work", "Personal", "Medical", "Military", "Disciplinary",
               "Other"],
    "Documentation": ["Not Required", "Pending", "Received"],
    "Excused?": ["Counted", "Excused"],
    "Makeup Type": ["Classroom", "PT", "Skills", "Admin Approved"],
    "Scoring Mode": ["Score", "Pass/Fail"],
    "Skill Result": ["Pass", "Fail", "Pending"],
    "Incident Direction": ["Positive", "Negative", "Informational"],
    "Severity": ["Informational", "Minor", "Moderate", "Major", "Critical"],
    "Resolution": ["Open", "In Review", "Resolved", "Closed"],
    "Assignment Type": ["Report", "Memo", "Essay", "Affidavit", "Other"],
    "Submission": ["Not Submitted", "Submitted", "Late", "Resubmitted"],
    "Attempt #": ["1", "2", "3", "4", "5"],
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
}

# ------------------------------------------------------------------ holidays
# (name, formula template over year cell {Y}) — observed-date logic from the
# coordinator's calendar workbook (Sat->Fri, Sun->Mon shifts where marked).
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
    ("Christmas Eve",
     "IF(WEEKDAY(DATE({Y},12,24),2)=6,DATE({Y},12,24)-1,"
     "IF(WEEKDAY(DATE({Y},12,24),2)=7,DATE({Y},12,24)+1,DATE({Y},12,24)))"),
    ("Christmas Day",
     "IF(WEEKDAY(DATE({Y},12,25),2)=6,DATE({Y},12,25)-1,"
     "IF(WEEKDAY(DATE({Y},12,25),2)=7,DATE({Y},12,25)+1,DATE({Y},12,25)))"),
]

# ---------------------------------------------------------------- instructors
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
    "Mike Saxion", "Justin Utley", "Jason Burton",
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
EVAL_QUESTIONS = [
    "Were the goals and objectives for the course provided at the beginning "
    "of the class?",
    "Was the material presented in a logical order?",
    "Was the course related to your job needs?",
    "Were real-life examples or situations used to relate the material?",
    "Will this course improve your job performance?",
    "Were the teaching aids, demonstrations, and handout materials of "
    "adequate quality?",
    "Was the instructor knowledgeable about the course material?",
    "Did the instructor present the information in a clear and "
    "understandable way?",
    "Did the instructor give you an opportunity to interact with the class "
    "or ask questions?",
    "Did you get satisfactory answers to your questions about the material?",
    "Would you recommend this course to others?",
]
EVAL_SCALE = "Check the appropriate box — 1 (lowest) to 5 (highest) — for each question."
EVAL_FOOTER = ("Please use the space below to elaborate on any of the above "
               "points or any personal comments you may have. Your responses "
               "will be used to shape and refine the quality and content of "
               "future schools.")
ACADEMY_ADDRESS = "711 W. Ferguson St., Tyler, TX 75702 — (903) 531-1018  Fax (903) 535-0102"

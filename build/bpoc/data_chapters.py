"""TCOLE BPOC #1000736 chapter list with module letters, TCOLE minimum hours,
and TPD delivered-hours targets, as taught at the Tyler Police Academy.

Source: coordinator's BPOC 7 hourly-calendar workbook (Control/Lists sheets).
MinHrs are TCOLE course-1000736 minimums; TPDHrs are Tyler's planned hours
(seed values — editable in ChapterMaster after build).
"""

# (module, chapter, name, tcole_min_hrs, tpd_hrs)
CHAPTERS = [
    ("A", "0",  "Administrative/Departmental Overview", 0, 0),
    ("B", "1",  "Professionalism and Ethics", 12, 12),
    ("B", "2",  "Professional Policing", 12, 12),
    ("B", "3",  "Fitness, Wellness, and Stress Management", 16, 16),
    ("B", "4",  "TCOLE Rules", 4, 4),
    ("B", "5",  "Multiculturalism and Human Relations", 8, 8),
    ("B", "6",  "Racial Profiling", 4, 4),
    ("C", "7",  "US, Texas Constitution and Rights", 10, 10),
    ("C", "8",  "Penal Code", 50, 60),
    ("C", "9",  "Code of Criminal Procedure", 12, 12),
    ("C", "10", "Arrest, Search, and Seizure", 40, 46),
    ("C", "11", "Asset Forfeiture", 4, 5),
    ("C", "12", "Identity Crimes", 4, 4),
    ("C", "13", "Consular Notification", 1, 1),
    ("C", "14", "Civil Process", 4, 4),
    ("D", "15", "Health and Safety Code and Controlled Substance Act", 12, 12),
    ("D", "16", "Alcoholic Beverage Code", 4, 5),
    ("E", "17", "Sexual Assault and Family Violence", 12, 12),
    ("E", "18", "Missing and Exploited Children", 8, 8),
    ("E", "19", "Child Safety Check Alert List", 1, 1),
    ("E", "20", "Victims of Crime", 10, 11),
    ("E", "21", "Human Trafficking", 4, 4),
    ("F", "22", "Traffic Code / Crash Investigation / TIM", 74, 74),
    ("F", "23", "Intoxicated Driver (SFST)", 24, 34),
    ("G", "24", "Written Communication", 16, 16),
    ("G", "25", "Verbal Communication/Public Interaction", 16, 16),
    ("G", "26", "Spanish", 16, 16),
    ("H", "27", "De-escalation Strategies", 8, 16),
    ("J", "28", "Force Options Theory", 28, 28),
    ("I", "29", "Crisis Intervention Training", 40, 40),
    ("I", "30", "Traumatic Brain Injury", 2, 2),
    ("J", "31", "Arrest and Control", 40, 48),
    ("K", "32", "Criminal Investigations", 40, 48),
    ("K", "33", "Juvenile Offenders", 10, 11),
    ("L", "34", "Professional Police Driving", 32, 51),
    ("M", "35", "Patrol Skills/Traffic Stops", 46, 62),
    ("M", "36", "Radio Communications /Amber-Silver Alert / TCIC-TLETS", 16, 16),
    ("M", "37", "Civilian Interaction Training", 2, 2),
    ("M", "38", "Interacting with Deaf and Hard of Hearing", 4, 4),
    ("M", "39", "Canine Encounters", 4, 4),
    ("N", "40", "Emergency Medical Assistance", 16, 16),
    ("O", "41", "Firearms", 48, 48),
    ("P", "42", "HazMat Awareness / ICS", 4, 4),
    ("Q", "43", "ALERRT Level 1", 16, 16),
]

REQUIRED_TCOLE_HOURS = 736

# Chapter-specific MANDATORY items per the 2025 IRG (course #1000736).
# Shown on ChapterMaster with a "Special Req Met?" audit column.
SPECIAL_REQS = {
    "2":  "Crime Stoppers Texas Course #22911 delivered here (not reported separately)",
    "22": "Students complete 4-hr National TIM Responder course; instructor holds TIM Instructor",
    "23": "Instructor: NHTSA SFST Instructor (1016) + TCOLE SFST Instructor Cert; students complete NHTSA SFST course",
    "29": "Lead instructor: 2+ yrs as Mental Health Officer or CIT member",
    "32": "Course 3232 Special Investigative Topics content (Occ. Code 1701.352)",
    "36": "DPS NCIC/TCIC Course #4800 completed, DPS-approved trainers (not reported separately)",
    "37": "Seven Step Violator Contact method must be taught",
    "39": "Practical role-play scenarios required for LOs 39.14/39.16/39.18/39.22",
    "40": "Full nationally accredited CPR/AED course (AHA Heartsaver or ARC) incl. Adult/Child/Infant",
    "41": "48-hr minimum; 70% minimum on BOTH courses of fire",
    "43": "ALERRT Level 1 (16 hrs, SB 1852) by certified ALERRT instructors (TtT 3315 + current cert)",
}

# TPD teaches some TCOLE chapters as separate sub-classes with their own
# hour targets. Scheduling uses THESE names; hours roll up to the parent
# chapter for TCOLE reconciliation. (name, parent chapter, TPD target hrs)
SUBTOPICS = [
    ("Traffic Code", "22", 50),
    ("Crash Investigation", "22", 12),
    ("TIM", "22", 12),
    ("Criminal Investigations - General Investigations", "32", 16),
    ("Criminal Investigations - Auto Theft", "32", 4),
    ("Criminal Investigations - Crime Scene", "32", 8),
    ("Criminal Investigations - Booking Operations", "32", 8),
    ("Criminal Investigations - Interview Techniques", "32", 4),
    ("Criminal Investigations - Case Management", "32", 4),
    ("Criminal Investigations - Media Relations", "32", 4),
]

# Chapters whose excess hours are reported under their OWN course number
# instead of the general Addendum to BPOC (#101), per the IRG.
SEPARATE_REPORT = {
    "31": "#2040 (Arrest and Control)",
    "34": "#2046 (Professional Police Driving)",
    "41": "#2055 (Firearms)",
}
ADDENDUM_COURSE = "#101 (Addendum to BPOC)"

# Program-level mandatory items (Audit sheet manual checklist).
PROGRAM_REQS = [
    ("Commission Rules distributed to all students (Rule 215.9)",
     "Copies of current TCOLE rules given to every student admitted"),
    ("Required rule review conducted (215.9)",
     "Reviewed 211.33, 221.1, 223.1, 223.3, 223.15, 223.17, 223.19 in class"),
    ("Approved TCLEDDS roster on file",
     "Required element of the training file"),
    ("Course reported at exactly 736 hours",
     "Excess hours reported under Addendum to BPOC (course #101), never inside the 736"),
    ("Non-BPOC classes reported separately within 30 days",
     "Any additional classes the academy teaches"),
    ("Facility standards met (in-person, scenario/skills adequate)",
     "Classroom supports role-play, defensive tactics, firearms demonstration"),
    ("Assessments document mastery of all objectives",
     "Learners are never passed on attendance alone"),
    ("State exam plan communicated (250 Q, 70%, 3 attempts max)",
     "Third failure requires enrolling in a new BPOC"),
    ("Five-year retention of training records",
     "Auditors sample past courses; files must be producible on request"),
    ("Prior academies' FINAL workbooks archived and locatable",
     "One file per academy; know where every finished class's file lives"),
    ("Digital evidence library organized per chapter",
     "LMS test PDFs, ExamSheet grade sheets, scanned spelling tests, "
     "sign-in scans, evals - one folder per chapter, ChapterPacket as cover"),
    ("Scanned documents established as legal originals",
     "Records management has confirmed scans serve as originals for "
     "training documents (incl. sign-ins w/ PID)"),
]

# Non-chapter schedule activities that appear on the hourly calendar and the
# daily sign-in sheets (kept as a dropdown list alongside chapters).
ACTIVITIES = [f"Test {i}" for i in range(1, 18)] + [
    "Review",
    "Digital Forensics",
    "Lunch",
    "PT - Base Level Evaluation",
    "PT",
    "PT - Ju Jitsu",
    "PT - Final",
    'PT - "The Murph"',
    "Final Test",
    "State Test",
    "Test Review",
    "Final Re-Test (If Needed)",
    "Graduation",
]

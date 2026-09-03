"""BPOC writing-assignment master: 40 assignments with prompts.

Source: "BPOC #7 — Writing Assignments (Updated)" PDF. Dates in the workbook
are COMPUTED from the schedule, not stored here:

  AssignedDate = first class day of (week when LinkedChapter is first taught
                 + assign_delay_weeks)
  DueDate      = last class day of (assigned week + due_offset_weeks), 1700

assign_delay_weeks/due_offset_weeks below reproduce the updated BPOC 7
timeline; per-assignment override columns in WritingMaster always win.
"""

# (num, title, linked_chapter, assign_delay_weeks, due_offset_weeks, prompt)
ASSIGNMENTS = [
    (1, "Code of Ethics", "1", 0, 0,
     "Discuss the ways the Code of Ethics affects the items posted on or to an "
     "Officer's social media sites."),
    (2, "Professional Policing", "2", 0, 0,
     "Discuss previous experiences with law enforcement, how the interaction "
     "correlates to the police models discussed, and how that interaction "
     "affected the quest for a law enforcement career."),
    (3, "Professional Policing", "4", 0, 0,
     "Discuss how TCOLE's role in policing supports the information discussed "
     "in Professional Policing."),
    (4, "Multiculturalism / Racial Profiling", "6", 0, 0,
     "You are the Chief/Sheriff. One of your officers/deputies is publicly "
     "accused of racially profiling during a traffic stop that ended in the "
     "death of the driver. How do you respond? (Press release? Video release? "
     "Who do you involve?)"),
    (5, "U.S. & Texas Constitution — Criminal Justice System", "7", 0, 0,
     "Identify all the courts within your jurisdiction, the Judges and the "
     "types of cases heard in the court. (local to highest court)"),
    (6, "Penal Code", "8", 0, 0,
     "Create a scenario that meets the lowest culpable mental state. Add a "
     "fact or two that supports the next culpable mental state. Continue this "
     "progression until all the culpable mental states are covered."),
    (7, "Courts & Offenses — A", "8", 1, 0,
     "Select a case from the Texas Court of Appeals of your jurisdiction. The "
     "decision must determine whether or not the evidence met the elements of "
     "the offense. Summarize the details of the offense, the issue on appeal, "
     "the position of the defendant and the state, the final decision and if "
     "there is a dissent, the argument of the dissent."),
    (8, "Courts & Offenses — B", "8", 1, 0,
     'Select a case from the Texas Court of Criminal Appeals that decides '
     'whether an offense is a "lesser included offense". Summarize the details '
     "of the offense, the issue on appeal, the position of the defendant and "
     "the state, the final decision and if there is a dissent, the argument "
     "of the dissent."),
    (9, "Arrest, Search, and Seizure — A", "10", 0, 0,
     "Select a recent (less than 2 years old) case from the U.S. 5th Circuit "
     "Court of Appeals that decides the legality/illegality of a search. "
     "Summarize the details of the offense, the issue on appeal, the position "
     "of the defendant and the state, the final decision and if there is a "
     "dissent, the argument of the dissent."),
    (10, "Arrest, Search, and Seizure — B", "10", 0, 0,
     "Select a recent case from the Texas Court of Appeals of your "
     "jurisdiction that decides the issue of possession. Summarize the "
     "details of the offense, the issue on appeal, the position of the "
     "defendant and the state, the final decision and if there is a dissent, "
     "the argument of the dissent."),
    (11, "Arrest, Search, and Seizure — C", "10", 1, 0,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of probable cause in an arrest/search "
     "or search warrant. Summarize the details of the offense, the issue on "
     "appeal, the position of the defendant and the state, the final decision "
     "and if there is a dissent, the argument of the dissent."),
    (12, "Arrest, Search, and Seizure — D", "10", 1, 0,
     "Select a case from the U.S. Supreme Court that decides the issue of "
     "detention/arrest. Summarize the details of the offense, the issue on "
     "appeal, the position of the defendant and the state, the final decision "
     "and if there is a dissent, the argument of the dissent."),
    (13, "Arrest, Search, and Seizure — E", "10", 2, 0,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of vehicle detention and searches. "
     "Summarize the details of the offense, the issue on appeal, the position "
     "of the defendant and the state, the final decision and if there is a "
     "dissent, the argument of the dissent."),
    (14, "Arrest, Search, and Seizure — F", "10", 2, 0,
     "Select a case from a Federal Court that decides the issue of "
     "culpability in an illegal transportation of aliens across the border. "
     "Summarize the details of the offense, the issue on appeal, the position "
     "of the defendant and the state, the final decision and if there is a "
     "dissent, the argument of the dissent."),
    (15, "Traffic / Crash — A", "22", 0, 0,
     "Interview drivers/witnesses in a vehicle collision. Write a narrative "
     "about the collision."),
    (16, "Traffic / Crash — B", "22", 0, 0,
     "Pick a traffic violation that you have observed. Write a detailed "
     "narrative about the violation."),
    (17, "Traffic / Crash — C", "22", 0, 0,
     "Prepare an arrest warrant affidavit (probable cause narrative) for a "
     "subject that you were unable to capture in a pursuit."),
    (18, "Traffic / Crash — D", "22", 0, 0,
     "Prepare a search warrant affidavit (probable cause narrative) for a "
     "vehicle involved in a failure to stop and render aid collision."),
    (19, "SFST", "23", 0, 1,
     "Prepare a case narrative, based on a wet lab test subject, for a DWI "
     "arrest. Complete all DPS forms (DIC-23, DIC-24, DIC-25, DIC-54, and "
     "DIC-55). NOTE: Assigned after wet lab. Wet lab must be completed before "
     "writing this narrative."),
    (20, "Written Communication — A", "24", 0, 1,
     'Submit an offense report narrative from the "Night Watch" painting.'),
    (21, "Written Communication — B", "24", 0, 1,
     "Prepare a full narrative on how to make a peanut butter and jelly "
     "sandwich."),
    (22, "Written Communication — C", "24", 0, 1,
     "Interview witnesses of a shooting (Clint Eastwood). Write a detailed "
     "narrative for a case report on the shooting."),
    (23, "Spanish", "26", 0, 1,
     "Complete a casual contact of a Spanish speaking individual. Document "
     "the contact on a Field Information form."),
    (24, "De-escalation", "27", 0, 1,
     "Review the article at https://www.calibrepress.com/2020/04/new-study-"
     "expert-vs-novice-use-of-force-decision-making/ — Discuss techniques in "
     "the article that correlate with the information taught during this "
     "portion of the academy."),
    (25, "Use of Force — A", "28", 0, 1,
     "Watch video. Submit a use of force / subject resistance report."),
    (26, "Use of Force — B", "28", 0, 1,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of excessive force. Summarize the "
     "details of the offense, the issue on appeal, the position of the "
     "defendant and the state, the final decision and if there is a dissent, "
     "the argument of the dissent."),
    (27, "Crisis Intervention", "29", 0, 1,
     "Complete a Peace Officer Emergency Detention form."),
    (28, "Arrest and Control — A", "31", 0, 1,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of resistance during an arrest. "
     "Summarize the details of the offense, the issue on appeal, the position "
     "of the defendant and the state, the final decision and if there is a "
     "dissent, the argument of the dissent."),
    # verbatim from the source PDF: the exact FEMA course VERSIONS matter
    # (IS-100.c / IS-200.c / IS-700.b) and a paraphrase to "training.fema.gov"
    # dropped them from the handout the cadets actually receive
    (29, "Arrest and Control — B", "31", 1, 1,
     "Complete the following FEMA online courses and submit your completion "
     "certificates: ICS 100 (https://training.fema.gov/is/courseoverview.aspx"
     "?code=IS-100.c), ICS 200 (https://training.fema.gov/is/courseoverview."
     "aspx?code=IS-200.c), ICS 700 (https://training.fema.gov/is/"
     "courseoverview.aspx?code=IS-700.b)"),
    (30, "Booking Procedures", "32", 0, 1,
     "Complete a booking form."),
    (31, "Interview", "32", 1, 2,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of Miranda warning (required/not "
     "required). Summarize the details of the offense, the issue on appeal, "
     "the position of the defendant and the state, the final decision and if "
     "there is a dissent, the argument of the dissent."),
    (32, "Crime Scene Investigation — A", "32", 1, 2,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of exigent circumstances for a search. "
     "Summarize the details of the offense, the issue on appeal, the position "
     "of the defendant and the state, the final decision and if there is a "
     "dissent, the argument of the dissent."),
    (33, "Crime Scene Investigation — B", "32", 2, 1,
     "Complete a crime scene sketch/inventory."),
    (34, "Professional Police Driving — A", "34", 0, 3,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of road blocks/pursuits and whether it "
     "was reasonable under the search and seizure statutes. Summarize the "
     "details of the offense, the issue on appeal, the position of the "
     "defendant and the state, the final decision and if there is a dissent, "
     "the argument of the dissent."),
    (35, "Professional Police Driving — B", "34", 1, 2,
     "Select an appeal case dealing with an officer's liability from actions "
     "that injured a third party during a pursuit or emergency vehicle "
     "operations. Summarize the details of the incident, the issue on appeal, "
     "the position of the defendant and the plaintiff, the final decision and "
     "if there is a dissent, the argument of the dissent."),
    (36, "Patrol Skills — A", "35", 0, 2,
     "Discuss the police style used by your agency, the patrol modes of your "
     "agency and patrol methods of your agency. Provide examples of how these "
     "items are used to provide effective police services."),
    (37, "Patrol Skills — B", "35", 1, 1,
     "Complete a residential/business security survey."),
    (38, "Radio Communications", "36", 0, 1,
     "Complete an Amber Alert request form or a Silver Alert request form."),
    (39, "Emergency Medical Assistance", "40", 0, 3,
     "Complete an emergency action plan that could be employed during the "
     "firearms training."),
    (40, "Firearms", "41", 0, 3,
     "Select a case from any court within your jurisdiction (Texas or "
     "Federal) that decides the issue of police officer liability when a "
     "firearm is used in an off-duty capacity. Summarize the details of the "
     "offense, the issue on appeal, the position of the defendant and the "
     "state, the final decision and if there is a dissent, the argument of "
     "the dissent."),
]

REQUIREMENTS_NOTE = (
    "Each assignment must be at least one (1) page, typed, single-spaced, "
    "12-point Arial or Times New Roman font. Due by 1700 on the due date. "
    "Late submissions require prior approval from the Training Coordinator. "
    "All case-law assignments require proper citation including court, year, "
    "and case name."
)

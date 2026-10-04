export const fields = [
  {
    "key": "FormKind",
    "label": "Report type",
    "section": 0,
    "block": "",
    "limit": 12,
    "date": false,
    "digits": null,
    "options": [
      "EVAL",
      "CHIEFEVAL",
      "FITREP"
    ],
    "multiline": false
  },
  {
    "key": "Title",
    "label": "Local draft title",
    "section": 7,
    "block": "",
    "limit": 160,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "MemberName",
    "label": "Local display name",
    "section": 7,
    "block": "",
    "limit": 120,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "PeriodStart",
    "label": "From",
    "section": 0,
    "block": "14",
    "limit": null,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "PeriodEnd",
    "label": "To",
    "section": 0,
    "block": "15",
    "limit": null,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "DraftNotes",
    "label": "Private draft notes",
    "section": 7,
    "block": "",
    "limit": null,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": true
  },
  {
    "key": "OfficialComments",
    "label": "Comments on Performance",
    "section": 5,
    "block": "",
    "limit": null,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": true,
    "blocks": {
      "EVAL": "43",
      "FITREP": "41",
      "CHIEFEVAL": "40"
    }
  },
  {
    "key": "OfficialName",
    "label": "Name (Last, First MI Suffix)",
    "section": 0,
    "block": "1",
    "limit": 27,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "RankRate",
    "label": "Grade/Rate",
    "section": 0,
    "block": "2",
    "limit": 5,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "Designator",
    "label": "Desig",
    "section": 0,
    "block": "",
    "limit": 12,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "UnitUIC",
    "label": "Unit ID Code (UIC)",
    "section": 0,
    "block": "6",
    "limit": 5,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ShipStation",
    "label": "Ship/Station",
    "section": 0,
    "block": "7",
    "limit": 18,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "Trait33",
    "label": "Trait33",
    "section": 4,
    "block": "33",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait34",
    "label": "Trait34",
    "section": 4,
    "block": "34",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait35",
    "label": "Trait35",
    "section": 4,
    "block": "35",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait36",
    "label": "Trait36",
    "section": 4,
    "block": "36",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait37",
    "label": "Trait37",
    "section": 4,
    "block": "37",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait38",
    "label": "Trait38",
    "section": 4,
    "block": "38",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "Trait39",
    "label": "Trait39",
    "section": 4,
    "block": "39",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "1.0",
      "2.0",
      "3.0",
      "4.0",
      "5.0"
    ],
    "multiline": false
  },
  {
    "key": "ObservationStatus",
    "label": "Not Observed Report (NOB)",
    "section": 0,
    "block": "16",
    "limit": 8,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Observed",
      "NOB"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "PromotionRecommendation",
    "label": "Individual (Promotion Recommendation)",
    "section": 5,
    "block": "45",
    "limit": 20,
    "date": false,
    "digits": null,
    "options": [
      "",
      "NOB",
      "Significant Problems",
      "Progressing",
      "Promotable",
      "Must Promote",
      "Early Promote"
    ],
    "multiline": false,
    "blocks": {
      "EVAL": "45",
      "FITREP": "42",
      "CHIEFEVAL": "41"
    }
  },
  {
    "key": "PromotionStatus",
    "label": "Promotion Status",
    "section": 0,
    "block": "8",
    "limit": 8,
    "date": false,
    "digits": null,
    "options": [
      "",
      "REGULAR",
      "FROCKED",
      "SELECTED",
      "SPOT"
    ],
    "multiline": false
  },
  {
    "key": "ReportingSeniorName",
    "label": "Reporting Senior (Last, FI MI)",
    "section": 1,
    "block": "22",
    "limit": 18,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorGrade",
    "label": "Grade",
    "section": 1,
    "block": "23",
    "limit": 5,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorDesignator",
    "label": "Desig",
    "section": 1,
    "block": "24",
    "limit": 5,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorTitle",
    "label": "Title",
    "section": 1,
    "block": "25",
    "limit": 14,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorUIC",
    "label": "UIC",
    "section": 1,
    "block": "26",
    "limit": 5,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "OccasionPeriodic",
    "label": "Periodic",
    "section": 0,
    "block": "10",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "OccasionDetachment",
    "label": "Detachment of Individual",
    "section": 0,
    "block": "11",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "OccasionFrocking",
    "label": "Detachment of Reporting Senior",
    "section": 0,
    "block": "12",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "OccasionSpecial",
    "label": "Special",
    "section": 0,
    "block": "13",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ReportRegular",
    "label": "Regular",
    "section": 0,
    "block": "17",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ReportConcurrent",
    "label": "Concurrent",
    "section": 0,
    "block": "18",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ReportOperationalCommander",
    "label": "Ops Cdr",
    "section": 0,
    "block": "19",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "PrimaryDuty",
    "label": "Primary duty abbreviation",
    "section": 2,
    "block": "29",
    "limit": 14,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "DutyDescription",
    "label": "Primary/Collateral/Watchstanding Duties/PFA",
    "section": 2,
    "block": "29",
    "limit": null,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": true
  },
  {
    "key": "MemberSSN",
    "label": "SSN",
    "section": 0,
    "block": "4",
    "limit": 9,
    "date": false,
    "digits": 9,
    "options": null,
    "multiline": false
  },
  {
    "key": "MemberDoDID",
    "label": "Member DoD ID",
    "section": 0,
    "block": "4",
    "limit": 10,
    "date": false,
    "digits": 10,
    "options": null,
    "multiline": false,
    "legacyHidden": true
  },
  {
    "key": "ComponentActive",
    "label": "ACT",
    "section": 0,
    "block": "5",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ComponentTAR",
    "label": "TAR",
    "section": 0,
    "block": "5",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ComponentInactive",
    "label": "INACT",
    "section": 0,
    "block": "5",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "ComponentATADSW",
    "label": "AT/ADOS",
    "section": 0,
    "block": "5",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true
  },
  {
    "key": "DateReported",
    "label": "Date Reported",
    "section": 0,
    "block": "9",
    "limit": 10,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "CounselingDate",
    "label": "Date Counseled",
    "section": 3,
    "block": "30",
    "limit": 10,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "CounselorName",
    "label": "Counselor",
    "section": 3,
    "block": "31",
    "limit": 20,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReadinessSummary",
    "label": "Physical Readiness",
    "section": 0,
    "block": "20",
    "limit": 4,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReadinessDetail",
    "label": "Physical Readiness \u2014 result",
    "section": 0,
    "block": "20",
    "limit": 15,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReadinessDate",
    "label": "Physical Readiness \u2014 date",
    "section": 0,
    "block": "20",
    "limit": 10,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "BilletSubcategory",
    "label": "Billet Subcategory (if any)",
    "section": 0,
    "block": "21",
    "limit": 10,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "AssignmentRecommendationA",
    "label": "Career Recommendations \u2014 first",
    "section": 5,
    "block": "41",
    "limit": 20,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "41",
      "FITREP": "40",
      "CHIEFEVAL": ""
    }
  },
  {
    "key": "AssignmentRecommendationB",
    "label": "Career Recommendations \u2014 second",
    "section": 5,
    "block": "41",
    "limit": 20,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "41",
      "FITREP": "40",
      "CHIEFEVAL": ""
    }
  },
  {
    "key": "RetentionYes",
    "label": "Recommended",
    "section": 5,
    "block": "47",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "RetentionNo",
    "label": "Not Recommended",
    "section": 5,
    "block": "47",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "StatementYes",
    "label": "I intend to submit a statement",
    "section": 6,
    "block": "51",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true,
    "blocks": {
      "EVAL": "51",
      "FITREP": "46",
      "CHIEFEVAL": "49"
    }
  },
  {
    "key": "StatementNo",
    "label": "I do not intend to submit a statement",
    "section": 6,
    "block": "51",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": [
      "",
      "Yes",
      "No"
    ],
    "multiline": false,
    "checkbox": true,
    "blocks": {
      "EVAL": "51",
      "FITREP": "46",
      "CHIEFEVAL": "49"
    }
  },
  {
    "key": "ReportingSeniorSSN",
    "label": "SSN",
    "section": 1,
    "block": "27",
    "limit": 9,
    "date": false,
    "digits": 9,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorDoDID",
    "label": "Reporting senior DoD ID",
    "section": 1,
    "block": "",
    "limit": 10,
    "date": false,
    "digits": 10,
    "options": null,
    "multiline": false,
    "legacyHidden": true
  },
  {
    "key": "ReportingSeniorAddress1",
    "label": "Address line 1",
    "section": 6,
    "block": "48",
    "limit": 30,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorAddress2",
    "label": "Address line 2",
    "section": 6,
    "block": "48",
    "limit": 30,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorCity",
    "label": "City",
    "section": 6,
    "block": "48",
    "limit": 15,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorState",
    "label": "State / region",
    "section": 6,
    "block": "48",
    "limit": 2,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorZip",
    "label": "ZIP / postal code",
    "section": 6,
    "block": "48",
    "limit": 9,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorPhone",
    "label": "Phone",
    "section": 6,
    "block": "48",
    "limit": 10,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "ReportingSeniorDSN",
    "label": "DSN",
    "section": 6,
    "block": "48",
    "limit": 7,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "blocks": {
      "EVAL": "48",
      "FITREP": "44",
      "CHIEFEVAL": "51"
    }
  },
  {
    "key": "RaterName",
    "label": "Signature of Rater \u2014 typed name & grade/rate",
    "section": 6,
    "block": "42",
    "limit": 28,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "RaterDate",
    "label": "Signature of Rater \u2014 date",
    "section": 6,
    "block": "42",
    "limit": 10,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "SeniorRaterName",
    "label": "Signature of Senior Rater \u2014 typed name & grade/rate",
    "section": 6,
    "block": "49",
    "limit": 28,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "SeniorRaterDate",
    "label": "Signature of Senior Rater \u2014 date",
    "section": 6,
    "block": "49",
    "limit": 10,
    "date": true,
    "digits": null,
    "options": null,
    "multiline": false,
    "forms": [
      "EVAL"
    ]
  },
  {
    "key": "MemberFirstName",
    "label": "Member first name",
    "section": 0,
    "block": "1",
    "limit": 27,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "MemberMiddleInitial",
    "label": "Member middle initial(s)",
    "section": 0,
    "block": "1",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "MemberLastName",
    "label": "Member last name",
    "section": 0,
    "block": "1",
    "limit": 27,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "MemberSuffix",
    "label": "Member suffix",
    "section": 0,
    "block": "1",
    "limit": 10,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorLastName",
    "label": "Reporting senior last name",
    "section": 1,
    "block": "22",
    "limit": 18,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorFirstInitial",
    "label": "Reporting senior first initial",
    "section": 1,
    "block": "22",
    "limit": 1,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "ReportingSeniorMiddleInitial",
    "label": "Reporting senior middle initial(s)",
    "section": 1,
    "block": "22",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "CounselorLastName",
    "label": "Counselor last name",
    "section": 3,
    "block": "31",
    "limit": 20,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "CounselorFirstInitial",
    "label": "Counselor first initial",
    "section": 3,
    "block": "31",
    "limit": 1,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  },
  {
    "key": "CounselorMiddleInitial",
    "label": "Counselor middle initial(s)",
    "section": 3,
    "block": "31",
    "limit": 3,
    "date": false,
    "digits": null,
    "options": null,
    "multiline": false
  }
];

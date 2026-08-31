import os
import re
import logging
import html
import hashlib
import base64
import csv
import time

from datetime import datetime
from pathlib import Path
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd
import pdfplumber

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# ============================================================
# HR JOB AUTOMATION
# GITHUB-SAFE VERSION
# ============================================================
#
# FEATURES
# ------------------------------------------------------------
# 1. Reads a local resume PDF
# 2. Reads job PDFs / Excel / CSV files
# 3. Can fetch job PDFs from Gmail
# 4. Detects QA / Testing jobs
# 5. Extracts required experience
# 6. Compares JD skills with resume skills
# 7. Calculates a realistic match score
# 8. Adds evidence strength:
#       STRONG / MEDIUM / WEAK
# 9. Creates an Excel review report
# 10. Supports TEST email sending
# 11. Supports LIVE recruiter sending only after explicit unlock
# 12. Prevents duplicate vacancy emails
# 13. Adds recruiter cooldown protection
#
# SECURITY
# ------------------------------------------------------------
# - No Gmail password is stored here
# - No App Password is stored here
# - No OAuth client secret is stored here
# - No OAuth token is stored here
# - No personal email is stored here
# - No phone number is stored here
# - No resume is stored here
#
# credentials.json and token.json must remain local and must
# never be committed to GitHub.
# ============================================================


# ============================================================
# 1. LOCAL USER CONFIGURATION
# ============================================================

# These values come from LOCAL environment variables.
# They are intentionally not hard-coded.

MY_NAME = os.getenv(
    "HR_AUTOMATION_NAME",
    "Candidate"
).strip()


MY_EMAIL = os.getenv(
    "HR_AUTOMATION_EMAIL",
    ""
).strip()


MY_PHONE = os.getenv(
    "HR_AUTOMATION_PHONE",
    ""
).strip()


MY_EXPERIENCE = float(
    os.getenv(
        "HR_AUTOMATION_EXPERIENCE",
        "5.0"
    )
)


# Optional.
# Leave empty if you do not want CTC mentioned in emails.

CURRENT_CTC = os.getenv(
    "HR_AUTOMATION_CURRENT_CTC",
    ""
).strip()


# ============================================================
# 2. SAFETY MODE
# ============================================================

# GitHub-safe default.
#
# If HR_AUTOMATION_MODE is not configured locally,
# this program starts in TEST mode.

MODE = os.getenv(
    "HR_AUTOMATION_MODE",
    "TEST"
).strip().upper()


# TEST emails default to the sender's own email.

TEST_EMAIL = os.getenv(
    "HR_AUTOMATION_TEST_EMAIL",
    MY_EMAIL
).strip()


# ------------------------------------------------------------
# LIVE MODE SECOND SAFETY LOCK
# ------------------------------------------------------------
#
# LIVE mode requires BOTH:
#
# HR_AUTOMATION_MODE=LIVE
#
# AND:
#
# HR_AUTOMATION_LIVE_UNLOCK=
# I_UNDERSTAND_THIS_SENDS_REAL_EMAILS
#
# Merely switching MODE to LIVE is therefore not enough.

LIVE_UNLOCK_VALUE = os.getenv(
    "HR_AUTOMATION_LIVE_UNLOCK",
    ""
).strip()


LIVE_SEND_UNLOCKED = (
    LIVE_UNLOCK_VALUE
    == "I_UNDERSTAND_THIS_SENDS_REAL_EMAILS"
)


# ============================================================
# 3. LIVE MODE QUALITY SETTINGS
# ============================================================

LIVE_ALLOWED_EVIDENCE = {
    "STRONG"
}


MIN_LIVE_MATCH_SCORE = 75


ONE_EMAIL_PER_UNIQUE_HR = True


HR_CONTACT_COOLDOWN_DAYS = 14


MAX_EMAILS_PER_RUN = 100


DELAY_BETWEEN_EMAILS_SECONDS = 1.5


MAX_SEND_RETRIES = 3


# ============================================================
# 4. FILE LOCATIONS
# ============================================================

RESUME_PATH = Path(
    "resume/My Resume.pdf"
)


JOB_FOLDER = Path(
    "job_files"
)


OUTPUT_FILE = Path(
    "job_review.xlsx"
)


SENT_LOG_FILE = Path(
    "sent_log.csv"
)


GOOGLE_CREDENTIALS_FILE = Path(
    "credentials.json"
)


GOOGLE_TOKEN_FILE = Path(
    "token.json"
)


GMAIL_FETCH_FOLDER = Path(
    "gmail_fetched_test"
)


PROCESSED_GMAIL_LOG = Path(
    "processed_gmail_messages.csv"
)


BACKLOG_MARKER_FILE = Path(
    "backlog_completed.flag"
)


# ============================================================
# 5. GMAIL CONFIGURATION
# ============================================================

JOB_SOURCE_EMAIL = os.getenv(
    "HR_AUTOMATION_JOB_SOURCE_EMAIL",
    "info@jobcurator.in"
).strip()


MAX_TEST_GMAIL_MESSAGES = 1


MAX_LIVE_GMAIL_MESSAGES = 20


GMAIL_SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]


# ============================================================
# 6. GENERAL SETTINGS
# ============================================================

TOP_RESULTS_TO_SHOW = 10


MIN_MATCH_SCORE = 45


logging.getLogger(
    "pdfminer"
).setLevel(
    logging.ERROR
)


logging.getLogger(
    "pdfminer.pdfpage"
).setLevel(
    logging.ERROR
)


logging.getLogger(
    "pdfminer.pdffont"
).setLevel(
    logging.ERROR
)


# ============================================================
# 7. EMAIL FILTERING
# ============================================================

BLOCKED_EXACT_EMAILS = {
    JOB_SOURCE_EMAIL.lower(),
}


BLOCKED_EMAIL_PREFIXES = {
    "noreply",
    "no-reply",
    "donotreply",
    "do-not-reply",
    "support",
    "help",
    "helpdesk",
    "admin",
    "webmaster",
}


RECRUITMENT_EMAIL_PREFIXES = {
    "hr",
    "career",
    "careers",
    "job",
    "jobs",
    "hiring",
    "hire",
    "recruiter",
    "recruitment",
    "recruiting",
    "talent",
    "talentacquisition",
    "talent-acquisition",
    "ta",
    "people",
    "peopleops",
}


EMAIL_PATTERN = re.compile(
    r"\b[A-Za-z0-9._%+-]+@"
    r"[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
)


# ============================================================
# 8. ROLE DETECTION
# ============================================================

ROLE_PATTERNS = [

    (
        "Senior SDET",
        r"\bsenior\s+sdet\b"
    ),

    (
        "SDET",
        r"\bsdet\b"
    ),

    (
        "Senior QA Automation Engineer",
        r"\bsenior\s+qa\s+automation\s+engineer\b"
    ),

    (
        "QA Automation Engineer",
        r"\bqa\s+automation\s+engineer\b"
    ),

    (
        "Automation Test Engineer",
        r"\bautomation\s+test\s+engineer\b"
    ),

    (
        "Test Automation Engineer",
        r"\btest\s+automation\s+engineer\b"
    ),

    (
        "Automation QA Engineer",
        r"\bautomation\s+qa\s+engineer\b"
    ),

    (
        "Senior QA Engineer",
        r"\bsenior\s+qa\s+engineer\b"
    ),

    (
        "QA Engineer",
        r"\bqa\s+engineer\b"
    ),

    (
        "Senior Test Engineer",
        r"\bsenior\s+test\s+engineer\b"
    ),

    (
        "Software Test Engineer",
        r"\bsoftware\s+test\s+engineer\b"
    ),

    (
        "Test Engineer",
        r"\btest\s+engineer\b"
    ),

    (
        "QA Analyst",
        r"\bqa\s+analyst\b"
    ),

    (
        "Test Analyst",
        r"\btest\s+analyst\b"
    ),

    (
        "Quality Engineer",
        r"\bquality\s+engineer\b"
    ),

    (
        "Quality Assurance Engineer",
        r"\bquality\s+assurance\s+engineer\b"
    ),

    (
        "Manual Tester",
        r"\bmanual\s+tester\b"
    ),

    (
        "Automation Tester",
        r"\bautomation\s+tester\b"
    ),

    (
        "Software Tester",
        r"\bsoftware\s+tester\b"
    ),

    (
        "QA Tester",
        r"\bqa\s+tester\b"
    ),
]


TESTING_SIGNAL_WORDS = [
    "software testing",
    "quality assurance",
    "manual testing",
    "automation testing",
    "test automation",
    "test cases",
    "defect",
    "regression",
    "smoke testing",
    "api testing",
    "postman",
    "playwright",
    "selenium",
    "pytest",
]


# ============================================================
# 9. SKILL DETECTION
# ============================================================

SKILL_PATTERNS = {

    "Playwright":
        r"\bplaywright\b",

    "Selenium":
        r"\bselenium\b",

    "Python":
        r"\bpython\b",

    "Java":
        r"\bjava\b",

    "JavaScript":
        r"\bjavascript\b|\bjava script\b",

    "TypeScript":
        r"\btypescript\b|\btype script\b",

    "PyTest":
        r"\bpytest\b",

    "API Testing":
        r"\bapi testing\b|\brest api\b|\bapi\b",

    "Postman":
        r"\bpostman\b",

    "SQL":
        r"\bsql\b",

    "Database Testing":
        r"\bdatabase testing\b|\bdb testing\b",

    "Manual Testing":
        r"\bmanual testing\b",

    "Automation Testing":
        r"\bautomation testing\b|\btest automation\b",

    "Functional Testing":
        r"\bfunctional testing\b",

    "Regression Testing":
        r"\bregression testing\b|\bregression\b",

    "Smoke Testing":
        r"\bsmoke testing\b|\bsmoke\b",

    "Sanity Testing":
        r"\bsanity testing\b|\bsanity\b",

    "UAT":
        r"\buat\b|\buser acceptance testing\b",

    "Agile":
        r"\bagile\b|\bscrum\b",

    "Jira":
        r"\bjira\b",

    "Azure DevOps":
        r"\bazure devops\b",

    "Git":
        r"\bgit\b",

    "GitHub":
        r"\bgithub\b",

    "CI/CD":
        r"\bci\s*/\s*cd\b|"
        r"\bcontinuous integration\b",

    "REST":
        r"\brest\b|\brestful\b",

    "JSON":
        r"\bjson\b",

    "JMeter":
        r"\bjmeter\b",

    "Performance Testing":
        r"\bperformance testing\b",

    "BrowserStack":
        r"\bbrowserstack\b",

    "Mobile Testing":
        r"\bmobile testing\b",

    "Accessibility Testing":
        r"\baccessibility testing\b",

    "POM":
        r"\bpage object model\b|\bpom\b",

    "Test Planning":
        r"\btest plan\b|\btest strategy\b",

    "RTM":
        r"\brtm\b|"
        r"\brequirement traceability matrix\b",

    "Defect Management":
        r"\bdefect management\b|\bbug tracking\b",
}


# ============================================================
# 10. HELPER FUNCTIONS
# ============================================================

def clean_text(text):

    if text is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(text)
    ).strip()


def normalise_email(email):

    return str(
        email
    ).strip().lower()


def get_email_prefix(email):

    return normalise_email(
        email
    ).split("@")[0]


def is_blocked_email(email):

    email = normalise_email(
        email
    )

    if email in BLOCKED_EXACT_EMAILS:
        return True

    prefix = get_email_prefix(
        email
    )

    if prefix in BLOCKED_EMAIL_PREFIXES:
        return True

    first_prefix_part = re.split(
        r"[._+-]",
        prefix
    )[0]

    return (
        first_prefix_part
        in BLOCKED_EMAIL_PREFIXES
    )


def email_confidence(email):

    prefix = get_email_prefix(
        email
    )

    prefix_parts = set(
        re.split(
            r"[._+-]",
            prefix
        )
    )

    if prefix in RECRUITMENT_EMAIL_PREFIXES:
        return "HIGH"

    if (
        prefix_parts
        & RECRUITMENT_EMAIL_PREFIXES
    ):
        return "HIGH"

    return "MEDIUM"


def email_confidence_score(email):

    if email_confidence(
        email
    ) == "HIGH":
        return 100

    return 70


def extract_valid_emails(text):

    emails = EMAIL_PATTERN.findall(
        str(text)
    )

    result = []

    for email in emails:

        email = normalise_email(
            email
        )

        if is_blocked_email(
            email
        ):
            continue

        if email not in result:
            result.append(
                email
            )

    return result


# ============================================================
# 11. ROLE FUNCTIONS
# ============================================================

def detect_role(text):

    for (
        canonical_role,
        pattern
    ) in ROLE_PATTERNS:

        if re.search(
            pattern,
            str(text),
            flags=re.IGNORECASE
        ):
            return canonical_role

    return None


def is_testing_job(text):

    if detect_role(
        text
    ):
        return True

    lower_text = str(
        text
    ).lower()

    signal_count = sum(
        signal in lower_text
        for signal
        in TESTING_SIGNAL_WORDS
    )

    return signal_count >= 3


def role_score(text):

    if detect_role(
        text
    ):
        return 100

    return 60


# ============================================================
# 12. SKILL FUNCTIONS
# ============================================================

def extract_skills(text):

    found = []

    for (
        skill,
        pattern
    ) in SKILL_PATTERNS.items():

        if re.search(
            pattern,
            str(text),
            flags=re.IGNORECASE
        ):
            found.append(
                skill
            )

    return found


def calculate_skill_scores(
    jd_skills,
    matched_skills
):

    unique_jd = set(
        jd_skills
    )

    unique_matched = set(
        matched_skills
    )

    if not unique_jd:
        return 35, 20

    coverage = (
        len(
            unique_matched
        )
        / len(
            unique_jd
        )
    ) * 100

    jd_skill_count = len(
        unique_jd
    )

    if jd_skill_count == 1:
        evidence_score = 35

    elif jd_skill_count == 2:
        evidence_score = 55

    elif jd_skill_count == 3:
        evidence_score = 75

    elif jd_skill_count == 4:
        evidence_score = 90

    else:
        evidence_score = 100

    return (
        coverage,
        evidence_score
    )


# ============================================================
# 13. EXPERIENCE FUNCTIONS
# ============================================================

def parse_experience_from_text(
    text
):

    text = str(
        text
    ).lower()

    range_match = re.search(
        r"(\d+(?:\.\d+)?)\s*"
        r"(?:-|–|—|to)\s*"
        r"(\d+(?:\.\d+)?)\s*"
        r"(?:years?|yrs?)",
        text
    )

    if range_match:

        return (
            float(
                range_match.group(1)
            ),
            float(
                range_match.group(2)
            )
        )

    plus_match = re.search(
        r"(\d+(?:\.\d+)?)\s*\+\s*"
        r"(?:years?|yrs?)",
        text
    )

    if plus_match:

        return (
            float(
                plus_match.group(1)
            ),
            None
        )

    minimum_match = re.search(
        r"(?:minimum|min\.?|at least)"
        r"\s*[:\-]?\s*"
        r"(\d+(?:\.\d+)?)\s*"
        r"(?:years?|yrs?)",
        text
    )

    if minimum_match:

        return (
            float(
                minimum_match.group(1)
            ),
            None
        )

    return (
        None,
        None
    )


def extract_experience(
    job_text
):

    lines = str(
        job_text
    ).splitlines()

    likely_lines = []

    for line in lines:

        lower_line = line.lower()

        if (
            "experience"
            in lower_line
            or re.search(
                r"\bexp\b",
                lower_line
            )
            or "yrs"
            in lower_line
            or "years"
            in lower_line
        ):
            likely_lines.append(
                line
            )

    for line in likely_lines:

        (
            min_exp,
            max_exp
        ) = parse_experience_from_text(
            line
        )

        if min_exp is not None:

            return (
                min_exp,
                max_exp
            )

    return parse_experience_from_text(
        job_text
    )


def experience_matches(
    min_exp,
    max_exp
):

    if min_exp is None:
        return False

    if MY_EXPERIENCE < min_exp:
        return False

    if (
        max_exp is not None
        and
        MY_EXPERIENCE
        > max_exp + 1
    ):
        return False

    return True


def experience_score(
    min_exp,
    max_exp
):

    if min_exp is None:
        return 0

    if MY_EXPERIENCE < min_exp:
        return 0

    if max_exp is None:
        return 95

    if (
        min_exp
        <= MY_EXPERIENCE
        <= max_exp
    ):
        return 100

    if (
        MY_EXPERIENCE
        == max_exp + 1
    ):
        return 80

    return 0


def format_experience(
    min_exp,
    max_exp
):

    if min_exp is None:
        return "Unknown"

    min_text = f"{min_exp:g}"

    if (
        max_exp is None
        or pd.isna(
            max_exp
        )
    ):
        return (
            f"{min_text}+ years"
        )

    return (
        f"{min_text}-"
        f"{max_exp:g} years"
    )


# ============================================================
# 14. PDF READING
# ============================================================

def read_pdf_pages(
    file_path
):

    pages = []

    with pdfplumber.open(
        file_path
    ) as pdf:

        for (
            page_number,
            page
        ) in enumerate(
            pdf.pages,
            start=1
        ):

            pages.append(
                {
                    "page_number":
                        page_number,

                    "text":
                        page.extract_text()
                        or ""
                }
            )

    return pages


def read_complete_pdf_text(
    file_path
):

    pages = read_pdf_pages(
        file_path
    )

    return "\n".join(
        page["text"]
        for page
        in pages
    )


def read_resume_skills():

    if not RESUME_PATH.exists():

        raise FileNotFoundError(
            f"Resume not found: "
            f"{RESUME_PATH}"
        )

    print(
        "\nReading resume..."
    )

    resume_text = (
        read_complete_pdf_text(
            RESUME_PATH
        )
    )

    skills = extract_skills(
        resume_text
    )

    print(
        "Resume skills detected:"
    )

    print(
        ", ".join(
            skills
        )
        if skills
        else
        "No known skills detected."
    )

    return skills


# ============================================================
# 15. JOB CONTEXT
# ============================================================

def get_email_line_context(
    page_text,
    email,
    before=10,
    after=8
):

    lines = str(
        page_text
    ).splitlines()

    email_lower = (
        email.lower()
    )

    matching_indexes = []

    for (
        index,
        line
    ) in enumerate(
        lines
    ):

        if (
            email_lower
            in line.lower()
        ):

            matching_indexes.append(
                index
            )

    contexts = []

    for index in matching_indexes:

        start = max(
            0,
            index - before
        )

        end = min(
            len(
                lines
            ),
            index + after + 1
        )

        contexts.append(
            "\n".join(
                lines[
                    start:end
                ]
            )
        )

    return contexts


# ============================================================
# 16. MATCHING + SCORING
# ============================================================

def calculate_match(
    job_text,
    resume_skills,
    min_exp,
    max_exp,
    email
):

    jd_skills = extract_skills(
        job_text
    )

    matched_skills = sorted(
        set(
            jd_skills
        )
        &
        set(
            resume_skills
        )
    )

    (
        skill_coverage_score,
        skill_evidence_score
    ) = calculate_skill_scores(
        jd_skills,
        matched_skills
    )

    exp_score = experience_score(
        min_exp,
        max_exp
    )

    r_score = role_score(
        job_text
    )

    e_score = (
        email_confidence_score(
            email
        )
    )

    final_score = (

        skill_coverage_score
        * 0.50

        + exp_score
        * 0.20

        + r_score
        * 0.15

        + skill_evidence_score
        * 0.10

        + e_score
        * 0.05
    )

    if (
        len(
            set(
                jd_skills
            )
        )
        == 1
    ):

        final_score = min(
            final_score,
            72
        )

    elif (
        len(
            set(
                jd_skills
            )
        )
        == 2
    ):

        final_score = min(
            final_score,
            82
        )

    return {

        "final_score":
            round(
                final_score,
                1
            ),

        "jd_skills":
            jd_skills,

        "matched_skills":
            matched_skills,

        "skill_coverage_score":
            round(
                skill_coverage_score,
                1
            ),

        "skill_evidence_score":
            skill_evidence_score,

        "experience_score":
            exp_score,

        "role_score":
            r_score,

        "email_score":
            e_score,
    }


def evidence_strength(
    jd_skills,
    matched_skills,
    role,
    email_conf,
    score
):

    jd_count = len(
        set(
            jd_skills
        )
    )

    matched_count = len(
        set(
            matched_skills
        )
    )

    if (
        jd_count >= 4
        and
        matched_count >= 3
        and
        role
        != "QA / Software Testing Role"
        and
        score >= 75
    ):
        return "STRONG"

    if (
        jd_count >= 2
        and
        matched_count >= 1
        and
        score >= 60
    ):
        return "MEDIUM"

    return "WEAK"


def analyse_job_context(
    job_text,
    email,
    source_file,
    page_number,
    resume_skills
):

    if is_blocked_email(
        email
    ):
        return None

    if not is_testing_job(
        job_text
    ):
        return None

    role = detect_role(
        job_text
    )

    if role is None:

        role = (
            "QA / Software Testing Role"
        )

    (
        min_exp,
        max_exp
    ) = extract_experience(
        job_text
    )

    if not experience_matches(
        min_exp,
        max_exp
    ):
        return None

    match = calculate_match(
        job_text,
        resume_skills,
        min_exp,
        max_exp,
        email
    )

    score = match[
        "final_score"
    ]

    if score < MIN_MATCH_SCORE:
        return None

    email_conf = (
        email_confidence(
            email
        )
    )

    evidence = evidence_strength(
        match["jd_skills"],
        match["matched_skills"],
        role,
        email_conf,
        score
    )

    return {

        "Email":
            email,

        "Email Confidence":
            email_conf,

        "Role":
            role,

        "Experience":
            format_experience(
                min_exp,
                max_exp
            ),

        "Minimum Experience":
            min_exp,

        "Maximum Experience":
            max_exp,

        "JD Skill Count":
            len(
                set(
                    match[
                        "jd_skills"
                    ]
                )
            ),

        "Matched Skill Count":
            len(
                set(
                    match[
                        "matched_skills"
                    ]
                )
            ),

        "JD Skills":
            ", ".join(
                match[
                    "jd_skills"
                ]
            ),

        "Matched Resume Skills":
            ", ".join(
                match[
                    "matched_skills"
                ]
            ),

        "Skill Coverage Score":
            match[
                "skill_coverage_score"
            ],

        "Skill Evidence Score":
            match[
                "skill_evidence_score"
            ],

        "Experience Score":
            match[
                "experience_score"
            ],

        "Role Score":
            match[
                "role_score"
            ],

        "Email Score":
            match[
                "email_score"
            ],

        "Match Score":
            score,

        "Evidence Strength":
            evidence,

        "Source PDF/File":
            source_file,

        "Page":
            page_number,

        "Status":
            "REVIEW",

        "Job Context":
            clean_text(
                job_text
            )
    }


# ============================================================
# 17. FILE SCANNING
# ============================================================

def scan_pdf(
    file_path,
    resume_skills
):

    print(
        f"Reading PDF: "
        f"{file_path.name}"
    )

    results = []

    for page in read_pdf_pages(
        file_path
    ):

        page_number = page[
            "page_number"
        ]

        page_text = page[
            "text"
        ]

        emails = extract_valid_emails(
            page_text
        )

        for email in emails:

            contexts = (
                get_email_line_context(
                    page_text,
                    email
                )
            )

            for context in contexts:

                result = analyse_job_context(
                    context,
                    email,
                    file_path.name,
                    page_number,
                    resume_skills
                )

                if result:
                    results.append(
                        result
                    )

    return results


def scan_excel(
    file_path,
    resume_skills
):

    print(
        f"Reading Excel: "
        f"{file_path.name}"
    )

    results = []

    sheets = pd.read_excel(
        file_path,
        sheet_name=None
    )

    for (
        sheet_name,
        dataframe
    ) in sheets.items():

        for (
            row_number,
            (_, row)
        ) in enumerate(
            dataframe.iterrows(),
            start=2
        ):

            values = [

                str(
                    value
                )

                for value
                in row.values

                if pd.notna(
                    value
                )
            ]

            row_text = (
                " | ".join(
                    values
                )
            )

            for email in extract_valid_emails(
                row_text
            ):

                result = analyse_job_context(
                    row_text,
                    email,
                    (
                        f"{file_path.name} "
                        f"/ {sheet_name}"
                    ),
                    row_number,
                    resume_skills
                )

                if result:

                    results.append(
                        result
                    )

    return results


def scan_csv(
    file_path,
    resume_skills
):

    print(
        f"Reading CSV: "
        f"{file_path.name}"
    )

    results = []

    dataframe = pd.read_csv(
        file_path
    )

    for (
        row_number,
        (_, row)
    ) in enumerate(
        dataframe.iterrows(),
        start=2
    ):

        values = [

            str(
                value
            )

            for value
            in row.values

            if pd.notna(
                value
            )
        ]

        row_text = (
            " | ".join(
                values
            )
        )

        for email in extract_valid_emails(
            row_text
        ):

            result = analyse_job_context(
                row_text,
                email,
                file_path.name,
                row_number,
                resume_skills
            )

            if result:

                results.append(
                    result
                )

    return results


# ============================================================
# 18. DUPLICATE REMOVAL
# ============================================================

def remove_duplicates(
    results
):

    best_records = {}

    for record in results:

        key = (

            record[
                "Email"
            ],

            record[
                "Role"
            ],

            record[
                "Minimum Experience"
            ],

            record[
                "Maximum Experience"
            ]
        )

        if key not in best_records:

            best_records[
                key
            ] = record

        elif (
            record[
                "Match Score"
            ]
            >
            best_records[
                key
            ][
                "Match Score"
            ]
        ):

            best_records[
                key
            ] = record

    return list(
        best_records.values()
    )


# ============================================================
# 19. SORTING
# ============================================================

def evidence_rank(
    value
):

    ranking = {

        "STRONG":
            3,

        "MEDIUM":
            2,

        "WEAK":
            1
    }

    return ranking.get(
        value,
        0
    )


def sort_results(
    dataframe
):

    dataframe = (
        dataframe.copy()
    )

    dataframe[
        "_Evidence Rank"
    ] = dataframe[
        "Evidence Strength"
    ].map(
        evidence_rank
    )

    dataframe = (
        dataframe.sort_values(

            by=[
                "_Evidence Rank",
                "Match Score",
                "Matched Skill Count"
            ],

            ascending=[
                False,
                False,
                False
            ]
        )
    )

    return dataframe.drop(
        columns=[
            "_Evidence Rank"
        ]
    )


# ============================================================
# 20. PRINT RESULTS
# ============================================================

def print_top_results(
    dataframe
):

    number_to_show = min(
        TOP_RESULTS_TO_SHOW,
        len(
            dataframe
        )
    )

    print(
        "\n"
        + "=" * 76
    )

    print(
        f"TOP {number_to_show} "
        f"JOB MATCHES"
    )

    print(
        "=" * 76
    )

    for (
        position,
        (_, row)
    ) in enumerate(
        dataframe.head(
            number_to_show
        ).iterrows(),
        start=1
    ):

        print(
            f"\n#{position}"
        )

        print(
            "-" * 76
        )

        print(
            "Role:",
            row[
                "Role"
            ]
        )

        print(
            "Email:",
            row[
                "Email"
            ]
        )

        print(
            "Experience:",
            row[
                "Experience"
            ]
        )

        print(
            "Evidence:",
            row[
                "Evidence Strength"
            ]
        )

        print(
            "Match Score:",
            f"{row['Match Score']}%"
        )

        print(
            "Matched Skills:",
            row[
                "Matched Resume Skills"
            ]
        )

        print(
            "Source:",
            row[
                "Source PDF/File"
            ]
        )

        print(
            "Page:",
            row[
                "Page"
            ]
        )

    print(
        "\n"
        + "=" * 76
    )


# ============================================================
# 21. GMAIL OAUTH
# ============================================================

def get_google_credentials():

    if not GOOGLE_CREDENTIALS_FILE.exists():

        raise FileNotFoundError(

            "\ncredentials.json was not found.\n\n"

            "Create a Google Cloud Desktop OAuth "
            "client and download the JSON file.\n"

            "Rename it to credentials.json and "
            "place it next to main.py.\n\n"

            "Never upload credentials.json "
            "or token.json to GitHub."
        )

    creds = None

    if GOOGLE_TOKEN_FILE.exists():

        try:

            creds = (
                Credentials
                .from_authorized_user_file(
                    str(
                        GOOGLE_TOKEN_FILE
                    ),
                    GMAIL_SCOPES
                )
            )

            if not creds.has_scopes(
                GMAIL_SCOPES
            ):
                creds = None

        except Exception:

            creds = None

    if (
        not creds
        or
        not creds.valid
    ):

        if (
            creds
            and
            creds.expired
            and
            creds.refresh_token
        ):

            creds.refresh(
                Request()
            )

        else:

            flow = (
                InstalledAppFlow
                .from_client_secrets_file(
                    str(
                        GOOGLE_CREDENTIALS_FILE
                    ),
                    GMAIL_SCOPES
                )
            )

            print(
                "\nGoogle OAuth "
                "authorization required."
            )

            creds = (
                flow.run_local_server(
                    port=0
                )
            )

        GOOGLE_TOKEN_FILE.write_text(
            creds.to_json(),
            encoding="utf-8"
        )

    return creds


def verify_authorized_google_account(
    creds
):

    service = build(
        "oauth2",
        "v2",
        credentials=creds,
        cache_discovery=False
    )

    user_info = (
        service
        .userinfo()
        .get()
        .execute()
    )

    authorized_email = normalise_email(
        user_info.get(
            "email",
            ""
        )
    )

    expected_email = normalise_email(
        MY_EMAIL
    )

    if (
        authorized_email
        != expected_email
    ):

        raise ValueError(

            "\nWrong Google account authorized.\n"

            f"Expected: {expected_email}\n"

            f"Authorized: "
            f"{authorized_email or 'unknown'}\n\n"

            "Delete token.json and run again "
            "with the correct Google account."
        )

    print(
        "\nGoogle account verified."
    )


def get_gmail_service():

    creds = (
        get_google_credentials()
    )

    verify_authorized_google_account(
        creds
    )

    return build(
        "gmail",
        "v1",
        credentials=creds,
        cache_discovery=False
    )


# ============================================================
# 22. GMAIL ATTACHMENT HELPERS
# ============================================================

def get_header_value(
    headers,
    header_name
):

    target = (
        header_name.lower()
    )

    for header in (
        headers
        or []
    ):

        if (
            str(
                header.get(
                    "name",
                    ""
                )
            ).lower()
            == target
        ):

            return str(
                header.get(
                    "value",
                    ""
                )
            )

    return ""


def iter_message_parts(
    part
):

    if not part:
        return

    yield part

    for child in (
        part.get(
            "parts",
            []
        )
        or []
    ):

        yield from (
            iter_message_parts(
                child
            )
        )


def decode_gmail_base64(
    data
):

    if not data:
        return b""

    padding = (
        "="
        * (
            -len(
                data
            )
            % 4
        )
    )

    return (
        base64
        .urlsafe_b64decode(
            (
                data
                + padding
            ).encode(
                "utf-8"
            )
        )
    )


def download_pdf_part(
    gmail_service,
    message_id,
    part,
    destination_folder
):

    filename = Path(
        str(
            part.get(
                "filename",
                ""
            )
        ).strip()
    ).name

    if not filename.lower().endswith(
        ".pdf"
    ):
        return None

    body = (
        part.get(
            "body",
            {}
        )
        or {}
    )

    attachment_id = body.get(
        "attachmentId"
    )

    inline_data = body.get(
        "data"
    )

    if attachment_id:

        attachment = (

            gmail_service
            .users()
            .messages()
            .attachments()
            .get(

                userId="me",

                messageId=
                    message_id,

                id=
                    attachment_id
            )
            .execute()
        )

        raw_bytes = (
            decode_gmail_base64(
                attachment.get(
                    "data",
                    ""
                )
            )
        )

    elif inline_data:

        raw_bytes = (
            decode_gmail_base64(
                inline_data
            )
        )

    else:

        return None

    if not raw_bytes:

        return None

    safe_path = (
        destination_folder
        / filename
    )

    if safe_path.exists():

        safe_path = (

            destination_folder

            / (
                f"{message_id[:8]}_"
                f"{filename}"
            )
        )

    safe_path.write_bytes(
        raw_bytes
    )

    return safe_path


# ============================================================
# 23. GMAIL PROCESSING STATE
# ============================================================

def read_processed_gmail_ids():

    if not (
        PROCESSED_GMAIL_LOG.exists()
    ):

        return set()

    try:

        dataframe = pd.read_csv(
            PROCESSED_GMAIL_LOG
        )

    except Exception:

        return set()

    if (
        "Gmail Message ID"
        not in dataframe.columns
    ):

        return set()

    return {

        str(
            value
        ).strip()

        for value
        in dataframe[
            "Gmail Message ID"
        ].dropna()

        if str(
            value
        ).strip()
    }


def read_last_processed_internal_date():

    if not (
        PROCESSED_GMAIL_LOG.exists()
    ):

        return None

    try:

        dataframe = pd.read_csv(
            PROCESSED_GMAIL_LOG
        )

    except Exception:

        return None

    if (
        "Internal Date"
        not in dataframe.columns
    ):

        return None

    values = pd.to_numeric(

        dataframe[
            "Internal Date"
        ],

        errors="coerce"

    ).dropna()

    if values.empty:
        return None

    return int(
        values.max()
    )


def append_processed_gmail_message(
    message_id,
    subject,
    date_value,
    internal_date
):

    file_exists = (
        PROCESSED_GMAIL_LOG.exists()
    )

    columns = [

        "Processed At",

        "Gmail Message ID",

        "Internal Date",

        "Subject",

        "Date"
    ]

    row = {

        "Processed At":
            datetime.now().isoformat(
                timespec="seconds"
            ),

        "Gmail Message ID":
            message_id,

        "Internal Date":
            str(
                internal_date
            ),

        "Subject":
            subject,

        "Date":
            date_value
    }

    with (
        PROCESSED_GMAIL_LOG.open(
            "a",
            newline="",
            encoding="utf-8"
        )
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=columns
        )

        if not file_exists:

            writer.writeheader()

        writer.writerow(
            row
        )


# ============================================================
# 24. FETCH JOB EMAILS
# ============================================================

def fetch_job_source_emails(
    gmail_service,
    max_messages,
    initial_backlog_phase
):

    query = (
        f"from:{JOB_SOURCE_EMAIL} "
        "has:attachment "
        "filename:pdf"
    )

    response = (

        gmail_service
        .users()
        .messages()
        .list(

            userId="me",

            q=query,

            maxResults=100
        )
        .execute()
    )

    message_refs = response.get(
        "messages",
        []
    )

    if not message_refs:

        print(
            "\nNo matching "
            "job-source emails found."
        )

        return (
            [],
            []
        )

    processed_ids = (
        read_processed_gmail_ids()
    )

    last_internal_date = (
        read_last_processed_internal_date()
    )

    detailed_messages = []

    for message_ref in message_refs:

        message_id = (
            message_ref[
                "id"
            ]
        )

        message = (

            gmail_service
            .users()
            .messages()
            .get(

                userId="me",

                id=message_id,

                format="full"
            )
            .execute()
        )

        internal_date = int(
            message.get(
                "internalDate",
                "0"
            )
        )

        if initial_backlog_phase:

            detailed_messages.append(
                message
            )

            continue

        if (
            last_internal_date
            is not None
            and
            internal_date
            <= last_internal_date
        ):

            continue

        if (
            message_id
            in processed_ids
        ):

            continue

        detailed_messages.append(
            message
        )

    if not detailed_messages:

        print(
            "\nNo new source "
            "emails to process."
        )

        return (
            [],
            []
        )

    if initial_backlog_phase:

        detailed_messages.sort(

            key=lambda message:
                int(
                    message.get(
                        "internalDate",
                        "0"
                    )
                ),

            reverse=True
        )

        selected_messages = [
            detailed_messages[0]
        ]

    else:

        detailed_messages.sort(

            key=lambda message:
                int(
                    message.get(
                        "internalDate",
                        "0"
                    )
                )
        )

        selected_messages = (
            detailed_messages[
                :max_messages
            ]
        )

    GMAIL_FETCH_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    for old_file in (
        GMAIL_FETCH_FOLDER.iterdir()
    ):

        if old_file.is_file():

            old_file.unlink()

    all_downloaded_pdfs = []

    source_messages = []

    for selected_message in (
        selected_messages
    ):

        message_id = (
            selected_message[
                "id"
            ]
        )

        internal_date = int(
            selected_message.get(
                "internalDate",
                "0"
            )
        )

        payload = selected_message.get(
            "payload",
            {}
        )

        headers = payload.get(
            "headers",
            []
        )

        sender = get_header_value(
            headers,
            "From"
        )

        subject = get_header_value(
            headers,
            "Subject"
        )

        date_value = get_header_value(
            headers,
            "Date"
        )

        if (
            JOB_SOURCE_EMAIL.lower()
            not in sender.lower()
        ):

            raise ValueError(
                "Unexpected Gmail "
                "source sender."
            )

        message_pdf_paths = []

        for part in iter_message_parts(
            payload
        ):

            downloaded_path = (
                download_pdf_part(

                    gmail_service,

                    message_id,

                    part,

                    GMAIL_FETCH_FOLDER
                )
            )

            if downloaded_path:

                message_pdf_paths.append(
                    downloaded_path
                )

                all_downloaded_pdfs.append(
                    downloaded_path
                )

        if not message_pdf_paths:

            continue

        source_messages.append(
            {

                "message_id":
                    message_id,

                "internal_date":
                    internal_date,

                "subject":
                    subject,

                "date":
                    date_value,

                "pdf_count":
                    len(
                        message_pdf_paths
                    )
            }
        )

        print(
            "\nJob source email selected."
        )

        print(
            "PDF attachments:",
            len(
                message_pdf_paths
            )
        )

    return (
        all_downloaded_pdfs,
        source_messages
    )


# ============================================================
# 25. BACKLOG POLICY
# ============================================================

def backlog_is_completed():

    return (
        BACKLOG_MARKER_FILE.exists()
    )


def mark_backlog_completed():

    BACKLOG_MARKER_FILE.write_text(

        "Initial local job backlog "
        "completed.\n",

        encoding="utf-8"
    )


def collect_job_files(
    fetched_pdfs
):

    combined = []

    seen_paths = set()

    def add_file(
        path
    ):

        path = Path(
            path
        )

        if (
            not path.exists()
            or
            not path.is_file()
        ):
            return

        if (
            path.suffix.lower()
            not in {
                ".pdf",
                ".xlsx",
                ".xls",
                ".csv"
            }
        ):
            return

        resolved = str(
            path.resolve()
        )

        if resolved in seen_paths:
            return

        seen_paths.add(
            resolved
        )

        combined.append(
            path
        )

    for path in fetched_pdfs:

        add_file(
            path
        )

    if (
        not backlog_is_completed()
        and
        JOB_FOLDER.exists()
    ):

        for path in (
            JOB_FOLDER.iterdir()
        ):

            add_file(
                path
            )

    return sorted(

        combined,

        key=lambda path:
            path.name.lower()
    )


# ============================================================
# 26. SCAN ALL JOBS
# ============================================================

def scan_all_jobs(
    job_files
):

    print(
        "\n"
        + "=" * 76
    )

    print(
        "HR JOB AUTOMATION"
    )

    print(
        f"MODE: {MODE}"
    )

    print(
        "=" * 76
    )

    resume_skills = (
        read_resume_skills()
    )

    all_results = []

    for file_path in job_files:

        extension = (
            file_path.suffix.lower()
        )

        try:

            if extension == ".pdf":

                all_results.extend(
                    scan_pdf(
                        file_path,
                        resume_skills
                    )
                )

            elif extension in {
                ".xlsx",
                ".xls"
            }:

                all_results.extend(
                    scan_excel(
                        file_path,
                        resume_skills
                    )
                )

            elif extension == ".csv":

                all_results.extend(
                    scan_csv(
                        file_path,
                        resume_skills
                    )
                )

        except Exception as error:

            print(
                f"Could not process "
                f"{file_path.name}: "
                f"{error}"
            )

    if not all_results:

        print(
            "\nNo suitable "
            "QA/testing jobs found."
        )

        return pd.DataFrame()

    clean_results = (
        remove_duplicates(
            all_results
        )
    )

    dataframe = pd.DataFrame(
        clean_results
    )

    dataframe = sort_results(
        dataframe
    )

    dataframe.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print_top_results(
        dataframe
    )

    print(
        f"\nReview file created: "
        f"{OUTPUT_FILE}"
    )

    return dataframe


# ============================================================
# 27. EMAIL CONTENT
# ============================================================

def top_matched_skills(
    skill_text,
    limit=6
):

    if (
        not skill_text
        or
        pd.isna(
            skill_text
        )
    ):

        return []

    skills = [

        skill.strip()

        for skill
        in str(
            skill_text
        ).split(",")

        if skill.strip()
    ]

    return skills[
        :limit
    ]


def build_email_subject(
    selected_job
):

    role = str(
        selected_job[
            "Role"
        ]
    )

    skills = top_matched_skills(

        selected_job[
            "Matched Resume Skills"
        ],

        limit=3
    )

    skills_part = (

        " | "
        + " | ".join(
            skills
        )

        if skills
        else ""
    )

    return (

        f"Application for {role} | "

        f"{MY_EXPERIENCE:g} Years"

        f"{skills_part}"
    )


def build_email_bodies(
    selected_job
):

    role = str(
        selected_job[
            "Role"
        ]
    )

    required_experience = str(
        selected_job[
            "Experience"
        ]
    )

    skills = top_matched_skills(

        selected_job[
            "Matched Resume Skills"
        ]
    )

    if skills:

        plain_skills = "\n".join(

            f"• {skill}"

            for skill
            in skills
        )

        html_skills = "".join(

            f"<li>"
            f"{html.escape(skill)}"
            f"</li>"

            for skill
            in skills
        )

    else:

        plain_skills = (
            "• Software Testing "
            "& Quality Assurance"
        )

        html_skills = (
            "<li>"
            "Software Testing "
            "&amp; Quality Assurance"
            "</li>"
        )

    phone_plain = ""

    phone_html = ""

    if MY_PHONE:

        phone_plain = (
            f"\nPhone: "
            f"{MY_PHONE}"
        )

        phone_html = (
            f"<br>Phone: "
            f"{html.escape(MY_PHONE)}"
        )

    ctc_plain = ""

    ctc_html = ""

    if CURRENT_CTC:

        ctc_plain = (
            f"\nCurrent CTC: "
            f"₹{CURRENT_CTC} LPA\n"
        )

        ctc_html = (
            "<p>"
            "<b>Current CTC:</b> "
            f"₹{html.escape(CURRENT_CTC)} "
            "LPA"
            "</p>"
        )

    plain_body = f"""Hello Hiring Team,

I am reaching out regarding the {role} opportunity.

I bring {MY_EXPERIENCE:g} years of professional experience, with relevant experience across software testing and quality engineering.
{ctc_plain}
Key skills aligned with this requirement:
{plain_skills}

The role mentions {required_experience} of experience, which aligns with my current experience level.

I have attached my updated resume for your consideration. I would welcome an opportunity to discuss how my experience and testing background can contribute to your team.

Regards,
{MY_NAME}{phone_plain}
Email: {MY_EMAIL}
"""

    html_body = f"""
<html>
<body style="
font-family: Arial, sans-serif;
font-size: 14px;
line-height: 1.55;
color: #222;
">

<p>Hello Hiring Team,</p>

<p>
I am reaching out regarding the
<b>{html.escape(role)}</b>
opportunity.
</p>

<p>
I bring
<b>{MY_EXPERIENCE:g} years of professional experience</b>,
with relevant experience across software testing
and quality engineering.
</p>

{ctc_html}

<p>
<b>Key skills aligned with this requirement:</b>
</p>

<ul>
{html_skills}
</ul>

<p>
The role mentions
<b>{html.escape(required_experience)}</b>
of experience, which aligns with my
current experience level.
</p>

<p>
I have attached my updated resume for your
consideration. I would welcome an opportunity
to discuss how my experience and testing
background can contribute to your team.
</p>

<p>
Regards,<br>
<b>{html.escape(MY_NAME)}</b>
{phone_html}<br>
Email: {html.escape(MY_EMAIL)}
</p>

</body>
</html>
"""

    return (
        plain_body,
        html_body
    )


def create_email_message(
    selected_job,
    recipient
):

    recipient = normalise_email(
        recipient
    )

    message = EmailMessage()

    message[
        "From"
    ] = formataddr(
        (
            MY_NAME,
            MY_EMAIL
        )
    )

    message[
        "To"
    ] = recipient

    message[
        "Subject"
    ] = build_email_subject(
        selected_job
    )

    (
        plain_body,
        html_body
    ) = build_email_bodies(
        selected_job
    )

    message.set_content(
        plain_body
    )

    message.add_alternative(
        html_body,
        subtype="html"
    )

    if not RESUME_PATH.exists():

        raise FileNotFoundError(
            f"Resume not found: "
            f"{RESUME_PATH}"
        )

    resume_bytes = (
        RESUME_PATH.read_bytes()
    )

    message.add_attachment(

        resume_bytes,

        maintype="application",

        subtype="pdf",

        filename=
            RESUME_PATH.name
    )

    return message


# ============================================================
# 28. GMAIL SEND
# ============================================================

def gmail_api_send(
    gmail_service,
    message
):

    raw_message = (
        base64
        .urlsafe_b64encode(
            message.as_bytes()
        )
        .decode(
            "utf-8"
        )
    )

    last_error = None

    for attempt in range(
        1,
        MAX_SEND_RETRIES + 1
    ):

        try:

            return (

                gmail_service
                .users()
                .messages()
                .send(

                    userId="me",

                    body={
                        "raw":
                            raw_message
                    }
                )
                .execute()
            )

        except HttpError as error:

            last_error = error

            status = getattr(
                error.resp,
                "status",
                None
            )

            retryable = (
                status
                in {
                    429,
                    500,
                    502,
                    503,
                    504
                }
            )

            if (
                retryable
                and
                attempt
                < MAX_SEND_RETRIES
            ):

                wait_seconds = (
                    2 ** attempt
                )

                print(
                    f"Temporary Gmail "
                    f"error {status}. "
                    f"Retrying..."
                )

                time.sleep(
                    wait_seconds
                )

                continue

            raise

    raise last_error


# ============================================================
# 29. SENT LOG
# ============================================================

SENT_LOG_COLUMNS = [

    "Timestamp",

    "Mode",

    "Recipient",

    "Intended HR",

    "Role",

    "Experience",

    "Matched Skills",

    "Vacancy Key",

    "Source",

    "Page",

    "Status",

    "Gmail Message ID",

    "Error"
]


def build_vacancy_key(
    selected_job
):

    email = normalise_email(
        selected_job[
            "Email"
        ]
    )

    role = clean_text(
        selected_job[
            "Role"
        ]
    ).lower()

    experience = clean_text(
        selected_job[
            "Experience"
        ]
    ).lower()

    skills = sorted(

        skill.strip().lower()

        for skill
        in str(
            selected_job[
                "Matched Resume Skills"
            ]
        ).split(",")

        if skill.strip()
    )

    context = clean_text(
        selected_job.get(
            "Job Context",
            ""
        )
    ).lower()

    fingerprint = "|".join(
        [
            email,
            role,
            experience,
            ",".join(
                skills
            ),
            context[:1200]
        ]
    )

    return (
        hashlib
        .sha256(
            fingerprint.encode(
                "utf-8"
            )
        )
        .hexdigest()
    )


def append_sent_log(
    *,
    mode,
    recipient,
    selected_job,
    status,
    gmail_message_id="",
    error=""
):

    file_exists = (
        SENT_LOG_FILE.exists()
    )

    row = {

        "Timestamp":
            datetime.now().isoformat(
                timespec="seconds"
            ),

        "Mode":
            mode,

        "Recipient":
            recipient,

        "Intended HR":
            selected_job[
                "Email"
            ],

        "Role":
            selected_job[
                "Role"
            ],

        "Experience":
            selected_job[
                "Experience"
            ],

        "Matched Skills":
            selected_job[
                "Matched Resume Skills"
            ],

        "Vacancy Key":
            build_vacancy_key(
                selected_job
            ),

        "Source":
            selected_job[
                "Source PDF/File"
            ],

        "Page":
            selected_job[
                "Page"
            ],

        "Status":
            status,

        "Gmail Message ID":
            gmail_message_id,

        "Error":
            error
    }

    with SENT_LOG_FILE.open(
        "a",
        newline="",
        encoding="utf-8"
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=
                SENT_LOG_COLUMNS
        )

        if not file_exists:

            writer.writeheader()

        writer.writerow(
            row
        )


def successful_live_log():

    if not SENT_LOG_FILE.exists():

        return pd.DataFrame()

    try:

        dataframe = pd.read_csv(
            SENT_LOG_FILE
        )

    except Exception:

        return pd.DataFrame()

    required = {
        "Mode",
        "Status"
    }

    if not required.issubset(
        dataframe.columns
    ):

        return pd.DataFrame()

    return dataframe[

        (
            dataframe[
                "Mode"
            ]
            == "LIVE"
        )

        &

        (
            dataframe[
                "Status"
            ]
            == "SENT"
        )
    ]


def successfully_sent_vacancy_keys():

    successful = (
        successful_live_log()
    )

    if (
        successful.empty
        or
        "Vacancy Key"
        not in successful.columns
    ):

        return set()

    return {

        str(
            value
        ).strip()

        for value
        in successful[
            "Vacancy Key"
        ].dropna()

        if str(
            value
        ).strip()
    }


def recently_contacted_hr_emails():

    successful = (
        successful_live_log()
    )

    if successful.empty:

        return set()

    timestamps = pd.to_datetime(

        successful[
            "Timestamp"
        ],

        errors="coerce"
    )

    cutoff = (

        pd.Timestamp.now()

        - pd.Timedelta(
            days=
                HR_CONTACT_COOLDOWN_DAYS
        )
    )

    recent = successful[
        timestamps
        >= cutoff
    ]

    return {

        normalise_email(
            email
        )

        for email
        in recent[
            "Intended HR"
        ].dropna()
    }


# ============================================================
# 30. SELECT JOBS
# ============================================================

def choose_test_job(
    dataframe
):

    if dataframe.empty:
        return None

    strong = dataframe[
        dataframe[
            "Evidence Strength"
        ]
        == "STRONG"
    ]

    if not strong.empty:

        return strong.iloc[
            0
        ]

    return dataframe.iloc[
        0
    ]


def choose_live_jobs(
    dataframe
):

    if dataframe.empty:

        return (
            dataframe.copy(),
            0
        )

    candidates = dataframe[

        dataframe[
            "Evidence Strength"
        ].isin(
            LIVE_ALLOWED_EVIDENCE
        )

    ].copy()

    candidates = candidates[

        candidates[
            "Match Score"
        ]
        >= MIN_LIVE_MATCH_SCORE
    ]

    candidates = sort_results(
        candidates
    )

    candidates[
        "_Vacancy Key"
    ] = candidates.apply(
        build_vacancy_key,
        axis=1
    )

    sent_keys = (
        successfully_sent_vacancy_keys()
    )

    if sent_keys:

        candidates = candidates[

            ~candidates[
                "_Vacancy Key"
            ].isin(
                sent_keys
            )
        ]

    recent_hrs = (
        recently_contacted_hr_emails()
    )

    if recent_hrs:

        candidates = candidates[

            ~candidates[
                "Email"
            ]
            .str.lower()
            .isin(
                recent_hrs
            )
        ]

    if ONE_EMAIL_PER_UNIQUE_HR:

        candidates = (
            candidates
            .drop_duplicates(
                subset=[
                    "Email"
                ],
                keep="first"
            )
        )

    total_before_cap = len(
        candidates
    )

    selected = candidates.head(
        MAX_EMAILS_PER_RUN
    )

    return (
        selected,
        total_before_cap
    )


# ============================================================
# 31. TEST SEND
# ============================================================

def run_test_send(
    dataframe,
    gmail_service
):

    if not MY_EMAIL:

        raise ValueError(
            "Sender email is not configured."
        )

    if not TEST_EMAIL:

        raise ValueError(
            "Test email is not configured."
        )

    # TEST is deliberately allowed only to
    # the sender's own account.

    if (
        normalise_email(
            TEST_EMAIL
        )
        !=
        normalise_email(
            MY_EMAIL
        )
    ):

        raise ValueError(

            "\nTEST SAFETY CHECK FAILED.\n"

            "The test recipient must "
            "match the sender email."
        )

    selected_job = (
        choose_test_job(
            dataframe
        )
    )

    if selected_job is None:

        print(
            "\nNo job available "
            "for test email."
        )

        return

    print(
        "\nTEST MODE"
    )

    print(
        "Actual recipient:",
        TEST_EMAIL
    )

    print(
        "Intended recruiter:",
        selected_job[
            "Email"
        ]
    )

    print(
        "Role:",
        selected_job[
            "Role"
        ]
    )

    message = (
        create_email_message(

            selected_job,

            recipient=
                TEST_EMAIL
        )
    )

    try:

        result = gmail_api_send(
            gmail_service,
            message
        )

        message_id = result.get(
            "id",
            ""
        )

        append_sent_log(

            mode="TEST",

            recipient=
                TEST_EMAIL,

            selected_job=
                selected_job,

            status=
                "TEST_SENT",

            gmail_message_id=
                message_id
        )

        print(
            "\nTEST EMAIL "
            "SENT SUCCESSFULLY"
        )

        print(
            "No recruiter email "
            "was sent."
        )

    except Exception as error:

        append_sent_log(

            mode="TEST",

            recipient=
                TEST_EMAIL,

            selected_job=
                selected_job,

            status=
                "FAILED",

            error=
                str(
                    error
                )
        )

        raise


# ============================================================
# 32. LIVE SEND
# ============================================================

def run_live_send(
    dataframe,
    gmail_service
):

    if not LIVE_SEND_UNLOCKED:

        raise RuntimeError(

            "\nLIVE sending is locked.\n"

            "Explicit local unlock "
            "is required."
        )

    (
        candidates,
        total_before_cap
    ) = choose_live_jobs(
        dataframe
    )

    if candidates.empty:

        print(
            "\nNo eligible recruiter "
            "applications to send."
        )

        return {
            "attempted": 0,
            "successful": 0,
            "failed": 0,
            "remaining": 0,
            "queue_fully_completed":
                True
        }

    print(
        "\n"
        + "=" * 76
    )

    print(
        "LIVE RECRUITER EMAIL QUEUE"
    )

    print(
        "=" * 76
    )

    print(
        "Eligible before run cap:",
        total_before_cap
    )

    print(
        "Sending this run:",
        len(
            candidates
        )
    )

    success_count = 0

    failed_count = 0

    attempted_count = 0

    stopped_early = False

    for (
        position,
        (_, selected_job)
    ) in enumerate(
        candidates.iterrows(),
        start=1
    ):

        recipient = normalise_email(
            selected_job[
                "Email"
            ]
        )

        attempted_count += 1

        print(
            f"\n[{position}/"
            f"{len(candidates)}] "
            f"{selected_job['Role']}"
        )

        message = (
            create_email_message(

                selected_job,

                recipient=
                    recipient
            )
        )

        try:

            result = gmail_api_send(
                gmail_service,
                message
            )

            message_id = (
                result.get(
                    "id",
                    ""
                )
            )

            append_sent_log(

                mode="LIVE",

                recipient=
                    recipient,

                selected_job=
                    selected_job,

                status=
                    "SENT",

                gmail_message_id=
                    message_id
            )

            success_count += 1

            print(
                "Sent successfully."
            )

        except Exception as error:

            failed_count += 1

            append_sent_log(

                mode="LIVE",

                recipient=
                    recipient,

                selected_job=
                    selected_job,

                status=
                    "FAILED",

                error=
                    str(
                        error
                    )
            )

            print(
                "Send failed:",
                error
            )

            status = None

            if isinstance(
                error,
                HttpError
            ):

                status = getattr(
                    error.resp,
                    "status",
                    None
                )

            if status in {
                403,
                429
            }:

                print(
                    "\nGmail sending limit "
                    "or rate response detected."
                )

                print(
                    "Stopping safely."
                )

                stopped_early = True

                break

        if (
            position
            < len(
                candidates
            )
        ):

            time.sleep(
                DELAY_BETWEEN_EMAILS_SECONDS
            )

    (
        _,
        remaining
    ) = choose_live_jobs(
        dataframe
    )

    queue_complete = (

        remaining == 0

        and
        failed_count == 0

        and
        not stopped_early
    )

    print(
        "\n"
        + "=" * 76
    )

    print(
        "LIVE RUN COMPLETE"
    )

    print(
        "Attempted:",
        attempted_count
    )

    print(
        "Successful:",
        success_count
    )

    print(
        "Failed:",
        failed_count
    )

    print(
        "Remaining:",
        remaining
    )

    print(
        "=" * 76
    )

    return {

        "attempted":
            attempted_count,

        "successful":
            success_count,

        "failed":
            failed_count,

        "remaining":
            remaining,

        "queue_fully_completed":
            queue_complete
    }


# ============================================================
# 33. CONFIG VALIDATION
# ============================================================

def validate_runtime_config():

    if not MY_NAME:

        raise ValueError(
            "HR_AUTOMATION_NAME "
            "is required."
        )

    if not MY_EMAIL:

        raise ValueError(

            "\nHR_AUTOMATION_EMAIL "
            "is not configured locally."
        )

    if "@" not in MY_EMAIL:

        raise ValueError(
            "Sender email does not "
            "look valid."
        )

    if MY_EXPERIENCE < 0:

        raise ValueError(
            "Experience cannot "
            "be negative."
        )

    if MODE not in {
        "TEST",
        "LIVE"
    }:

        raise ValueError(
            'MODE must be TEST or LIVE.'
        )

    if MODE == "TEST":

        if (
            normalise_email(
                TEST_EMAIL
            )
            !=
            normalise_email(
                MY_EMAIL
            )
        ):

            raise ValueError(

                "\nTEST recipient must "
                "match sender account."
            )

    if (
        MODE == "LIVE"
        and
        not LIVE_SEND_UNLOCKED
    ):

        raise RuntimeError(

            "\nLIVE MODE IS LOCKED.\n"

            "Explicit local unlock "
            "is required."
        )


# ============================================================
# 34. MAIN WORKFLOW
# ============================================================

def run():

    validate_runtime_config()

    gmail_service = (
        get_gmail_service()
    )

    initial_backlog_phase = (
        not backlog_is_completed()
    )

    if MODE == "TEST":

        gmail_limit = (
            MAX_TEST_GMAIL_MESSAGES
        )

    else:

        gmail_limit = (
            MAX_LIVE_GMAIL_MESSAGES
        )

    (
        fetched_pdfs,
        source_messages
    ) = fetch_job_source_emails(

        gmail_service,

        max_messages=
            gmail_limit,

        initial_backlog_phase=
            initial_backlog_phase
    )

    selected_files = (
        collect_job_files(
            fetched_pdfs
        )
    )

    if not selected_files:

        print(
            "\nNo job files "
            "available to analyse."
        )

        return

    dataframe = scan_all_jobs(
        selected_files
    )

    if MODE == "TEST":

        if dataframe.empty:

            print(
                "\nNo test email sent "
                "because no suitable job "
                "was found."
            )

            return

        run_test_send(
            dataframe,
            gmail_service
        )

        return

    # --------------------------------------------------------
    # LIVE MODE
    # --------------------------------------------------------

    if dataframe.empty:

        print(
            "\nNo suitable jobs "
            "were found."
        )

        # Source scanning completed successfully,
        # even though nothing was eligible.

        for source in source_messages:

            append_processed_gmail_message(

                source[
                    "message_id"
                ],

                source[
                    "subject"
                ],

                source[
                    "date"
                ],

                source[
                    "internal_date"
                ]
            )

        if initial_backlog_phase:

            mark_backlog_completed()

        return

    result = run_live_send(
        dataframe,
        gmail_service
    )

    if result[
        "queue_fully_completed"
    ]:

        for source in source_messages:

            append_processed_gmail_message(

                source[
                    "message_id"
                ],

                source[
                    "subject"
                ],

                source[
                    "date"
                ],

                source[
                    "internal_date"
                ]
            )

        if initial_backlog_phase:

            mark_backlog_completed()

        print(
            "\nSource state updated "
            "successfully."
        )

    else:

        print(
            "\nSource state was NOT "
            "advanced because the "
            "LIVE queue did not "
            "complete successfully."
        )


# ============================================================
# 35. START
# ============================================================

if __name__ == "__main__":

    run()
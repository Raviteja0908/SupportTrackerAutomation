import re


def resolve_incident_type(category_type: str, subject: str) -> str:
    cat = (category_type or "").lower()
    subj = (subject or "").lower()

    if "deployment/property config" in cat:
        if "deployed" in subj or "deployment" in subj:
            if "prd" in subj or "prod" in subj or "uat" in subj:
                return "SR-Deployment"
        return "SR-Property configuration"

    if "file process/data check" in cat:
        return "SR-File process"

    if "daily/weekly task" in cat:
        if "daily task hyparchive" in subj or "severes warnings and errors in eai/es aws/es symphony" in subj:
            return "SR-Daily task"
        return "SR-Weekly task"

    if "exception" in cat:
        # INC-Failed SOP is only for failed/skipped exception rows with
        # interface codes ending in 001 (e.g., VMI001, CS001, DE001).
        if re.search(r"\b[a-z]{2,}\d*001\b", subj) and ("failed" in subj or "skipped" in subj):
            return "INC-Failed SOP"
        return "INC-Exceptions"

    if "maintenance/patch" in cat:
        return "SR-Patch"

    return ""

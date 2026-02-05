def resolve_environment(text: str) -> str:
    if not text:
        return ""
    s = text.lower()

    force_prod_phrases = [
        "daily task hyparchive",
        "severes warnings and errors in eai/es aws/es symphony",
        "files from es to grp",
    ]
    if any(p in s for p in force_prod_phrases):
        return "PROD"

    # Non-prod markers map to UAT
    if "non-prod" in s or "non prod" in s or "test" in s:
        return "UAT"

    # Environment markers
    if "prod" in s or "prd" in s or "production" in s:
        return "PROD"
    if "uat" in s:
        return "UAT"
    if "qa" in s:
        return "QA"
    if "dev" in s or "development" in s:
        return "DEV"

    # Interface environment codes
    if "fcp" in s:
        return "PROD"
    if "fct" in s:
        return "UAT"
    if "fcq" in s:
        return "QA"
    if "fcd" in s:
        return "DEV"
    return ""

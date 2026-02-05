def resolve_interface_code(description: str) -> str:
    if not description:
        return ""
    marker = "-->"
    idx = description.find(marker)
    if idx >= 0:
        return description[:idx].strip()
    return description.strip()

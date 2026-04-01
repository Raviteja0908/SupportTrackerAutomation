def resolve_service_request(category_type: str) -> str:
    if not category_type:
        return "ServiceRequest"
    s = category_type.lower()
    if "exception internal" in s or "exception external" in s:
        return "Incident"
    return "ServiceRequest"

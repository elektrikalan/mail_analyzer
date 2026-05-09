from app.services.keyword_service import analyze_keywords
from app.services.regex_service import analyze_regex


def analyze_mail_content(body: str, custom_keywords: list = None, custom_regex: dict = None):
    """
    Analyzes mail body based on default rules and optional custom parameters.
    """
    body = body or ""
    risk_score = 0
    matched_rules = []

    default_keywords = ["password", "vpn", "admin", "confidential", "token"]
    default_regex = {
        "ipv4": r"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b",
        "email": r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
        "jwt": r"eyJ[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+"
    }

    keywords = list(set(default_keywords + (custom_keywords or [])))
    regex_rules = {**default_regex, **(custom_regex or {})}

    keyword_matches = analyze_keywords(body, keywords)
    regex_matches = analyze_regex(body, regex_rules)

    for kw in keyword_matches:
        risk_score += 10
        matched_rules.append(f"keyword:{kw}")

    for regex_match in regex_matches:
        risk_score += 20
        matched_rules.append(f"regex:{regex_match['rule']}")

    return {
        "risk_score": min(risk_score, 100),
        "matched_rules": matched_rules
    }
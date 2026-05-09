import re

REGEX_RULES = {
    "ipv4": r"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b",
    "email": r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
    "jwt": r"eyJ[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+"
}

def analyze_regex(text: str, regex_rules: dict = None):
    """Return regex matches for the provided text using the given rule set."""
    regex_rules = regex_rules or REGEX_RULES
    text = text or ""
    findings = []

    for name, pattern in regex_rules.items():
        try:
            matches = re.findall(pattern, text, flags=re.IGNORECASE)
            if matches:
                findings.append({
                    "rule": name,
                    "matches": matches
                })
        except re.error:
            continue

    return findings
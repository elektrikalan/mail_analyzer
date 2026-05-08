import re

REGEX_RULES = {
    "ipv4": r"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b",
    "email": r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
    "jwt": r"eyJ[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+"
}

def analyze_regex(text: str):

    findings = []

    for name, pattern in REGEX_RULES.items():

        matches = re.findall(pattern, text)

        if matches:
            findings.append({
                "rule": name,
                "matches": matches
            })

    return findings
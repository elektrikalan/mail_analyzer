KEYWORDS = [
    "password",
    "vpn",
    "admin",
    "confidential",
    "token"
]

def analyze_keywords(text: str):

    findings = []

    lower_text = text.lower()

    for keyword in KEYWORDS:

        if keyword in lower_text:
            findings.append(keyword)

    return findings
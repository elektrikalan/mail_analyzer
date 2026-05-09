KEYWORDS = [
    "password",
    "vpn",
    "admin",
    "confidential",
    "token"
]

def analyze_keywords(text: str, keywords: list = None):
    """Return a list of matched keywords from the provided text."""
    keywords = keywords or KEYWORDS
    lower_text = (text or "").lower()
    findings = []

    for keyword in keywords:
        if keyword and keyword.lower() in lower_text:
            findings.append(keyword)

    return findings
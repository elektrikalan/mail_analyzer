import re

def analyze_mail_content(body: str, custom_keywords: list = None, custom_regex: dict = None):
    """
    Analyzes mail body based on default rules and optional custom parameters.
    """
    risk_score = 0
    matched_rules = []
    
    # Default Rules
    default_keywords = ["password", "vpn", "admin", "confidential", "token"]
    default_regex = {
        "ipv4": r"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b",
        "email": r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
        "jwt": r"eyJ[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+"
    }
    
    # Merge rules
    keywords = list(set(default_keywords + (custom_keywords or [])))
    regex_rules = {**default_regex, **(custom_regex or {})}
    
    # Keyword analysis
    lower_body = body.lower()
    for kw in keywords:
        if kw.lower() in lower_body:
            risk_score += 10
            matched_rules.append(f"keyword:{kw}")
            
    # Regex analysis
    for name, pattern in regex_rules.items():
        try:
            matches = re.findall(pattern, body)
            if matches:
                risk_score += 20
                matched_rules.append(f"regex:{name}")
        except:
            continue
            
    return {
        "risk_score": min(risk_score, 100), # Cap at 100
        "matched_rules": matched_rules
    }
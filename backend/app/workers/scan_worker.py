from sqlalchemy.orm import Session
from app.models.mail import Mail
from app.services.analyzer_service import analyze_mail_content

def process_mail(db: Session, mail_data: dict):
    """
    Saves mail data to DB. If analysis is not provided, it runs a default analysis.
    """
    risk_score = mail_data.get("risk_score")
    matched_rules = mail_data.get("matched_rules")
    
    # If not already analyzed, run default analysis
    if risk_score is None:
        analysis = analyze_mail_content(mail_data["body"])
        risk_score = analysis["risk_score"]
        matched_rules = analysis["matched_rules"]

    db_mail = Mail(
        message_id=mail_data["message_id"],
        subject=mail_data["subject"],
        sender=mail_data["sender"],
        recipients=[], # Can be expanded if needed
        body=mail_data["body"],
        folder_path=mail_data["folder_path"],
        pst_source=mail_data.get("store", "Outlook"), # Use store name (e.g. PST filename)
        attachments=mail_data["attachments"],
        risk_score=risk_score,
        matched_rules=matched_rules,
        received_at=mail_data["received_at"]
    )

    # Use merge to avoid unique constraint errors on message_id if re-scanning
    db.merge(db_mail)
    db.commit()
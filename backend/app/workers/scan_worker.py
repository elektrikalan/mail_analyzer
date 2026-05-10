from sqlalchemy.orm import Session
from ..models.mail import Mail
from ..services.analyzer_service import analyze_mail_content


def process_mail(db: Session, mail_data: dict):
    risk_score = mail_data.get("risk_score")
    matched_rules = mail_data.get("matched_rules")

    if risk_score is None:
        analysis = analyze_mail_content(mail_data.get("body", ""))
        risk_score = analysis["risk_score"]
        matched_rules = analysis["matched_rules"]

    db_mail = Mail(
        message_id=mail_data.get("message_id", "N/A"),
        subject=mail_data.get("subject", ""),
        sender=mail_data.get("sender", ""),
        recipients=[],
        body=mail_data.get("body", ""),
        folder_path=mail_data.get("folder_path", ""),
        pst_source=mail_data.get("store", "Outlook"),
        attachments=mail_data.get("attachments", []),
        risk_score=risk_score,
        matched_rules=matched_rules,
        received_at=mail_data.get("received_at"),
    )
    db.merge(db_mail)
    db.commit()
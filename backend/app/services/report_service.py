import pandas as pd
from sqlalchemy.orm import Session
from app.models.mail import Mail
import os
from datetime import datetime

def generate_excel_report(db: Session):
    # Fetch all mails from DB
    mails = db.query(Mail).all()
    
    if not mails:
        return None
        
    data = []
    for m in mails:
        data.append({
            "ID": m.id,
            "Subject": m.subject,
            "Sender": m.sender,
            "Received At": m.received_at,
            "Folder": m.folder_path,
            "Risk Score": m.risk_score,
            "Matched Rules": ", ".join(m.matched_rules) if m.matched_rules else ""
        })
        
    df = pd.DataFrame(data)
    
    # Ensure reports directory exists
    reports_dir = os.path.join(os.getcwd(), "reports")
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)
        
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"mail_analysis_{timestamp}.xlsx"
    filepath = os.path.join(reports_dir, filename)
    
    df.to_excel(filepath, index=False)
    
    return filepath

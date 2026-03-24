import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Email Automation", layout="centered")

st.title("📧 Weekly Class Email Automation")

# =========================
# 📂 Upload Excel
# =========================
uploaded_file = st.file_uploader("Upload Student Excel File", type=["xlsx"])

# =========================
# 🎯 Select Batch
# =========================
selected_batch = st.selectbox("Select Batch", ["10:00 AM", "11:00 AM", "12:00 PM"])

# =========================
# 🗓️ Week Input
# =========================
week_range = st.text_input("Enter Week Range (e.g. 12 March - 19 March)")

# =========================
# 🧠 Dynamic Topics Input
# =========================
st.subheader("📚 Enter Completed Topics")
completed_topics = st.text_area(
    "Write completed topics (one per line)",
    height=150,
    placeholder="Example:\nOOP\nClasses\nInheritance"
)

st.subheader("🚀 Enter Upcoming Topics")
upcoming_topics = st.text_area(
    "Write upcoming topics (one per line)",
    height=150,
    placeholder="Example:\nPLR\nClassification\nMetrics"
)

# =========================
# 📎 Attachment
# =========================
attachment_file = st.file_uploader("Upload Attachment (Optional)", type=["pdf", "docx", "pptx"])

# =========================
# 🔐 Email Credentials
# =========================
your_email = st.text_input("Your Email")
your_password = st.text_input("App Password", type="password")

# =========================
# ✉️ Create Email
# =========================
def create_email(name):
    completed_list = "\n".join([f"• {t}" for t in completed_topics.split("\n") if t.strip()])
    upcoming_list = "\n".join([f"• {t}" for t in upcoming_topics.split("\n") if t.strip()])

    return f"""
Dear {name},

We hope that you are learning well and enjoying your classes. This email is to bring to your notice that in the last week, we have completed the following topics:-

{completed_list}

And in the coming week, we are expecting to cover the following topics:-

{upcoming_list}

We hope that you are satisfied with what you have been taught till now. If you are facing any issues, feel free to let us know.

Looking forward to moving ahead positively.

A no reply to this email will be considered as "Acknowledged" from your side.

Thanks & Regards  

Aftab Alam  
Data Science & AI/ML Trainer
"""

# =========================
# 🚀 SEND EMAIL BUTTON
# =========================
if st.button("🚀 Send Emails"):

    if uploaded_file is None:
        st.error("Please upload Excel file")
    elif not completed_topics or not upcoming_topics:
        st.error("Please enter topics")
    elif not your_email or not your_password:
        st.error("Enter email credentials")
    else:
        df = pd.read_excel(uploaded_file)

        # ✅ SUBJECT (NO BATCH NAME)
        subject = f"WEEKLY CLASS SUMMARY {week_range}"

        filtered_df = df[df["Batch"] == selected_batch]

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(your_email, your_password)

            for _, row in filtered_df.iterrows():
                msg = MIMEMultipart()
                msg['From'] = your_email
                msg['To'] = row['Email']
                msg['Subject'] = subject

                body = create_email(row['Name'])
                msg.attach(MIMEText(body, 'plain'))

                # 📎 Attachment
                if attachment_file:
                    part = MIMEApplication(attachment_file.read(), Name=attachment_file.name)
                    part['Content-Disposition'] = f'attachment; filename="{attachment_file.name}"'
                    msg.attach(part)

                server.send_message(msg)

            server.quit()

            st.success(f"✅ Emails sent to {selected_batch} batch!")

        except Exception as e:
            st.error(f"❌ Error: {e}")
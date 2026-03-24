import streamlit as st
import pandas as pd
import smtplib
import uuid
from email.utils import formatdate, make_msgid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

st.set_page_config(page_title="Email Automation", layout="centered")

st.title("📧 Weekly Class Email Automation")

# =========================
# 📂 Upload Excel
# =========================
uploaded_file = st.file_uploader("Upload Student Excel File", type=["xlsx"])

df = None
selected_batch = None




if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    if "Batch" not in df.columns:
        st.error("❌ 'Batch' column not found in Excel")
    else:
        # 🔥 FIX: Normalize Batch column
        def format_batch(x):
            try:
                return pd.to_datetime(str(x)).strftime("%I:%M %p")
            except:
                return str(x).strip()

        df["Batch"] = df["Batch"].apply(format_batch)

        # ✅ Now safe to sort
        batch_list = sorted(df["Batch"].dropna().unique().tolist())

        selected_batch = st.selectbox("Select Batch", batch_list)
else:
    st.info("Please upload Excel file")

# =========================
# 🗓️ Week Input
# =========================
week_range = st.text_input("Enter Week Range (e.g. 12 March - 19 March)")

# =========================
# 🧠 Topics Input
# =========================
st.subheader("📚 Completed Topics")
completed_topics = st.text_area("Completed Topics", height=150, key="completed")

st.subheader("🚀 Upcoming Topics")
upcoming_topics = st.text_area("Upcoming Topics", height=150, key="upcoming")

# =========================
# 📎 Attachment
# =========================
attachment_file = st.file_uploader("Upload Attachment (Optional)", type=["pdf", "docx", "pptx"])

# =========================
# 📩 CC OPTION
# =========================
cc_input = st.text_input("Enter CC Emails (comma separated)")
cc_list = [email.strip() for email in cc_input.split(",") if email.strip()]

# =========================
# 🔐 Email Credentials
# =========================
your_email = st.text_input("Your Email")
your_password = st.text_input("App Password", type="password")

# =========================
# ✉️ Email Function
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

We hope that you are satisfied with what you have been taught till now.

Looking forward to moving ahead positively.

A no reply to this email will be considered as "Acknowledged" from your side.

Thanks & Regards  

Aftab Alam  
Data Science & AI/ML Trainer
"""

# =========================
# 👀 PREVIEW SYSTEM
# =========================
st.subheader("👀 Preview Email")

if df is not None and selected_batch:
    filtered_df = df[df["Batch"] == selected_batch]

    if not filtered_df.empty:
        student_name = st.selectbox(
            "Select Student for Preview",
            filtered_df["Name"].tolist()
        )

        if st.button("👀 Generate Preview"):
            if not completed_topics or not upcoming_topics:
                st.warning("Please enter topics first")
            else:
                subject_preview = f"{student_name} | Weekly Class Summary ({week_range})"

                st.markdown(f"**Subject:** {subject_preview}")
                st.markdown(f"**To:** {student_name}")
                if cc_list:
                    st.markdown(f"**CC:** {', '.join(cc_list)}")

                preview_body = create_email(student_name)
                st.text_area("Email Preview", preview_body, height=400)
    else:
        st.warning("No students in this batch")

# =========================
# 🚀 SEND EMAILS
# =========================
if st.button("🚀 Send Emails"):

    if df is None:
        st.error("Upload Excel file")
    elif not selected_batch:
        st.error("Select batch")
    elif not completed_topics or not upcoming_topics:
        st.error("Enter topics")
    elif not your_email or not your_password:
        st.error("Enter email credentials")
    else:
        filtered_df = df[df["Batch"] == selected_batch]

        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(your_email, your_password)

            count = 0

            for _, row in filtered_df.iterrows():

                student_email = str(row["Email"]).strip()
                student_name = row["Name"]

                if not student_email or student_email.lower() == "nan":
                    continue

                # ✅ UNIQUE SUBJECT PER STUDENT (FINAL FIX)
                subject = f"{student_name} | Weekly Class Summary ({week_range})"

                msg = MIMEMultipart()
                msg['From'] = your_email
                msg['To'] = student_email
                msg['Subject'] = subject

                # ✅ CC
                if cc_list:
                    msg['Cc'] = ", ".join(cc_list)

                # 🔥 HIDDEN UNIQUE HEADERS (NO THREADING)
                msg['Message-ID'] = make_msgid()
                msg['Date'] = formatdate(localtime=True)
                msg['X-Unique-ID'] = str(uuid.uuid4())
                msg['X-Mailer'] = "Aftab-AI-Mailer"

                body = create_email(student_name)
                msg.attach(MIMEText(body, 'plain'))

                # 📎 Attachment fix
                if attachment_file:
                    attachment_file.seek(0)
                    part = MIMEApplication(attachment_file.read(), Name=attachment_file.name)
                    part['Content-Disposition'] = f'attachment; filename="{attachment_file.name}"'
                    msg.attach(part)

                recipients = [student_email] + cc_list
                server.sendmail(your_email, recipients, msg.as_string())

                count += 1

            server.quit()

            st.success(f"✅ {count} emails sent successfully!")

        except Exception as e:
            st.error(f"❌ Error: {e}")

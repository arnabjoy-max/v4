# 🤖 FANUC FIR Report Generator

A web application for generating FANUC Field Installation Reports (FIR) with a professional UI, automatic filename generation, Excel population, and email delivery.

---

## ✨ Features

- **Smart Filename Generation** — Automatically creates filenames in FANUC convention: `FIR-XXXXXX-YYMMDD-EmpNo(Type).xlsm`
- **Full Form Input** — All fields from the original FIR Excel template
- **7-Day Work Log** — Enter daily hours (Straight, OT, Double Time, Travel, Working, Wait)
- **Excel Population** — Fills the original `.xlsm` template and lets you download it
- **Email Delivery** — Send the FIR directly to multiple recipients via SMTP
- **Beautiful FANUC-branded UI** — Dark industrial theme with yellow accents

---

## 🚀 Deploy to Streamlit Cloud (FREE — No App Store Required)

### Step 1: Fork / Upload to GitHub

1. Create a free account at [github.com](https://github.com)
2. Create a new repository (e.g. `fanuc-fir-app`)
3. Upload these files to the repository:
   - `app.py`
   - `requirements.txt`
   - `FIR_template.xlsm`
   - `README.md`
   - `.streamlit/config.toml`

### Step 2: Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **"New app"**
4. Select your repository, branch (`main`), and set **Main file path** to `app.py`
5. Click **"Deploy!"**

Your app will be live at a URL like:
```
https://your-username-fanuc-fir-app-app-xxxxxx.streamlit.app
```

**That's it — completely free, no credit card, no app store!**

---

## 💻 Run Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

Open your browser at `http://localhost:8501`

---

## 📂 File Structure

```
fanuc-fir-app/
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── FIR_template.xlsm       # Original FANUC FIR Excel template (keep_vba)
├── README.md               # This file
└── .streamlit/
    └── config.toml         # Streamlit theme configuration
```

---

## 📧 Email Configuration

The app supports sending FIR reports via email. Supported providers:

| Provider | SMTP Host | Port |
|----------|-----------|------|
| Gmail | smtp.gmail.com | 587 |
| Outlook | smtp.office365.com | 587 |
| Yahoo | smtp.mail.yahoo.com | 587 |

**For Gmail:** Use an [App Password](https://support.google.com/accounts/answer/185833) (not your regular password). Go to: Google Account → Security → 2-Step Verification → App passwords.

---

## 📋 FIR Filename Convention

```
FIR-[ProjectNum]-[YYMMDD]-[EmpNo]([Type]).xlsm

Example:
FIR-3011187-260216-1112(O).xlsm
         │       │      │   └── Report type: (I)=Installation, (P)=Preliminary, (O)=Other
         │       │      └────── Employee number
         │       └───────────── Date: YYMMDD
         └───────────────────── Project number
```

---

## 🔧 Hour Types

| Code | Meaning |
|------|---------|
| S | Straight time |
| OT | Overtime (typically after 8 straight hours) |
| DT | Double Time (weekends/holidays) |
| TT | Travel Time |
| W | Working |
| Wait | Wait time |

---

## 📝 License

For internal FANUC America Corporation use. Not for redistribution.

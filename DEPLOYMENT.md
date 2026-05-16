# Deployment Instructions

## Local Setup
```bash
pip install -r requirements.txt
echo "OPENAI_API_KEY=your_key_here" > .env
streamlit run app.py
```

## Streamlit Cloud Deployment

### Step 1: Push code to GitHub
```bash
git add .
git commit -m "Update privacy settings and remove API key input"
git push
```

### Step 2: Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io
2. Sign in with GitHub
3. Click "New app"
4. Select your repo: `jahnavi-janu/table-extractor-excel`
5. Branch: `main`
6. Main file path: `app.py`

### Step 3: Add Secrets (CRITICAL)
In Streamlit Cloud dashboard:
1. Go to your app settings ⚙️
2. Click "Secrets"
3. Add:
   ```
   OPENAI_API_KEY = "sk-proj-your-actual-key-here"
   ```
4. Save

**Do NOT share or commit actual API keys to GitHub!**

### Step 4: Redeploy
- Click "Rerun" or wait for auto-deploy

---

## Data Security

**What gets sent to OpenAI:**
- Only extracted table images (PNG format)
- NOT the original PDF/Word file

**What stays local:**
- File uploads (in memory, auto-deleted)
- Extracted tables (in memory, auto-deleted based on settings)
- User's downloaded Excel file (user's device)

**Privacy options enabled by default:**
- PII masking (masks names, emails, phone numbers, IDs)
- Auto-delete after download

---

## Troubleshooting

### "OpenAI API key not configured"
- Check Streamlit Cloud secrets are set correctly
- Wait 1-2 minutes for deployment to refresh

### "No response from API"
- Verify API key is valid
- Check OpenAI account has available credits
- Check internet connectivity

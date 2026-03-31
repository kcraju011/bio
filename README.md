# BioAttend – Setup Guide

## Files
| File | Purpose |
|------|---------|
| `index.html` | Frontend (deploy to Vercel) |
| `Code.gs` | Google Apps Script backend (paste into Apps Script) |

---

## Step 1 – Google Apps Script Setup

1. Open your Google Spreadsheet (or create a new one).
2. Go to **Extensions → Apps Script**.
3. Delete any existing code and paste the contents of **`Code.gs`**.
4. Click 💾 **Save**.
5. Click **Deploy → New Deployment**:
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy** → copy the **Web App URL** (looks like `https://script.google.com/macros/s/AKfy.../exec`).

> ⚠️ Every time you edit `Code.gs`, create a **new deployment** and copy the new URL.

---

## Step 2 – Connect Frontend to Backend

Open `index.html` in VS Code and find this line near the top of the `<script>` tag:

```js
const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
```

Replace `YOUR_APPS_SCRIPT_URL_HERE` with the URL you copied in Step 1.

---

## Step 3 – Deploy to Vercel

```bash
# Install Vercel CLI (once)
npm install -g vercel

# Inside the folder containing index.html
vercel

# Follow prompts – choose "Other" framework, root = ./
# Vercel will give you a URL like https://bioattend.vercel.app
```

---

## Step 4 – Test Biometric on Phone

1. Open the Vercel URL in Chrome/Safari on your phone.
2. **Create Account** → fill details → tap **Create Account**.
3. Once the account is created, tap **Register Fingerprint / Face ID**.
   - Your phone will prompt for fingerprint/Face ID.
   - This uses the **WebAuthn API** built into modern browsers.
4. Next time, go to **Sign In**, enter your email, tap **Sign in with Fingerprint / Face ID**.

> 📱 WebAuthn biometric requires **HTTPS** – Vercel provides this automatically.
> 🖥 Localhost will NOT work for biometric; use the Vercel URL on your phone.

---

## Spreadsheet Structure (auto-created)

### `Users` sheet
| UserID | FullName | Email | PasswordHash | DOB | Mobile | Institution | Department | MarkFromAnywhere | BiometricCredentialId | CreatedAt |

### `Attendance` sheet
| AttendanceID | UserID | FullName | Email | Timestamp | Date | Time | Method |

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "Network error – check URL config" | Make sure you pasted the correct Apps Script URL |
| Biometric button does nothing | Open browser console; WebAuthn needs HTTPS & a supported browser |
| "No biometric registered" | Register biometric first from the Create Account tab |
| Attendance not appearing | Check the `Attendance` sheet tab in your Google Spreadsheet |

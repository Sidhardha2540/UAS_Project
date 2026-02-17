# Outlook Mail Reader & BEO PDF Processor

This project fetches Inbox emails from Outlook, finds PDF attachments, validates them with an AI agent (signed Hospitality form + BEO), and saves valid BEO PDFs into a folder structure—either on your machine or in **OneDrive** so you can access them from any device.

---

## What you need before starting

- **Python 3.9+** installed on your computer.
- A **Microsoft account** (personal Outlook.com or work/school) that has the emails you want to process.
- An **OpenAI API key** (for BEO validation). Get one at [platform.openai.com](https://platform.openai.com).

**When do I need a work or school account?** For reading mail and saving BEO PDFs **on your computer** (no `SAVE_TO_ONEDRIVE`), a **personal Microsoft account** (e.g. Outlook.com) is enough. If you set `SAVE_TO_ONEDRIVE=true`, both personal and work/school accounts work—files are saved to that account’s OneDrive.

---

## Step 1: Register the app in Azure

This gives the script permission to read your mail and (optionally) write to OneDrive.

1. Open **[Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)** and sign in.
2. Click **New registration**.
3. Fill in:
   - **Name**: e.g. `UAS Outlook Reader`.
   - **Supported account types**:  
     - Choose **"Accounts in any organizational directory and personal Microsoft accounts"** if you might use both personal and work accounts.  
     - Or **"Personal Microsoft accounts only"** if you only use Outlook.com.
   - **Redirect URI**: leave blank.
4. Click **Register**.
5. On the app’s **Overview** page, copy the **Application (client) ID**. You will put this in `.env` as `CLIENT_ID`.
6. In the left menu, go to **API permissions**.
7. Click **Add a permission**.
8. Choose **Microsoft Graph** → **Delegated permissions**.
9. Add these permissions (search or scroll to find them):
   - **Mail.Read** — to read your Inbox.
   - **User.Read** — to sign you in.
   - **Files.ReadWrite.All** — only if you will save BEO folders to **OneDrive** (so the app can create folders and upload files).
10. Click **Add permissions**.
11. In the left menu, go to **Authentication**.
12. Click **Add a platform** → **Mobile and desktop applications**.
13. Leave the suggested redirect URI (e.g. `http://localhost`) and click **Configure**.

You’re done with Azure. Keep the **Application (client) ID** for Step 4.

---

## Step 2: Install the project

1. Open a terminal (or PowerShell) and go to the folder where you want the project, for example:
   ```bash
   cd C:\Users\YourName\Projects
   ```
2. Clone the repo (or unzip the project folder):
   ```bash
   git clone https://github.com/Sidhardha2540/UAS_Project.git
   cd UAS_Project
   ```
3. Create and activate a virtual environment:
   - **Windows:**
     ```bash
     python -m venv .venv
     .venv\Scripts\activate
     ```
   - **macOS/Linux:**
     ```bash
     python -m venv .venv
     source .venv/bin/activate
     ```
4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

You should see no errors. The project is ready to configure.

---

## Step 3: Configure your `.env` file

1. In the project root folder, create a file named `.env` (no filename before the dot, extension is `.env`).
2. Copy the contents of `.env.example` into `.env` (or create `.env` with the lines below).
3. Edit `.env` and set each value as follows.

| Variable | What to put | Required? |
|----------|-------------|-----------|
| **CLIENT_ID** | The **Application (client) ID** you copied from Azure in Step 1. | **Yes** |
| **OPENAI_API_KEY** | Your OpenAI API key (starts with `sk-...`). Needed for BEO validation. | **Yes** for BEO |
| **SAVE_TO_ONEDRIVE** | Set to `true` to save BEO folders to your **OneDrive**. Use `false` or leave empty to save **on your computer** only. | No (default: local) |
| **BEO_BASE_PATH** | Leave empty to use the default folder `beo_output` in the project. Or set a full path, e.g. `C:\BEO_Files`. Used **only when SAVE_TO_ONEDRIVE is not true**. | No |

**Example `.env` (saving to OneDrive):**

```env
CLIENT_ID=12345678-1234-1234-1234-123456789abc
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxxxxx
SAVE_TO_ONEDRIVE=true
BEO_BASE_PATH=
```

**Example `.env` (saving only on your computer):**

```env
CLIENT_ID=12345678-1234-1234-1234-123456789abc
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxxxxx
SAVE_TO_ONEDRIVE=false
BEO_BASE_PATH=
```

Save the file. Do not commit `.env` to git (it should be in `.gitignore`).

---

## Step 4: Run the script

### First time: sign in

1. In the same terminal (with the virtual environment activated), run:
   - To **only list** your Inbox (no BEO processing):
     ```bash
     python get_mails.py --list
     ```
   - To **run the full BEO pipeline** (process PDFs and save them):
     ```bash
     python get_mails.py
     ```
2. The script will print a **URL** and a **code**.
3. Open the URL in your browser.
4. Sign in with your Microsoft account (personal or work/school; if you set `SAVE_TO_ONEDRIVE=true`, files go to that account’s OneDrive).
5. Enter the code shown in the terminal and approve the requested permissions.
6. After that, the script continues: with `--list` it prints your messages; without it, it processes PDFs and saves them (locally or to OneDrive).

### Later runs

- Just run the same command again. The script uses cached sign-in; you won’t be asked to sign in unless you delete the token cache or it expires.

### What each command does

| Command | What it does |
|--------|----------------|
| `python get_mails.py --list` | Fetches your Inbox and prints each message (subject, from, date, preview). No PDF processing. |
| `python get_mails.py` | Fetches Inbox, finds PDF attachments, runs the BEO validation agent on each, and saves valid PDFs. If `SAVE_TO_ONEDRIVE=true`, creates the folder structure in your OneDrive and prints a “Saved to OneDrive:” link. Otherwise saves under `BEO_BASE_PATH` and prints “Saved:” with the file path. |

### Folder structure created

Valid BEO PDFs are saved in this structure:

- **In OneDrive:**  
  In your OneDrive root:  
  `Year` → `month` → `day` → `BEO_number - client name` → PDF file.
- **On your computer:**  
  Under `BEO_BASE_PATH` (or `beo_output`):  
  `Year/month/day/BEO_number - client name/document.pdf`.

Example: BEO date 01/01/2026, number 12345, client “Acme Corp” →  
`2026` → `1` → `1` → `12345 - Acme Corp` → PDF.

---

## Project structure

```
UAS_Project/
├── .env                 # Your secrets (CLIENT_ID, OPENAI_API_KEY, SAVE_TO_ONEDRIVE, etc.) – do not commit
├── .env.example         # Template for .env
├── README.md            # This file
├── beo_processor.py      # PDF text extraction, AI agent, folder logic, OneDrive upload
├── get_mails.py         # Main script: mail + attachments + BEO pipeline
├── requirements.txt     # Python dependencies
└── token_cache.bin      # Cached sign-in (created on first run) – do not commit
```

---

## Security

- Do **not** commit `.env` or `token_cache.bin` (they are in `.gitignore`).
- Keep your **CLIENT_ID** and **OPENAI_API_KEY** private.
- To sign in again from scratch (e.g. switch account), delete `token_cache.bin` and run the script again.

---

## Troubleshooting

| Problem | What to do |
|--------|------------|
| **CLIENT_ID not set** | Create a `.env` file in the project root and set `CLIENT_ID=` to your Azure Application (client) ID. |
| **OPENAI_API_KEY is not set** | Add `OPENAI_API_KEY=sk-...` to `.env` (required for BEO processing). |
| **AADSTS50059** / tenant error | In Azure, set “Supported account types” to include the account type you use (e.g. “Accounts in any organizational directory and personal Microsoft accounts”). |
| **Application not found** | Confirm the app in Azure supports the account type you’re signing in with (personal and/or work/school). |
| **Permission denied** (mail) | In Azure → API permissions, add **Mail.Read** and **User.Read** and grant admin consent if required. |
| **Permission denied** (OneDrive / files) | In Azure → API permissions, add **Files.ReadWrite.All** (delegated). Sign in with the account whose OneDrive you want to use. |
| **403 / Forbidden when uploading to OneDrive** | Ensure **Files.ReadWrite.All** is granted in Azure and that you consented. Use the same Microsoft account that owns the OneDrive you want to write to. |

If you follow the steps above in order (Azure → install → `.env` → run), the script should list mail and, when not using `--list`, process PDFs and save them locally or to OneDrive as configured.

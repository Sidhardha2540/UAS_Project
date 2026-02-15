# Outlook Mail Reader

A Python script that fetches Inbox emails from Outlook (jmovva25@outlook.com) using Microsoft Graph API and prints them to the terminal.

## Features

- Uses Microsoft Graph API to access Outlook mailbox
- Device Code flow for authentication (no web server required)
- Token caching so you only sign in once
- Prints subject, sender, date, and body preview for each message
- Handles pagination to fetch all Inbox messages

## Prerequisites

- Python 3.9+
- A Microsoft personal account (outlook.com)
- Azure App Registration with Microsoft Graph permissions

## Azure App Registration

1. Go to [Azure Portal - App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
3. Fill in:
   - **Name**: `UAS Outlook Reader` (or any name)
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts" or "Personal Microsoft accounts only"
   - **Redirect URI**: Leave blank
4. Click **Register**
5. Copy the **Application (client) ID** from the overview page
6. Go to **API permissions** → **Add a permission**
   - Select **Microsoft Graph** → **Delegated permissions**
   - Add: `Mail.Read`, `User.Read`
   - Click **Add permissions**
7. Go to **Authentication** → **Add a platform**
   - Select **Mobile and desktop applications**
   - Add redirect URI: `http://localhost` (or use the suggested one)
   - Click **Configure**

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Sidhardha2540/UAS_Project.git
   cd UAS_Project
   ```

2. Create a virtual environment (recommended):
   ```bash
   python -m venv .venv
   .venv\Scripts\activate   # Windows
   source .venv/bin/activate  # macOS/Linux
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Create a `.env` file in the project root:
   ```bash
   copy .env.example .env   # Windows
   cp .env.example .env    # macOS/Linux
   ```

5. Edit `.env` and add your Azure Application (client) ID:
   ```
   CLIENT_ID=your-actual-client-id-from-azure
   ```

## Usage

Run the script:

```bash
python get_mails.py
```

**First run**: The script prints a URL and code. Open the URL in a browser, sign in with jmovva25@outlook.com, enter the code, and authorize the app. The script then fetches and prints all Inbox messages.

**Later runs**: Uses cached tokens. No sign-in needed; mails are printed immediately.

## Output

Each message is printed with:
- Subject
- From (name and email)
- Received date
- Body preview

## Project Structure

```
UAS_Project/
├── .env                 # Your CLIENT_ID (create from .env.example, do not commit)
├── .env.example         # Template for .env
├── .gitignore
├── README.md
├── get_mails.py         # Main script
├── requirements.txt
└── token_cache.bin      # Cached auth tokens (created on first run, do not commit)
```

## Security Notes

- Never commit `.env` or `token_cache.bin` (they are in `.gitignore`)
- Keep your Azure client ID private
- Token cache is stored locally; delete it to force re-authentication

## Troubleshooting

| Error | Solution |
|-------|----------|
| `CLIENT_ID not set` | Create `.env` file with your Azure Application (client) ID |
| `AADSTS50059` | Use authority `consumers` (already set for personal accounts) |
| `Application not found` | Verify app supports personal Microsoft accounts in Azure |
| `Permission denied` | Ensure `Mail.Read` and `User.Read` are added in API permissions |

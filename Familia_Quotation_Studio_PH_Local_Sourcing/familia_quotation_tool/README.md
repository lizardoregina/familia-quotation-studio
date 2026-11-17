# Familia Quotation Studio — FREE ONLINE WEBSITE EDITION

This edition is for a **free pilot**: employees can visit a website from anywhere, protected by **one shared company password**. Standard products and a list of generated quotations are saved in a **shared Supabase PostgreSQL database**. The quotation cart is still local to each employee browser. Online product sourcing requires an optional SerpApi key (often paid). The Excel export uses the quotation reference supplied by Familia.

## Step 1: Create your free Supabase database
1. Register at https://supabase.com and create a **Free** project named `familia-quotation`.
2. In its dashboard open **SQL Editor** > **New query**.
3. Open the included `SUPABASE_SETUP.sql` file. Paste its contents and click **Run**.
4. Find **Project URL** in **Integrations > Data API** or the **Connect** dialog.
5. Find or create a **Secret key** (begins `sb_secret_`) in **Settings > API Keys**. It is a server secret; **never** place it in GitHub, chat, website code, or employee instructions. You will enter it only into Render's secret environment variables.

## Step 2: Upload the website files to a PRIVATE GitHub repository
1. Register at https://github.com. Create a repository named `familia-quotation-studio` with visibility **Private**.
2. In the repository click **Add file > Upload files**.
3. Upload the **contents** of this application folder: `app.py`, `requirements.txt`, `render.yaml`, `SUPABASE_SETUP.sql`, `README.md`, and both **folders** `assets` and `templates`, preserving folder hierarchy. The `assets/quotation_template.xlsx` file must be present. GitHub's browser uploader supports dropping files and folders. Do not upload your actual supplier cost files, API keys, `familia.db`, or other private data.
4. Click **Commit changes**.

## Step 3: Deploy for free on Render
1. Register at https://render.com and authorize access to your GitHub repo.
2. From Dashboard select **New > Web Service** (not paid Blueprint).
3. Select your repository, choose runtime **Python 3**, instance **Free**, build command `pip install -r requirements.txt` and start command `gunicorn app:app --bind 0.0.0.0:$PORT`.
4. In **Environment** set the following confidential settings (don't put actual values in repository):

   | Key | Value |
   | --- | --- |
   | `SUPABASE_URL` | Your project's URL, e.g. `https://abcxyz.supabase.co` |
   | `SUPABASE_SECRET_KEY` | Your Supabase **secret** key (`sb_secret_...`) |
   | `APP_PASSWORD` | A unique, strong company password |
   | `SECRET_KEY` | A separate randomly generated long secret, e.g. from a secure password manager |

5. Click **Deploy Web Service**. Once the build completes, use the issued `https://...onrender.com` address.
6. Log in using the company password. Import a **test** price list, refresh from another browser, create an Excel quotation, and verify the **Shared quotation history** section.

Optional supplier search: add `SERPAPI_KEY` as another Render environment variable after obtaining a key from https://serpapi.com. Searches use Google Shopping with the Philippines and Metro Manila settings, then filter for PHP-denominated listings from Philippine-facing stores/marketplaces. They return at most 3 qualifying listings ranked by item price. **These filters cannot verify Philippine shipping origin** (marketplace items can ship from abroad), nor do they cover all local suppliers. Employees must verify the seller is physically sourcing/dispatching locally, exact product, stock, VAT, warranty, shipping cost, and final landed price before quoting.

## Price formula
VAT-inclusive selling unit price = supplier VAT-exclusive cost × (1 + markup%/100) × (1 + VAT%/100). Defaults are 35% markup and 12% VAT, editable per quotation item. Normalize any VAT-inclusive supplier cost first so VAT is not added twice.

## Shared versus local data
- Shared database: **catalog products** and **quotation-export history** (date/customer/contact/amount/count plus the full quotation payload).
- Per-browser only: **unfinished quotation cart**. A browser wipe will erase drafts.
- Excel export: downloaded to the employee's computer, not stored as a file in Supabase; save it in an approved shared company folder.

## Security and reliability limitations
- This is a **pilot, not production-grade secure multiuser software**. Everyone uses the **same password**; there are no unique employee accounts, permissions, change history, quote approvals, or rate-limited centrally managed login protection. Anyone with the password can access/edit supplier pricing. Do **not** use it for confidential operational data until individual authentication, permissions, audit logs, and backup/recovery are added.
- Render Free goes to sleep after about 15 minutes of inactivity; initial requests can be slow. Render's filesystem resets on restarts. The SQL tables exist in Supabase, **not** Render, so stored products and history survive application restart.
- Supabase Free may pause databases with low usage; paused databases preserve data but must be reactivated. **Supabase Free has no managed automatic backups**. Periodic manual exports or a paid backup plan are essential for critical information.
- Costs are 0 within both providers' free allowances, but there are no uptime guarantees; a paid setup will be needed for dependable business operations.
- The online product sourcing API might involve third-party charges.

## Local development
`pip install -r requirements.txt` then `python app.py`. If `SUPABASE_URL` and `SUPABASE_SECRET_KEY` aren't defined, local development uses SQLite `familia.db` for catalog only. The free cloud deployment deliberately refuses to start without its external database and sign-in secrets.

## Philippines-only sourcing policy (updated)
Online comparison targets the Philippines (gl=ph, Metro Manila) and accepts only peso-priced shopping listings from Philippine-facing stores and recognizable PH marketplaces. The comparison is by displayed item price, not delivered cost. Listings have an explicit warning that the seller/ship-from origin is unverified. To enforce a strict **domestic fulfillment only** procurement policy, your purchasing team must verify each listing before placing an order; Google Shopping data does not reliably expose warehouse country.

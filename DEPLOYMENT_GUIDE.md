# NP Refrigeration Inventory — Cloud Deployment Guide

This makes your app accessible from anywhere (phone or laptop) via a web
link, with data stored permanently in a free cloud database (Supabase) so
nothing is lost when the app restarts or is redeployed.

---

## STEP 1 — Create a free Supabase project

1. Go to https://supabase.com and sign up (free).
2. Click "New Project". Pick any name, set a database password (save it
   somewhere safe), choose the region closest to you.
3. Wait ~2 minutes for it to finish setting up.

## STEP 2 — Create the storage table

1. In your Supabase project, open the **SQL Editor** (left sidebar).
2. Click "New query", paste this, and click "Run":

```sql
create table app_storage (
    key text primary key,
    value text not null
);
```

That's it — this single table stores your encrypted inventory, sales, and
user/login data.

## STEP 3 — Get your API credentials

1. In Supabase, go to **Project Settings → API**.
2. Copy the **Project URL** (looks like `https://xxxxx.supabase.co`).
3. Copy the **anon public key** (a long string).

You'll need both of these in Step 5.

## STEP 4 — Put the code on GitHub

1. Create a free GitHub account if you don't have one (https://github.com).
2. Create a new repository (e.g. "shop-inventory"), and upload all the
   files from this folder (`app.py`, `requirements.txt`, etc).
   - Easiest way: on the new repo page, click "uploading an existing file"
     and drag the files in.

## STEP 5 — Deploy on Streamlit Community Cloud (free)

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click "New app", choose your repository and `app.py` as the main file.
3. Before clicking deploy, click "Advanced settings" → "Secrets" and paste:

```toml
SUPABASE_URL = "https://xxxxx.supabase.co"
SUPABASE_KEY = "your-anon-key-here"
```

(Use the values you copied in Step 3.)

4. Click "Deploy". After a minute or two, you'll get a permanent link like
   `https://your-app-name.streamlit.app`.

## STEP 6 — Use it on your father's phone/laptop

- Open the link in any browser.
- On a phone, use the browser menu → "Add to Home Screen" to make it behave
  like an app icon.
- Log in with the same username/password system as before — first run will
  create the default user (check `app.py` for the default credentials, and
  change the password after first login).

---

## Notes

- **Data persistence**: All inventory, sales, and user data now lives in
  Supabase, not on the server's disk — so redeploying, restarting, or even
  switching hosting providers will NOT lose data.
- **Security**: Data is still encrypted (Fernet) before being stored, same
  as your original app. The encryption key itself is stored in the same
  Supabase table.
- **Free tier limits**: Supabase free tier and Streamlit Community Cloud
  free tier are both more than enough for a single shop's inventory data.
- **Local testing**: To run locally before deploying, create a real
  `.streamlit/secrets.toml` file (copy `secrets.toml.example` and fill in
  your real Supabase URL/key — do NOT upload this file with real keys to a
  public GitHub repo).

# AT&T Billing Processor

A reconciliation tool for processing AT&T billing data — handles MDN-based lookups,
Telzeq/NAL customer-specific exports, and working copy generation from raw vendor
Excel files.

## Local setup

```bash
npm install
npm start
```

Visit `http://localhost:3000`.

By default the app uses the password `changeme`. **Change this before sharing the
link with anyone** — see below.

## Setting the shared password

The app is protected with a simple shared login (HTTP Basic Auth — your browser
will show a username/password prompt; the username can be anything, only the
password is checked).

- **Locally:** set the `APP_PASSWORD` environment variable before starting:
  ```bash
  APP_PASSWORD=yourpassword npm start
  ```
- **On Railway (or any host):** set `APP_PASSWORD` as an environment variable in
  your project's settings. Never commit the real password into the code.

Share the password with your team (Gopal, Camilo) through a separate, secure
channel — not in the same place as the link itself.

## Deploying (Railway — free tier)

1. Push this project to a GitHub repository (see steps below).
2. Go to [railway.app](https://railway.app) and sign in with GitHub.
3. **New Project → Deploy from GitHub repo** → select this repo.
4. Railway auto-detects Node.js and runs `npm install` then `npm start`.
5. In the project's **Variables** tab, add:
   - `APP_PASSWORD` = (a real password of your choice)
6. Once deployed, Railway gives you a public URL (Settings → Networking →
   Generate Domain). Share that URL + the password with your team.

## Pushing to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

(Create the empty repo on GitHub first, without a README, then run the above
from this project folder.)

## Notes

- The `uploads/` folder is used as scratch space for incoming/outgoing Excel
  files during processing and is cleared automatically — it is not committed
  to git (see `.gitignore`).
- This app has no database; nothing persists between sessions except what you
  download.

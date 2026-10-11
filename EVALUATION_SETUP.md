# FriendZone evaluation setup

The standalone study lives in `evaluation/`. GitHub Pages publishes that folder, while Supabase stores de-identified responses.

## 1. Create the Supabase table

1. Create or open a Supabase project.
2. Open **SQL Editor**.
3. Run [`supabase/evaluation_schema.sql`](supabase/evaluation_schema.sql).

The SQL enables row-level security. Anonymous visitors can insert one response payload, but they cannot select, update or delete table rows. Researchers can inspect and export responses through the authenticated Supabase dashboard.

## 2. Add GitHub Actions secrets

In `49079986/DECO-7285`, open **Settings → Secrets and variables → Actions** and add:

- `EVALUATION_SUPABASE_URL`: Supabase **Project URL**
- `EVALUATION_SUPABASE_ANON_KEY`: Supabase **publishable/anon key**

The publishable/anon key is designed to appear in browser code. Never put the `service_role` key in GitHub Pages or any client-side file. Data protection comes from the RLS policy in the SQL migration.

## 3. Enable GitHub Pages

Open **Settings → Pages** and select **GitHub Actions** as the source. Merge or push the evaluation files to `main`, or run **Deploy FriendZone evaluation** manually from the Actions tab.

Expected URL:

```text
https://49079986.github.io/DECO-7285/
```

If the two secrets are missing, the page remains usable for testing and downloading answers, but online submission is disabled.

## 4. Test locally

Generate a local configuration and serve the repository root:

```powershell
$env:EVALUATION_SUPABASE_URL = "https://YOUR_PROJECT.supabase.co"
$env:EVALUATION_SUPABASE_ANON_KEY = "YOUR_PUBLISHABLE_ANON_KEY"
node scripts/write-evaluation-config.mjs
python -m http.server 4173
```

Then open `http://localhost:4173/evaluation/`. The generated `evaluation/config.js` is ignored by Git so credentials are not accidentally committed.

## 5. Export and withdraw responses

Use **Supabase → Table Editor → evaluation_responses → Export CSV** for analysis. Participants receive both a response code and receipt code. To locate a withdrawal request, filter by either code. Delete or correct rows only from the authenticated Supabase dashboard.

## Collected fields

- Random response and receipt codes
- Questionnaire version and language
- Device category and prior prototype use
- Permission for de-identified quotations
- Three task completion reports and comments
- Ten SUS responses and calculated 0–100 SUS score
- Open feedback about prompts, leaderboard, simulated NFC/privacy and post-meetup continuity
- Server-generated submission time

The form does not intentionally collect names, student numbers, precise location, prototype click history, audio or video.


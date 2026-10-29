# Cleanearth Site Audit

Field app for Cleanearth water and wastewater site audits.

Each client project runs through five phases, each with a standard checklist the team can edit:

| Phase | What it holds |
| --- | --- |
| 1. Pre-audit visit | Questionnaire, equipment, chemicals, product lines, field readings, observations |
| 2. Proposal | Checklist and costing table (totals by phase) |
| 3. Full audit | Samples and sample codes, lab results, flow log, and an analysis of results against discharge limits, removal efficiencies and measured loads |
| 4. Remediation and retrofit | Options register (cost, priority, status) and costing |
| 5. Monitoring and aftercare | Ongoing results, field readings and samples |

The Overview shows the current phase, checklist progress and every finding, including limit exceedances and measured overloads.

Data is shared live between the team through a free Supabase project, and the app keeps working offline on site.

- **App:** this repository, served by GitHub Pages at `https://cscleanearth-lang.github.io/site-audit/`
- **Data and photos:** your Supabase project (not stored in this repository)
- **SharePoint and Teams:** download the workbook from the Overview, save it into the synced SharePoint **Incoming** folder, and the Power Automate flow posts it to Teams

Setup takes about 20 minutes and needs no IT admin.

## 1. Create the Supabase project (free)

1. Go to <https://supabase.com>, click **Start your project** and sign up (a GitHub or email account works).
2. Click **New project**. Name it `cleanearth-site-audit`, set a strong database password (save it somewhere safe), pick the region closest to Ghana (for example *West EU (London)* or *Central EU (Frankfurt)*), and create it. Wait about two minutes.

## 2. Create the database

1. In the project, open **SQL Editor > New query**.
2. Open `schema.sql` from this repository, copy all of it, paste it in and click **Run**. You should see *Success*. It is safe to run again.

This creates the `records` table, the photo storage, live updates between phones, and the rules that only signed-in team members can read or change anything.

## 3. Lock sign-up and add the team

1. Open **Authentication > Sign In / Providers**. Under **User Signups**, switch **Allow new users to sign up** off and save. Only people you add can then use the app.
2. Open **Authentication > Users > Add user > Create new user**. Enter a colleague's email and a password, tick **Auto Confirm User**, and click **Create user**. Repeat for each person.
3. Give each person their email and password. To reset a password later, open the user in the same list.

## 4. Connect the app

1. In Supabase, open **Project Settings > API** (or **Data API**). Copy the **Project URL** and the **anon public** key.
2. In this repository, edit `config.js` and paste them in:

```js
supabaseUrl: 'https://xxxxxxxx.supabase.co',
supabaseAnonKey: 'eyJ...'
```

3. Commit the change. The anon key is designed to be public. The database rules from step 2 block anyone who is not signed in.

## 5. Turn on GitHub Pages

1. In this repository, open **Settings > Pages**.
2. Set **Source** to *Deploy from a branch*, **Branch** to `main`, folder `/ (root)`, and click **Save**.
3. After a minute the app is live at `https://cscleanearth-lang.github.io/site-audit/`.

On phones, open that address in Safari or Chrome, sign in once, then choose **Add to Home Screen**. The app then opens like a normal app, stays signed in, and loads without signal.

## Good to know

- **Free plan limits:** 500 MB of data and 1 GB of photo storage, which is plenty for many audits. Supabase **pauses a free project after 7 days with no activity**. If the app says it cannot connect after a quiet spell, open the project in Supabase and click **Restore**. Upgrading to the Pro plan removes the pause.
- **Offline:** entries made without signal are kept on the phone and sent automatically when the connection returns. Sign in once with signal before going to site.
- **Backups:** download the workbook from the Overview after each visit and keep it in SharePoint.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | The app |
| `ce-backend.js` | Sign-in, data and photo storage (Supabase) |
| `config.js` | Your Supabase project address and public key |
| `schema.sql` | One-time database setup, run in the Supabase SQL Editor |
| `sw.js` | Keeps the app available offline |
| `vendor/` | Supabase client 2.45.4 and ExcelJS 4.4.0 |

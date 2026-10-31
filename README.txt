KITCHEN RUN
=============================
Files (keep them together in one folder):
  index.html      the page
  styles.css      all styling (colours are the variables at the top of the file)
  auth.js         sign-up / sign-in with email codes, approval check, cloud workspace loading and saving
  team.js         the admin Team tab (loaded only for the admin)
  supabase_setup.sql  the database tables and privacy rules (already applied)
  app.js          all logic: customer data, freeze/due dates, production sheet, orders, export
  logo.svg        the Anino Cloud logo
  menu_data.json  the monthly Simple Diet / Macro / Keto menu rotation

Put the whole folder on any web host (cPanel public_html, Netlify, Vercel, GitHub Pages, S3...).
It is a static site: no server, database or build step. It must be served over http(s),
not opened from a file:// path, because the page loads menu_data.json.

Things to know
- Customer data is saved to each person's account in Supabase (not the browser).
- The page loads two things from the internet: Google Fonts and the SheetJS library
  (cdnjs.cloudflare.com, version 0.18.5) used to read and write Excel files. To host everything yourself,
  download xlsx.full.min.js 0.18.5, save it next to index.html and change the script tag in index.html.
- The logo is the file logo.svg. Replace it with another SVG or PNG (same name, or edit the logo.svg references).
- Menus: edit menu_data.json to change dishes. Keep the same structure.

Accounts (Supabase)
- People sign up or sign in with their email and a one-time code. New sign-ups wait as Pending until the admin
  (aninocloud@gmail.com) approves them in the Team tab. Each approved person has a private workspace; nobody sees
  anyone else's customers. The admin sees everyone in Team and can open, edit, block or reset any workspace.
- Data is saved to the Supabase project automatically (the "Saved to your account" label in the top bar).
- auth.js holds the public project URL and publishable key. These are safe in a web page because the database
  rules (see supabase_setup.sql) enforce who can read and write what. Never put the secret/service key in the site.

One-time Supabase dashboard settings (needed for the email code to work)
1. Authentication > Email Templates > "Magic Link" (used for returning users) and "Confirm signup" (new users):
   edit the body so it shows the code, e.g.  Your Kitchen Run code is: {{ .Token }}
   (Without {{ .Token }}, people receive a link instead of a code.)
2. Authentication > Providers > Email: keep "Enable Email provider" on. Optionally set the code length/expiry.
3. Authentication > SMTP Settings: Supabase's built-in mailer sends only a few emails per hour. Add a real SMTP
   sender (Resend, SendGrid, your own mailbox) before more than a handful of people sign up.

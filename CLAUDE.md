# Hamsa Tourism — Project Rules

Project: Arabic (RTL) offers site for Hajj/Umrah/Tayseer. Static frontend on GitHub Pages,
backend = Google Apps Script + Google Sheets (Offers, Users, Subscriptions).

## ALWAYS
- Keep the site RTL Arabic, mobile-first, brand colors (gold #f9c614, navy #0e2a38, sky #dceef9).
- Enforce every permission in Apps Script (server side). The browser is never trusted.
- Store passwords only as salted hashes; never log or return passwords/hashes.
- Keep admin email/password only in Script Properties.
- Escape all user/offer text before inserting into HTML.
- Explain changes in simple Arabic before applying them.

## ASK FIRST
- Changing the data model (sheet columns), login system, or roles.
- Adding payments or any third-party service.
- Adding or removing a field that collects personal data.
- Migrating to Supabase or any other backend.

## NEVER
- Put secrets, admin credentials or API keys in frontend code or GitHub.
- Trust the browser for permissions, prices, or payment status.
- Delete user data or offers without explicit confirmation.
- Create an admin account through public registration.

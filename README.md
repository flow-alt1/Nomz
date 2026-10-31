# Nomz Kitchen website

```
index.html          the website
admin.html          menu admin (yourdomain.com/admin.html, not linked publicly)
config.js           all settings: API URL, Web3Forms key, hours, store links
menu.js             built-in menu (shown instantly, and the backup if the API is down)
assets/             logo, favicons, and logo-parts/ (pieces for the hero animation)
apps-script/Code.gs backend for the admin panel (Google Sheet)
```

## 1. Contact / catering form
Paste the Web3Forms access key into `config.js` → `WEB3FORMS_KEY`.
Submissions are emailed to the address the key was created with. After sending,
the form is replaced by a thank-you message.

## 2. Admin panel (add dishes without touching code)
1. Create a new Google Sheet (e.g. "Nomz Menu"), then **Extensions → Apps Script**.
2. Replace the code with `apps-script/Code.gs` and save.
3. Select `setup` in the function menu and **Run**. Approve the permissions.
   This creates the **Menu** tab, fills it with the 11 current dishes, and creates
   an admin password. Open the **Execution log** to see it.
   Change it any time: **Project Settings → Script properties → ADMIN_PASSWORD**.
4. **Deploy → New deployment → Web app**. Execute as: **Me**. Who has access: **Anyone**.
5. Copy the URL ending in `/exec` into `config.js` → `API_URL`.
6. Open `/admin.html`, sign in, and use **Add a dish**: paste the square.site
   product link, tap **Get details** (name, Persian name, description, ingredients
   and photo are pulled from Square), add the price and category, then save.
   The live site picks it up within about 5 minutes (the public menu is cached).

If you later edit Code.gs, use **Deploy → Manage deployments → Edit → New version**
so the `/exec` URL stays the same.

## 3. Hosting
Upload the whole folder to Netlify, Vercel, Cloudflare Pages or any static host,
and point nomzkitchen.com at it. If nomzkitchen.com is connected to Square Online
as a custom domain, disconnect it in Square first.

## How ordering works
Customers build their whole order on nomzkitchen.com (saved in their browser).
At checkout, each dish opens in one shared Square tab, where they add it to the
Square cart, then they pay on Square. A browser can't put items into another
site's cart, so a one-click hand-off needs Square API access (see notes).

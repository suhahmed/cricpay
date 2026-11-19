# cricpay: cricket tournament payment tracker (Excel-backed)

A web UI for `data/2025_Cricket_Tournaments.xlsx`. **There is no database: the Excel file is the database.**

## Pages
- **Tournaments**: one card per tournament, each with collected/outstanding totals.
  - **+ New tournament** creates a new sheet in the workbook. You set the short name (sheet tab), tournament name, ground, fee (required), squad size, number of matches (1–15), players per match and optional opponent names. The per-player cost is calculated from the fee. All the formulas are set up for you, with room for 30 players and an organiser summary.
  - **+ Add players** fills the next empty player rows on any tournament sheet and copies the row formulas. Anything typed in *Already paid* is logged as an opening payment.
  - **Edit matches** (new-style sheets) renames matches and changes the fee or players per match. Everyone who played is re-charged automatically.
- **Advances**: each tournament can have an advance per player (default ₹3,000) and an expected total cost per player (default ₹3,500–4,000).
  - Set them when creating a tournament, or later with **Settings** on its card. They're saved as `Advance:` and `Expected cost:` lines in the tournament box on the sheet.
  - The overview shows an **Advance** column: ✓ when paid, or **Collect ₹X**, which opens Record payment with the shortfall filled in.
  - There's also an *Advances collected* card and an **Advance pending** status filter.
  - Every payment is tagged **Advance / Match fees / Other** in the new *Type* column of the Payments sheet.
- **Organiser & kitty**: a strip on the overview (for all tournaments, or the one you filter to), plus a box on every tournament card. It shows:
  - **Organiser fee**, **Paid to organiser** and **Balance to organiser**
  - **Collected from players** and **Still to collect** (fee − collected)
  - **Kitty in hand** (collected − paid out), shown as **Short (fronted)** when more was paid out than collected
  - Use **+ Pay organiser** to record money paid to the organiser. It goes into the next empty organiser-payment cell on the tournament sheet, with the date next to it. It's also logged on the **Organiser Payments** sheet, with Undo on Payment history.
  - It reads the organiser block your sheets already have (the `Total − SUM(paid)` balance formula). Set or change the fee in **Settings**, and a sheet without a block gets one added. New tournaments get the block automatically.
- **Payment overview** (`/`): every player in every tournament, with amount due, paid so far, outstanding and status.
  - The six match boxes are clickable. ✓ means played. Clicking switches Yes/No and the amounts recalculate from the sheet's own formulas.
    - On sheets where Yes/No is typed in (GurukulMarvel, UCC, UrbanFarms), the click writes Yes/No into that cell.
    - On BRC-Rebelstars, Yes/No is calculated from the match line-up blocks, so the click adds the player to (or removes them from) that match's line-up.
  - **+ Record payment** or **+ Payment** on a row: amount, date, mode (UPI/Cash/Bank) and a note.
    - The amount is added to the player's *Paid* cell on their tournament sheet (or the last instalment cell feeding it, e.g. *Clearance* on GurukulMarvel).
    - It's also logged as a row on a **Payments** sheet that the app creates in the workbook.
  - **Export CSV** exports the filtered table.
- **Payment history**: all payments recorded from the website. **Undo** subtracts the amount again and deletes the log row.
- **Sheet view** (`/sheets`): edit any non-formula cell of the workbook directly (names, costs, instalments, dates…).

- Each change in the browser (Y/N attendance, amounts, Paid status, names, dates) is written **straight into the matching cell** of the Excel file.
- Formula cells (balances, totals, "To be paid", etc.) are **never overwritten**. The app calculates them with the workbook's own formulas, so the UI shows the same numbers Excel does. In the UI they are marked as calculated cells (hover to see the formula).
- If you edit the Excel file in Excel and save it, the web page picks up the change within about 15 seconds.
- Every 10 minutes of editing, a backup copy goes to `data/backups/` (the last 40 are kept).
- The **Excel** button downloads the current file.

## Roles & access
Everyone signs in. Nothing is visible without an account.

| Role | Scope | Can see | Can change | Manages access |
|---|---|---|---|---|
| **Admin** | whole website | every tournament, plus the **All tournaments** overview | everything | everyone (admins, captains, players) |
| **Captain** | per tournament | their tournaments | their tournaments only (payments, matches, players, settings, organiser) | the **players** of their tournaments |
| **Player** | per tournament | their tournaments | nothing | – |

- **Overview:** pick a tournament with the switcher in the page title, and the overview, organiser & kitty strip, and payment history all show only that tournament. **All tournaments** appears only for admins.
- The server enforces all of this. A player's browser never even receives other tournaments' data. The full Excel download and the log sheets are admin only.
- **First account:** the first account becomes the **admin**. If an edit PIN was set before, it's needed to create it.
- **Sign-ups:** players sign up with an optional "which team are you in?" note and wait. An admin or a captain gives them access to a tournament, which approves them.
- **Team & access page:**
  - **Admins:** add members, make someone an admin, set per-tournament access (Player / Captain), disable, delete, reset passwords, open or close sign-ups.
  - **Captains:** add players to their tournaments, give or remove player access, approve sign-ups into their tournament, reset passwords for players who are only in their tournaments.
- **New tournaments:** admins and captains can create them. A captain who creates one becomes its captain.
- **Audit:** each payment records who entered it, in the *Recorded by* column of the Payments and Organiser Payments sheets.
- **Security:**
  - Passwords are stored as salted PBKDF2 hashes in `data/cricpay-users.json`, never in Excel.
  - Sessions last 30 days.
  - Five wrong passwords lock that login for 60 seconds.
  - Changes need the app's own request header, which blocks cross-site form attacks.
- **Locked out?** In a Bash console, in the `cricket-tracker` folder:
  - `python manage_users.py list`
  - `python manage_users.py reset-password <username>`
  - `python manage_users.py make-admin <username>`
  - `python manage_users.py create-admin <username> "<Full name>"`

## Match costs (from the tournament fee)
- **Match cost = tournament fee ÷ number of matches.** Example: ₹42,000 ÷ 6 = ₹7,000 per match.
- **Per player = match cost ÷ players per match**, rounded to the rupee. Example: ₹7,000 ÷ 12 = ₹583.
- Each match's cost is shared among the players marked **Yes** for it. If 12 play, each pays ₹583. If 10 play, each pays ₹700. If 14 play, each pays ₹500.
- Marking someone as played or not played re-splits it for everyone.
- Change the fee or the players per match (each match can differ) in **Tournaments → Settings**, and every amount recalculates.
- On the sheet, they're the yellow rows under the player table: *Rate per player (fee ÷ matches ÷ players)*, *Players per match*, *Match cost (fee ÷ matches)*, *Cost per player (shared)*. They read the fee from the organiser box.
- **Custom amount:** captains and admins can click **✎** next to a player's amount to set a fixed amount, which replaces their match-based share (extras still apply). Clear it to go back. It's stored in the *Custom amount* column.
- Tournament sheets created by the website before this version are upgraded automatically on first start, and a backup named `before-fee-based-cost-upgrade-…xlsx` is saved first.
- The four original 2025 sheets keep their own formulas.

## Run it on your laptop

1. Install Python 3.10+ from python.org (on Windows, tick "Add Python to PATH").
2. After it starts, open the site and create the admin account on the sign-in page.
3. Double-click `start-windows.bat`, then open http://localhost:8000.

> Windows note: if the Excel file is open in Excel, Windows locks it and saves from the web will fail with a clear message. Close it in Excel and the next save goes through.

**Adding a new tournament:** copy an existing sheet in Excel, keep the same layout (`SN | Player | Match 1…` table, `SN | Player | Cost | Status` match blocks, `Sno | Name` instalments, a `Tournament:` info box), save, and it appears as a new tab.

---

## Making it public: pick one

### Option A: Keep the Excel file on your laptop and share a public link (5 minutes)
Best if you're fine with the laptop staying on while people use it.

1. Install Cloudflare's free `cloudflared` tool: https://developers.cloudflare.com/cloudflare-one/networks/connectors/cloudflare-tunnel/downloads/
   (Windows: `winget install --id Cloudflare.cloudflared`)
2. Start the tracker (above), then in a second terminal run:
   ```
   cloudflared tunnel --url http://localhost:8000
   ```
3. It prints a link like `https://random-words.trycloudflare.com`. Share it on the team WhatsApp. No account is needed.
   - The link changes each time you restart it. For a permanent URL, create a free Cloudflare account, add a domain and set up a *named tunnel*. (ngrok's free plan, which includes one static domain, is another option: `ngrok http 8000`.)

### Option B: Host it 24×7 for free on PythonAnywhere (laptop can be off)
The Excel file then lives on PythonAnywhere's disk (persistent storage, unlike Render/Vercel/Heroku-style free hosts, which wipe files on restart).

1. Sign up for a free account at https://www.pythonanywhere.com. Your site will be `https://YOURUSERNAME.pythonanywhere.com`.
2. **Files** tab → upload this folder as `/home/YOURUSERNAME/cricket-tracker/` (or upload the zip and run `unzip` in a Bash console).
3. **Consoles** → Bash:
   ```
   cd ~/cricket-tracker && pip3.11 install --user -r requirements.txt
   ```
4. **Web** tab → *Add a new web app* → *Manual configuration* → Python 3.11.
5. Open the *WSGI configuration file* link and replace its contents with `pythonanywhere_wsgi.py` (put in your username), then open the site and create the admin account. Click **Reload**.
6. Free-tier caveat: once a month, log in and click **"Run until 1 month from today"** on the Web tab, or the site goes to sleep. Use the site's **Excel** button whenever you want a copy of the file locally.

### Option C: Your own server or VPS
`pip install -r requirements.txt` then `PORT=8000 python app.py` (it uses the waitress production server), with nginx/Caddy for HTTPS in front.

### Same public link every time (Mac, free): ngrok
ngrok's free plan gives your account **one fixed address** (like `https://something-random.ngrok-free.app`). You can't pick the name, but it stays the same every time.
1. Sign up at https://ngrok.com, then copy your **authtoken** from the dashboard (*Getting Started → Your Authtoken*).
2. Download ngrok for macOS from https://ngrok.com/download (Apple chip = ARM64, Intel = x86-64). Unzip it so the `ngrok` file sits in **Downloads**.
3. One time, in Terminal: `~/Downloads/ngrok config add-authtoken YOUR_TOKEN`
4. Every time: `bash ~/Documents/cricket-tracker/start-public.sh`. This starts cricpay, keeps the Mac awake and opens the same public link. Ctrl+C stops everything.
- Free plan limits: 20,000 requests and 1 GB a month. Open pages check for updates every 30 s, and only while visible. Visitors see a one-time ngrok "Visit site" page, which is then remembered for 7 days.

## Settings (environment variables)
| Variable | Default | Meaning |
|---|---|---|
| `EDIT_PIN` | *(empty)* | Optional. Only used to protect creation of the first (admin) account. |
| `SECURE_COOKIES` | *(empty)* | Set to `1` on an https site (the PythonAnywhere WSGI file does this). |
| `TRACKER_XLSX` | `data/2025_Cricket_Tournaments.xlsx` | Which Excel file to use |
| `PORT` | `8000` | Port to listen on |

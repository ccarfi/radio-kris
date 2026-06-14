# Radio KRIS — Setup

Two files:
- **`Code.gs`** — the backend, pasted into the Sheet's Apps Script.
- **`index.html`** — the app, opened in any browser.

The password is **`extricity`** (stored as a salted hash, never plaintext).

---

## 1. Get a YouTube Data API key
1. Go to <https://console.cloud.google.com/> → create/select a project.
2. **APIs & Services → Library → "YouTube Data API v3" → Enable.**
3. **APIs & Services → Credentials → Create credentials → API key.** Copy it.
   - (Optional but recommended: restrict the key to the YouTube Data API.)
   - Free quota ≈ 10,000 units/day; each search ≈ 100 units → ~100 searches/day. Raise in Cloud if needed.

## 2. Add the backend to your Sheet
1. Open the Sheet → **Extensions → Apps Script.**
2. Delete the starter `Code.gs` content and paste in **`Code.gs`** from this folder.
3. **Project Settings (⚙) → Script Properties → Add property:**
   - Name: `YT_API_KEY`  Value: *(your key from step 1)*
4. Save.

## 3. Deploy the backend as a Web App
1. In Apps Script: **Deploy → New deployment.**
2. Type: **Web app.**
3. **Execute as: Me.** **Who has access: Anyone.**
4. **Deploy**, authorize when prompted, and **copy the Web app URL** (ends in `/exec`).

> Re-deploying a *new version*: use **Manage deployments → edit (pencil) → Version: New version** so the URL stays the same.

## 4. Point the app at the backend
1. Open **`index.html`** in a text editor.
2. Find the `CONFIG` block near the bottom and set:
   ```js
   url: 'https://script.google.com/macros/s/XXXXXXXX/exec',
   ```
   (Leave `authSalt` / `authHash` as-is unless you rotate the password — see below.)
3. Save.

## 5. Run it
- **Locally:** double-click `index.html` (or drag into a browser).
- **Share with others:** host the single file anywhere static — GitHub Pages, Netlify drop, etc. — and send the link. Everyone uses password `extricity`.

## 6. First use
- Each station = one **tab** in your Sheet (e.g. "Extricity Classics"). Add tracks via **＋ ADD** (in-app search).
- A `Metadata` tab and `Presence` tab are created automatically on first run.
- Tap **▶ TUNE IN** to join the broadcast (required once per device for mobile audio).

---

## Rotating the password
The plaintext word lives nowhere. To change it:
```bash
python3 -c "import secrets,hashlib; s=secrets.token_hex(16); \
print('salt',s); print('hash',hashlib.sha256((s+'NEWPASSWORD').encode()).hexdigest())"
```
Put the new `salt`/`hash` into **both** `Code.gs` (`AUTH_SALT`/`AUTH_HASH`) and `index.html` (`CONFIG.authSalt`/`authHash`), re-deploy the script, and re-share the file.

## Honest security note
The password is obfuscation, not a lock: the hash travels on every request (visible in devtools) and is replayable. Keep nothing sensitive in the station. See §6 of the spec.

## Troubleshooting
- **"Set CONFIG.url…"** — you didn't paste the `/exec` URL into `index.html`.
- **"unauthorized"** — salt/hash mismatch between the two files, or wrong password.
- **Search says "YT_API_KEY not set"** — add it in Script Properties (step 2.3).
- **Nothing plays / "TAP TO RESYNC"** — mobile blocked autoplay; tap TUNE IN again.
- **CORS errors** — make sure the deployment is **Anyone** access and you're using the `/exec` (not `/dev`) URL.

# SMKPB Mail Merge — Outlook add-in

Paste a table, write one email with fill-in fields like `{{First Name}}`, preview each person's version, and send everyone their own copy from your smkpb.org mailbox. Each person gets an individual email (nobody sees anyone else's address), and every email lands in your Sent Items.

It runs as a side panel in **Outlook on the web** and the **new Outlook**. Emails go straight from Outlook to Microsoft; nothing passes through GitHub or any other server.

---

## One-time setup (about 20 minutes)

You'll do four things: put the files on GitHub Pages, point the manifest at them, register the add-in with Microsoft, and install it in Outlook.

### 1. Put the files on GitHub Pages

1. Sign in at [github.com](https://github.com) (create a free account if you don't have one).
2. Click **New repository**. Name it exactly `smkpb-mail-merge`, set it to **Public**, and click **Create repository**.
3. On the new repository's page, click **uploading an existing file**, drag in everything from this folder (including the `assets` and `vendor` folders), and click **Commit changes**.
4. Go to **Settings → Pages**. Under **Build and deployment**, set Source to **Deploy from a branch**, Branch to **main** and **/(root)**, then **Save**.
5. After a minute or two, `https://smcka7.github.io/smkpb-mail-merge/` shows a short "SMKPB Mail Merge" page. That means it's live.

The repository is public, but it holds no passwords or keys. The Application ID from step 3 isn't a secret.

### 2. Point the manifest at your GitHub Pages address

Open `manifest.xml` in any text editor and replace every `smcka7` with your GitHub username (there are a dozen; use Find & Replace). Save it, and upload the updated `manifest.xml` to the repository too.

### 3. Register the add-in with Microsoft

This tells Microsoft that the add-in is allowed to send mail as you.

1. Go to [entra.microsoft.com](https://entra.microsoft.com) and sign in with your smkpb.org account.
2. Open **Applications → App registrations → New registration**.
3. Fill in:
   - **Name:** `SMKPB Mail Merge`
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** platform **Single-page application (SPA)**, value `brk-multihub://smcka7.github.io`
     (domain only, with no `https://` and no folder name after it)
4. Click **Register**. On the overview page, copy the **Application (client) ID**. You'll paste it into the add-in.
5. Open **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, tick **Mail.Send**, and click **Add permissions**.
6. If you see **Grant admin consent for …**, click it. If you don't, that's fine: Outlook will ask you to approve the first time you sign in.

**If you use Microsoft 365 through GoDaddy:** GoDaddy manages these accounts, and some GoDaddy plans don't give you full admin rights. If entra.microsoft.com won't let you register an app, or sign-in later says an admin must approve, that limit is the cause. Your options are to ask GoDaddy support to allow it, or to move the account off GoDaddy's management ("defederating"), which is a bigger change.

### 4. Install it in Outlook

1. Open Outlook on the web at [outlook.office.com](https://outlook.office.com) with your smkpb.org account.
2. Go to [aka.ms/olksideload](https://aka.ms/olksideload). It opens the add-ins window in Outlook.
3. Choose **My add-ins → Add a custom add-in → Add from file**, pick your edited `manifest.xml`, and confirm.
4. Open any email. Click **Apps** in the toolbar (or the **…** menu), then **SMKPB Mail Merge**. Click the pin icon at the top of the panel to keep it open as you move between emails.
5. In the panel, open **Settings** at the bottom, paste the Application (client) ID, and click **Sign in**. Approve the permission when Microsoft asks.

---

## Using it

1. **Recipients.** In Excel, select your table including the header row, copy, and paste it into the box. You can also upload a CSV. The add-in finds the email column on its own; change it if it guessed wrong. Optional CC and BCC columns are supported, and a cell can hold several addresses separated by `;`.
2. **Message.** Write the subject and body. Click a field name to insert it, or type `{{Column Name}}` yourself. Add a fallback for blank cells with a bar: `Hi {{First Name|there}},`. Links you type become clickable.
3. **Signature and attachment (optional).** Your Outlook signature isn't added to emails sent this way, so the panel has its own signature box (saved for next time). One attachment of up to 2.5 MB can go to everyone.
4. **Preview.** Step through each person's email with the arrows. Rows with a missing, invalid, or repeated address are flagged and skipped. Blank fields are flagged too.
5. **Send.** Click **Send test to me** first; it sends the person on screen's version to you with "[Test]" in the subject. Then click **Send to N people** and confirm.

**While sending:** emails go out at about 25 a minute (Microsoft allows 30 a minute per mailbox), so 100 people takes about 4 minutes. Keep the panel open. **Stop** halts after the current email.

**After sending:** the panel lists each person's result. **Retry failed** resends only the ones that didn't go. **Copy results** copies a table you can paste into Excel.

**No double sends:** the panel remembers who already got a given subject and body on this computer. Running the same message again skips them and only sends to the new or failed rows. To send someone the same message again on purpose, click **Send to them again** under the Send button.

## Good to know

- **Limits:** Exchange Online caps a mailbox at 30 messages a minute and 10,000 recipients a day. For cold outreach, much smaller daily batches (a few dozen) keep your domain's reputation healthy.
- **Commercial email (CAN-SPAM):** for marketing emails, include your business's physical address and a way to opt out. The signature box is a good place for both. Honor opt-out requests within 10 business days.
- **Privacy:** your table is only kept in this browser, on this computer. Each email goes directly from Outlook to Microsoft Graph.

## Troubleshooting

| What you see | What to do |
|---|---|
| "Add your Application (client) ID…" | Settings → paste the ID from step 3.4 → Save. |
| "Microsoft doesn't recognize that Application (client) ID" | Recopy it from the app's Overview page in Entra. Copy the Application ID, not the Object ID. |
| "missing this add-in's redirect URI" | In Entra → your app → Authentication, check the SPA redirect is exactly `brk-multihub://smcka7.github.io`. |
| "Permission to send mail hasn't been granted" or "admin approval required" | Step 3.5–3.6. On GoDaddy-managed accounts, see the GoDaddy note in step 3. |
| "This version of Outlook can't sign in add-ins" | Use Outlook on the web or the new Outlook. Classic desktop Outlook may not support this sign-in. |
| The add-in doesn't appear under Apps | Make sure the GitHub Pages address in step 1.5 loads, then reinstall the manifest. |
| Changes you uploaded to GitHub don't show up | GitHub Pages takes a minute to update. Close and reopen the panel, or refresh Outlook. |

## Files

| File | What it is |
|---|---|
| `manifest.xml` | Tells Outlook where the add-in lives. This is the file you install. |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The side panel. |
| `merge.js` | Table parsing, field filling, checks, and the paced send loop. |
| `config.js` | Optional settings: a built-in client ID, send pace, and attachment limit. |
| `vendor/msal-browser.min.js` | Microsoft's sign-in library (MSAL.js 5.27, MIT license). |
| `assets/` | Icons. |

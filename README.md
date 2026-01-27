# Reply Copilot for Outlook

- Auto-drafts replies on **Reply** with **TL;DR**, **direct tone**, **medium length**, **"Thanks"** closing.
- Learns lightweight style signals on **Send** (client-side, RoamingSettings).
- Optional connection to Azure OpenAI / OpenAI from the Taskpane (no server required).

## Host
Published via **GitHub Pages** from branch `main` / folder `root`.

Live files:
- `https://<user>.github.io/reply-copilot-addin/commands.html`
- `https://<user>.github.io/reply-copilot-addin/taskpane.html`

## Install
1. Ensure `icon-32.png` and `icon-80.png` exist in repo root (any small PNG is fine).
2. Download `manifest.xml` locally.
3. Outlook → **Get Add-ins → My add-ins → Upload custom add-in → Add from file** → pick `manifest.xml`.
4. Open any email → **Reply** → the draft should appear.

## LLM (optional)
Open taskpane → choose **Azure OpenAI** or **OpenAI** → paste keys → Save → regenerate draft.

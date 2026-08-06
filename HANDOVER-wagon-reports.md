# Fox wagon earnings reports — handover

Written 04/08/2026. Covers both daily wagon earnings automations, what is live,
what is outstanding, and the things that will bite anyone picking this up cold.

---

## 1. What the two reports are

### Lancashire (Fox Brothers Lancashire)

Took over from Simon Colderley when he left the group end of July 2026. It must
go out as he had it.

| | |
|---|---|
| Trigger | Mel Vose emails `Daily wagon earning master Mel.xlsx`, daily around 12:30 to 14:15. Subject is the filename. Daniel is cc'd. |
| Output | 13 page landscape PDF `Fox-Group-Daily-Wagon-Earnings-YYYY-MM-DD.pdf`, plus Mel's master with the COVER tab rebuilt onto it |
| To | mel@foxbrothers.co.uk, paulfox@foxbrothers.co.uk, mark.hierons@foxgroup.co, darren@foxbrothers.co.uk, samuel@foxbrothers.co.uk, katie@foxbrothers.co.uk, stuartsweet@foxbrothers.co.uk |
| Cc | barry.hope@foxgroup.co, mike.yates@foxgroup.co, liam@foxgroup.co, Richard.Kirwin@foxgroup.co, daniel.walsh@kfltd.uk |
| Code | `lancs_auto.py` (trigger and send), `lancs_pack.py` (the PDF), `lancs_data.py` (the figures), `lancs_cover.py` (cover tab), `lancs_inject.py` (transplant) |

Data comes from Mel's master for everything except drivers, which come from the
portal load sheets in `fox-portal/reports/*.json` where division is
"Fox Brothers Lancashire". Not Samsara.

### Leyland (Fox Brothers Leyland)

| | |
|---|---|
| Trigger | Katie Ward emails `Wagon earnings - DD/MM/YY`. Emmy Duckworth covers when Katie is off, Paul sometimes forwards them on. |
| Output | The master updated to the latest day, with the COVER tab, emailed to Paul |
| To | paulfox@foxbrothers.co.uk |
| Cc | daniel.walsh@kfltd.uk |
| Code | `wagon_auto.py` (trigger and send), `wagon_master_fill.py` (fills the master), `tidy_master.py` (formatting), plus the shared cover modules above |

Parts and tyres come from the workshop transaction report that lands at 18:00,
not from the database.

**These two must never mix.** Leyland only accepts mail from Katie, Emmy or
Paul, deliberately not from Simon, who sent the Lancashire pack. On top of that
every sheet is scored against the Leyland fleet registrations and rejected below
70 per cent. A real Leyland sheet scores 100 per cent, a Lancashire one scores 0.

---

## 2. Where it runs

Coolify app **fox-workshop**, uuid `atny5ap3f0lnv5g0kfnxeqzx`.

| Task | uuid | Schedule |
|---|---|---|
| lancs-pack-auto | `a18nysdrmdqh5ovtlk3j47jd` | every 15 minutes |
| wagon-master-auto | `o106gwdqxkufc577gjxbstah` | every 15 minutes |
| daily-6pm-report | `e12km9imhe1ht4xrhhn5dia0` | 17:00 and 18:00 |
| reply-agent | `ior2eqamazoxrgxhr6s2de7r` | several times daily |

Repo: `https://github.com/danielwalsh-ai/FoxWorkshop.git`

**There is no persistent volume.** State is wiped on every redeploy, and that is
by design. Both jobs rebuild their position from Gmail:

- Lancashire looks for the newest pack we sent that is not prefixed `[TEST`, and
  only acts on mail from Mel newer than that.
- Leyland recovers the master itself from the newest `Daily wagon earnings *.xlsx`
  we emailed, and treats everything Katie sent before it as already done.

Without that, a redeploy would re-send the whole run to the full distribution list.

---

## 3. The cover sheet

Paul asked for it on 03/08/2026. It is a new first tab so it is the first thing
seen on opening. One bar chart per highlighted row, month by month, with a trend
line over the bars.

Settings that were argued over and are correct as they stand:

- Starts January 2026, not 2025
- Four charts across the page, not one long column
- Bars all one colour, Fox navy `1A2646`. Trend line red `C00000`
- The trendline equation and R squared are hidden. They must be written as `0`,
  because leaving them out means Excel applies its own default and shows them
- The month figure sits inside the end of each bar, white and bold, so the trend
  line does not cover it
- No frozen panes anywhere. A frozen column clips the charts, and a frozen row
  leaves month headers floating over nothing
- Section names come off the block headers on the DAILY tab
- Fox logo top right

Rows are picked up from Paul's dark blue highlighting (`FF0070C0`) on the DAILY
tab. Leyland has no highlighting, so there it falls back to matching by meaning.

---

## 4. Where things stood on 04/08/2026

**Lancashire.** Live and running. Schedule is on. The last pack, Friday 31 July's
figures, went to the full list on Monday 03/08/2026 at 14:45. The daily email now
carries the master with the cover on it.

**Leyland.** Live and running, but with nothing to do. The last run sheet was
Friday 31/07/2026 carrying Wednesday 29 July's figures. Nothing arrived Monday
03/08 or Tuesday 04/08, so 30 July, 31 July and 3 August are all missing. That is
their end, not ours. **Chase Katie or Emmy.**

The Leyland master with its cover was emailed to Paul on 04/08/2026 asking finance
to fill that version out instead of the old one.

---

## 5. Outstanding

1. **The 37 row inserts on Leyland DAILY.** Instructions are in
   `Leyland rows to insert.xlsx` in Downloads. Daniel inserts them in Excel, which
   is the only safe way, because Excel updates the references in 9 other sheets and
   31 charts automatically. Doing it in code means rewriting every one of those by
   hand. Excel will not launch under the automation session, so it cannot be scripted
   from here.

   **After the inserts, the automation must be re-pointed before it runs again.**
   It writes to fixed row numbers and they will all be wrong. While in there, switch
   the row lookup to work off the labels so this cannot happen again.

2. **Paul's "fix the dashboard up a bit"** is still undefined. Ask him what he means.

3. **Finance divergence.** If finance fill out their own copy daily while the
   automation fills ours, the two masters drift apart. It has worked so far because
   Daniel folds their edits back in by hand. Worth sorting properly.

4. **Tell Paul his dragged back totals start January 2026, not January 2025.**

5. **Section total naming.** Lancashire calls each block's total TOTAL EARNINGS,
   Leyland calls it HOOKS TOTAL, 8 W TOTAL and so on. Same row, same figure, and
   Leyland's is far better filled in, 337 days against 106. Left alone deliberately.

---

## 6. Things learned the hard way

**Never re-save these workbooks through openpyxl.** Mel's carries 180 chart parts
and 16 drawings, Leyland's 31. openpyxl destroys them on save. Everything is done
by editing the raw XML inside the zip.

**Excel opens on whichever sheet carries `tabSelected`, not on `activeTab`.**
Setting activeTab alone leaves it opening wherever the last person saved. Strip
tabSelected off every sheet and put it on the cover.

**Strip the old cover before injecting a new one.** The Leyland master round trips
through our own email, so it comes back already covered. Injecting again leaves the
old worksheet, drawing and 44 chart parts orphaned, and Excel reads orphans as a
damaged file. Verified stable over three strip and inject cycles, same size and part
counts every time.

**Style indices have to be shifted.** A cell's `s=` is an index into that workbook's
own styles.xml, so the cover's indices mean nothing inside Mel's file. Its fonts,
fills and number formats are appended to hers and every index moved by the offset.

**definedName localSheetId is a position, not an id.** Inserting a sheet at the
front shifts every sheet after it, so any defined name scoped at or beyond that
point has to move with it, or Excel repairs the file and drops the drawings.

**"Enable Editing" on an emailed workbook is Excel's Protected View.** It applies
to any file that arrives by email and there is nothing in the file causing it.

**Build chart pages from the rendered reference, not the extracted text.** Three
of Simon's charts were rebuilt wrong because they were read from PDF text rather
than looked at.

**Night work registrations are 8 characters** (PJ21OWEN). A 7 character filter let
them through as their own wagons.

**Guard the date window.** Katie's 07.07.2026 sheet had 2025 in C2, which would
have landed the figures in last year's column. There is a 90 day window and a
filename cross check.

**The external Postgres on port 5433 accepts a connection then swallows it.**
Never solved. Sidestepped by working from the emailed attachments instead, which
turned out better anyway.

---

## 7. Credentials

`fox-report/.env` is **gitignored**, so it is not in the repo and not on GitHub.
If this folder is deleted without copying it, these are gone:

```
GMAIL_USER, GMAIL_APP_PASSWORD          the account both reports send from
COOLIFY_URL, COOLIFY_TOKEN              deploys and scheduled tasks
WORKSHOP_DATABASE_URL                   the workshop Postgres
ANTHROPIC_API_KEY, RESEND_API_KEY
AUTOVOLT_URL, AUTOVOLT_EMAIL, AUTOVOLT_PASSWORD
MOT_API_KEY, MOT_CLIENT_ID, MOT_CLIENT_SECRET, MOT_SCOPE, MOT_TOKEN_URL
EMAIL_FROM, EMAIL_TO
```

**Copy `.env` somewhere safe before deleting anything.** The same values are set
as environment variables on the Coolify app, so the running service will not stop
if the local copy is lost, but they cannot be read back out easily.

Also local only and not in the repo: `state/`, `base_report.xlsx`, and the
generated PDFs. None of those matter, they rebuild themselves.

---

## 8. Everything else in this folder

The other projects sitting alongside this one, each its own git repo:

`fox-portal`, `fox-movement-order`, `fox-vor`, `fox-marina-hub`, `fox-samsara-map`,
`fox-merch`, `fox-spotter`, `fox-waiting-time`, `cw-staff`, `tranzparts-staff`,
`fw-classic-cars`, `preston-basketball`, `barge-pricing`, `label-detector`,
`Kevin-MIT`.

Before deleting any of them, check two things in each: that `git status` is clean
and pushed, and whether it has its own `.env`. The code is safe on GitHub. The
`.env` files are not.

# NCD Mobile — Finance Contact Enrichment Routine

Goal: for each of the **164 NCD-issuer companies** in `NCD Mobile.xlsx`, find the
**direct finance contact** — ideally the **CFO or Treasury head**, else the finance
team — with their **mobile, email**, and a **fallback** (landline / IR desk) when no
person-level contact exists. Output goes to a **new** file, `NCD Mobile - Enriched.xlsx`;
the original is never modified.

This is a **resumable** routine. Two passes are scheduled tonight (**9:00 PM** and
**1:00 AM**). The 1 AM pass **continues where 9 PM stopped** — it does not redo companies
already completed. If the 9 PM pass finishes all 164, the 1 AM pass simply reports "nothing
left" and exits.

---

## Files (state lives in the repo, so any session can resume)

| File | Role |
|---|---|
| `NCD Mobile.xlsx` | Source of truth — 164 companies. **Never edit.** |
| `contacts.json` | Accumulated research + per-company `status`. **This is the resume state.** |
| `build_enriched.py` | Renders `NCD Mobile - Enriched.xlsx` (color-coded) from `contacts.json`. Idempotent. |
| `NCD Mobile - Enriched.xlsx` | Deliverable, rebuilt after every batch. |
| `README.md` | This spec. |

Each company object in `contacts.json`:

```json
{
  "row": 0, "company": "...", "cin": "...", "sector": "...",
  "existing_reg_email": "...", "existing_phone": "...",
  "status": "pending | done | partial | not_found",
  "contact": {                      // PRIMARY contact (CFO / Treasury preferred)
    "name": null, "designation": null, "mobile": null, "email": null,
    "fallback": null, "confidence": "high|medium|low",
    "source": null, "notes": null, "run": null
  },
  "additional": []                  // ANY other finance contacts found — same shape
}
```

**Capture everything you find.** The master should hold *all* contacts per company, not
just one. Put the best CFO/Treasury contact in `contact`; add every other useful finance
person (deputy CFO, VP Finance, Treasury manager, IR head, Company Secretary for debt, …)
as its own object in the `additional` list. Each additional contact carries its own
`confidence`, and shows in the "Additional Contacts" column with an inline `(HIGH/MED/LOW)`
tag.

---

## What each scheduled run must do

1. **Clone / pull** this repo.
2. **Read `contacts.json`.** Select companies where `status == "pending"` (or `partial`).
   Skip everything already `done` / `not_found`. **This is what makes it resumable.**
3. Work through pending companies **in order** (row 0 first — they are sorted by rating count,
   most important first). For each, research the finance contact:
   - **Priority:** CFO → Treasury head → other finance leadership → finance/IR desk.
   - **Sources:** company website / investor-relations page, latest annual report, NCD offer
     document & debenture-trustee filings, MCA / DIN director records, credible news, LinkedIn.
   - Capture `name`, `designation`, `mobile`, `email`, and a `fallback` (direct finance
     landline or IR email) when a personal contact isn't available.
   - Set `confidence` and `source` (see rules below). Set `run` to `"2026-08-13-2100"` or
     `"2026-08-14-0100"`.
   - Update `status`: `done` (person-level contact found), `partial` (only fallback found),
     or `not_found` (nothing credible).
4. **Save `contacts.json`**, then run `python build_enriched.py`.
5. **Commit & push** both `contacts.json` and `NCD Mobile - Enriched.xlsx` **frequently**
   (e.g. every ~15 companies) so progress survives if the run is cut short.
6. Stop when: no pending companies remain, **or** the time/token budget is nearly spent.
   Always commit before stopping. Leave the rest `pending` for the next pass.

---

## Confidence → color (surety), applied by `build_enriched.py`

| Confidence | Color | Use when |
|---|---|---|
| `high` | 🟢 GREEN | Verified: official site / annual report / MCA / offer doc, or **2+ sources agree** |
| `medium` | 🟡 YELLOW | Plausible: a single source, LinkedIn, or an **inferred** email pattern |
| `low` | 🔴 RED | Weak / **fallback**: generic landline used instead of a person, or unverified |

### Hard rules
- **Never fabricate a phone number or email.** If unsure, leave the field blank.
- An **inferred** email (e.g. `firstname.lastname@domain`) is **medium at best**, and only when
  the domain pattern is confirmed from a real address. Note it in `notes`.
- Prefer a **real fallback** (finance-desk landline) over a guessed mobile. Fallback ⇒ `low`.
- Keep the row's original data intact; only fill the new contact fields.

---

## Manual run (for testing outside the schedule)

```bash
python build_enriched.py     # rebuild the deliverable from current contacts.json
```

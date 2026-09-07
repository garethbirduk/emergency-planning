#!/usr/bin/env python3
"""Task-module library (build steps 1.3 + 2.1).

Turns the reader's Model + a trigger (death) date into dated Tasks:

- **Base tasks** — universal first-days steps (register the death, funeral
  director, locate the will…) emitted for every household, no tab needed.
- **Module builders** — one per tab. A module's tasks are generated iff its
  tab has data (AC-C1) — the reader already guarantees that by only putting
  populated tabs in ``model.modules``.
- **`joint?` / `owner` branching (AC-C3)** — a joint asset follows the
  survivor path (fewer steps); a sole asset owned by the deceased follows
  the probate path (more steps); a sole asset owned by someone else needs
  no action and is reported as a skip note, not silently dropped.
- **Death-relative date math** — every task is an offset in days from the
  trigger date; rerunning with a corrected date moves everything (AC-I3
  groundwork).
- **Stable task keys** — ``<row planKey>/<task-id>`` (or ``base/<task-id>``)
  so calendar items keep their identity across reruns (AC-I2 groundwork).

The deceased is the ``household`` row with role ``planholder``; an ``owner``
cell matches them by full name or first name, case-insensitive. Blank owner
on a sole asset is treated as the deceased's (conservative: more tasks).

Usage:  python tools/tasks.py <workbook.xlsx> --date YYYY-MM-DD
"""

import argparse
import re
from dataclasses import dataclass
from datetime import date, timedelta

from reader import Model, Row, read_workbook
from roles import describe as describe_roles
from roles import parse as parse_roles


@dataclass
class Task:
    plan_key: str      # stable identity: "<tab>/<row-slug>/<task-id>"
    due: date
    title: str
    module: str        # "base" or the source tab
    details: list      # extra human-readable lines (pointers etc.)
    # External-surface identity (SPEC §13). Local outputs use the fields
    # above; the .ics / Google use only these.
    visibility: str = "full"   # full | generic | none
    external_key: str = ""     # UID — no name-derived text when redacted
    external_title: str = ""   # SUMMARY — redacted when not "full"
    row_ref: str = ""          # "#3" reference (redacted rows, for legends)
    row_label: str = ""        # real label (redacted rows, for legends)
    # Ownership (SPEC §14): the functional role a task belongs to, and the
    # person it resolved to (empty when nobody is listed to assign to).
    role: str = "admin"
    assignee: str = ""
    # This step typically requires presenting a certified copy of the
    # death certificate — marked 📜 in the checklist and counted for
    # the how-many-copies estimate on the register-death task.
    needs_cert: bool = False

    def __post_init__(self):
        self.external_key = self.external_key or self.plan_key
        self.external_title = self.external_title or self.title


@dataclass
class Context:
    trigger: date
    deceased: str | None   # planholder's name from the household tab


def _day(ctx: Context, offset: int) -> date:
    return ctx.trigger + timedelta(days=offset)


def find_deceased(model: Model) -> str | None:
    for row in model.modules.get("household", []):
        role = (row.fields.get("role") or "").lower()
        if role == "planholder":
            return row.fields.get("name")
    return None


def owner_is_deceased(owner, deceased) -> bool:
    """Blank owner or no known planholder → assume the deceased's."""
    if not owner or not deceased:
        return True
    owner = owner.strip().lower()
    deceased = deceased.strip().lower()
    return owner in (deceased, deceased.split()[0])


# --- base (universal) tasks --------------------------------------------------

BASE_TASKS = [
    # (task-id, day offset, default role, title, details)
    ("verify-death", 0, "medical",
     "Obtain the medical certificate of cause of death",
     ["From the hospital or GP — needed to register the death."]),
    ("tell-family", 1, "local",
     "Tell close family; secure the home, pets and post", []),
    ("funeral-director", 2, "organiser",
     "Choose and contact a funeral director",
     ["Check the wishes tab / letter of wishes before agreeing details."]),
    ("locate-will", 3, "admin", "Locate the will and letter of wishes",
     ["See the key_documents tab for locations and custodians."]),
    ("register-death", 5, "admin",
     "Register the death (legal deadline: 5 days in England & Wales)",
     []),   # the how-many-copies advice is computed in derive_tasks
    ("tell-us-once", 5, "admin",
     "Use the government 'Tell Us Once' service",
     ["Notifies HMRC, DWP, DVLA, passport office and council in one go."]),
    ("assess-probate", 30, "admin",
     "Assess whether probate is needed; speak to the solicitor",
     ["Sole assets usually need a grant of probate; joint assets usually "
      "pass to the survivor."]),
]


def base_tasks(ctx: Context) -> list:
    return [Task(plan_key=f"base/{task_id}", due=_day(ctx, offset),
                 title=title, module="base", details=list(details),
                 role=role,
                 # the probate application itself wants a certificate
                 needs_cert=(task_id == "assess-probate"))
            for task_id, offset, role, title, details in BASE_TASKS]


# --- module builders ---------------------------------------------------------

# Generic per-tab noun for redacted aliases: "Notify building society #3"
# says what to do without saying to whom.
GENERIC_LABEL = {
    "banks": "bank", "building_societies": "building society",
    "cash_isas": "cash ISA", "shares_isas": "shares ISA",
    "share_portfolios": "shareholding", "pensions": "pension",
    "property": "property", "other_assets": "asset",
    "digital": "online service", "funeral_contacts": "contact",
    "wishes": "recorded wish", "key_documents": "document",
    "access_pointers": "access item", "funeral": "funeral arrangement",
    "professionals": "professional contact",
    "life_insurance": "life-insurance policy",
}


# Steps where UK practice typically wants a certified copy of the death
# certificate presented (institutions each want an original and return
# it). Central so the estimate and the 📜 markers can't drift.
CERT_NEEDED = {
    ("banks", "notify"), ("building_societies", "notify"),
    ("cash_isas", "notify"), ("shares_isas", "notify"),
    ("share_portfolios", "notify"), ("pensions", "notify"),
    ("life_insurance", "notify"), ("life_insurance", "start-claim"),
    ("life_insurance", "claim-benefits"),
    ("property", "notify-lender"), ("property", "land-registry-djp"),
}


def _alias(*parts) -> str:
    """Join alias parts, dropping blanks and (case-insensitive)
    duplicates — 'Monzo' + 'Monzo' is 'Monzo', not 'Monzo — Monzo'."""
    seen, kept = set(), []
    for part in parts:
        if part and part.strip() and part.strip().lower() not in seen:
            seen.add(part.strip().lower())
            kept.append(part.strip())
    return " — ".join(kept)


def _row_task(row: Row, task_id: str, ctx: Context, offset: int,
              template: str, alias: str = "", details=(),
              role: str = "admin") -> Task:
    """Build a task from one template with a single ``{alias}`` slot.

    The alias is the ONLY row data a title may carry. The full form is
    the row's name plus its generic reference — "Monzo (bank #1)" — so
    full and redacted labels visibly correspond; external surfaces get
    the full form only for a ``full`` row, otherwise just the reference
    ("bank #1"), built from the tab's fixed noun + the row id.
    Redaction by construction, never by stripping: a template plus a
    vocabulary-only alias cannot leak. A template without an
    ``{alias}`` slot carries no row data and is identical everywhere."""
    ref = row.external_ref
    noun = GENERIC_LABEL.get(row.tab, row.tab)
    generic_ref = f"{noun} #{ref}"
    # Blank name fields are fine — the row is then known purely by its
    # generic reference ("Notify life-insurance policy #1").
    title = template.format(
        alias=f"{alias} ({generic_ref})" if alias else generic_ref)
    if row.visibility == "full":
        redacted = {"external_title": title}
    else:
        generic = f"{noun} #{ref}"
        redacted = {
            "visibility": row.visibility,
            "external_key": f"{row.tab}/{ref}/{task_id}",
            "external_title": template.format(alias=generic),
            "row_ref": ref,
            "row_label": row.label,
        }
    return Task(plan_key=f"{row.plan_key}/{task_id}", due=_day(ctx, offset),
                title=title, module=row.tab,
                details=[d for d in details if d], role=role,
                needs_cert=(row.tab, task_id) in CERT_NEEDED, **redacted)


def _pointer(row: Row) -> str:
    loc = row.fields.get("location_of_details")
    return f"Details: {loc}" if loc else ""


def _deposit_account_tasks(row: Row, ctx: Context, inst: str):
    """Shared bank / building-society account logic (same UK process)."""
    label = row.fields.get("account_label") or ""
    alias = _alias(inst, label)
    if (row.fields.get("joint?") or "").lower() == "joint":
        # Survivor path: account passes to the survivor, no probate needed.
        yield _row_task(row, "notify", ctx, 7,
                        "Notify {alias}", alias,
                        ["Joint account — ask for it to move to the "
                         "survivor's sole name; take a death certificate.",
                         _pointer(row)])
        yield _row_task(row, "confirm-survivor", ctx, 30,
                        "Confirm {alias} is now in the survivor's sole "
                        "name", alias,
                        ["Check standing orders and direct debits still "
                         "work."])
    else:
        # Probate path: account freezes until the grant.
        yield _row_task(row, "notify", ctx, 7,
                        "Notify {alias}", alias,
                        ["Sole account — ask for the bereavement team; "
                         "the account will be frozen.", _pointer(row)])
        yield _row_task(row, "dod-balance", ctx, 14,
                        "Request date-of-death balance for {alias}", alias,
                        ["Needed for probate and inheritance-tax forms."])
        yield _row_task(row, "close-or-transfer", ctx, 90,
                        "After probate: close or transfer {alias}", alias,
                        ["Send the grant of probate; distribute per the "
                         "will."])


def banks_tasks(row: Row, ctx: Context):
    yield from _deposit_account_tasks(
        row, ctx, row.fields.get("bank") or "")


def building_societies_tasks(row: Row, ctx: Context):
    yield from _deposit_account_tasks(
        row, ctx, row.fields.get("society") or "")


def cash_isas_tasks(row: Row, ctx: Context):
    provider = row.fields.get("provider") or ""
    label = row.fields.get("account_label") or ""
    alias = _alias(provider, label)
    yield _row_task(row, "notify", ctx, 7,
                    "Notify {alias}", alias,
                    ["Ask for the bereavement team; the ISA's tax-free "
                     "status continues while the estate is administered.",
                     _pointer(row)])
    yield _row_task(row, "aps", ctx, 30,
                    "Ask about the APS allowance for {alias}", alias,
                    ["A surviving spouse can inherit an extra ISA allowance "
                     "equal to the balance (Additional Permitted "
                     "Subscription) — time-limited, so ask early."])
    yield _row_task(row, "close-or-transfer", ctx, 90,
                    "After probate: close or transfer {alias}", alias,
                    ["Send the grant of probate; distribute per the will."])


def shares_isas_tasks(row: Row, ctx: Context):
    provider = row.fields.get("provider") or ""
    label = row.fields.get("account_label") or ""
    alias = _alias(provider, label)
    yield _row_task(row, "notify", ctx, 7,
                    "Notify {alias}", alias,
                    ["Ask for the bereavement team; investments can usually "
                     "stay invested during administration.", _pointer(row)])
    yield _row_task(row, "dod-valuation", ctx, 14,
                    "Request date-of-death valuation for {alias}", alias,
                    ["Needed for probate and inheritance-tax forms."])
    yield _row_task(row, "aps", ctx, 30,
                    "Ask about the APS allowance for {alias}", alias,
                    ["A surviving spouse can inherit an extra ISA allowance "
                     "equal to the value (Additional Permitted "
                     "Subscription)."])
    yield _row_task(row, "transfer-or-sell", ctx, 90,
                    "After probate: transfer or sell {alias}", alias,
                    ["The survivor can often transfer in-specie into their "
                     "own account rather than selling."])


def share_portfolios_tasks(row: Row, ctx: Context):
    reg = row.fields.get("registrar_or_platform") or ""
    label = row.fields.get("holding_label") or ""
    alias = _alias(reg, label)
    if (row.fields.get("joint?") or "").lower() == "joint":
        # Survivor path: joint holdings pass by survivorship.
        yield _row_task(row, "notify", ctx, 7,
                        "Notify {alias}", alias,
                        ["Joint holding — ask for it to move to the "
                         "survivor's sole name; take a death certificate.",
                         _pointer(row)])
        yield _row_task(row, "confirm-survivor", ctx, 30,
                        "Confirm {alias} is in the survivor's sole name",
                        alias,
                        ["Check the dividend mandate still points at a "
                         "live bank account."])
    else:
        yield _row_task(row, "notify", ctx, 7,
                        "Notify {alias}", alias,
                        ["Sole holding — ask for the bereavement/"
                         "deceased-holder team.", _pointer(row)])
        yield _row_task(row, "dod-valuation", ctx, 14,
                        "Request date-of-death valuation for {alias}",
                        alias,
                        ["Share prices at the date of death; needed for "
                         "probate and inheritance-tax forms."])
        yield _row_task(row, "transfer-or-sell", ctx, 90,
                        "After probate: transfer or sell {alias}", alias,
                        ["Certificated holdings need the certificates — "
                         "see the pointer above if given."])


def pensions_tasks(row: Row, ctx: Context):
    provider = row.fields.get("provider") or ""
    label = row.fields.get("pension_label") or ""
    ptype = (row.fields.get("type") or "").lower()
    alias = _alias(provider, label)
    if ptype == "state":
        # DWP is notified via Tell Us Once (a base task); confirm + claim.
        # No {alias} slot: these titles carry no row data on any surface.
        yield _row_task(row, "confirm-dwp", ctx, 7,
                        "Confirm 'Tell Us Once' has notified the DWP "
                        "(State Pension)",
                        details=["Stops payments — overpayments are "
                                 "clawed back from the estate."])
        yield _row_task(row, "bereavement-benefits", ctx, 30,
                        "Ask the DWP about bereavement/survivor benefits",
                        details=["Bereavement Support Payment must "
                                 "usually be claimed within 3 months for "
                                 "the full amount."])
    else:
        notify_details = ["Ask the bereavement team what death benefits "
                          "apply and what documents they need.",
                          _pointer(row)]
        if ptype == "annuity":
            notify_details.insert(
                1, "Ask whether the annuity has a spouse's pension or a "
                   "guarantee period, or simply stops.")
        yield _row_task(row, "notify", ctx, 7,
                        "Notify {alias}", alias, notify_details)
        if (row.fields.get("nomination_in_place?") or "").upper() == "Y":
            claim_detail = ("A nomination/expression of wish is in place — "
                            "benefits should pay to the nominee, usually "
                            "outside the estate.")
        else:
            claim_detail = ("No nomination recorded — the trustees decide "
                            "at their discretion; ask the solicitor who "
                            "should claim.")
        yield _row_task(row, "claim-benefits", ctx, 30,
                        "Claim death benefits from {alias}", alias,
                        [claim_detail])


def other_assets_tasks(row: Row, ctx: Context):
    label = row.fields.get("asset_label") or ""
    atype = (row.fields.get("type") or "").lower()
    joint = (row.fields.get("joint?") or "").lower() == "joint"
    where = row.fields.get("location")
    loc_detail = f"Location: {where}" if where else ""
    if atype == "vehicle":
        yield _row_task(row, "dvla-insurer", ctx, 3,
                        "Tell DVLA and the insurer about {alias}", label,
                        ["A policy in the deceased's sole name may lapse — "
                         "nobody should drive it until cover is confirmed.",
                         loc_detail, _pointer(row)])
        if not joint:
            yield _row_task(row, "transfer-or-sell", ctx, 60,
                            "Transfer or sell {alias}", label,
                            ["Update the V5C with DVLA; the vehicle is "
                             "usually part of the estate."])
    else:
        yield _row_task(row, "insurance-valuation", ctx, 14,
                        "Check insurance and valuation for {alias}", label,
                        ["Keep cover in force; a valuation may be needed "
                         "for the estate.", loc_detail, _pointer(row)])
        if not joint:
            yield _row_task(row, "estate-inventory", ctx, 30,
                            "Include {alias} in the estate inventory",
                            label,
                            ["Sole items form part of the estate for "
                             "probate and inheritance tax."])


def property_tasks(row: Row, ctx: Context):
    label = row.fields.get("property_label") or ""
    insurer = row.fields.get("insurer") or ""
    deeds = row.fields.get("location_of_deeds")
    joint = (row.fields.get("joint?") or "").lower() == "joint"
    yield _row_task(row, "notify-insurer", ctx, 7,
                    "Notify insurer {alias}", _alias(insurer, label),
                    ["If the property is now unoccupied, check the policy's "
                     "unoccupied-property terms."])
    if (row.fields.get("mortgage?") or "").upper() == "Y":
        yield _row_task(row, "notify-lender", ctx, 14,
                        "Notify the mortgage lender for {alias}", label,
                        ["Ask how payments continue during the estate "
                         "administration."])
    if joint:
        # Survivor path: title passes by survivorship; just update the
        # register.
        yield _row_task(row, "land-registry-djp", ctx, 30,
                        "Send Land Registry form DJP to remove the "
                        "deceased from the title of {alias}", label,
                        ["Free; needs an official copy of the death "
                         "certificate.",
                         f"Deeds: {deeds}" if deeds else ""])
        if (row.fields.get("tenancy") or "").lower() == "tenants_in_common":
            yield _row_task(row, "tic-share", ctx, 30,
                            "Tenants in common: check the will for the "
                            "deceased's share of {alias}", label,
                            ["The share does NOT pass automatically to the "
                             "survivor — speak to the solicitor."])
    else:
        # Probate path: property is part of the estate.
        yield _row_task(row, "secure-property", ctx, 7,
                        "Secure and maintain {alias} until probate", label,
                        ["Keep insurance valid; clear post; check regularly "
                         "if empty."])
        yield _row_task(row, "transfer-or-sell", ctx, 90,
                        "After probate: transfer or sell {alias}", label,
                        ["Solicitor handles the Land Registry transfer "
                         "(AS1) or sale.",
                         f"Deeds: {deeds}" if deeds else ""])


def digital_tasks(row: Row, ctx: Context):
    service = row.fields.get("service") or ""
    action = (row.fields.get("action_on_death") or "").lower()
    plans = {
        # action -> (day offset, title template, detail)
        "cancel": (30, "Cancel {alias}",
                   "Stop the payments too — check bank statements for the "
                   "direct debit or card charge."),
        "memorialise": (30, "Memorialise {alias}",
                        "Most platforms have a dedicated memorialisation / "
                        "legacy-contact process for a death."),
        "transfer": (60, "Transfer {alias}",
                     "Start early — transfers (e.g. domains) lapse if the "
                     "renewal date passes first."),
        "archive": (30, "Archive {alias}",
                    "Download or export anything worth keeping before "
                    "closing it."),
        "keep": (14, "Keep {alias} running",
                 "Move billing and ownership across so it doesn't lapse "
                 "with the deceased's cards or email account."),
    }
    offset, template, detail = plans.get(
        action, (30, "Decide what to do with {alias}",
                 "No action_on_death recorded in the workbook — choose "
                 "cancel / memorialise / transfer / archive / keep."))
    yield _row_task(row, "action", ctx, offset, template, service,
                    [detail, _pointer(row)])


def wishes_tasks(row: Row, ctx: Context):
    topic = row.fields.get("topic") or ""
    wish = row.fields.get("wish")
    documented = row.fields.get("documented_where")
    urgent = topic.lower() == "organ_donation"
    details = [f"Wish: {wish}" if wish else "",
               f"Documented: {documented}" if documented else ""]
    if urgent:
        details.insert(0, "Organ/tissue donation is time-critical — tell "
                          "the hospital team immediately.")
    yield _row_task(row, "honour", ctx, 0 if urgent else 1,
                    "Check recorded wish: {alias}", topic, details,
                    role="medical" if urgent else "organiser")


def key_documents_tasks(row: Row, ctx: Context):
    document = row.fields.get("document") or ""
    location = row.fields.get("location")
    custodian = row.fields.get("custodian")
    details = [f"Location: {location}" if location else "",
               f"Custodian: {custodian}" if custodian else ""]
    if document.lower().startswith("lpa"):
        details.append("An LPA ends at death — it cannot be used from now "
                       "on; the executor acts under the will instead.")
    elif document.lower() == "letter_of_wishes":
        details.append("Not legally binding, but it guides the executor "
                       "and family — read it before decisions are made.")
    yield _row_task(row, "locate", ctx, 3, "Locate: {alias}", document,
                    details)


# A pointer cell that appears to contain an ACTUAL secret (SPEC §4:
# pointers, never secrets). Matches "password: x", "PIN is 1234" etc., but
# not prose like "password manager" or "never the password itself".
SECRET_RE = re.compile(
    r"\b(password|passphrase|pin|seed(?:\s+phrase)?|security\s+answer)\b"
    r"\s*(?:is|was|[:=])\s*\S", re.IGNORECASE)


def access_pointers_tasks(row: Row, ctx: Context):
    what = row.fields.get("what") or ""
    where = row.fields.get("where_to_find")
    custodian = row.fields.get("custodian")
    yield _row_task(row, "locate", ctx, 1, "Locate access: {alias}", what,
                    [f"Where: {where}" if where else "",
                     f"Custodian: {custodian}" if custodian else ""])


def find_secret(row: Row):
    """Return the offending cell text if a row seems to hold a real secret."""
    for value in row.fields.values():
        if isinstance(value, str) and SECRET_RE.search(value):
            return value
    return None


def funeral_tasks(row: Row, ctx: Context):
    """A populated funeral tab means the funeral is pre-arranged: contact
    that provider instead of choosing one (the generic base task is
    removed in derive_tasks — SPEC §3)."""
    provider = row.fields.get("provider") or ""
    package = row.fields.get("package")
    venue = row.fields.get("venue")
    hint = row.fields.get("contact_hint")
    paperwork = row.fields.get("location_of_paperwork")
    details = [f"Package: {package}" if package else "",
               f"Venue: {venue}" if venue else ""]
    if (row.fields.get("prepaid?") or "").upper() == "Y":
        details.append("Pre-paid plan — the package is already paid for; "
                       "don't arrange or pay for anything twice.")
    details += [f"How to reach them: {hint}" if hint else "",
                f"Paperwork: {paperwork}" if paperwork else ""]
    yield _row_task(row, "contact-provider", ctx, 2,
                    "Contact the pre-arranged funeral director: {alias}",
                    provider, details, role="organiser")


def professionals_tasks(row: Row, ctx: Context):
    """Type-driven: each kind of professional gets its own follow-up."""
    name = row.fields.get("name") or ""
    ptype = (row.fields.get("type") or "").lower()
    hint = row.fields.get("contact_hint")
    hint_detail = f"How to reach them: {hint}" if hint else ""
    if ptype == "solicitor":
        yield _row_task(row, "notify", ctx, 3,
                        "Tell {alias} (solicitor) about the death", name,
                        ["They may hold the will or deeds — see the "
                         "key_documents tab.", hint_detail])
        yield _row_task(row, "engage-probate", ctx, 30,
                        "Engage {alias} for probate, or decide to handle "
                        "it yourselves", name,
                        ["Pairs with the 'assess whether probate is "
                         "needed' step — take the asset list along."])
    elif ptype == "accountant":
        yield _row_task(row, "notify", ctx, 7,
                        "Tell {alias} (accountant) about the death", name,
                        [hint_detail])
        yield _row_task(row, "final-tax", ctx, 60,
                        "Ask {alias} to settle tax affairs to the date of "
                        "death", name,
                        ["Income tax up to the date of death, plus the "
                         "estate's own returns if administration runs "
                         "on."])
    elif ptype in ("financial_adviser", "adviser"):
        yield _row_task(row, "notify", ctx, 7,
                        "Tell {alias} (financial adviser) about the death",
                        name, [hint_detail])
        yield _row_task(row, "asset-list", ctx, 14,
                        "Ask {alias} for a list of every account and "
                        "policy they know of", name,
                        ["Advisers often know of assets the family "
                         "doesn't — cross-check against this plan's "
                         "tabs."])
    else:
        yield _row_task(row, "notify", ctx, 7,
                        "Tell {alias} about the death", name,
                        [hint_detail])


def life_insurance_tasks(row: Row, ctx: Context):
    """Policies and death-in-service — the classic 'nobody knew it
    existed' asset (SPEC §3)."""
    provider = row.fields.get("provider") or ""
    label = row.fields.get("policy_label") or ""
    alias = _alias(provider, label)
    ltype = (row.fields.get("type") or "").lower()
    nominated = (row.fields.get("nomination_in_place?") or "").upper() == "Y"
    if ltype == "death_in_service":
        if nominated:
            nom_detail = ("An expression of wish is in place — the "
                          "benefit should pay to the nominee, usually "
                          "outside the estate.")
        else:
            nom_detail = ("No expression of wish recorded — the scheme "
                          "trustees decide at their discretion; ask the "
                          "solicitor who should claim.")
        yield _row_task(row, "notify", ctx, 3,
                        "Tell {alias} (employer) about the "
                        "death-in-service benefit", alias,
                        ["Ask HR even if unsure a benefit exists — "
                         "death-in-service is the classic benefit nobody "
                         "knew about.", nom_detail, _pointer(row)])
        yield _row_task(row, "claim-benefits", ctx, 30,
                        "Claim the death-in-service benefit from {alias}",
                        alias,
                        ["Usually a multiple of salary, paid by the "
                         "scheme trustees outside the estate — no grant "
                         "of probate needed."])
    else:
        if nominated:
            trust_detail = ("Nominated beneficiary / written in trust — "
                            "should pay out directly, without waiting "
                            "for probate.")
        else:
            trust_detail = ("Check whether the policy is written in "
                            "trust; if not, it pays into the estate and "
                            "may need the grant of probate.")
        yield _row_task(row, "start-claim", ctx, 7,
                        "Notify {alias} and start the claim", alias,
                        ["Ask what they need — usually a death "
                         "certificate and the policy paperwork.",
                         trust_detail, _pointer(row)])
        yield _row_task(row, "chase-claim", ctx, 45,
                        "Chase the life-insurance payout from {alias}",
                        alias,
                        ["A payout can ease the estate's cash flow while "
                         "probate is still in progress."])


def funeral_contacts_tasks(row: Row, ctx: Context):
    name = row.fields.get("name") or ""
    invite_to = (row.fields.get("invite_to") or "both").lower()
    hint = row.fields.get("contact_hint")
    reach = f"How to reach them: {hint}" if hint else ""
    # invite_to is dropdown-constrained vocabulary, safe in a template.
    # Same task id either way: flipping the value updates the calendar
    # event in place rather than minting a new one.
    if invite_to == "notify_only":
        yield _row_task(row, "invite", ctx, 5,
                        "Tell {alias} the news", name,
                        ["Marked notify-only — let them know, but they "
                         "are not on the funeral or wake invite list.",
                         reach],
                        role="comms")
        return
    what = {"funeral": "the funeral", "wake": "the wake"}.get(
        invite_to, "the funeral and wake")
    yield _row_task(row, "invite", ctx, 5,
                    f"Invite {{alias}} to {what}", name, [reach],
                    role="comms")


# Asset-style modules get `joint?`/`owner` branching (a tab without a
# `joint?` column is implicitly sole, so owner branching still applies);
# other modules emit for every row regardless of owner.
ASSET_BUILDERS = {
    "banks": banks_tasks,
    "building_societies": building_societies_tasks,
    "cash_isas": cash_isas_tasks,
    "shares_isas": shares_isas_tasks,
    "share_portfolios": share_portfolios_tasks,
    "pensions": pensions_tasks,
    "property": property_tasks,
    "other_assets": other_assets_tasks,
    "digital": digital_tasks,
    "life_insurance": life_insurance_tasks,
}
OTHER_BUILDERS = {
    "funeral": funeral_tasks,
    "professionals": professionals_tasks,
    "funeral_contacts": funeral_contacts_tasks,
    "wishes": wishes_tasks,
    "key_documents": key_documents_tasks,
    "access_pointers": access_pointers_tasks,
}


def _assignable_people(model: Model) -> tuple:
    """(people, unrecognised) where people is (name, legal_role,
    relationship, held_roles) for everyone except the planholder, in
    sheet order, and unrecognised is (name, token) for every
    crisis_roles value that matches no known role. The legal role also
    counts as held (SPEC §14).

    Roles are an enum (``roles.py``): a misspelt one used to make the
    role silently unheld, which sent its tasks to the fallback person
    while the plan still read as correct. Now it's a note, not a shrug.
    """
    people, unrecognised = [], []
    for row in model.modules.get("household", []):
        name = row.fields.get("name")
        legal = (row.fields.get("role") or "").strip().lower()
        if not name or legal == "planholder":
            continue
        rel = (row.fields.get("relationship") or "").strip().lower()
        crisis, unknown = parse_roles(row.fields.get("crisis_roles"))
        unrecognised += [(name, token) for token in unknown]
        held = set(crisis)
        if legal:
            held.add(legal)
        people.append((name, legal, rel, held))
    return people, unrecognised


def assign_owners(model: Model, tasks, notes) -> None:
    """Resolve each task's role to a person (AC-R1). Unheld roles fall back
    executor → spouse/partner (by relationship — 'role' is legal standing
    only) → first person, with one visible note per role — never silently.
    Nobody listed → tasks stay unowned."""
    people, unrecognised = _assignable_people(model)
    for who, token in unrecognised:
        notes.append(
            f"household: {who} lists the crisis role '{token}', which is "
            f"not one of {describe_roles()} — it was ignored, so nobody "
            "counts as holding it. Fix the spelling on the household tab "
            "(the wizard offers the roles as tick-boxes).")
    if not people:
        return
    # legal == "spouse" survives for workbooks made before role became
    # legal-standing-only; new ones record it in the relationship column.
    fallback = next(
        (n for n, legal, _, _ in people if legal == "executor"),
        None) or next(
        (n for n, legal, rel, _ in people
         if legal == "spouse"
         or any(w in rel for w in ("spouse", "husband", "wife",
                                   "partner"))),
        None) or people[0][0]
    noted = set()
    for task in tasks:
        holder = next((n for n, _, _, held in people if task.role in held),
                      None)
        if holder is None:
            holder = fallback
            if task.role not in noted:
                noted.add(task.role)
                notes.append(
                    f"role '{task.role}': held by nobody on the household "
                    f"tab — its tasks default to {fallback}")
        task.assignee = holder


def derive_tasks(model: Model, trigger: date):
    """Return (tasks, notes): dated tasks + skip notes for transparency."""
    ctx = Context(trigger=trigger, deceased=find_deceased(model))
    tasks = base_tasks(ctx)
    notes = []

    # A populated funeral tab pre-arms the funeral: its contact-provider
    # task replaces the generic "choose a funeral director" base task
    # (SPEC §3) — noted, never silent.
    if "funeral" in model.modules:
        tasks = [t for t in tasks if t.plan_key != "base/funeral-director"]
        notes.append(
            "funeral: a pre-arranged funeral is recorded — the generic "
            "'choose and contact a funeral director' task is replaced by "
            "contacting the provider on the funeral tab")

    def rejected_as_secret(row: Row) -> bool:
        # Never echo the offending cell content — it may BE the secret.
        if find_secret(row):
            notes.append(
                f"{row.plan_key}: WARNING — a cell looks like it contains "
                "an actual secret, so the row was left OUT of the plan. "
                "Move the secret to the password manager and keep only a "
                "pointer in the workbook (see PRIVACY.md).")
            return True
        return False

    def warn_if_id_missing(row: Row) -> None:
        # A redacted row with no id gets an opaque hash ref — works, but
        # the human should fill the id in (SPEC §13).
        if row.visibility != "full" and not row.row_id:
            notes.append(
                f"{row.plan_key}: WARNING — no id set; calendar entries "
                f"use opaque ref #{row.external_ref}. Fill the id column "
                "in for a readable reference.")

    for tab, builder in ASSET_BUILDERS.items():
        for row in model.modules.get(tab, []):
            if rejected_as_secret(row):
                continue
            sole = (row.fields.get("joint?") or "").lower() != "joint"
            owner = row.fields.get("owner")
            if sole and not owner_is_deceased(owner, ctx.deceased):
                notes.append(f"{row.plan_key}: sole asset owned by {owner} "
                             f"(not the deceased) — no action needed")
                continue
            warn_if_id_missing(row)
            tasks.extend(builder(row, ctx))
    for tab, builder in OTHER_BUILDERS.items():
        for row in model.modules.get(tab, []):
            if rejected_as_secret(row):
                continue
            warn_if_id_missing(row)
            tasks.extend(builder(row, ctx))
    tasks.sort(key=lambda t: (t.due, t.plan_key))
    assign_owners(model, tasks, notes)

    # How many certified death-certificate copies to order: estimated
    # from the steps in THIS plan that need one (marked 📜), instead of
    # the traditional folklore number.
    cert_count = sum(1 for t in tasks if t.needs_cert)
    for t in tasks:
        if t.plan_key == "base/register-death":
            if cert_count:
                recommend = max(3, min(cert_count, 10))
                t.details.append(
                    f"Order {recommend} certified copies of the death "
                    f"certificate — {cert_count} step(s) in this plan "
                    "(marked 📜 in the checklist) each want one; "
                    "institutions return them, so they can be reused.")
            else:
                t.details.append(
                    "Order 3 certified copies of the death certificate "
                    "— institutions want originals, and extras save "
                    "waiting for returns.")
    return tasks, notes


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Derive dated crisis-plan tasks from a workbook.")
    ap.add_argument("workbook", help="path to the .xlsx (sample only in dev)")
    ap.add_argument("--date", required=True, metavar="YYYY-MM-DD",
                    help="trigger (death) date")
    args = ap.parse_args()

    trigger = date.fromisoformat(args.date)
    model = read_workbook(args.workbook)
    tasks, notes = derive_tasks(model, trigger)

    print(f"Plan for trigger date {trigger} — {len(tasks)} tasks\n")
    for t in tasks:
        print(f"{t.due}  [{t.module}] {t.title}")
        print(f"          key: {t.plan_key}")
        for d in t.details:
            print(f"          - {d}")
    if notes:
        print("\nSkipped (no action needed):")
        for n in notes:
            print(f"  {n}")


if __name__ == "__main__":
    main()

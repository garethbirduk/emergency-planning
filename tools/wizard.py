#!/usr/bin/env python3
"""Onboarding wizard (build step 5.1, SPEC §15).

A form-style interview that produces a household's filled workbook, so
nobody has to face a 17-tab spreadsheet cold. Design rules (SPEC §15):

- **Local-only.** A stdlib HTTP server on 127.0.0.1 serves one page to
  your own browser; nothing leaves the machine (answers ARE the
  sensitive data — a cloud form would leak the plan before it exists).
- **An authoring front-end, not a new store.** The output IS a normal
  workbook, written once via the template generator; the engine and
  privacy model are untouched. Visibility columns get the safe
  defaults (generic / not AI-visible) — tune them in the spreadsheet.
- **Write-once, never edit.** An existing file is never overwritten;
  re-running the wizard writes a fresh file to adopt by hand.
- **No drift.** The form is generated from make_template's TABS
  definitions — columns and dropdowns come from the same source as the
  spreadsheet itself.
- **Finishes with the calendar (optional).** Tick the box, give family
  emails, and the wizard runs the one-time gsetup after writing the
  workbook (browser consent appears if not yet authorised).

Usage:  python tools/wizard.py        (opens the form in your browser)
"""

import html
import json
import re
import shutil
import subprocess
import sys
import threading
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import make_template
from augment import INSTITUTION_COLUMN, QUERY_CONTEXT
from make_template import TABS, TAB_ORDER
from openpyxl import Workbook
from reader import slugify
from roles import CRISIS_ROLES, format_roles
from roles import parse as parse_roles
from tasks import GENERIC_LABEL

REPO_ROOT = Path(__file__).resolve().parents[1]
META = {"id", "online_visibility", "ai_visible?"}
# Columns holding several values from a fixed set: rendered as tick-boxes
# (one posted value each), stored comma-separated in one cell. The list
# comes from roles.py so the form and the engine can't drift apart.
MULTI_ENUM = {("household", "crisis_roles"): CRISIS_ROLES}

# Which columns compose a row's on-calendar alias, per tab — a preview
# approximation of the task builders' alias choices (tasks.py), used
# only for the live "shown on the shared calendar as" hint.
# Built-in pre-crisis checklist suggestions: the wizard seeds these
# into a fresh workbook's preparation tab (tickable, closeable as "not
# needed", extendable with custom items/groups). Deliberately pointer-
# shaped (nothing invites writing a secret) and judgment-free; several
# items are "get X into the crisis tabs" and are worded to match those
# tabs so a future 5.4 cross-link can auto-clear them.
PREP_CATALOG = {
    "Write or update the will": [
        "List all assets and roughly value the estate",
        "Decide the beneficiaries and what each should receive",
        "Talk it through with the beneficiaries — no surprises later",
        "Choose the executor(s) and ask if they're willing",
        "Ask the solicitor about inheritance-tax implications",
        "Draft it with a solicitor (or a reputable will-writing "
        "service)",
        "Sign with two independent witnesses; store it and record the "
        "location in key_documents",
    ],
    "Paperwork & legal": [
        "Set up and register LPAs — finance and health (they take "
        "weeks)",
        "Write a letter of wishes to sit with the will",
        "Gather certificates (birth, marriage) into one known place",
    ],
    "Finances": [
        "List every account, pension and policy into the workbook",
        "Find passbooks, share certificates and policy documents",
        "Review pension and life-policy nominations / expressions of "
        "wish",
        "Check for lost accounts and unclaimed pensions",
        "Audit direct debits — what is actually being paid for?",
    ],
    "Security & digital": [
        "Get every password into the password manager",
        "Fix reused passwords, then weak ones",
        "Turn on 2FA for email and banking",
        "Set legacy contacts (Google Inactive Account Manager, Apple "
        "Legacy Contact)",
        "Record where the password-manager emergency kit lives",
    ],
    "Professional services": [
        "Choose a solicitor and record them",
        "Pick a preferred funeral director; consider a pre-paid plan",
        "Decide whether an accountant or adviser review is needed",
        "Record GP and pharmacy details",
    ],
    "Health & medical": [
        "Know where the medical records are held; keep a pointer in "
        "access_pointers",
        "Keep a medication list with the GP and in the safe — record "
        "WHERE, not the list itself",
        "Discuss DNACPR / treatment wishes; record them on the wishes "
        "tab and tell the GP",
        "Check the organ-donor register decision and record it",
        "Discuss hospital-vs-home care preferences and write them "
        "down",
    ],
    "Arrange cleaner": [
        "Advertise for a cleaner",
        "Interview candidates",
        "Agree terms and start",
    ],
    "Home accessibility": [
        "Assess stairs, steps and trip hazards",
        "Consider a stairlift (if wanted)",
        "Remove loose rugs; add grab rails",
    ],
    "Home & help": [
        "Consider a personal alarm / falls pendant and a keysafe",
        "Investigate home-care agencies",
        "Investigate residential-care options and costs (research "
        "only)",
    ],
    "Family & communication": [
        "Agree who holds which crisis role",
        "Tell the family where the workbook and printed plan live",
        "Set an annual review date",
    ],
}

ALIAS_COLS = {
    "banks": ["bank", "account_label"],
    "building_societies": ["society", "account_label"],
    "cash_isas": ["provider", "account_label"],
    "shares_isas": ["provider", "account_label"],
    "share_portfolios": ["registrar_or_platform", "holding_label"],
    "pensions": ["provider", "pension_label"],
    "property": ["property_label"],
    "other_assets": ["asset_label"],
    "digital": ["service"],
    "funeral_contacts": ["name"],
    "wishes": ["topic"],
    "key_documents": ["document"],
    "access_pointers": ["what"],
    "funeral": ["provider"],
    "professionals": ["name"],
    "life_insurance": ["provider", "policy_label"],
}

# Plain string on purpose (NOT an f-string): real braces, no doubling.
SCRIPT_JS = """
function addRow(tab, scaffold) {
  const tpl = document.getElementById('tpl_' + tab);
  const container = document.getElementById('rows_' + tab);
  const idx = container.children.length;
  const div = document.createElement('div');
  div.className = 'row';
  div.innerHTML = tpl.innerHTML.replaceAll('__IDX__', '__' + idx + '__');
  container.appendChild(div);
  div.addEventListener('input',
                       () => { refreshPreview(div, tab);
                               updateCount(tab); });
  div.addEventListener('change',
                       () => { refreshPreview(div, tab);
                               updateCount(tab); });
  // Button-added rows get a unique scaffold name in the first column,
  // so adding four rows yields four distinct rows even before editing.
  if (scaffold && ALIAS[tab]) {
    const first = div.querySelector(
        '[name="' + tab + '__' + idx + '__' + ALIAS[tab][0] + '"]');
    if (first && !first.value) first.value = tab + '_' + (idx + 1);
  }
  refreshPreview(div, tab);
  updateCount(tab);
}

function removeRow(btn) {
  const row = btn.closest('.row');
  const tab = row.parentNode.id.slice(5);
  row.remove();
  updateCount(tab);
}

// Live "[n]" in the section heading: rows with any real content.
function updateCount(tab) {
  const el = document.getElementById('count_' + tab);
  if (!el) return;
  let n = 0;
  for (const row of document.getElementById('rows_' + tab).children) {
    const fields = row.querySelectorAll(
        'input:not([type=checkbox]):not([type=hidden]), select');
    for (const f of fields) {
      if (f.value && f.value.trim()) { n += 1; break; }
    }
  }
  el.textContent = '[' + n + ']';
}

// Live preview of the row's calendar label: generic "bank account #1"
// until the details checkbox is ticked, then the real alias.
function refreshPreview(div, tab) {
  const el = div.querySelector('.gpreview');
  if (!el) return;
  const idx = Array.prototype.indexOf.call(div.parentNode.children, div);
  const field = c => div.querySelector(
      '[name="' + tab + '__' + idx + '__' + c + '"]');
  const box = field('online_visibility');
  const idField = field('id');
  const num = (idField && idField.value) ? idField.value : (idx + 1);
  const ref = (GENERIC[tab] || tab) + ' #' + num;
  const noneField = field('vis_none');
  if (noneField && noneField.value === 'none'
      && !(box && box.checked)) {
    el.textContent = 'kept off the calendar entirely (' + ref
                     + ' in the local checklist)';
    return;
  }
  if (box && box.checked) {
    // Same format as the engine: deduped name + generic reference,
    // e.g. "Monzo (bank #1)" — never "Monzo — Monzo".
    const seen = new Set();
    const parts = (ALIAS[tab] || []).map(field)
        .map(f => f && f.value.trim()).filter(Boolean)
        .filter(v => !seen.has(v.toLowerCase())
                     && seen.add(v.toLowerCase()));
    el.textContent = parts.length
        ? parts.join(' \\u2014 ') + ' (' + ref + ')'
        : ref + ' (fill the row in for the full label)';
  } else {
    el.textContent = ref;
  }
}

// Sections start with no rows — the first "+ add" creates tab_1, so
// there is never an unnamed empty row. Household is the exception: it
// gets a starter row for the plan-for name mirror to fill.
addRow('household');

// Editing an existing workbook: populate the rows we loaded. Values
// are marked dirty so the planholder mirror never overwrites them.
function fillRow(tab, idx, row) {
  const div = document.getElementById('rows_' + tab).children[idx];
  const field = c => div.querySelector(
      '[name="' + tab + '__' + idx + '__' + c + '"]');
  for (const [col, val] of Object.entries(row)) {
    if (col.startsWith('_')) continue;
    const fs = div.querySelectorAll(
        '[name="' + tab + '__' + idx + '__' + col + '"]');
    if (!fs.length) continue;
    if (fs[0].type === 'checkbox') {
      // Multi-value enum (crisis_roles): the loader already reduced the
      // cell to canonical names, so an exact match is right here.
      const want = String(val || '').split(',').map(s => s.trim());
      for (const f of fs) {
        f.checked = want.includes(f.value);
        f.dataset.dirty = '1';
      }
    } else {
      fs[0].value = val;
      fs[0].dataset.dirty = '1';
    }
  }
  // Never drop a value silently: say what the workbook held that no
  // role matched, since saving this form will drop it.
  const warn = div.querySelector('.rolewarn');
  if (warn && row._badroles && row._badroles.length) {
    warn.textContent = '⚠ ignored, not a known role: '
        + row._badroles.join(', ') + ' — tick what was meant.';
  }
  const idf = field('id');
  if (idf) idf.value = row._id || '';
  const vis = field('online_visibility');
  if (vis) vis.checked = row._vis === 'full';
  const noneF = field('vis_none');
  if (noneF && row._vis === 'none') noneF.value = 'none';
  const ai = field('ai_visible?');
  if (ai) ai.checked = !!row._ai;
  refreshPreview(div, tab);
  updateCount(tab);
}
if (PREFILL) {
  for (const [tab, rows] of Object.entries(PREFILL)) {
    if (tab === 'preparation') continue;   // bespoke pane handles it
    for (let i = 0; i < rows.length; i++) {
      if (document.getElementById('rows_' + tab).children.length <= i)
        addRow(tab);
      fillRow(tab, i, rows[i]);
    }
  }
}

// Enter never submits the form (far too destructive mid-typing — only
// the submit button writes); in the checklist add-boxes it adds the
// item/group instead.
document.querySelector('form').addEventListener('submit',
                                                () => { leaving = true; });
document.querySelector('form').addEventListener('keydown', e => {
  if (e.key !== 'Enter' || e.target.tagName !== 'INPUT') return;
  e.preventDefault();
  if (e.target.classList.contains('padd'))
    prepAddCustom(e.target.parentNode.querySelector('button'));
  else if (e.target.id === 'prep_newgroup') prepNewGroup();
});

// --- pre-crisis / crisis panes -----------------------------------------

function showPane(which) {
  for (const p of document.querySelectorAll('.pane'))
    p.classList.remove('active');
  for (const b of document.querySelectorAll('.tabs button'))
    b.classList.remove('active');
  document.getElementById('pane_' + which).classList.add('active');
  document.getElementById('tabbtn_' + which).classList.add('active');
}

// --- pre-crisis checklist ----------------------------------------------
// Items live as preparation__N__* fields; ticking sets status=done,
// closing (x) sets status=not_needed and moves the item to the bottom
// list, restore brings it back. All hidden columns (type, priority,
// notes) round-trip untouched through edits.

let prepIdx = 0;

function prepGroupBody(group) {
  const key = 'pg_' + group;
  let fs = document.getElementById(key);
  if (!fs) {
    fs = document.createElement('fieldset');
    fs.className = 'pgroup';
    fs.id = key;
    fs.innerHTML = '<legend><b></b> ' +
      '<button type="button" title="whole group not needed" ' +
      'onclick="prepCloseGroup(this)">×</button></legend>' +
      '<div class="pbody"></div>' +
      '<p><input class="padd" size="30" placeholder="new item"> ' +
      '<button type="button" onclick="prepAddCustom(this)">+ add ' +
      'item</button></p>';
    fs.querySelector('legend b').textContent = group || 'General';
    if (!group) fs.querySelector('legend button').style.display = 'none';
    document.getElementById('prep_active').appendChild(fs);
  }
  return fs.querySelector('.pbody');
}

function prepStatus(li) {
  const closed = li.closest('#prep_closed') !== null;
  const done = li.querySelector('.pdone').checked;
  li.querySelector('.pstatus').value =
      closed ? 'not_needed' : (done ? 'done' : '');
}

// The not-needed area mirrors the group structure, so a whole group
// can be restored in one go; empty group boxes tidy themselves away.
function prepClosedBody(group) {
  const key = 'pgc_' + group;
  let fs = document.getElementById(key);
  if (!fs) {
    fs = document.createElement('fieldset');
    fs.className = 'pgroup';
    fs.id = key;
    fs.innerHTML = '<legend><b></b> ' +
      '<button type="button" onclick="prepRestoreGroup(this)">' +
      '↩ restore group</button></legend>' +
      '<div class="pbody"></div>';
    fs.querySelector('legend b').textContent = group || 'General';
    document.getElementById('prep_closed').appendChild(fs);
  }
  return fs.querySelector('.pbody');
}

function prepTidyClosed() {
  for (const fs of document.querySelectorAll('#prep_closed .pgroup'))
    if (!fs.querySelector('.pitem')) fs.remove();
}

function prepRestoreGroup(btn) {
  const fs = btn.closest('.pgroup');
  for (const li of [...fs.querySelectorAll('.pitem')]) {
    prepGroupBody(li.querySelector('[name$="__group"]').value)
        .appendChild(li);
    prepStatus(li);
  }
  prepTidyClosed();
}

function prepItem(group, item, extra) {
  extra = extra || {};
  const i = prepIdx++;
  const li = document.createElement('div');
  li.className = 'pitem';
  const n = c => 'preparation__' + i + '__' + c;
  li.innerHTML =
    '<input type="checkbox" class="pdone">' +
    '<input name="' + n('item') + '" size="57">' +
    '<input name="' + n('owner') + '" size="10" placeholder="owner">' +
    '<input name="' + n('notes') + '" size="28" ' +
    'placeholder="notes / outcome">' +
    '<input type="hidden" name="' + n('group') + '">' +
    '<input type="hidden" name="' + n('type') + '">' +
    '<input type="hidden" name="' + n('priority') + '">' +
    '<input type="hidden" name="' + n('id') + '">' +
    '<input type="hidden" class="pstatus" name="' + n('status') + '">' +
    '<button type="button" class="pclose" title="not needed">×' +
    '</button>' +
    '<button type="button" class="prestore" title="restore">↩ restore' +
    '</button>';
  const set = (c, v) => li.querySelector('[name="' + n(c) + '"]')
                          .value = v || '';
  set('item', item); set('group', group);
  set('owner', extra.owner); set('type', extra.type);
  set('priority', extra.priority); set('notes', extra.notes);
  set('id', extra.id);
  const status = (extra.status || '').toLowerCase();
  li.querySelector('.pdone').checked = status === 'done';
  const grp = () => li.querySelector('[name="' + n('group') + '"]')
                      .value;
  const home = status === 'not_needed'
      ? prepClosedBody(group) : prepGroupBody(group);
  home.appendChild(li);
  prepStatus(li);
  li.querySelector('.pdone').addEventListener('change',
                                              () => prepStatus(li));
  li.querySelector('.pclose').addEventListener('click', () => {
    prepClosedBody(grp()).appendChild(li);
    prepStatus(li);
  });
  li.querySelector('.prestore').addEventListener('click', () => {
    prepGroupBody(grp()).appendChild(li);
    prepStatus(li);
    prepTidyClosed();
  });
  return li;
}

function prepCloseGroup(btn) {
  const fs = btn.closest('.pgroup');
  for (const li of [...fs.querySelectorAll('.pitem')]) {
    prepClosedBody(li.querySelector('[name$="__group"]').value)
        .appendChild(li);
    prepStatus(li);
  }
  fs.remove();   // recreated automatically on restore
}

function prepAddCustom(btn) {
  const input = btn.parentNode.querySelector('.padd');
  const group = btn.closest('.pgroup')
                   .querySelector('legend b').textContent;
  if (input.value.trim())
    prepItem(group === 'General' ? '' : group, input.value.trim());
  input.value = '';
}

function prepNewGroup() {
  const input = document.getElementById('prep_newgroup');
  if (input.value.trim()) prepGroupBody(input.value.trim());
  input.value = '';
}

// Load the workbook's checklist, then merge in any catalog defaults it
// hasn't seen yet (new suggestions added after the workbook was saved)
// — matched by group+item text, so closed items STAY closed and only
// genuinely-new defaults appear as fresh suggestions.
const seeded = new Set();
if (PREFILL && PREFILL.preparation) {
  for (const row of PREFILL.preparation) {
    prepItem(row.group || '', row.item || '',
             {owner: row.owner, type: row.type, priority: row.priority,
              notes: row.notes, status: row.status, id: row._id});
    seeded.add(((row.group || '') + '|' + (row.item || ''))
               .toLowerCase());
  }
}
for (const [group, items] of Object.entries(PREPCAT))
  for (const item of items)
    if (!seeded.has((group + '|' + item).toLowerCase()))
      prepItem(group, item, {});

// --- saving without leaving the form ----------------------------------
// The interview is long. Save whenever you like and carry on typing;
// the workbook is written, the form stays exactly where it was.
let dirty = false;
let leaving = false;
document.addEventListener('input', () => { dirty = true; }, true);
document.addEventListener('change', () => { dirty = true; }, true);
window.addEventListener('beforeunload', e => {
  if (!dirty || leaving) return;
  e.preventDefault();           // typed answers are the expensive part
  e.returnValue = '';
});

function saveStatus(msg, bad) {
  const el = document.getElementById('savemsg');
  el.textContent = msg;
  el.className = bad ? 'savemsg bad' : 'savemsg';
}

// Hold on to the ids the server just issued and point the form at the
// file it wrote. Without this the next save would be refused as
// "already exists", and any id it re-issued would move a calendar
// event to a new identity.
function applySaved(data) {
  for (const [tab, map] of Object.entries(data.ids)) {
    for (const [idx, id] of Object.entries(map)) {
      const f = document.querySelector(
          '[name="' + tab + '__' + idx + '__id"]');
      if (f) f.value = id;
    }
  }
  for (const [tab, maxid] of Object.entries(data.maxids)) {
    const f = document.querySelector('[name="maxid_' + tab + '"]');
    if (f) f.value = maxid;
  }
  document.querySelector('[name="editing"]').value = data.path;
  const out = document.getElementById('out_dir');
  if (out) {
    out.readOnly = true;
    out.title = 'Saves now overwrite ' + data.path;
  }
}

async function saveNow() {
  const btn = document.getElementById('savebtn');
  btn.disabled = true;
  saveStatus('Saving…');
  try {
    const resp = await fetch('/save', {
      method: 'POST',
      headers: {'Content-Type': 'application/x-www-form-urlencoded'},
      body: new URLSearchParams(
          new FormData(document.querySelector('form'))),
    });
    const data = await resp.json();
    if (!data.ok) { saveStatus(data.error, true); return; }
    if (data.reload) {          // a row was added for you; re-read it
      leaving = true;
      location = '/?load=' + encodeURIComponent(data.path);
      return;
    }
    applySaved(data);
    dirty = false;
    saveStatus('Saved ' + data.saved_at + ' to ' + data.path
               + (data.backup ? ' (previous version kept as '
                                + data.backup + ')' : ''));
  } catch (err) {
    saveStatus('Save failed: ' + err, true);
  } finally {
    btn.disabled = false;
  }
}

// Native pickers, courtesy of the local server (browsers won't hand a
// web page an absolute path; the server can). Empty = cancelled.
// The dialog is a window on THIS machine, and the page can't do
// anything until it's answered — so say so, in case it still ends up
// behind the browser on some setup.
function pickBusy(on, msg) {
  for (const b of document.querySelectorAll('.pickbtn')) b.disabled = on;
  const el = document.getElementById('pickmsg');
  el.textContent = on ? (msg || 'Look for the folder window — it opens '
                                + 'in front of the browser.') : '';
}
async function pick(kind, msg) {
  pickBusy(true, msg);
  try {
    const resp = await fetch('/pick?kind=' + kind);
    const text = (await resp.text()).trim();
    if (text.startsWith('ERROR:')) {
      const typed = window.prompt(
          'Could not open a picker on this machine:\\n\\n'
          + text.slice(6).trim()
          + '\\n\\nPaste the full path of the workbook instead:');
      return (typed || '').trim();
    }
    return text;
  } finally {
    pickBusy(false);
  }
}
async function pickLoad() {
  const path = await pick('workbook',
      'Look for the "choose the workbook" window — it opens in front '
      + 'of the browser.');
  if (path) location = '/?load=' + encodeURIComponent(path);
}

// Live-mirror "who is this plan for?" into the first household row
// (name / relationship 'self' / role 'planholder') until the user
// edits those fields themselves. The server still validates either way.
const planFor = document.querySelector('[name="plan_for"]');
function syncPlanholder() {
  const name = document.querySelector('[name="household__0__name"]');
  if (!name || name.dataset.dirty) return;
  name.value = planFor.value;
  const rel = document.querySelector(
      '[name="household__0__relationship"]');
  if (rel && !rel.dataset.dirty) rel.value = 'self';
  const role = document.querySelector('[name="household__0__role"]');
  if (role && !role.dataset.dirty) role.value = 'planholder';
  updateCount('household');
}
planFor.addEventListener('input', syncPlanholder);
for (const f of ['name', 'relationship', 'role']) {
  const el = document.querySelector('[name="household__0__' + f + '"]');
  if (el) el.addEventListener('input',
                              e => e.target.dataset.dirty = '1');
}
"""

TAB_BLURB = {
    "household": "Everyone involved. The first row fills itself from "
                 "the name at the top (the 'planholder' — the person "
                 "this plan is for); add spouse, executor, and the "
                 "wider circle who'll do the work.",
    "preparation": "What to do BEFORE anything happens: actions (make "
                   "a will, vet cleaners), decisions and living "
                   "preferences ('no stairlift'). Track status here; "
                   "stays local, never on any calendar.",
    "wishes": "Funeral, burial/cremation, organ donation, DNACPR — what is "
              "wanted and where that's written down.",
    "funeral": "Only if a funeral is pre-arranged. Service date/venue can "
               "be left blank until known.",
    "key_documents": "Where the will, deeds, LPAs and certificates live.",
    "access_pointers": "WHERE logins live (e.g. 'password manager', "
                       "'home safe') — never the secrets themselves.",
    "funeral_contacts": "Who should be told and invited.",
    "professionals": "Solicitor, accountant, adviser.",
    "banks": "One row per account.",
    "building_societies": "One row per account.",
    "cash_isas": "One row per ISA.",
    "shares_isas": "One row per ISA.",
    "share_portfolios": "Shareholdings and platforms.",
    "pensions": "Workplace, personal, state, annuities.",
    "property": "Homes and other property.",
    "other_assets": "Vehicles, valuables, collections.",
    "life_insurance": "Policies AND death-in-service through an employer "
                      "— the classic benefit nobody knew existed.",
    "digital": "Domains, subscriptions, accounts online.",
}


# --- workbook writing --------------------------------------------------------

def _cell(form: dict, tab: str, i: int, col: str) -> str:
    """One form field -> one cell. A multi-value enum arrives as one
    posted value per ticked box; it goes into the cell in canonical
    order, so the workbook never carries a spelling the engine can't
    read (roles.py)."""
    posted = form.get(f"{tab}__{i}__{col}", [])
    if (tab, col) in MULTI_ENUM:
        return format_roles(v for v in posted if v.strip())
    return posted[0].strip() if posted else ""


def build_rows(form: dict, kept_indexes: dict = None) -> dict:
    """Collect ``tab__index__column`` form fields into rows per tab.
    A row counts only if at least one field is filled (same presence
    rule as the reader).

    Pass ``kept_indexes`` to receive ``{tab: [form index per kept row]}``
    alongside the rows: saving in place has to tell the page which of its
    rows got which id, and empty rows are dropped on the way through, so
    the two lists only line up if the caller is told what was kept.
    """
    rows_by_tab = {}
    for tab, spec in TABS.items():
        data_cols = [c[0] for c in spec["columns"] if c[0] not in META]
        indexes = set()
        prefix = re.compile(re.escape(tab) + r"__(\d+)__")
        for key in form:
            m = prefix.match(key)
            if m:
                indexes.add(int(m.group(1)))
        rows, kept = [], []
        for i in sorted(indexes):
            values = tuple(_cell(form, tab, i, col) for col in data_cols)
            # Presence = typed data. Tick-boxes don't make a row exist:
            # roles ticked against a nameless person are nothing.
            if not any(v for v, col in zip(values, data_cols)
                       if (tab, col) not in MULTI_ENUM):
                continue
            # Meta columns: the row's stable id (hidden field — kept
            # through edits), the consent checkbox ('full'), and the
            # hidden carrier that preserves a 'none' row through a
            # checkbox that can't express it. Unticked = safe defaults.
            row_id = form.get(f"{tab}__{i}__id", [""])[0].strip()
            vis = form.get(f"{tab}__{i}__online_visibility", [""])[0]
            none = form.get(f"{tab}__{i}__vis_none", [""])[0]
            ai = form.get(f"{tab}__{i}__ai_visible?", [""])[0]
            vis_final = "full" if vis == "full" else (
                "none" if none == "none" else None)
            rows.append(values + (row_id or None, vis_final, ai or None))
            kept.append(i)
        if rows:
            rows_by_tab[tab] = rows
            if kept_indexes is not None:
                kept_indexes[tab] = kept
    return rows_by_tab


def assign_ids(rows_by_tab: dict, form: dict) -> None:
    """Give id-less (new) rows the next never-used id per tab. The
    hidden maxid_<tab> field carries the highest id the loaded workbook
    ever issued, so a deleted row's number is never re-used."""
    for tab, rows in rows_by_tab.items():
        used = [int(r[-3]) for r in rows
                if r[-3] and str(r[-3]).isdigit()]
        raw = form.get(f"maxid_{tab}", ["0"])[0]
        next_id = max(used + [int(raw) if raw.isdigit() else 0],
                      default=0) + 1
        for i, row in enumerate(rows):
            if not row[-3]:
                rows[i] = row[:-3] + (str(next_id),) + row[-2:]
                next_id += 1


def _jsonable(value):
    """Workbook cell -> form-safe string (dates/times normalised)."""
    import datetime as _dt
    if isinstance(value, _dt.datetime):
        return (value.date().isoformat() if not value.time().minute
                and not value.time().hour and not value.time().second
                else value.isoformat(sep=" "))
    if isinstance(value, _dt.date):
        return value.isoformat()
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M")
    return str(value)


def load_prefill(raw_path: str):
    """Read an existing workbook for editing. Returns
    (error, plan_for, out_dir, prefill, maxids)."""
    from reader import read_workbook
    path = Path(raw_path).resolve()
    if path.is_relative_to(REPO_ROOT):
        return ("REFUSED: that file is inside the repo working tree — "
                "only workbooks outside it can be edited.",
                "", "", None, None)
    if not path.is_file() or path.suffix.lower() != ".xlsx":
        return (f"Not a workbook: {path}", "", "", None, None)
    model = read_workbook(path)
    prefill, maxids = {}, {}
    plan_for = ""
    for tab, rows in model.modules.items():
        entries = []
        for r in rows:
            entry = {col: _jsonable(val) for col, val in r.fields.items()
                     if col not in META and val is not None}
            entry["_id"] = r.row_id or ""
            entry["_vis"] = r.visibility
            entry["_ai"] = r.ai_visible
            for _tab, col in (k for k in MULTI_ENUM if k[0] == tab):
                # Reduce a hand-typed cell to the enum so the tick-boxes
                # match exactly; anything unrecognised is carried through
                # to be shown, not quietly deleted on the next save.
                known, unknown = parse_roles(entry.get(col))
                entry[col] = format_roles(known)
                entry["_badroles"] = unknown
            entries.append(entry)
            if (tab == "household"
                    and (r.fields.get("role") or "").lower()
                    == "planholder"):
                plan_for = r.fields.get("name") or ""
        prefill[tab] = entries
        maxids[tab] = max(
            [int(r.row_id) for r in rows
             if r.row_id and r.row_id.isdigit()], default=0)
    return (None, plan_for or path.stem, str(path), prefill, maxids)


def ensure_planholder(rows_by_tab: dict, plan_for: str):
    """The person the plan is for must anchor the household tab as its
    ``planholder`` row — the engine derives 'the deceased' from it, and
    later phases hang their pre-death preparation tasks off it.

    No planholder row -> add one for plan_for (never silent: the row is
    labelled and the success page says so). A conflicting or duplicate
    planholder -> error string (returned, nothing written)."""
    rows = list(rows_by_tab.get("household", []))
    # household data columns: name, relationship, role, crisis_roles, ...
    planholders = [r for r in rows if len(r) > 2
                   and (r[2] or "").strip().lower() == "planholder"]
    if len(planholders) > 1:
        return ("Only one household row can be 'planholder' — the plan "
                "is for exactly one person.")
    if planholders:
        listed = (planholders[0][0] or "").strip()
        if listed.lower() != plan_for.lower():
            return (f"The plan is for {plan_for!r} but the household "
                    f"planholder row says {listed!r} — make them match "
                    "(the planholder IS the person this plan is for).")
        return None
    rows.insert(0, (plan_for, "self — the person this plan is for",
                    "planholder", "", "Y", "",
                    "Added by the wizard: the plan is for this person",
                    None, None, None))   # meta: id / visibility / ai
    rows_by_tab["household"] = rows
    return None


def write_workbook(path: Path, rows_by_tab: dict) -> None:
    """Write a fresh workbook through the template generator (same
    protection, dropdowns, instructions and id pre-fill as the sample).
    Never overwrites (checked by the caller)."""
    wb = Workbook()
    make_template.build_instructions(wb)
    for name in TAB_ORDER:
        spec = dict(TABS[name], rows=rows_by_tab.get(name, []))
        make_template.build_tab(wb, name, spec)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def initial_dir(start: str = "") -> str:
    """Where a picker dialog should open: the nearest existing ancestor
    of what the form currently says, else Documents\\crisis-plan, else
    Documents, else home."""
    if start:
        p = Path(start)
        while True:
            if p.is_dir():
                return str(p)
            if p.parent == p:
                break
            p = p.parent
    for candidate in (Path.home() / "Documents" / "crisis-plan",
                      Path.home() / "Documents"):
        if candidate.is_dir():
            return str(candidate)
    return str(Path.home())


def target_path(out_dir: str, household: str) -> Path:
    """The workbook goes in a per-person subfolder of the chosen root
    (e.g. Documents\\crisis-plan\\pat\\pat.xlsx), so several
    people's plans can share one root — every generated file (plan .md
    files, .ics, backups) lands beside the workbook in that subfolder.
    Browsing straight to the person's own folder doesn't nest another
    level."""
    slug = slugify(household)
    base = Path(out_dir)
    if base.name.lower() != slug:
        base = base / slug
    return base / f"{slug}.xlsx"


def check_target(path: Path):
    """-> error string or None. Write-once + outside-the-repo."""
    resolved = path.resolve()
    if resolved.is_relative_to(REPO_ROOT):
        return (f"REFUSED: {resolved} is inside the repo working tree — "
                "a household's workbook belongs OUTSIDE it (see "
                "PRIVACY.md).")
    if resolved.exists():
        return (f"{resolved} already exists. The wizard never overwrites "
                "— pick a different household name or folder, or move "
                "the old file first.")
    return None


# --- form rendering ----------------------------------------------------------

def _input_html(tab: str, col: str, options, strict: bool) -> str:
    name = f"{tab}__IDX__{col}"
    label = html.escape(col)
    if (tab, col) in MULTI_ENUM:
        # An enum, so it's tick-boxes, not typing: a misspelt crisis role
        # used to leave the role held by nobody with the plan still
        # reading as correct (roles.py). All the boxes share one field
        # name; build_rows joins what's ticked back into the cell.
        boxes = "".join(
            f'<label class="chk"><input type="checkbox" name="{name}"'
            f' value="{html.escape(role)}"> <b>{html.escape(role)}</b>'
            f" — {html.escape(blurb)}</label>"
            for role, blurb in MULTI_ENUM[(tab, col)])
        return (f'<label>{label}<br><span class="chks">{boxes}'
                '<span class="chk rolewarn"></span></span></label>')
    if options and strict:
        opts = "".join(f'<option value="{html.escape(o)}">{html.escape(o)}'
                       "</option>" for o in options)
        field = (f'<select name="{name}"><option value=""></option>'
                 f"{opts}</select>")
    elif options:
        dl = f"dl_{tab}_{re.sub(r'[^a-z0-9]', '', col)}"
        opts = "".join(f'<option value="{html.escape(o)}">'
                       for o in options)
        field = (f'<input name="{name}" list="{dl}" size="18">'
                 f'<datalist id="{dl}">{opts}</datalist>')
    else:
        size = 34 if col in ("notes",) or "location" in col else 20
        field = f'<input name="{name}" size="{size}">'
    return f"<label>{label}<br>{field}</label>"


def render_form(editing: str = "", plan_for_value: str = "",
                prefill=None, maxids=None) -> str:
    sections = []
    for tab in TAB_ORDER:
        if tab == "preparation":   # bespoke pre-crisis pane, not a grid
            continue
        spec = TABS[tab]
        fields = "".join(
            _input_html(tab, col, options, strict)
            for col, _w, options, strict in spec["columns"]
            if col not in META)
        fields += (f'<input type="hidden" name="{tab}__IDX__id">'
                   '<button type="button" class="rm" title="remove this '
                   'row" onclick="removeRow(this)">'
                   '× remove row</button>')
        # household rows never become events; preparation is local-only
        # by design — neither gets calendar/AI consent checkboxes.
        if tab not in ("household", "preparation"):
            fields += f"""
<span class="chks">
<label class="chk"><input type="checkbox"
 name="{tab}__IDX__online_visibility" value="full">
 put this row's real details on the Google calendar</label>
<input type="hidden" name="{tab}__IDX__vis_none">
<span class="chk">→ shown on the shared calendar as:
 <b class="gpreview"></b></span>"""
            # The AI-lookup consent only exists where the lookup does:
            # tabs augment.py can query, wording what it actually sends.
            if tab in INSTITUTION_COLUMN:
                fields += f"""
<label class="chk"><input type="checkbox"
 name="{tab}__IDX__ai_visible?" value="Y">
 allow the optional AI lookup to send this row's
 <b>{html.escape(INSTITUTION_COLUMN[tab])}</b> name — nothing else —
 to find: {html.escape(QUERY_CONTEXT[tab])}. It shows the exact query
 and asks before going online.</label>"""
            fields += "\n</span>"
        sections.append(f"""
<details {'open' if tab == 'household' else ''}>
<summary><span class="cnt" id="count_{tab}">[0]</span> <b>{tab}</b> —
 {html.escape(TAB_BLURB.get(tab, ''))}</summary>
<template id="tpl_{tab}">{fields}</template>
<div id="rows_{tab}"></div>
<button type="button" onclick="addRow('{tab}', true)">+ add another
 row</button>
</details>""")
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Crisis-plan wizard</title>
<style>
 body {{ font-family: system-ui, sans-serif; max-width: 60rem;
        margin: 2rem auto; padding: 0 1rem; }}
 details {{ border: 1px solid #ccc; border-radius: 6px;
           padding: .6rem 1rem; margin: .6rem 0; }}
 summary {{ cursor: pointer; }}
 .row {{ border-top: 1px dashed #ddd; padding: .5rem 0;
        display: flex; flex-wrap: wrap; gap: .6rem; }}
 label {{ font-size: .8rem; color: #444; }}
 .chks {{ flex-basis: 100%; }}
 .cnt {{ color: #888; font-family: monospace; }}
 .tabs {{ margin: 1rem 0; }}
 .tabs button {{ font-size: 1rem; padding: .4rem 1rem; }}
 .tabs button.active {{ font-weight: bold; border-color: #66c; }}
 .pane {{ display: none; }}
 .pane.active {{ display: block; }}
 .pitem {{ margin: .2rem 0; display: flex; gap: .4rem;
          align-items: center; }}
 .pgroup {{ border: 1px solid #ccc; border-radius: 6px;
           padding: .4rem 1rem; margin: .6rem 0; }}
 #prep_closed .pitem {{ opacity: .55; }}
 #prep_closed .pdone, #prep_closed .pclose,
 #prep_active .prestore {{ display: none; }}
 .chk {{ display: block; margin-top: .15rem; color: #666; }}
 .rolewarn {{ color: #a33; }}
 .pickmsg {{ color: #a33; margin-left: .5rem; }}
 .savemsg {{ color: #262; margin-left: .5rem; }}
 .savemsg.bad {{ color: #a33; }}
 .top {{ background: #f4f4f8; border-radius: 6px; padding: 1rem; }}
 /* Stays put while the long form scrolls under it: on a 17-section
    interview, Save has to be reachable without hunting for it. */
 .toolbar {{ position: sticky; top: 0; z-index: 10; display: flex;
            gap: .5rem; align-items: center; flex-wrap: wrap;
            background: #fff; border-bottom: 1px solid #ccc;
            padding: .6rem 0; margin-bottom: 1rem; }}
 .toolbar button {{ font-size: 1rem; padding: .3rem .9rem; }}
 .tbtitle {{ font-size: 1.2rem; margin-right: .5rem; }}
 button[type=submit] {{ font-size: 1.1rem; padding: .5rem 1.5rem;
                        margin-top: 1rem; }}
</style></head><body>
<div class="toolbar">
<b class="tbtitle">Crisis-plan wizard</b>
<button type="button" class="pickbtn" onclick="pickLoad()">Open…</button>
<button type="button" id="savebtn" onclick="saveNow()">Save</button>
<span id="savemsg" class="savemsg">{
 ('Editing ' + html.escape(editing)) if editing else ''}</span>
<span id="pickmsg" class="pickmsg"></span>
</div>
<p>This builds the plan for <b>one person</b>: what happens when they
die. Their assets, their wishes, and the people around them. (A couple
makes one each — most answers will be similar.) The person also appears
in the household section as the <i>planholder</i>; a later version adds
their own pre-death preparation tasks (write the will, record wishes,
consolidate accounts) against that row.</p>
<p>Answer what you can — every section is optional and everything can be
edited in the spreadsheet afterwards. Skip whole sections that don't
apply. <b>Pointers, never secrets:</b> write where a login or document
lives, never a password or account number. Nothing here leaves this
computer.</p>
<p>Each row has two consent checkboxes. Left unticked (the safe
default), the shared calendar shows only anonymous labels like "bank
account&nbsp;#1" and the optional AI lookup sees nothing. Tick per row
to opt up. (A third level — keep a row off the calendar entirely — is
available in the spreadsheet's <code>online_visibility</code> column.)
</p>
<form method="post" action="/create">
<div class="top">
<input type="hidden" name="editing" value="{html.escape(editing)}">
{''.join(f'<input type="hidden" name="maxid_{t}" '
         f'value="{(maxids or {}).get(t, 0)}">' for t in TAB_ORDER)}
{'<input type="hidden" name="plan_for" '
 'value="' + html.escape(plan_for_value) + '">'
 if editing else
 '<label>Who is this plan for?<br>'
 '<input name="plan_for" required size="24" '
 'value="' + html.escape(plan_for_value) + '"></label>'}
<input type="hidden" name="out_dir" id="out_dir"
 value="{html.escape(str(Path(editing).parent) if editing else
         str(Path.home() / 'Documents' / 'crisis-plan'))}">
<p><label><input type="checkbox" name="google_setup" value="on">
 After writing the workbook, run the one-time Google Calendar setup
 (needs tools/GOOGLE-SETUP.md steps 1–2 done first)</label><br>
<label>Family emails to share the calendar with (comma-separated)<br>
<input name="google_emails" size="48"></label></p>
</div>
<div class="tabs">
<button type="button" id="tabbtn_pre" class="active"
 onclick="showPane('pre')">Pre-crisis — preparing now</button>
<button type="button" id="tabbtn_crisis"
 onclick="showPane('crisis')">Crisis — when it happens</button>
</div>
<div class="pane active" id="pane_pre">
<p>{html.escape(TAB_BLURB['preparation'])}</p>
<p>Tick what's done. The <b>×</b> closes an item (or a whole group) as
"not needed" — it drops to the bottom and can be restored. Add your own
items and groups; owners are optional.</p>
<div id="prep_active"></div>
<p><input id="prep_newgroup" size="24" placeholder="new group name">
<button type="button" onclick="prepNewGroup()">+ add group</button></p>
<h3>Not needed</h3>
<div id="prep_closed"></div>
</div>
<div class="pane" id="pane_crisis">
{''.join(sections)}
</div>
<button type="submit">Finish — write the workbook and generate the
 first plan</button>
</form>
<script>
const ALIAS = {json.dumps(ALIAS_COLS)};
const GENERIC = {json.dumps(GENERIC_LABEL)};
const PREFILL = {json.dumps(prefill)};
const PREPCAT = {json.dumps(PREP_CATALOG)};
{SCRIPT_JS}
</script>
</body></html>"""


def render_result(path: Path, rows_by_tab: dict, gsetup_output,
                  artifacts, backup=None) -> str:
    counts = "".join(f"<li><b>{t}</b>: {len(r)} row(s)</li>"
                     for t, r in rows_by_tab.items())
    gsetup_html = ""
    if gsetup_output is not None:
        gsetup_html = ("<h2>Google Calendar setup</h2><pre>"
                       + html.escape(gsetup_output) + "</pre>")
    if artifacts:
        links = "".join(
            f'<li><a href="/file?p={urllib.parse.quote(str(f))}" '
            f'target="_blank">{html.escape(f.name)}</a>'
            f" — {html.escape(desc)}</li>"
            for f, desc in artifacts)
        artifacts_html = f"""
<h2>Your first plan (generated just now, offline)</h2>
<p>Saved alongside the workbook. Trigger date is a placeholder (today)
— the dates shift when a real run gets the real date. Click to
read:</p>
<ul>{links}</ul>"""
    else:
        artifacts_html = ""
    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>Workbook written</title></head>
<body style="font-family: system-ui, sans-serif; max-width: 60rem;
 margin: 2rem auto; padding: 0 1rem;">
<h1>Workbook written ✔</h1>
<p><code>{html.escape(str(path))}</code></p>
{f'<p>Edited in place — the previous version is backed up as '
 f'<code>backups\\{html.escape(backup.name)}</code>.</p>'
 if backup else ''}
<p>This is the plan for one person — they're on the household tab as
the <b>planholder</b> (added automatically if you didn't list them).</p>
<p><a href="/?load={urllib.parse.quote(str(path))}">← Keep editing this
workbook</a> — the wizard is still running, so you can go back, change
answers and save again without starting over.</p>
<ul>{counts}</ul>
{artifacts_html}
<h2>Next steps</h2>
<ol>
<li>Open the workbook and check it — the instructions tab explains every
 column. Your per-row checkbox choices landed in the
 <code>online_visibility</code> / <code>ai_visible?</code> columns;
 fine-tune them there any time (including <code>none</code> to keep a
 row off calendars entirely).</li>
<li>Generate a plan any time:<br>
 <code>python tools\\generate.py -f "{html.escape(str(path))}"
 --date YYYY-MM-DD</code></li>
<li>Keep the workbook out of the repo, back it up, and re-run the
 generator after any edit.</li>
</ol>
{gsetup_html}
<p><a href="/done"><b>Finish — close the wizard</b></a> (it stays
running until then so the links above keep working).</p>
</body></html>"""


# --- server ------------------------------------------------------------------

class WizardHandler(BaseHTTPRequestHandler):
    def _send(self, code: int, body: str):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/done":
            self._send(200, "<p>Wizard closed — you can close this "
                            "window.</p>")
            threading.Thread(target=self.server.shutdown,
                             daemon=True).start()
            return
        if parsed.path == "/file":
            self._serve_artifact(parsed)
            return
        if parsed.path == "/pick":
            self._pick(parsed)
            return
        load = urllib.parse.parse_qs(parsed.query).get("load", [""])[0]
        if load.strip():
            error, plan_for, editing, prefill, maxids = load_prefill(
                load.strip())
            if error:
                self._send(400, f"<p>{html.escape(error)}</p>"
                                '<p><a href="/">back</a></p>')
                return
            self._send(200, render_form(editing=editing,
                                        plan_for_value=plan_for,
                                        prefill=prefill, maxids=maxids))
            return
        self._send(200, render_form())

    def _pick(self, parsed):
        """Open a NATIVE workbook picker on this machine (the server
        runs locally, so it can do what the browser won't: hand back
        an absolute path). Empty response = cancelled."""
        import os
        kind = urllib.parse.parse_qs(parsed.query).get("kind", [""])[0]
        if kind != "workbook":
            self._send(400, "unknown picker kind")
            return
        start = initial_dir(
            urllib.parse.parse_qs(parsed.query).get("start", [""])[0])
        fake = os.environ.get("CRISISPLAN_PICK_FAKE")   # test hook
        if fake is not None:
            chosen = fake
        else:
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                # Windows only puts a dialog in front if its owner has
                # been through an event loop first: setting -topmost and
                # going straight into the dialog opened it BEHIND the
                # browser, and since this server is single-threaded the
                # page froze too — the picker looked like it had gone.
                root.attributes("-topmost", True)
                root.update()
                root.lift()
                root.focus_force()
                chosen = filedialog.askopenfilename(
                    parent=root, initialdir=start,
                    title="Choose the workbook to edit",
                    filetypes=[("Crisis-plan workbook", "*.xlsx")])
                root.destroy()
            except Exception as exc:
                # No Tk, or a display that won't give it a window. Say
                # so — a button that does nothing reads as a broken
                # wizard; the path box still works.
                chosen = f"ERROR: {type(exc).__name__}: {exc}"
        data = (chosen or "").encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _serve_artifact(self, parsed):
        """Serve one generated plan file as plain text. Only files from
        this session's run directory — nothing else on disk."""
        run_dir = getattr(self.server, "run_dir", None)
        raw = urllib.parse.parse_qs(parsed.query).get("p", [""])[0]
        target = Path(raw).resolve() if raw else None
        if (run_dir is None or target is None
                or not target.is_relative_to(run_dir)
                or target.suffix.lower() != ".md"   # never the workbook
                or not target.is_file()):
            self._send(403, "<p>Not a file from this wizard run.</p>")
            return
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _prepare(self, form: dict, kept_indexes: dict = None):
        """Shared by 'save' and 'write the workbook': validate the form
        and work out where it goes. -> (path, rows_by_tab, editing,
        error) with error as ready-to-send HTML, everything else None
        when it's set."""
        plan_for = form.get("plan_for", [""])[0].strip()
        out_dir = form.get("out_dir", [""])[0].strip()
        editing = form.get("editing", [""])[0].strip()
        if not plan_for or (not editing and not out_dir):
            return (None, None, "", "<p>The person's name and the folder "
                    "are required — go back and fill them in.</p>")
        if editing:
            path = Path(editing).resolve()
            if path.is_relative_to(REPO_ROOT):
                return (None, None, editing,
                        "<p>REFUSED: cannot edit a workbook inside the "
                        "repo working tree.</p>")
            if not path.is_file():
                return (None, None, editing,
                        f"<p>{html.escape(str(path))} no longer exists "
                        "— nothing to edit.</p>")
        else:
            path = target_path(out_dir, plan_for)
            error = check_target(path)
            if error:
                return (None, None, editing,
                        f"<p>{html.escape(error)}</p>")
        rows_by_tab = build_rows(form, kept_indexes)
        error = ensure_planholder(rows_by_tab, plan_for)
        if error:
            return (None, None, editing, f"<p>{html.escape(error)}</p>")
        assign_ids(rows_by_tab, form)
        return (path, rows_by_tab, editing, None)

    def _write(self, path: Path, rows_by_tab: dict, editing: str):
        """Write the workbook, backing the old one up first. -> the
        backup path, or None.

        One backup per file per wizard session: the copy worth keeping is
        the state before you started, and saving every few minutes should
        not bury it under fifty near-identical files.
        """
        backup = None
        backed_up = getattr(self.server, "backed_up", None)
        if backed_up is None:
            backed_up = self.server.backed_up = set()
        if editing and path.is_file() and str(path) not in backed_up:
            from datetime import datetime
            backup_dir = path.parent / "backups"
            backup_dir.mkdir(exist_ok=True)
            backup = backup_dir / (
                f"{path.stem}-backup-"
                f"{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx")
            shutil.copyfile(path, backup)
            backed_up.add(str(path))
        write_workbook(path, rows_by_tab)
        return backup

    def _save(self, form: dict):
        """Save without leaving the form (POST /save). Answers JSON so
        the page can carry on where it was — including the ids just
        issued, which the form must hold on to: an id that changes
        between saves changes a calendar event's identity."""
        from datetime import datetime
        kept_indexes = {}
        path, rows_by_tab, editing, error = self._prepare(
            form, kept_indexes)
        if error:
            self._send_json({"ok": False,
                             "error": re.sub(r"<[^>]+>", "", error)})
            return
        # ensure_planholder can insert a row the page doesn't have; the
        # id map would be off by one, so the page reloads instead.
        inserted = len(rows_by_tab.get("household", [])) > len(
            kept_indexes.get("household", []))
        backup = self._write(path, rows_by_tab, editing)
        self.server.saved_path = path
        print(f"Saved: {path}")
        ids = {}
        for tab, rows in rows_by_tab.items():
            if tab in kept_indexes and not (tab == "household" and inserted):
                ids[tab] = {str(i): (r[-3] or "")
                            for i, r in zip(kept_indexes[tab], rows)}
        self._send_json({
            "ok": True,
            "path": str(path),
            "reload": inserted,
            "saved_at": datetime.now().strftime("%H:%M:%S"),
            "backup": backup.name if backup else "",
            "ids": ids,
            "maxids": {tab: max([int(r[-3]) for r in rows
                                 if r[-3] and str(r[-3]).isdigit()],
                                default=0)
                       for tab, rows in rows_by_tab.items()},
        })

    def _send_json(self, payload: dict):
        data = json.dumps(payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        form = urllib.parse.parse_qs(
            self.rfile.read(length).decode("utf-8"))
        if urllib.parse.urlparse(self.path).path == "/save":
            self._save(form)
            return
        path, rows_by_tab, editing, error = self._prepare(form)
        if error:
            self._send(400, error)
            return
        plan_for = form.get("plan_for", [""])[0].strip()
        backup = self._write(path, rows_by_tab, editing)

        gsetup_output = None
        if form.get("google_setup", [""])[0] == "on":
            emails = form.get("google_emails", [""])[0].strip()
            result = subprocess.run(
                [sys.executable, str(Path(__file__).parent / "gsetup.py"),
                 "--emails", emails, "--household", plan_for],
                capture_output=True, text=True, encoding="utf-8",
                errors="replace")
            gsetup_output = result.stdout + (
                ("\n" + result.stderr) if result.returncode else "")

        artifacts = self._first_run(path)
        self._send(200, render_result(path, rows_by_tab, gsetup_output,
                                      artifacts, backup))
        print(f"Workbook written: {path}")
        self.server.wizard_done = True
        # The server stays up so the artifact links work; /done ends it.

    def _first_run(self, workbook: Path):
        """Generate the first plan (offline, placeholder date) so the
        success page can link to the checklist / preview files."""
        from datetime import date
        # The plan files live in the workbook's own folder (outside the
        # repo by construction — the workbook already is).
        run_dir = workbook.parent
        result = subprocess.run(
            [sys.executable, str(Path(__file__).parent / "generate.py"),
             "-f", str(workbook), "--date", date.today().isoformat(),
             "--mode", "local", "--out", str(run_dir)],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace")
        if result.returncode != 0 or not run_dir.exists():
            return []
        self.server.run_dir = run_dir.resolve()
        described = {
            "checklist.md": "the master checklist, everyone's tasks",
            "google-preview.md": "exactly what a calendar push would "
                                 "publish — read before any real push",
            "funeral-details.md": "the share-with-friends funeral note",
            "preparation.md": "what to do BEFORE — actions, "
                              "preferences, completeness check",
        }
        artifacts = []
        for f in sorted(run_dir.glob("*.md")):
            desc = described.get(f.name)
            if desc is None and f.name.startswith("checklist-"):
                desc = "personal crisis checklist"
            if desc is None and f.name.startswith("preparation-"):
                desc = "personal preparation list"
            artifacts.append((f, desc or "generated file"))
        return artifacts


def make_server(port: int = 0) -> HTTPServer:
    server = HTTPServer(("127.0.0.1", port), WizardHandler)
    server.wizard_done = False
    return server


def main() -> None:
    server = make_server()
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print("Crisis-plan wizard running (local only — nothing leaves this "
          "machine).")
    print(f"If the browser doesn't open, go to: {url}")
    print("Click 'Finish' on the final page (or Ctrl+C here) to close.")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nAbandoned — nothing was written."
              if not server.wizard_done else "")
    finally:
        server.server_close()
    if server.wizard_done:
        print("Done — see the browser page for next steps.")


if __name__ == "__main__":
    main()

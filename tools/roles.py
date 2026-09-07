#!/usr/bin/env python3
"""Canonical crisis roles (SPEC §14) — the enum, in one place.

``crisis_roles`` on the household tab shares the work out, and the engine
branches on it: a task's role decides whose personal checklist it lands
on. That made a free-text cell quietly dangerous. "organizer" is not
"organiser", so the role ended up held by nobody, the fallback handed the
work to the executor, and **the plan still looked completely correct** —
the worst failure this tool can have, because nobody re-reads a plan that
looks fine.

So the set lives here and every surface reads it from here:

- the **wizard** renders one checkbox per role, so typing one wrong is
  not possible;
- the **template instructions** are generated from this list, so they
  can't drift from what the engine understands;
- :func:`parse` normalises whatever a human typed in Excel and *reports*
  what it doesn't recognise, so an unknown value is a visible note in the
  plan rather than a silently unheld role.

Unknown values are never guessed at and never silently dropped: only the
spelling variants in :data:`ALIASES` — the same word in a different form —
are accepted without comment.
"""

import re

# The five functional roles the task modules assign work to. Order is the
# order they're offered in and rendered in; the blurb is the wizard label
# and the instructions text, so the human sees the same words everywhere.
CRISIS_ROLES = [
    ("admin", "paperwork: banks, probate, the estate"),
    ("organiser", "the funeral, the wishes, the logistics"),
    ("local", "on the ground: house, keys, neighbours, pets"),
    ("medical", "GP, hospital, medical and end-of-life conversations"),
    ("comms", "telling people, and keeping everyone updated"),
]

ROLE_NAMES = [name for name, _ in CRISIS_ROLES]
ROLE_BLURB = dict(CRISIS_ROLES)

# The SAME role in a different form — a US spelling, a plural, the long
# word. Accepted quietly because the human plainly meant the role.
# Anything not here and not a role name is reported, never assumed.
ALIASES = {
    "organizer": "organiser",
    "organisor": "organiser",
    "organizor": "organiser",
    "organising": "organiser",
    "organizing": "organiser",
    "administrator": "admin",
    "administration": "admin",
    "admins": "admin",
    "paperwork": "admin",
    "communications": "comms",
    "communication": "comms",
    "comm": "comms",
    "medic": "medical",
    "locals": "local",
}

# Commas are the documented separator; the rest are what people actually
# type when they're filling a spreadsheet at speed.
SEPARATORS = re.compile(r"[,;/\n]+")


def _normalise(token: str) -> str:
    """Lower-case, collapse spacing/underscores/hyphens, apply aliases."""
    flat = re.sub(r"[\s_\-]+", " ", token.strip().lower())
    return ALIASES.get(flat, flat)


def parse(text) -> tuple:
    """``"admin, Organizer, hlper"`` -> ``(["admin", "organiser"],
    ["hlper"])``.

    Returns (held, unknown): canonical role names in the order listed,
    de-duplicated, plus the raw tokens that matched no role — kept
    verbatim so a warning can quote back exactly what was typed.
    """
    held, unknown = [], []
    for raw in SEPARATORS.split(str(text or "")):
        token = raw.strip()
        if not token:
            continue
        name = _normalise(token)
        if name in ROLE_BLURB:
            if name not in held:
                held.append(name)
        elif token not in unknown:
            unknown.append(token)
    return held, unknown


def format_roles(names) -> str:
    """Canonical cell text: known roles, canonical order, comma-space.

    Takes either one name per item (a ticked box each) or items that are
    themselves comma-separated (a cell, or a form posted by hand), so the
    caller never has to know which shape it has.
    """
    held, _ = parse(", ".join(str(n) for n in names))
    return ", ".join(role for role in ROLE_NAMES if role in held)


def describe() -> str:
    """One-line human list of the roles, for instructions and warnings."""
    return ", ".join(ROLE_NAMES)


if __name__ == "__main__":
    for name, blurb in CRISIS_ROLES:
        print(f"{name:<10} {blurb}")

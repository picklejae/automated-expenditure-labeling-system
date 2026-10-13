"""SQLite storage for the label memory, amount history and section list.

The database is a single local file. It holds three tables:

    labels       commitment item + description  -> section
                 (the "memory": once you label a line, it is labeled forever)
    sof_history  every line ever seen, with its last known amount
                 (used to flag lines whose amount changed since the last report)
    sections     the section names offered in the UI dropdown
"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from typing import Iterable, Iterator

# Lines charged to the unit by mistake (another unit's purchase). They are
# remembered like any other label, but kept out of the unit's totals.
EXCLUDED_SECTION = "Not our unit"

DEFAULT_SECTIONS = [
    "Command Group",
    "HQ Co",
    "Alpha Co",
    "Bravo Co",
    "Charlie Co",
    "Maintenance",
    "Medical",
    "Supply",
    "Other",
    EXCLUDED_SECTION,
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS labels (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    commitment_item TEXT NOT NULL,
    description     TEXT NOT NULL,
    section         TEXT NOT NULL,
    updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (commitment_item, description)
);

CREATE TABLE IF NOT EXISTS sof_history (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    obligation_doc  TEXT NOT NULL,
    commitment_item TEXT NOT NULL,
    description     TEXT NOT NULL,
    amount          REAL,
    imported_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (obligation_doc, commitment_item, description)
);

CREATE TABLE IF NOT EXISTS sections (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);
"""


class LabelStore:
    """Thin wrapper around the SQLite file. One instance per database path."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init()

    @contextmanager
    def _conn(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init(self) -> None:
        with self._conn() as conn:
            conn.executescript(SCHEMA)
            conn.executemany(
                "INSERT OR IGNORE INTO sections (name) VALUES (?)",
                [(name,) for name in DEFAULT_SECTIONS],
            )

    # ---- labels -----------------------------------------------------------

    def get_labels(self) -> dict[tuple[str, str], str]:
        """Return every known label as {(commitment_item, description): section}.

        Loading the whole table once is much faster than one query per line
        when a report has 1,000+ rows.
        """
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT commitment_item, description, section FROM labels"
            ).fetchall()
        return {(r["commitment_item"], r["description"]): r["section"] for r in rows}

    def save_label(self, commitment_item: str, description: str, section: str) -> None:
        self.save_labels([(commitment_item, description, section)])

    def save_labels(self, rows: Iterable[tuple[str, str, str]]) -> None:
        """Bulk upsert (commitment_item, description, section) rows."""
        with self._conn() as conn:
            conn.executemany(
                """
                INSERT INTO labels (commitment_item, description, section)
                VALUES (?, ?, ?)
                ON CONFLICT (commitment_item, description)
                DO UPDATE SET section = excluded.section,
                              updated_at = CURRENT_TIMESTAMP
                """,
                list(rows),
            )

    # ---- amount history ---------------------------------------------------

    def get_history(self) -> dict[tuple[str, str, str], float]:
        """Return {(obligation_doc, commitment_item, description): last_amount}."""
        with self._conn() as conn:
            rows = conn.execute(
                "SELECT obligation_doc, commitment_item, description, amount FROM sof_history"
            ).fetchall()
        return {
            (r["obligation_doc"], r["commitment_item"], r["description"]): r["amount"]
            for r in rows
        }

    def save_history(self, rows: Iterable[tuple[str, str, str, float]]) -> None:
        """Upsert (obligation_doc, commitment_item, description, amount) rows."""
        with self._conn() as conn:
            conn.executemany(
                """
                INSERT INTO sof_history (obligation_doc, commitment_item, description, amount)
                VALUES (?, ?, ?, ?)
                ON CONFLICT (obligation_doc, commitment_item, description)
                DO UPDATE SET amount = excluded.amount,
                              imported_at = CURRENT_TIMESTAMP
                """,
                list(rows),
            )

    # ---- sections ---------------------------------------------------------

    def get_sections(self) -> list[str]:
        with self._conn() as conn:
            rows = conn.execute("SELECT name FROM sections ORDER BY id").fetchall()
        return [r["name"] for r in rows]

    def add_section(self, name: str) -> None:
        with self._conn() as conn:
            conn.execute("INSERT OR IGNORE INTO sections (name) VALUES (?)", (name,))

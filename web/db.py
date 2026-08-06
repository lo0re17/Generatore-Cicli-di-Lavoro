"""Accesso SQLite per il Generatore Cicli web.

Scelta architetturale: il DB tiene *anagrafiche e tracciabilita'*
(clienti, PN, utenti, catalogo controlli, log generazioni), mentre i
**template restano su filesystem** in ``templates/<PN>/`` come nella
desktop (template.xlsx + config.json). Cosi' il motore openpyxl e i
wizard esistenti si riusano invariati e i template restano modificabili
a mano senza passare dal DB.
"""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "instance" / "database.db"

RUOLI_VALIDI = ("ufficio_tecnico", "qualita", "magazzino_produzione", "admin")
TIPI_DOCUMENTO = ("odl", "riq", "cqc")


def connessione() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS utente (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    nome_visualizzato TEXT NOT NULL DEFAULT '',
    ruolo TEXT NOT NULL CHECK (ruolo IN
        ('ufficio_tecnico', 'qualita', 'magazzino_produzione', 'admin'))
);

CREATE TABLE IF NOT EXISTS cliente (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    indirizzo TEXT NOT NULL DEFAULT '',
    contatti TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS anagrafica_pn (
    pn TEXT PRIMARY KEY,
    cliente_id INTEGER REFERENCES cliente(id) ON DELETE SET NULL,
    descrizione TEXT NOT NULL DEFAULT '',
    codice_disegno TEXT NOT NULL DEFAULT '',
    revisione TEXT NOT NULL DEFAULT '',
    unita_misura TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS catalogo_controlli (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    categoria TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS log_generazione (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    quando TEXT NOT NULL,
    utente_id INTEGER REFERENCES utente(id) ON DELETE SET NULL,
    username TEXT NOT NULL DEFAULT '',
    pn TEXT NOT NULL DEFAULT '',
    tipi_documento TEXT NOT NULL DEFAULT '',
    ordine_interno TEXT NOT NULL DEFAULT '',
    quantita INTEGER NOT NULL DEFAULT 0,
    file_generati TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS progressivo_riq (
    anno TEXT PRIMARY KEY,
    ultimo INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS ultima_compilazione (
    pn TEXT PRIMARY KEY,
    dati TEXT NOT NULL,
    quando TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS fornitore (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    contatti TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS log_operazione (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    quando TEXT NOT NULL,
    username TEXT NOT NULL DEFAULT '',
    azione TEXT NOT NULL,
    dettaglio TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS log_errore (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    quando TEXT NOT NULL,
    username TEXT NOT NULL DEFAULT '',
    percorso TEXT NOT NULL DEFAULT '',
    messaggio TEXT NOT NULL DEFAULT '',
    dettaglio TEXT NOT NULL DEFAULT ''
);
"""

CONTROLLI_INIZIALI = [
    ("Controllo materiale", "materiale"),
    ("Controllo trattamento termico", "processo"),
    ("Controllo rivestimento", "processo"),
    ("Controllo dimensionale", "dimensionale"),
    ("Test non distruttivi", "test"),
    ("Test elettrico", "test"),
    ("Test funzionale", "test"),
    ("Test vibrazionale", "test"),
    ("Test termico", "test"),
    ("Test idraulico", "test"),
]


def _migra_ruolo_admin(conn: sqlite3.Connection) -> None:
    """Su DB creati prima del ruolo 'admin' il CHECK della tabella utente
    non lo ammette: ricostruisce la tabella con lo schema aggiornato."""
    riga = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='utente'"
    ).fetchone()
    if riga and "'admin'" in (riga["sql"] or ""):
        return
    conn.executescript("""
        ALTER TABLE utente RENAME TO utente_vecchia;
        CREATE TABLE utente (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            nome_visualizzato TEXT NOT NULL DEFAULT '',
            ruolo TEXT NOT NULL CHECK (ruolo IN
                ('ufficio_tecnico', 'qualita', 'magazzino_produzione', 'admin'))
        );
        INSERT INTO utente SELECT * FROM utente_vecchia;
        DROP TABLE utente_vecchia;
    """)


def inizializza() -> None:
    with connessione() as conn:
        conn.executescript(SCHEMA)
        _migra_ruolo_admin(conn)

        gia_presenti = conn.execute(
            "SELECT COUNT(*) AS n FROM catalogo_controlli").fetchone()["n"]
        if not gia_presenti:
            conn.executemany(
                "INSERT INTO catalogo_controlli (nome, categoria) VALUES (?, ?)",
                CONTROLLI_INIZIALI)

        admin = conn.execute(
            "SELECT id FROM utente WHERE username = 'admin'").fetchone()
        if not admin:
            from werkzeug.security import generate_password_hash
            conn.execute(
                "INSERT INTO utente (username, password_hash, nome_visualizzato, ruolo) "
                "VALUES (?, ?, ?, ?)",
                ("admin", generate_password_hash("admin26"), "Amministratore", "admin"))


# --------------------------------------------------------------------- utenti
def trova_utente(username: str) -> sqlite3.Row | None:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM utente WHERE username = ?", (username,)).fetchone()


def crea_utente(username: str, password_hash: str, nome_visualizzato: str,
                ruolo: str) -> None:
    if ruolo not in RUOLI_VALIDI:
        raise ValueError(f"Ruolo non valido: {ruolo}")
    with connessione() as conn:
        conn.execute(
            "INSERT INTO utente (username, password_hash, nome_visualizzato, ruolo) "
            "VALUES (?, ?, ?, ?)",
            (username, password_hash, nome_visualizzato, ruolo))


def utenti() -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute(
            "SELECT id, username, nome_visualizzato, ruolo FROM utente "
            "ORDER BY username").fetchall()


def utente_per_id(utente_id: int) -> sqlite3.Row | None:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM utente WHERE id = ?", (utente_id,)).fetchone()


def aggiorna_password_utente(utente_id: int, password_hash: str) -> None:
    with connessione() as conn:
        conn.execute("UPDATE utente SET password_hash = ? WHERE id = ?",
                     (password_hash, utente_id))


def elimina_utente(utente_id: int) -> None:
    with connessione() as conn:
        conn.execute("DELETE FROM utente WHERE id = ?", (utente_id,))


# -------------------------------------------------------------------- clienti
def clienti() -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute("""
            SELECT c.*, (SELECT COUNT(*) FROM anagrafica_pn a
                         WHERE a.cliente_id = c.id) AS n_pn
            FROM cliente c ORDER BY c.nome
        """).fetchall()


def cliente_per_nome(nome: str) -> sqlite3.Row | None:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM cliente WHERE nome = ?", (nome.strip(),)).fetchone()


def crea_cliente(nome: str, indirizzo: str = "", contatti: str = "",
                 note: str = "") -> int:
    with connessione() as conn:
        cur = conn.execute(
            "INSERT INTO cliente (nome, indirizzo, contatti, note) "
            "VALUES (?, ?, ?, ?)", (nome.strip(), indirizzo, contatti, note))
        return int(cur.lastrowid)


def trova_o_crea_cliente(nome: str) -> int:
    esistente = cliente_per_nome(nome)
    return esistente["id"] if esistente else crea_cliente(nome)


def aggiorna_cliente(cliente_id: int, nome: str, indirizzo: str,
                     contatti: str, note: str) -> None:
    with connessione() as conn:
        conn.execute(
            "UPDATE cliente SET nome = ?, indirizzo = ?, contatti = ?, note = ? "
            "WHERE id = ?", (nome.strip(), indirizzo, contatti, note, cliente_id))


def elimina_cliente(cliente_id: int) -> None:
    with connessione() as conn:
        conn.execute("DELETE FROM cliente WHERE id = ?", (cliente_id,))


# ------------------------------------------------------------- anagrafica PN
def pn_tutti(cliente_id: int | None = None,
             solo_non_assegnati: bool = False) -> list[sqlite3.Row]:
    sql = """
        SELECT a.*, c.nome AS cliente_nome
        FROM anagrafica_pn a LEFT JOIN cliente c ON c.id = a.cliente_id
    """
    parametri: tuple = ()
    if solo_non_assegnati:
        sql += " WHERE a.cliente_id IS NULL"
    elif cliente_id is not None:
        sql += " WHERE a.cliente_id = ?"
        parametri = (cliente_id,)
    sql += " ORDER BY a.pn"
    with connessione() as conn:
        return conn.execute(sql, parametri).fetchall()


def pn_singolo(pn: str) -> sqlite3.Row | None:
    with connessione() as conn:
        return conn.execute("""
            SELECT a.*, c.nome AS cliente_nome
            FROM anagrafica_pn a LEFT JOIN cliente c ON c.id = a.cliente_id
            WHERE a.pn = ?
        """, (pn,)).fetchone()


def conta_pn_non_assegnati() -> int:
    with connessione() as conn:
        return conn.execute(
            "SELECT COUNT(*) AS n FROM anagrafica_pn WHERE cliente_id IS NULL"
        ).fetchone()["n"]


def salva_pn(pn: str, cliente_id: int | None, descrizione: str = "",
             codice_disegno: str = "", revisione: str = "",
             unita_misura: str = "") -> None:
    """Inserisce o aggiorna un PN in anagrafica (upsert)."""
    with connessione() as conn:
        conn.execute("""
            INSERT INTO anagrafica_pn
                (pn, cliente_id, descrizione, codice_disegno, revisione, unita_misura)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(pn) DO UPDATE SET
                cliente_id = excluded.cliente_id,
                descrizione = excluded.descrizione,
                codice_disegno = excluded.codice_disegno,
                revisione = excluded.revisione,
                unita_misura = excluded.unita_misura
        """, (pn.strip(), cliente_id, descrizione, codice_disegno,
              revisione, unita_misura))


def assegna_cliente_a_pn(pn: str, cliente_id: int | None) -> None:
    with connessione() as conn:
        conn.execute("UPDATE anagrafica_pn SET cliente_id = ? WHERE pn = ?",
                     (cliente_id, pn))


def elimina_pn(pn: str) -> None:
    with connessione() as conn:
        conn.execute("DELETE FROM anagrafica_pn WHERE pn = ?", (pn,))


# ---------------------------------------------------------- catalogo controlli
def controlli() -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM catalogo_controlli ORDER BY categoria, nome").fetchall()


def crea_controllo(nome: str, categoria: str = "") -> int:
    with connessione() as conn:
        cur = conn.execute(
            "INSERT INTO catalogo_controlli (nome, categoria) VALUES (?, ?)",
            (nome.strip(), categoria.strip()))
        return int(cur.lastrowid)


def elimina_controllo(controllo_id: int) -> None:
    with connessione() as conn:
        conn.execute("DELETE FROM catalogo_controlli WHERE id = ?", (controllo_id,))


# ------------------------------------------------------------ log generazioni
def registra_generazione(utente_id: int | None, username: str, pn: str,
                         tipi_documento: str, ordine_interno: str,
                         quantita: int, file_generati: list[str]) -> None:
    with connessione() as conn:
        conn.execute("""
            INSERT INTO log_generazione
                (quando, utente_id, username, pn, tipi_documento,
                 ordine_interno, quantita, file_generati)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (datetime.now().isoformat(timespec="seconds"), utente_id, username,
              pn, tipi_documento, ordine_interno, quantita,
              "\n".join(file_generati)))


def storico(limite: int = 100) -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM log_generazione ORDER BY id DESC LIMIT ?",
            (limite,)).fetchall()


# ---------------------------------------------------------- progressivo RIQ
def prossimo_progressivo_riq(anno: str, quanti: int = 1) -> list[int]:
    """Riserva ``quanti`` numeri progressivi RIQ per l'anno dato.

    Usato quando il RIQ ha numerazione indipendente dall'ODL.
    """
    with connessione() as conn:
        riga = conn.execute(
            "SELECT ultimo FROM progressivo_riq WHERE anno = ?", (anno,)).fetchone()
        ultimo = riga["ultimo"] if riga else 0
        nuovo = ultimo + quanti
        conn.execute("""
            INSERT INTO progressivo_riq (anno, ultimo) VALUES (?, ?)
            ON CONFLICT(anno) DO UPDATE SET ultimo = excluded.ultimo
        """, (anno, nuovo))
        return list(range(ultimo + 1, nuovo + 1))


def imposta_progressivo_riq(anno: str, ultimo: int) -> None:
    with connessione() as conn:
        conn.execute("""
            INSERT INTO progressivo_riq (anno, ultimo) VALUES (?, ?)
            ON CONFLICT(anno) DO UPDATE SET ultimo = excluded.ultimo
        """, (anno, ultimo))


# ---------------------------------------------------------- ultima compilazione
def salva_ultima_compilazione(pn: str, dati: dict) -> None:
    """Ricorda i valori dell'ultima generazione per un PN, per proporli
    come default alla prossima apertura del form (in qualunque sessione)."""
    with connessione() as conn:
        conn.execute("""
            INSERT INTO ultima_compilazione (pn, dati, quando) VALUES (?, ?, ?)
            ON CONFLICT(pn) DO UPDATE SET dati = excluded.dati, quando = excluded.quando
        """, (pn, json.dumps(dati, ensure_ascii=False),
              datetime.now().isoformat(timespec="seconds")))


def ultima_compilazione(pn: str) -> dict | None:
    with connessione() as conn:
        riga = conn.execute(
            "SELECT dati FROM ultima_compilazione WHERE pn = ?", (pn,)).fetchone()
    return json.loads(riga["dati"]) if riga else None


# ------------------------------------------------------------------ fornitori
def fornitori() -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute("SELECT * FROM fornitore ORDER BY nome").fetchall()


def crea_fornitore(nome: str, contatti: str = "", note: str = "") -> int:
    with connessione() as conn:
        cur = conn.execute(
            "INSERT INTO fornitore (nome, contatti, note) VALUES (?, ?, ?)",
            (nome.strip(), contatti, note))
        return int(cur.lastrowid)


def aggiorna_fornitore(fornitore_id: int, nome: str, contatti: str, note: str) -> None:
    with connessione() as conn:
        conn.execute(
            "UPDATE fornitore SET nome = ?, contatti = ?, note = ? WHERE id = ?",
            (nome.strip(), contatti, note, fornitore_id))


def elimina_fornitore(fornitore_id: int) -> None:
    with connessione() as conn:
        conn.execute("DELETE FROM fornitore WHERE id = ?", (fornitore_id,))


# --------------------------------------------------------------------- log
def registra_operazione(username: str, azione: str, dettaglio: str = "") -> None:
    with connessione() as conn:
        conn.execute(
            "INSERT INTO log_operazione (quando, username, azione, dettaglio) "
            "VALUES (?, ?, ?, ?)",
            (datetime.now().isoformat(timespec="seconds"), username, azione, dettaglio))


def log_operazioni(limite: int = 200) -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM log_operazione ORDER BY id DESC LIMIT ?",
            (limite,)).fetchall()


def registra_errore(username: str, percorso: str, messaggio: str, dettaglio: str = "") -> None:
    with connessione() as conn:
        conn.execute(
            "INSERT INTO log_errore (quando, username, percorso, messaggio, dettaglio) "
            "VALUES (?, ?, ?, ?, ?)",
            (datetime.now().isoformat(timespec="seconds"), username, percorso,
             messaggio, dettaglio))


def log_errori(limite: int = 200) -> list[sqlite3.Row]:
    with connessione() as conn:
        return conn.execute(
            "SELECT * FROM log_errore ORDER BY id DESC LIMIT ?",
            (limite,)).fetchall()

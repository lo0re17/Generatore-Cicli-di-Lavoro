"""Crea utenti di prova, uno per ruolo, per collaudare lo scheletro.

Uso:
    python seed_utenti.py
"""

from __future__ import annotations

from werkzeug.security import generate_password_hash

import db

UTENTI_PROVA = [
    ("ut", "Ufficio Tecnico", "ufficio_tecnico"),
    ("qualita", "Qualità", "qualita"),
    ("produzione", "Magazzino e Produzione", "magazzino_produzione"),
]


def main() -> None:
    db.inizializza()
    for username, nome, ruolo in UTENTI_PROVA:
        if db.trova_utente(username) is not None:
            print(f"Utente '{username}' già presente, salto.")
            continue
        db.crea_utente(username, generate_password_hash("prova123"), nome, ruolo)
        print(f"Creato utente '{username}' (ruolo {ruolo}), password: prova123")


if __name__ == "__main__":
    main()

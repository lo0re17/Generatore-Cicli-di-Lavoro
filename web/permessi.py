"""Permessi per ruolo.

Ufficio Tecnico: redige template e fasi.
Qualita': redige, compila e revisiona.
Magazzino e Produzione: solo compila (genera cicli dai template esistenti).
Admin: tutto, piu' la pagina Configurazioni e DB (utenti, operatori, log...).
"""

from __future__ import annotations

AZIONI_PER_RUOLO = {
    "ufficio_tecnico": {"redige", "compila"},
    "qualita": {"redige", "compila", "revisiona"},
    "magazzino_produzione": {"compila"},
    "admin": {"redige", "compila", "revisiona", "amministra"},
}

ETICHETTE_RUOLO = {
    "ufficio_tecnico": "Ufficio Tecnico",
    "qualita": "Qualità",
    "magazzino_produzione": "Magazzino e Produzione",
    "admin": "Amministratore",
}


def puo(ruolo: str, azione: str) -> bool:
    return azione in AZIONI_PER_RUOLO.get(ruolo, set())

"""Anagrafica operatori (con firma .png) e loghi azienda/clienti.

Dati salvati in una cartella (di default ``anagrafica/`` accanto all'app):
    operatori.json   -> {"operatori": [{sigla, nome, reparto, firma, e_cq, firma_cq}]}
    firme/*.png      -> immagini delle firme (operatore e CQ)
    loghi.json       -> {"azienda": "loghi/..png", "clienti": {"Nome": "loghi/..png"}}
    loghi/*.png
"""

from __future__ import annotations

import json
from dataclasses import dataclass, asdict, field
from pathlib import Path


@dataclass
class Operatore:
    sigla: str
    nome: str
    reparto: str = ""
    matricola: str = ""
    firma: str = ""            # percorso relativo alla cartella anagrafica
    e_cq: bool = False         # abilitato come firmatario Controllo Qualita' (RIQ)
    firma_cq: str = ""         # percorso firma CQ (puo' differire dalla firma standard)

    def etichetta(self) -> str:
        parti = [self.nome or self.sigla]
        if self.sigla and self.nome:
            parti[0] = f"{self.nome} ({self.sigla})"
        if self.reparto:
            parti.append(f"[{self.reparto}]")
        return " ".join(parti)


class Anagrafica:
    def __init__(self, cartella: str | Path):
        self.cartella = Path(cartella)
        self.operatori: list[Operatore] = []
        self.logo_azienda: str = ""
        self.loghi_clienti: dict[str, str] = {}
        self.carica()

    # ------------------------------------------------------------- percorsi
    @property
    def file_operatori(self) -> Path:
        return self.cartella / "operatori.json"

    @property
    def file_loghi(self) -> Path:
        return self.cartella / "loghi.json"

    def percorso_firma(self, op: Operatore) -> Path | None:
        if not op.firma:
            return None
        p = (self.cartella / op.firma)
        return p if p.is_file() else None

    def percorso_firma_cq(self, op: Operatore) -> Path | None:
        if not op.firma_cq:
            return None
        p = (self.cartella / op.firma_cq)
        return p if p.is_file() else None

    def operatori_cq(self) -> list[Operatore]:
        return [o for o in self.operatori if o.e_cq]

    def percorso_logo_cliente(self, cliente: str) -> Path | None:
        rel = self.loghi_clienti.get(cliente)
        if not rel:
            return None
        p = self.cartella / rel
        return p if p.is_file() else None

    # ------------------------------------------------------------- carica/salva
    def carica(self) -> None:
        if self.file_operatori.is_file():
            dati = json.loads(self.file_operatori.read_text(encoding="utf-8"))
            self.operatori = [Operatore(**o) for o in dati.get("operatori", [])]
        if self.file_loghi.is_file():
            dati = json.loads(self.file_loghi.read_text(encoding="utf-8"))
            self.logo_azienda = dati.get("azienda", "")
            self.loghi_clienti = dati.get("clienti", {})

    def salva(self) -> None:
        self.cartella.mkdir(parents=True, exist_ok=True)
        self.file_operatori.write_text(
            json.dumps({"operatori": [asdict(o) for o in self.operatori]},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.file_loghi.write_text(
            json.dumps({"azienda": self.logo_azienda, "clienti": self.loghi_clienti},
                       ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    # ------------------------------------------------------------- utilita'
    def operatore(self, sigla: str) -> Operatore | None:
        for o in self.operatori:
            if o.sigla == sigla:
                return o
        return None

    def aggiungi_o_aggiorna(self, op: Operatore) -> None:
        for i, o in enumerate(self.operatori):
            if o.sigla == op.sigla:
                self.operatori[i] = op
                return
        self.operatori.append(op)

    def rimuovi(self, sigla: str) -> None:
        self.operatori = [o for o in self.operatori if o.sigla != sigla]

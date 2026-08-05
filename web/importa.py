"""Import massivo dell'anagrafica articoli da export del gestionale.

L'export reale non ha una colonna cliente e mescola articoli veri con
manodopera e servizi, quindi:
  - la **mappatura colonne e' libera** (l'utente dice quale colonna e' il
    PN, quale la descrizione, ecc.): funziona con export di formato diverso;
  - si filtra per **tipo articolo** (colonna Tp: A = articolo, L = manodopera,
    S = servizio), di default solo gli articoli;
  - il **cliente e' opzionale**: le righe senza cliente restano non
    assegnate e si abbinano dopo.
"""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

# Nell'export il codice porta la revisione accodata dopo degli spazi:
#   "70007927-P002 2" -> PN "70007927-P002" rev "2"
#   "610410243     03" -> PN "610410243"     rev "03"
#   "MOD-STP-1490  --" -> PN "MOD-STP-1490"  rev "--"
_RE_CODICE_REV = re.compile(r"^(?P<pn>.+?)\s+(?P<rev>\d{1,3}|-{1,3})$")

# Tipi articolo del gestionale (colonna "Tp")
TIPI_ARTICOLO = {
    "A": "Articolo",
    "L": "Manodopera",
    "S": "Servizio",
}


@dataclass
class RigaImport:
    pn: str
    revisione: str = ""
    descrizione: str = ""
    cliente: str = ""
    tipo: str = ""
    unita_misura: str = ""
    codice_alternativo: str = ""
    esito: str = ""          # nuovo | aggiornato | escluso
    motivo: str = ""


@dataclass
class Anteprima:
    intestazioni: list[str]
    righe: list[RigaImport] = field(default_factory=list)
    esclusi: int = 0
    avvisi: list[str] = field(default_factory=list)


def leggi_tabella(percorso: str | Path) -> tuple[list[str], list[list[str]]]:
    """Legge .xlsx/.csv/.tsv e ritorna (intestazioni, righe di testo)."""
    percorso = Path(percorso)
    if percorso.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(percorso, data_only=True)
        ws = wb.active
        righe = [[("" if c is None else str(c).strip()) for c in riga]
                 for riga in ws.iter_rows(values_only=True)]
    else:
        testo = percorso.read_text(encoding="utf-8-sig", errors="replace")
        delimitatore = "\t" if "\t" in testo.split("\n")[0] else ";"
        if delimitatore == ";" and testo.count(",") > testo.count(";"):
            delimitatore = ","
        righe = [[c.strip() for c in r]
                 for r in csv.reader(io.StringIO(testo), delimiter=delimitatore)]

    righe = [r for r in righe if any(c for c in r)]
    if not righe:
        return [], []

    # L'export del gestionale ha una prima colonna vuota: la si tiene,
    # l'utente mappera' solo le colonne che servono.
    intestazioni = righe[0]
    intestazioni = [h or f"(colonna {i + 1})" for i, h in enumerate(intestazioni)]
    return intestazioni, righe[1:]


def _valore(riga: list[str], indice: int | None) -> str:
    if indice is None or indice < 0 or indice >= len(riga):
        return ""
    return riga[indice].strip()


def separa_pn_revisione(codice: str) -> tuple[str, str]:
    """Divide 'PN<spazi>REV' nelle sue due parti; rev vuota se assente."""
    codice = " ".join(codice.split())  # normalizza gli spazi multipli
    m = _RE_CODICE_REV.match(codice)
    if m:
        return m.group("pn").strip(), m.group("rev").strip()
    return codice, ""


def costruisci_anteprima(
    intestazioni: list[str],
    righe: list[list[str]],
    col_pn: int,
    col_descrizione: int | None = None,
    col_cliente: int | None = None,
    col_tipo: int | None = None,
    col_um: int | None = None,
    col_cod_alt: int | None = None,
    tipi_ammessi: tuple[str, ...] = ("A",),
    pn_esistenti: set[str] | None = None,
    separa_revisione: bool = True,
) -> Anteprima:
    """Applica mappatura e filtri, e classifica ogni riga."""
    pn_esistenti = pn_esistenti or set()
    anteprima = Anteprima(intestazioni=intestazioni)
    visti: set[str] = set()

    for riga in righe:
        codice = _valore(riga, col_pn)
        if not codice:
            continue
        pn, revisione = (separa_pn_revisione(codice) if separa_revisione
                         else (" ".join(codice.split()), ""))

        tipo = _valore(riga, col_tipo).upper()
        if col_tipo is not None and tipi_ammessi and tipo not in tipi_ammessi:
            anteprima.esclusi += 1
            continue

        if pn in visti:
            anteprima.esclusi += 1
            continue
        visti.add(pn)

        voce = RigaImport(
            pn=pn,
            revisione=revisione,
            descrizione=_valore(riga, col_descrizione),
            cliente=_valore(riga, col_cliente),
            tipo=tipo,
            unita_misura=_valore(riga, col_um),
            codice_alternativo=_valore(riga, col_cod_alt),
        )
        voce.esito = "aggiornato" if pn in pn_esistenti else "nuovo"
        anteprima.righe.append(voce)

    if anteprima.esclusi:
        anteprima.avvisi.append(
            f"{anteprima.esclusi} righe escluse (tipo non ammesso o PN duplicato).")
    if not anteprima.righe:
        anteprima.avvisi.append("Nessuna riga importabile con questa mappatura.")
    return anteprima


def indovina_colonne(intestazioni: list[str]) -> dict[str, int | None]:
    """Propone una mappatura leggendo i nomi delle intestazioni."""
    def cerca(*chiavi: str) -> int | None:
        for i, h in enumerate(intestazioni):
            testo = h.lower()
            if any(k in testo for k in chiavi):
                return i
        return None

    return {
        "col_pn": cerca("codice", "articolo", "pn", "part"),
        "col_descrizione": cerca("descrizione", "descr"),
        "col_cliente": cerca("cliente", "committente"),
        "col_tipo": cerca("tp", "tipo"),
        "col_um": cerca("um", "unità", "unita"),
        "col_cod_alt": cerca("alternativo", "alt"),
    }

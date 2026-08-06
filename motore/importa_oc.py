"""Lettura degli Ordini Cliente (OC) esportati dal gestionale come
"Stampa elenco su Excel" in formato SpreadsheetML (.xml).

Non genera cicli in autonomo (la corrispondenza Commessa -> PN richiede
conferma umana, i codici del gestionale non coincidono sempre col PN
usato nei template): estrae le righe in forma tabellare cosi' l'utente
puo' scegliere quella giusta e usarla per precompilare cliente e dati
del ciclo nel form di generazione.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

_NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

# Intestazioni attese -> chiave interna. Le colonne non elencate qui
# vengono ignorate.
_COLONNE = {
    "Tipo": "tipo",
    "Serie": "serie",
    "Numero": "numero",
    "Data doc": "data_doc",
    "Cliente/fornitore": "cliente",
    "Note": "note",
    "Denominazione": "denominazione",
    "Rif.doc.numero": "rif_doc_numero",
    "Rif.doc.data": "rif_doc_data",
    "Data trasporto": "data_trasporto",
    "Buyer": "buyer",
    "Commessa": "commessa",
    "Totale Documento": "totale",
}


@dataclass
class RigaOC:
    tipo: str = ""
    serie: str = ""
    numero: str = ""
    data_doc: str = ""
    cliente: str = ""
    note: str = ""
    denominazione: str = ""
    rif_doc_numero: str = ""
    rif_doc_data: str = ""
    data_trasporto: str = ""
    buyer: str = ""
    commessa: str = ""
    totale: str = ""


def _testo_cella(cella: ET.Element) -> str:
    dato = cella.find("ss:Data", _NS)
    return (dato.text or "").strip() if dato is not None and dato.text else ""


def leggi_oc(percorso: str | Path) -> list[RigaOC]:
    """Legge il primo foglio del file OC e ritorna le righe (dati, non
    intestazione), usando l'ordine delle colonne trovato in intestazione."""
    albero = ET.parse(percorso)
    tabella = albero.getroot().find(".//ss:Worksheet/ss:Table", _NS)
    if tabella is None:
        return []

    righe_xml = tabella.findall("ss:Row", _NS)
    if not righe_xml:
        return []

    def _celle_con_indice(riga: ET.Element) -> dict[int, str]:
        """Le celle usano ss:Index per saltare colonne vuote: ricostruisce
        la posizione reale di ciascuna cella."""
        celle: dict[int, str] = {}
        indice = 0
        for cella in riga.findall("ss:Cell", _NS):
            attr_indice = cella.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if attr_indice:
                indice = int(attr_indice) - 1
            celle[indice] = _testo_cella(cella)
            indice += 1
        return celle

    intestazione = _celle_con_indice(righe_xml[0])
    posizione_chiave: dict[int, str] = {}
    for idx, testo in intestazione.items():
        chiave = _COLONNE.get(testo)
        if chiave:
            posizione_chiave[idx] = chiave

    righe: list[RigaOC] = []
    for riga_xml in righe_xml[1:]:
        celle = _celle_con_indice(riga_xml)
        valori = {chiave: celle.get(idx, "") for idx, chiave in posizione_chiave.items()}
        if not any(valori.values()):
            continue
        righe.append(RigaOC(**valori))
    return righe

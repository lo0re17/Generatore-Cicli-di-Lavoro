"""Gestione dell'Ordine Interno progressivo.

L'Ordine Interno identifica in modo univoco un ciclo di lavoro.
Formato tipico: ``NNNN-YY`` (es. ``6935-26`` = progressivo 6935, anno 26).
Generando N cicli si incrementa di 1 la sola parte progressiva,
mantenendo l'anno e l'eventuale zero-padding.
"""

from __future__ import annotations

import re

# Progressivo + separatore + coda (anno). Es: 6935-26, 00042-25
_RE_NNNN_YY = re.compile(r"^(\s*)(\d+)(-\d+)(\s*)$")
# Fallback: incrementa l'ultimo gruppo di cifre presente nella stringa.
_RE_ULTIMO_NUM = re.compile(r"(\d+)(?!.*\d)")
# Progressivo-anno con anno inserito per esteso (4 cifre): 6935-2026
_RE_ANNO_ESTESO = re.compile(r"^(\s*\d+\s*-\s*)(\d{4})(\s*)$")


def normalizza(base: str) -> str:
    """Se l'anno e' stato digitato per esteso (4 cifre) lo riduce alle
    ultime 2, secondo la convenzione dell'Ordine Interno (NNNN-YY).

    >>> normalizza("6935-2026")
    '6935-26'
    >>> normalizza("6935-26")
    '6935-26'
    """
    m = _RE_ANNO_ESTESO.match(str(base))
    if m:
        pre, anno, post = m.groups()
        return f"{pre}{anno[-2:]}{post}"
    return base


def incrementa(base: str, passo: int = 1) -> str:
    """Restituisce l'Ordine Interno incrementato di ``passo``.

    >>> incrementa("6935-26")
    '6936-26'
    >>> incrementa("6935-26", 3)
    '6938-26'
    >>> incrementa("00042-25")
    '00043-25'
    >>> incrementa("125")
    '126'
    """
    base = str(base)

    m = _RE_NNNN_YY.match(base)
    if m:
        pre, prog, coda, post = m.groups()
        larghezza = len(prog)
        nuovo = str(int(prog) + passo).zfill(larghezza)
        return f"{pre}{nuovo}{coda}{post}"

    # Nessun formato NNNN-YY riconosciuto: incrementa l'ultimo numero.
    m = _RE_ULTIMO_NUM.search(base)
    if m:
        prog = m.group(1)
        larghezza = len(prog)
        nuovo = str(int(prog) + passo).zfill(larghezza)
        return base[: m.start(1)] + nuovo + base[m.end(1) :]

    raise ValueError(
        f"Ordine Interno '{base}' non contiene una parte numerica da incrementare."
    )


def serie(base: str, quantita: int) -> list[str]:
    """Genera la serie di ``quantita`` Ordini Interni a partire da ``base``.

    >>> serie("6935-26", 4)
    ['6935-26', '6936-26', '6937-26', '6938-26']
    """
    if quantita < 1:
        return []
    return [base if i == 0 else incrementa(base, i) for i in range(quantita)]


if __name__ == "__main__":
    import doctest

    risultati = doctest.testmod(verbose=False)
    print(f"doctest: {risultati.attempted} test, {risultati.failed} falliti")
    print("Esempio serie 6935-26 x5:", serie("6935-26", 5))

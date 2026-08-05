"""Crea dati di esempio nell'anagrafica: un operatore con firma .png generata.

Solo per sviluppo/test: l'operatore reale caricherà le proprie firme scansionate.
"""

from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

RADICE = Path(__file__).resolve().parent.parent
ANAG = RADICE / "anagrafica"
FIRME = ANAG / "firme"


def crea_firma_png(testo: str, percorso: Path) -> None:
    """Genera una finta firma corsiva su sfondo trasparente."""
    font = None
    for nome in ("segoesc.ttf", "brushsci.ttf", "arial.ttf"):
        try:
            font = ImageFont.truetype(f"C:/Windows/Fonts/{nome}", 96)
            break
        except OSError:
            continue
    if font is None:
        font = ImageFont.load_default()

    tmp = Image.new("RGBA", (1400, 400), (0, 0, 0, 0))
    d = ImageDraw.Draw(tmp)
    d.text((20, 20), testo, font=font, fill=(20, 30, 120, 255))
    bbox = tmp.getbbox()
    ritaglio = tmp.crop(bbox) if bbox else tmp
    percorso.parent.mkdir(parents=True, exist_ok=True)
    ritaglio.save(percorso)


def main() -> None:
    FIRME.mkdir(parents=True, exist_ok=True)
    crea_firma_png("N. Del Vecchio", FIRME / "NDV.png")

    operatori = {
        "operatori": [
            {"sigla": "NDV", "nome": "Nicola Del Vecchio", "reparto": "RT",
             "firma": "firme/NDV.png"},
        ]
    }
    (ANAG / "operatori.json").write_text(
        json.dumps(operatori, ensure_ascii=False, indent=2), encoding="utf-8")

    if not (ANAG / "loghi.json").exists():
        (ANAG / "loghi.json").write_text(
            json.dumps({"azienda": "", "clienti": {}}, ensure_ascii=False, indent=2),
            encoding="utf-8")

    print("Anagrafica di esempio creata in", ANAG)
    print(" - firma:", FIRME / "NDV.png")


if __name__ == "__main__":
    main()

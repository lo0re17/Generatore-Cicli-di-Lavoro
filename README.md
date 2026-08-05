# Generatore Cicli di Lavoro — F&N compositi

Tool per generare i **cicli di lavorazione** (Excel) partendo da un *layout per
ogni PN*. L'operatore (CQ / resp. produzione) sceglie il PN, compila i campi
variabili e genera **N cicli** con l'Ordine Interno che si incrementa da solo.

## Come funziona (approccio a segnaposto)

Ogni PN ha una cartella in `templates/<PN>/` con:

- `template.xlsx` — il ciclo con i campi variabili scritti come token `{{...}}`
  (es. `{{COMMESSA}}`, `{{ORDINE_INTERNO}}`, `{{LOTTO_3V000107}}`);
- `config.json` — descrive i campi (etichetta, tipo, obbligatorietà, gruppo) e
  il pattern del nome file.

Il motore copia il template e sostituisce i token. Lo **stesso token** può
comparire in più celle: viene sostituito ovunque → i **lotti** inseriti una
volta si propagano tra distinta base e testo delle fasi. Immagini/logo e
impostazioni di stampa (A4 orizzontale, scala 78%, area A1:AP55) sono preservate.

## Struttura

```
GeneratoreCicli/
  app.py                interfaccia operatore (tkinter, 3 schede)
  motore/
    ordine_interno.py   incremento NNNN-YY (+1 sul progressivo)
    generatore.py       sostituzione token, celle tipizzate, batch,
                        firma testo/immagine e data esecuzione per fase
    fasi.py             rilevamento tabella fasi (data/firma per riga)
    rileva_campi.py     riconoscimento automatico campi (wizard Nuovo PN)
    anagrafica.py       operatori (firme .png) e loghi
  templates/<PN>/
    template.xlsx       ciclo tokenizzato
    config.json         definizione campi
  anagrafica/
    operatori.json      sigla, nome, reparto, firma
    firme/*.png         immagini firma
    loghi.json, loghi/  logo azienda e clienti
  strumenti/
    crea_template_*.py  bootstrap template via script (alternativa al wizard)
    seed_anagrafica.py  dati anagrafica di esempio
    build_exe.ps1       build dell'eseguibile standalone
  dist/GeneratoreCicli/ eseguibile distribuibile (exe + dati accanto)
  output/               cicli generati
  test_motore.py        test end-to-end
```

## Uso dall'interfaccia

1. **Campi del ciclo** — scegli il PN, compila i campi (i lotti si propagano
   ovunque compaiano nel ciclo).
2. **Date e firme (opz.)** — per digitalizzare un processo già eseguito:
   data di esecuzione + firma operatore (testo o immagine .png) nelle fasi
   selezionate. Operatori e firme si gestiscono da "👥 Operatori…".
3. **Genera** — numero di cicli (l'Ordine Interno si incrementa da solo),
   eventuale ultimo ciclo con n. pezzi diverso, generazione immediata o
   **coda multi-PN** in un unico flusso.

**Nuovo PN**: pulsante "➕ Nuovo PN…" → scegli un ciclo compilato →
il riconoscimento automatico propone i campi variabili → rivedi → crea.

## Distribuzione

`strumenti\build_exe.ps1` produce `dist\GeneratoreCicli\`: copiala sul PC
dell'operatore e lancia `GeneratoreCicli.exe`. Non servono Python né Excel.
`templates/` e `anagrafica/` accanto all'exe si aggiornano senza ricompilare.

## Tipi di campo (`config.json`)

| tipo             | comportamento                                             |
|------------------|-----------------------------------------------------------|
| `testo`          | testo semplice                                            |
| `intero`         | scritto come numero (es. Nr. Pezzi)                       |
| `data`           | scritto come data Excel (mantiene il formato della cella) |
| `ordine_interno` | come `testo`, ma guida l'incremento del batch             |
| `lotto`          | testo, tipicamente con propagazione in più celle          |

## Uso da codice

```python
from motore import generatore as G
config = G.ConfigPN.da_file("templates/70009309-P002/config.json")
valori = {"{{COMMESSA}}": "SPRBHAL1-875", "{{ORDINE_INTERNO}}": "6935-26", ...}
G.genera_batch(config, valori, quantita=4, cartella_out="output",
               pezzi_override={3: 2})   # ultimo ciclo del lotto con 2 pezzi
```

## Aggiungere un nuovo PN

Oggi: duplicare `strumenti/crea_template_70009309.py`, adattare la mappatura
celle→token per il nuovo ciclo, eseguirlo. In arrivo (Fase 3): un **wizard**
grafico che marca i campi variabili da un ciclo compilato, con suggerimento
automatico.

## Requisiti

Python 3.10+ e `openpyxl` (`pip install -r requirements.txt`). Excel non è
necessario per generare; serve solo per l'eventuale export PDF/stampa diretta.

## Test

```
python test_motore.py
```

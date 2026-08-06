"""Generatore Cicli — applicazione web (prototipo locale).

Sostituisce l'app desktop tkinter mantenendo lo stesso motore openpyxl.
  - login con ruoli (Ufficio Tecnico / Qualita' / Magazzino e Produzione)
  - anagrafica clienti e PN, con import massivo dall'export del gestionale
  - generazione ODL (+ RIQ associato) in batch, con coda multi-PN
  - flag N/A per singolo campo, ultimo ciclo con pezzi diversi
  - storico delle generazioni

I template restano su filesystem in ``templates/<PN>/`` (template.xlsx +
config.json): il DB tiene anagrafiche, utenti e tracciabilita'.
"""

from __future__ import annotations

import functools
import io
import sqlite3
import sys
import tempfile
import traceback
import zipfile
from datetime import datetime
from pathlib import Path

from flask import (Flask, after_this_request, flash, redirect,
                    render_template, request, send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash

BASE = Path(__file__).resolve().parent
PROGETTO = BASE.parent
CARTELLA_TEMPLATES = PROGETTO / "templates"
CARTELLA_ANAGRAFICA = PROGETTO / "anagrafica"
CARTELLA_WIZARD = BASE / "instance" / "wizard"

sys.path.insert(0, str(PROGETTO))
sys.path.insert(0, str(BASE))

from motore import generatore as G  # noqa: E402
from motore import fasi as F  # noqa: E402
from motore import ordine_interno as OI  # noqa: E402
from motore import riq as RIQ  # noqa: E402
from motore import rileva_campi as RC  # noqa: E402
from motore import rileva_riq as RCR  # noqa: E402
from motore.anagrafica import Anagrafica, Operatore  # noqa: E402

import db  # noqa: E402
import importa  # noqa: E402
import permessi  # noqa: E402

import os  # noqa: E402

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-cambiare-in-produzione")

anagrafica_operatori = Anagrafica(CARTELLA_ANAGRAFICA)


# --------------------------------------------------------------------------- #
# Utilita'
# --------------------------------------------------------------------------- #
def template_disponibili() -> dict[str, Path]:
    """{PN: percorso config.json} per i PN che hanno un template ODL."""
    risultato: dict[str, Path] = {}
    if not CARTELLA_TEMPLATES.exists():
        return risultato
    for cartella in sorted(CARTELLA_TEMPLATES.iterdir()):
        cfg = cartella / "config.json"
        if cfg.is_file():
            try:
                risultato[G.ConfigPN.da_file(cfg).pn] = cfg
            except Exception:
                traceback.print_exc()
    return risultato


def config_riq_di(pn: str) -> RIQ.ConfigRIQ | None:
    cfg = RIQ.percorso_config_riq(CARTELLA_TEMPLATES / pn)
    if not cfg.is_file():
        return None
    try:
        return RIQ.ConfigRIQ.da_file(cfg)
    except Exception:
        traceback.print_exc()
        return None


def fasi_di(config: G.ConfigPN) -> list[F.Fase]:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(config.percorso_template())
        return F.rileva_fasi(wb.active)
    except Exception:
        traceback.print_exc()
        return []


def login_richiesto(vista):
    @functools.wraps(vista)
    def wrapper(*args, **kwargs):
        if "utente_id" not in session:
            return redirect(url_for("login"))
        return vista(*args, **kwargs)
    return wrapper


def serve_azione(azione: str):
    """Blocca la vista se il ruolo in sessione non puo' fare l'azione."""
    def decoratore(vista):
        @functools.wraps(vista)
        def wrapper(*args, **kwargs):
            if "utente_id" not in session:
                return redirect(url_for("login"))
            if not permessi.puo(session.get("ruolo", ""), azione):
                flash(f"Il tuo ruolo non è abilitato a questa operazione ({azione}).")
                return redirect(url_for("dashboard"))
            return vista(*args, **kwargs)
        return wrapper
    return decoratore


@app.context_processor
def variabili_comuni():
    ruolo = session.get("ruolo", "")
    return {
        "puo_redigere": permessi.puo(ruolo, "redige"),
        "puo_compilare": permessi.puo(ruolo, "compila"),
        "puo_revisionare": permessi.puo(ruolo, "revisiona"),
        "puo_amministrare": permessi.puo(ruolo, "amministra"),
        "etichetta_ruolo": permessi.ETICHETTE_RUOLO.get(ruolo, ruolo),
    }


def registra_azione(azione: str, dettaglio: str = "") -> None:
    db.registra_operazione(session.get("username", ""), azione, dettaglio)


@app.errorhandler(Exception)
def gestisci_errore_non_previsto(e):
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e
    dettaglio = traceback.format_exc()
    traceback.print_exc()
    db.registra_errore(session.get("username", ""), request.path, str(e), dettaglio)
    return render_template("errore.html", messaggio=str(e)), 500


# --------------------------------------------------------------------------- #
# Autenticazione
# --------------------------------------------------------------------------- #
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        utente = db.trova_utente(request.form.get("username", "").strip())
        password = request.form.get("password", "")
        if utente is None or not check_password_hash(utente["password_hash"], password):
            flash("Utente o password non validi.")
            return render_template("login.html")
        session.clear()
        session["utente_id"] = utente["id"]
        session["username"] = utente["username"]
        session["nome_visualizzato"] = utente["nome_visualizzato"]
        session["ruolo"] = utente["ruolo"]
        return redirect(url_for("dashboard"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# --------------------------------------------------------------------------- #
# Generazione
# --------------------------------------------------------------------------- #
@app.route("/")
@login_richiesto
def dashboard():
    templates = template_disponibili()
    filtro = request.args.get("cliente", "")

    if filtro == "non_assegnati":
        pn_ammessi = {r["pn"] for r in db.pn_tutti(solo_non_assegnati=True)}
    elif filtro.isdigit():
        pn_ammessi = {r["pn"] for r in db.pn_tutti(cliente_id=int(filtro))}
    else:
        pn_ammessi = None  # tutti

    pn_lista = sorted(templates if pn_ammessi is None
                      else {p: c for p, c in templates.items() if p in pn_ammessi})

    pn_scelto = request.args.get("pn") or (pn_lista[0] if pn_lista else None)
    config = (G.ConfigPN.da_file(templates[pn_scelto])
              if pn_scelto in templates else None)

    valori_salvati = session.get("form_valori", {}).get(pn_scelto, {})
    if not valori_salvati and pn_scelto:
        valori_salvati = db.ultima_compilazione(pn_scelto) or {}

    fasi_pn = fasi_di(config) if config else []
    operatori = anagrafica_operatori.operatori
    config_riq_pn = config_riq_di(pn_scelto) if pn_scelto else None

    return render_template(
        "dashboard.html",
        clienti=db.clienti(),
        n_non_assegnati=db.conta_pn_non_assegnati(),
        filtro_cliente=filtro,
        pn_lista=pn_lista,
        pn_scelto=pn_scelto,
        config=config,
        anagrafica_pn=db.pn_singolo(pn_scelto) if pn_scelto else None,
        config_riq=config_riq_pn,
        coda=session.get("coda", []),
        valori_salvati=valori_salvati,
        fasi=fasi_pn,
        operatori=operatori,
    )


def _raccogli_valori(config: G.ConfigPN) -> dict[str, object]:
    """Legge i campi dal form, rispettando i flag N/A per singolo campo.
    I campi fissi prendono il default dal config, non dal form."""
    valori: dict[str, object] = {}
    for c in config.campi:
        if c.fisso:
            valori[c.token] = c.default
            continue
        if request.form.get(f"na__{c.token}"):
            testo = request.form.get(f"na_testo__{c.token}", "").strip() or "-"
            valori[c.token] = G.NonApplicabile(testo)
        else:
            valori[c.token] = request.form.get(c.token, "")
    return valori


def _salva_form_in_sessione(pn: str) -> None:
    """Salva tutti i campi del form corrente in sessione, indicizzati per PN."""
    dati = {}
    for chiave, valore in request.form.items():
        if chiave not in ("pn", "fase"):
            dati[chiave] = valore
    form_valori = session.get("form_valori", {})
    form_valori[pn] = dati
    session["form_valori"] = form_valori


def _pulisci_form_sessione(pn: str) -> None:
    """Rimuove i valori salvati per un PN dopo generazione riuscita."""
    form_valori = session.get("form_valori", {})
    form_valori.pop(pn, None)
    session["form_valori"] = form_valori


def _costruisci_interventi(config: G.ConfigPN) -> list[dict]:
    """Legge dal form le date di esecuzione per fase e costruisce gli interventi."""
    if not request.form.get("compila_fasi"):
        return []
    fasi = fasi_di(config)
    if not fasi:
        return []
    data_globale = request.form.get("data_esecuzione", "").strip() or None
    op_sigla_globale = request.form.get("operatore_globale", "").strip()
    op_globale = anagrafica_operatori.operatore(op_sigla_globale) if op_sigla_globale else None

    interventi: list[dict] = []
    for i, fase in enumerate(fasi):
        if not request.form.get(f"fase_{i}"):
            continue
        data = request.form.get(f"fase_data_{i}", "").strip() or data_globale
        op_sigla = request.form.get(f"fase_op_{i}", "").strip()
        op = (anagrafica_operatori.operatore(op_sigla) if op_sigla
              else op_globale)
        firma_testo = None
        firma_img = None
        if op:
            percorso = anagrafica_operatori.percorso_firma(op)
            if percorso:
                firma_img = str(percorso)
            else:
                firma_testo = op.nome or op.sigla
        if data is None and firma_testo is None and firma_img is None:
            continue
        interventi.append({
            "cella_data": fase.cella_data,
            "data": data,
            "cella_firma": fase.cella_firma,
            "firma_testo": firma_testo,
            "firma_immagine": firma_img,
        })
    return interventi


def _leggi_lavoro() -> tuple[dict, str | None]:
    """Valida il form di generazione e ritorna il 'lavoro' da eseguire."""
    templates = template_disponibili()
    pn = request.form.get("pn", "")
    if pn not in templates:
        return {}, "PN non valido."

    _salva_form_in_sessione(pn)

    config = G.ConfigPN.da_file(templates[pn])

    valori = _raccogli_valori(config)
    errori = G.valida(config, valori)
    if errori:
        return {}, " · ".join(errori)

    try:
        quantita = int(request.form.get("quantita", "1"))
        if quantita < 1:
            raise ValueError
    except ValueError:
        return {}, "Numero di cicli non valido."

    pezzi_override: dict[int, str] = {}
    if request.form.get("ultimo_diverso"):
        pezzi = request.form.get("pezzi_ultimo", "").strip()
        if not pezzi:
            return {}, "Indica il numero di pezzi dell'ultimo ciclo."
        pezzi_override[quantita - 1] = pezzi

    campo_oi = config.campo_per_tipo("ordine_interno")
    ordine_interno = str(valori.get(campo_oi.token, "")) if campo_oi else ""

    interventi = _costruisci_interventi(config)

    riq_valori_extra: dict[str, str] = {}
    config_riq = config_riq_di(pn)
    if config_riq and request.form.get("con_riq"):
        for campo in config_riq.campi:
            chiave_form = f"riq_{campo.token}"
            v = request.form.get(chiave_form, "").strip()
            if v:
                riq_valori_extra[campo.token] = v

    return {
        "pn": pn,
        "quantita": quantita,
        "valori": {k: (v.testo if isinstance(v, G.NonApplicabile) else v)
                   for k, v in valori.items()},
        "na": [k for k, v in valori.items() if isinstance(v, G.NonApplicabile)],
        "pezzi_override": pezzi_override,
        "con_riq": bool(request.form.get("con_riq")),
        "riq_progressivo": request.form.get("riq_progressivo", "come_odl"),
        "ordine_interno": ordine_interno,
        "interventi_data": interventi,
        "riq_valori_extra": riq_valori_extra,
        "riq_data_esito": request.form.get("riq_data_esito", ""),
        "riq_varianza": request.form.get("riq_varianza", "0.05"),
    }, None


def _esegui_lavoro(lavoro: dict, cartella: Path) -> list[Path]:
    """Genera i file di un lavoro (ODL + eventuale RIQ) nella cartella data."""
    templates = template_disponibili()
    config = G.ConfigPN.da_file(templates[lavoro["pn"]])

    valori: dict[str, object] = dict(lavoro["valori"])
    for token in lavoro.get("na", []):
        valori[token] = G.NonApplicabile(str(lavoro["valori"][token]))

    interventi = [G.InterventoFase(**d) for d in lavoro.get("interventi_data", [])]

    prodotti = G.genera_batch(config, valori, lavoro["quantita"], cartella,
                              pezzi_override=lavoro.get("pezzi_override") or {},
                              interventi=interventi or None)

    if lavoro.get("con_riq"):
        prodotti += _genera_riq(lavoro, config, valori, cartella)
    return prodotti


def _genera_riq(lavoro: dict, config: G.ConfigPN, valori: dict,
                cartella: Path) -> list[Path]:
    """Genera i RIQ associati, ereditando i dati dall'ODL.

    I campi comuni (commessa, PN, lotti, date...) non si reinseriscono:
    si copiano dai valori dell'ODL confrontando i token.
    """
    config_riq = config_riq_di(lavoro["pn"])
    if config_riq is None:
        return []

    campo_oi = config.campo_per_tipo("ordine_interno")
    base_oi = str(lavoro.get("ordine_interno") or "").strip()
    serie_oi = OI.serie(base_oi, lavoro["quantita"]) if base_oi else []

    progressivi_riq: list[str] = []
    if lavoro.get("riq_progressivo") == "indipendente" and base_oi:
        anno = base_oi.split("-")[-1] if "-" in base_oi else ""
        numeri = db.prossimo_progressivo_riq(anno or "00", lavoro["quantita"])
        progressivi_riq = [f"{n:04d}-{anno}" if anno else str(n) for n in numeri]

    campo_oi_riq = config_riq.campo_per_tipo("ordine_interno")

    riq_extra = lavoro.get("riq_valori_extra", {})
    data_esito = lavoro.get("riq_data_esito", "") or None

    try:
        varianza = float(lavoro.get("riq_varianza", "0.05"))
    except (ValueError, TypeError):
        varianza = 0.05
    righe_misura = RIQ.leggi_righe_misura(config_riq)
    if righe_misura:
        RIQ.genera_valori_calibro(righe_misura, varianza)

    prodotti: list[Path] = []
    for i in range(lavoro["quantita"]):
        valori_riq: dict[str, object] = {}
        for campo in config_riq.campi:
            if campo.token in riq_extra:
                valori_riq[campo.token] = riq_extra[campo.token]
            elif campo.token in valori:
                valori_riq[campo.token] = valori[campo.token]
            elif campo.token == "{{PN}}":
                valori_riq[campo.token] = lavoro["pn"]
            elif campo.token == "{{DESCRIZIONE}}":
                valori_riq[campo.token] = config.descrizione
            elif campo.token == "{{QUANTITA_PRODOTTE}}":
                valori_riq[campo.token] = str(lavoro["quantita"])
            else:
                valori_riq[campo.token] = campo.default

        if campo_oi_riq is not None:
            if progressivi_riq:
                valori_riq[campo_oi_riq.token] = progressivi_riq[i]
            elif serie_oi:
                valori_riq[campo_oi_riq.token] = serie_oi[i]

        righe_i = [RIQ.RigaMisura(**vars(r)) for r in righe_misura]
        if righe_i:
            RIQ.genera_valori_calibro(righe_i, varianza)

        nome = RIQ.nome_file_riq(config_riq, valori_riq)
        prodotti.append(RIQ.genera_riq(
            config_riq, valori_riq, cartella / nome,
            righe_misura=righe_i or None,
            data_esito=data_esito,
        ))
    return prodotti


def _consegna(prodotti: list[Path], nome_zip: str):
    """Manda al browser un singolo xlsx o uno zip se sono piu' file."""
    if len(prodotti) == 1:
        return send_file(
            io.BytesIO(prodotti[0].read_bytes()), as_attachment=True,
            download_name=prodotti[0].name,
            mimetype="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in prodotti:
            z.write(p, arcname=p.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name=nome_zip,
                     mimetype="application/zip")


@app.route("/genera", methods=["POST"])
@serve_azione("compila")
def genera():
    lavoro, errore = _leggi_lavoro()
    if errore:
        flash(errore)
        return redirect(url_for("dashboard", pn=request.form.get("pn")))

    with tempfile.TemporaryDirectory() as tmp:
        try:
            prodotti = _esegui_lavoro(lavoro, Path(tmp))
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            db.registra_errore(session.get("username", ""), request.path,
                               str(e), traceback.format_exc())
            flash(f"Errore in generazione: {e}")
            return redirect(url_for("dashboard", pn=lavoro["pn"]))

        db.registra_generazione(
            session.get("utente_id"), session.get("username", ""),
            lavoro["pn"], "odl+riq" if lavoro.get("con_riq") else "odl",
            lavoro.get("ordine_interno", ""), lavoro["quantita"],
            [p.name for p in prodotti])
        dati_form = {k: v for k, v in request.form.items() if k not in ("pn", "fase")}
        db.salva_ultima_compilazione(lavoro["pn"], dati_form)
        _pulisci_form_sessione(lavoro["pn"])
        return _consegna(prodotti, f"{lavoro['pn']}_documenti.zip")


# --------------------------------------------------------------------------- #
# Coda multi-PN
# --------------------------------------------------------------------------- #
@app.route("/coda/aggiungi", methods=["POST"])
@serve_azione("compila")
def coda_aggiungi():
    lavoro, errore = _leggi_lavoro()
    if errore:
        flash(errore)
        return redirect(url_for("dashboard", pn=request.form.get("pn")))
    coda = session.get("coda", [])
    coda.append(lavoro)
    session["coda"] = coda
    flash(f"Aggiunto alla coda: {lavoro['pn']} ×{lavoro['quantita']}.")
    return redirect(url_for("dashboard", pn=lavoro["pn"]))


@app.route("/coda/rimuovi/<int:indice>", methods=["POST"])
@serve_azione("compila")
def coda_rimuovi(indice: int):
    coda = session.get("coda", [])
    if 0 <= indice < len(coda):
        coda.pop(indice)
        session["coda"] = coda
    return redirect(url_for("dashboard"))


@app.route("/coda/svuota", methods=["POST"])
@serve_azione("compila")
def coda_svuota():
    session["coda"] = []
    return redirect(url_for("dashboard"))


@app.route("/coda/genera", methods=["POST"])
@serve_azione("compila")
def coda_genera():
    coda = session.get("coda", [])
    if not coda:
        flash("La coda è vuota.")
        return redirect(url_for("dashboard"))

    with tempfile.TemporaryDirectory() as tmp:
        prodotti: list[Path] = []
        try:
            for lavoro in coda:
                prodotti += _esegui_lavoro(lavoro, Path(tmp))
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            db.registra_errore(session.get("username", ""), request.path,
                               str(e), traceback.format_exc())
            flash(f"Errore in generazione: {e}")
            return redirect(url_for("dashboard"))

        for lavoro in coda:
            db.registra_generazione(
                session.get("utente_id"), session.get("username", ""),
                lavoro["pn"], "odl+riq" if lavoro.get("con_riq") else "odl",
                lavoro.get("ordine_interno", ""), lavoro["quantita"], [])
        session["coda"] = []
        return _consegna(prodotti, "coda_documenti.zip")


# --------------------------------------------------------------------------- #
# Clienti
# --------------------------------------------------------------------------- #
@app.route("/clienti")
@login_richiesto
def clienti():
    return render_template("clienti.html", clienti=db.clienti())


@app.route("/clienti/nuovo", methods=["POST"])
@serve_azione("redige")
def clienti_nuovo():
    nome = request.form.get("nome", "").strip()
    if not nome:
        flash("Il nome del cliente è obbligatorio.")
    elif db.cliente_per_nome(nome):
        flash(f"Il cliente '{nome}' esiste già.")
    else:
        db.crea_cliente(nome, request.form.get("indirizzo", ""),
                        request.form.get("contatti", ""),
                        request.form.get("note", ""))
        flash(f"Cliente '{nome}' creato.")
    return redirect(url_for("clienti"))


@app.route("/clienti/<int:cliente_id>/modifica", methods=["POST"])
@serve_azione("redige")
def clienti_modifica(cliente_id: int):
    db.aggiorna_cliente(cliente_id, request.form.get("nome", ""),
                        request.form.get("indirizzo", ""),
                        request.form.get("contatti", ""),
                        request.form.get("note", ""))
    flash("Cliente aggiornato.")
    return redirect(url_for("clienti"))


@app.route("/clienti/<int:cliente_id>/elimina", methods=["POST"])
@serve_azione("redige")
def clienti_elimina(cliente_id: int):
    db.elimina_cliente(cliente_id)
    flash("Cliente eliminato. I suoi PN restano, senza cliente assegnato.")
    return redirect(url_for("clienti"))


# --------------------------------------------------------------------------- #
# Anagrafica PN
# --------------------------------------------------------------------------- #
@app.route("/pn")
@login_richiesto
def pn_elenco():
    filtro = request.args.get("cliente", "")
    if filtro == "non_assegnati":
        righe = db.pn_tutti(solo_non_assegnati=True)
    elif filtro.isdigit():
        righe = db.pn_tutti(cliente_id=int(filtro))
    else:
        righe = db.pn_tutti()

    templates = set(template_disponibili())
    return render_template("pn.html", righe=righe, clienti=db.clienti(),
                           filtro_cliente=filtro, con_template=templates,
                           n_non_assegnati=db.conta_pn_non_assegnati())


@app.route("/pn/salva", methods=["POST"])
@serve_azione("redige")
def pn_salva():
    pn = request.form.get("pn", "").strip()
    if not pn:
        flash("Il PN è obbligatorio.")
        return redirect(url_for("pn_elenco"))
    cliente_id = request.form.get("cliente_id", "")
    db.salva_pn(pn, int(cliente_id) if cliente_id.isdigit() else None,
                request.form.get("descrizione", ""),
                request.form.get("codice_disegno", ""),
                request.form.get("revisione", ""),
                request.form.get("unita_misura", ""))
    flash(f"PN '{pn}' salvato.")
    return redirect(url_for("pn_elenco"))


@app.route("/pn/assegna", methods=["POST"])
@serve_azione("redige")
def pn_assegna():
    """Assegna un cliente a piu' PN selezionati (assegnazione in blocco)."""
    cliente_id = request.form.get("cliente_id", "")
    scelti = request.form.getlist("pn_scelti")
    if not scelti:
        flash("Nessun PN selezionato.")
        return redirect(url_for("pn_elenco"))
    valore = int(cliente_id) if cliente_id.isdigit() else None
    for pn in scelti:
        db.assegna_cliente_a_pn(pn, valore)
    flash(f"{len(scelti)} PN aggiornati.")
    return redirect(url_for("pn_elenco", cliente=request.form.get("filtro", "")))


@app.route("/pn/<path:pn>/elimina", methods=["POST"])
@serve_azione("redige")
def pn_elimina(pn: str):
    db.elimina_pn(pn)
    flash(f"PN '{pn}' rimosso dall'anagrafica (il template resta sul disco).")
    return redirect(url_for("pn_elenco"))


# --------------------------------------------------------------------------- #
# Nuovo PN — wizard di riconoscimento campi da un ciclo compilato
# --------------------------------------------------------------------------- #
@app.route("/pn/nuovo", methods=["GET", "POST"])
@serve_azione("redige")
def pn_nuovo():
    if request.method == "GET":
        return render_template("nuovo.html", fase="scelta_file")

    fase = request.form.get("fase", "")
    if fase == "revisione":
        return _wizard_revisione()
    if fase == "crea":
        return _wizard_crea()
    return redirect(url_for("pn_nuovo"))


def _wizard_revisione():
    caricato = request.files.get("file")
    if caricato is None or not caricato.filename:
        flash("Scegli un ciclo compilato da analizzare.")
        return redirect(url_for("pn_nuovo"))

    CARTELLA_WIZARD.mkdir(parents=True, exist_ok=True)
    percorso = CARTELLA_WIZARD / ("sorgente" + Path(caricato.filename).suffix)
    caricato.save(percorso)

    try:
        proposta = RC.analizza(percorso)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        flash(f"Impossibile analizzare il file: {e}")
        return redirect(url_for("pn_nuovo"))

    if not proposta.pn:
        flash("PN non riconosciuto automaticamente: inseriscilo qui sotto.")
    return render_template("nuovo.html", fase="revisione",
                           percorso=str(percorso), proposta=proposta,
                           clienti=db.clienti())


def _wizard_crea():
    percorso = Path(request.form.get("percorso", ""))
    if not percorso.is_file():
        flash("File analizzato non più disponibile: ricaricalo.")
        return redirect(url_for("pn_nuovo"))

    proposta = RC.analizza(percorso)
    proposta.pn = request.form.get("pn", "").strip() or proposta.pn
    proposta.descrizione = request.form.get("descrizione", proposta.descrizione)

    for i, campo in enumerate(proposta.campi):
        campo.incluso = bool(request.form.get(f"incluso_{i}"))
        campo.fisso = bool(request.form.get(f"fisso_{i}"))
        etichetta = request.form.get(f"etichetta_{i}", "").strip()
        if etichetta:
            campo.etichetta = etichetta
        campo.obbligatorio = bool(request.form.get(f"obbligatorio_{i}"))
        campo.cella = request.form.get(f"cella_{i}", campo.cella).strip().upper()

    def _riesponi(messaggio: str):
        flash(messaggio)
        return render_template("nuovo.html", fase="revisione",
                               percorso=str(percorso), proposta=proposta,
                               clienti=db.clienti())

    if not proposta.pn:
        return _riesponi("Il PN è obbligatorio.")
    if (CARTELLA_TEMPLATES / proposta.pn / "template.xlsx").is_file():
        return _riesponi(f"Esiste già un template per il PN '{proposta.pn}'.")

    try:
        RC.costruisci_template(proposta, CARTELLA_TEMPLATES)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        db.registra_errore(session.get("username", ""), request.path,
                           str(e), traceback.format_exc())
        return _riesponi(f"Errore nella creazione del template: {e}")

    cliente_scelto = request.form.get("cliente_id", "")
    db.salva_pn(proposta.pn,
                int(cliente_scelto) if cliente_scelto.isdigit() else None,
                proposta.descrizione)

    registra_azione("crea_template_odl", proposta.pn)
    flash(f"Template creato per PN '{proposta.pn}'.")
    return redirect(url_for("pn_elenco"))


# --------------------------------------------------------------------------- #
# Nuovo RIQ — wizard di riconoscimento campi da un RIQ compilato
# --------------------------------------------------------------------------- #
@app.route("/riq/nuovo", methods=["GET", "POST"])
@serve_azione("redige")
def riq_nuovo():
    if request.method == "GET":
        return render_template("nuovo_riq.html", fase="scelta_file")

    fase = request.form.get("fase", "")
    if fase == "revisione":
        return _wizard_riq_revisione()
    if fase == "crea":
        return _wizard_riq_crea()
    return redirect(url_for("riq_nuovo"))


def _wizard_riq_revisione():
    caricato = request.files.get("file")
    if caricato is None or not caricato.filename:
        flash("Scegli un RIQ compilato da analizzare.")
        return redirect(url_for("riq_nuovo"))

    CARTELLA_WIZARD.mkdir(parents=True, exist_ok=True)
    percorso = CARTELLA_WIZARD / ("sorgente_riq" + Path(caricato.filename).suffix)
    caricato.save(percorso)

    try:
        proposta = RCR.analizza_riq(percorso)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        flash(f"Impossibile analizzare il file: {e}")
        return redirect(url_for("riq_nuovo"))

    if not proposta.pn:
        flash("PN non riconosciuto automaticamente: inseriscilo qui sotto.")
    return render_template("nuovo_riq.html", fase="revisione",
                           percorso=str(percorso), proposta=proposta)


def _wizard_riq_crea():
    percorso = Path(request.form.get("percorso", ""))
    if not percorso.is_file():
        flash("File analizzato non più disponibile: ricaricalo.")
        return redirect(url_for("riq_nuovo"))

    proposta = RCR.analizza_riq(percorso)
    proposta.pn = request.form.get("pn", "").strip() or proposta.pn
    proposta.descrizione = request.form.get("descrizione", proposta.descrizione)

    for i, campo in enumerate(proposta.campi):
        campo.incluso = bool(request.form.get(f"incluso_{i}"))
        campo.fisso = bool(request.form.get(f"fisso_{i}"))
        etichetta = request.form.get(f"etichetta_{i}", "").strip()
        if etichetta:
            campo.etichetta = etichetta
        campo.obbligatorio = bool(request.form.get(f"obbligatorio_{i}"))
        campo.cella = request.form.get(f"cella_{i}", campo.cella).strip().upper()

    def _riesponi(messaggio: str):
        flash(messaggio)
        return render_template("nuovo_riq.html", fase="revisione",
                               percorso=str(percorso), proposta=proposta)

    if not proposta.pn:
        return _riesponi("Il PN è obbligatorio.")
    if (CARTELLA_TEMPLATES / proposta.pn / "template_riq.xlsx").is_file():
        return _riesponi(f"Esiste già un template RIQ per il PN '{proposta.pn}'.")

    try:
        RCR.costruisci_template_riq(proposta, CARTELLA_TEMPLATES)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        db.registra_errore(session.get("username", ""), request.path,
                           str(e), traceback.format_exc())
        return _riesponi(f"Errore nella creazione del template RIQ: {e}")

    registra_azione("crea_template_riq", proposta.pn)
    flash(f"Template RIQ creato per PN '{proposta.pn}'.")
    return redirect(url_for("pn_elenco"))


# --------------------------------------------------------------------------- #
# Backup database
# --------------------------------------------------------------------------- #
@app.route("/backup")
@serve_azione("redige")
def backup_database():
    nome = f"backup_generatore_cicli_{datetime.now():%Y%m%d_%H%M%S}.db"
    percorso_tmp = BASE / "instance" / f"_tmp_{nome}"
    percorso_tmp.parent.mkdir(parents=True, exist_ok=True)

    origine = sqlite3.connect(db.DB_PATH)
    destinazione = sqlite3.connect(percorso_tmp)
    with destinazione:
        origine.backup(destinazione)
    origine.close()
    destinazione.close()

    @after_this_request
    def pulisci(response):
        percorso_tmp.unlink(missing_ok=True)
        return response

    return send_file(percorso_tmp, as_attachment=True, download_name=nome)


# --------------------------------------------------------------------------- #
# Configurazioni e DB (solo admin)
# --------------------------------------------------------------------------- #
@app.route("/configurazioni")
@serve_azione("amministra")
def configurazioni():
    return render_template("configurazioni.html")


@app.route("/configurazioni/utenti", methods=["GET", "POST"])
@serve_azione("amministra")
def configurazioni_utenti():
    if request.method == "POST":
        azione = request.form.get("azione", "")

        if azione == "crea":
            username = request.form.get("username", "").strip().lower()
            password = request.form.get("password", "").strip()
            nome = request.form.get("nome_visualizzato", "").strip()
            ruolo = request.form.get("ruolo", "")
            if not username or not password or ruolo not in db.RUOLI_VALIDI:
                flash("Utente, password e ruolo sono obbligatori.")
            elif db.trova_utente(username):
                flash(f"Esiste già un utente '{username}'.")
            else:
                db.crea_utente(username, generate_password_hash(password), nome, ruolo)
                registra_azione("crea_utente", username)
                flash(f"Utente '{username}' creato.")

        elif azione == "reset_password":
            utente_id = int(request.form.get("utente_id", 0) or 0)
            password = request.form.get("password", "").strip()
            if not password:
                flash("Indica la nuova password.")
            else:
                db.aggiorna_password_utente(utente_id, generate_password_hash(password))
                registra_azione("reset_password", f"utente_id={utente_id}")
                flash("Password aggiornata.")

        elif azione == "elimina":
            utente_id = int(request.form.get("utente_id", 0) or 0)
            if utente_id == session.get("utente_id"):
                flash("Non puoi eliminare il tuo stesso account.")
            else:
                db.elimina_utente(utente_id)
                registra_azione("elimina_utente", f"utente_id={utente_id}")
                flash("Utente eliminato.")

        return redirect(url_for("configurazioni_utenti"))

    return render_template("configurazioni_utenti.html", utenti=db.utenti(),
                           ruoli=db.RUOLI_VALIDI, etichette_ruolo=permessi.ETICHETTE_RUOLO)


@app.route("/configurazioni/operatori", methods=["GET", "POST"])
@serve_azione("amministra")
def configurazioni_operatori():
    if request.method == "POST":
        azione = request.form.get("azione", "")

        if azione == "salva":
            sigla = request.form.get("sigla", "").strip().upper()
            if not sigla:
                flash("La sigla è obbligatoria.")
                return redirect(url_for("configurazioni_operatori"))

            esistente = anagrafica_operatori.operatore(sigla)
            op = Operatore(
                sigla=sigla,
                nome=request.form.get("nome", "").strip(),
                reparto=request.form.get("reparto", "").strip(),
                matricola=request.form.get("matricola", "").strip(),
                e_cq=bool(request.form.get("e_cq")),
                firma=esistente.firma if esistente else "",
                firma_cq=esistente.firma_cq if esistente else "",
            )
            file_firma = request.files.get("firma")
            if file_firma and file_firma.filename:
                cartella_firme = CARTELLA_ANAGRAFICA / "firme"
                cartella_firme.mkdir(parents=True, exist_ok=True)
                nome_file = f"{sigla}.png"
                file_firma.save(cartella_firme / nome_file)
                op.firma = f"firme/{nome_file}"

            anagrafica_operatori.aggiungi_o_aggiorna(op)
            anagrafica_operatori.salva()
            registra_azione("salva_operatore", sigla)
            flash(f"Operatore '{sigla}' salvato.")

        elif azione == "elimina":
            sigla = request.form.get("sigla", "")
            anagrafica_operatori.rimuovi(sigla)
            anagrafica_operatori.salva()
            registra_azione("elimina_operatore", sigla)
            flash(f"Operatore '{sigla}' rimosso.")

        return redirect(url_for("configurazioni_operatori"))

    return render_template("configurazioni_operatori.html",
                           operatori=anagrafica_operatori.operatori)


@app.route("/configurazioni/fornitori", methods=["GET", "POST"])
@serve_azione("amministra")
def configurazioni_fornitori():
    if request.method == "POST":
        azione = request.form.get("azione", "")

        if azione == "crea":
            nome = request.form.get("nome", "").strip()
            if not nome:
                flash("Il nome è obbligatorio.")
            else:
                db.crea_fornitore(nome, request.form.get("contatti", ""),
                                  request.form.get("note", ""))
                registra_azione("crea_fornitore", nome)
                flash(f"Fornitore '{nome}' creato.")

        elif azione == "modifica":
            fornitore_id = int(request.form.get("fornitore_id", 0) or 0)
            nome = request.form.get("nome", "").strip()
            db.aggiorna_fornitore(fornitore_id, nome, request.form.get("contatti", ""),
                                  request.form.get("note", ""))
            registra_azione("modifica_fornitore", nome)
            flash("Fornitore aggiornato.")

        elif azione == "elimina":
            fornitore_id = int(request.form.get("fornitore_id", 0) or 0)
            db.elimina_fornitore(fornitore_id)
            registra_azione("elimina_fornitore", f"id={fornitore_id}")
            flash("Fornitore eliminato.")

        return redirect(url_for("configurazioni_fornitori"))

    return render_template("configurazioni_fornitori.html", fornitori=db.fornitori())


@app.route("/configurazioni/log")
@serve_azione("amministra")
def configurazioni_log():
    return render_template("configurazioni_log.html",
                           operazioni=db.log_operazioni(), errori=db.log_errori())


# --------------------------------------------------------------------------- #
# Import anagrafica
# --------------------------------------------------------------------------- #
@app.route("/importa", methods=["GET", "POST"])
@serve_azione("redige")
def importa_anagrafica():
    if request.method == "GET":
        return render_template("importa.html", fase="scelta_file")

    fase = request.form.get("fase", "")

    if fase == "mappatura":
        return _importa_mappatura()
    if fase == "conferma":
        return _importa_conferma()
    return redirect(url_for("importa_anagrafica"))


def _percorso_temporaneo(nome: str) -> Path:
    cartella = BASE / "instance" / "import"
    cartella.mkdir(parents=True, exist_ok=True)
    return cartella / nome


def _importa_mappatura():
    """Passo 1→2: salva il file caricato e mostra mappatura + anteprima."""
    caricato = request.files.get("file")
    if caricato is None or not caricato.filename:
        flash("Scegli un file da importare.")
        return redirect(url_for("importa_anagrafica"))

    percorso = _percorso_temporaneo("ultimo_import" + Path(caricato.filename).suffix)
    caricato.save(percorso)

    try:
        intestazioni, righe = importa.leggi_tabella(percorso)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        flash(f"Impossibile leggere il file: {e}")
        return redirect(url_for("importa_anagrafica"))

    if not intestazioni:
        flash("Il file sembra vuoto.")
        return redirect(url_for("importa_anagrafica"))

    proposta = importa.indovina_colonne(intestazioni)
    return _mostra_anteprima(percorso, intestazioni, righe, proposta,
                             tipi=("A",), separa_rev=True)


def _importa_conferma():
    """Passo 2→3: rilegge con la mappatura scelta e (se richiesto) salva."""
    percorso = Path(request.form.get("percorso", ""))
    if not percorso.is_file():
        flash("File di import non più disponibile: ricaricalo.")
        return redirect(url_for("importa_anagrafica"))

    intestazioni, righe = importa.leggi_tabella(percorso)

    def indice(nome: str) -> int | None:
        valore = request.form.get(nome, "")
        return int(valore) if valore.lstrip("-").isdigit() and int(valore) >= 0 else None

    mappatura = {
        "col_pn": indice("col_pn"),
        "col_descrizione": indice("col_descrizione"),
        "col_cliente": indice("col_cliente"),
        "col_tipo": indice("col_tipo"),
        "col_um": indice("col_um"),
        "col_cod_alt": indice("col_cod_alt"),
    }
    tipi = tuple(request.form.getlist("tipi")) or ("A",)
    separa_rev = bool(request.form.get("separa_revisione"))

    if mappatura["col_pn"] is None:
        flash("Devi indicare quale colonna contiene il PN.")
        return _mostra_anteprima(percorso, intestazioni, righe, mappatura,
                                 tipi, separa_rev)

    if not request.form.get("conferma"):
        return _mostra_anteprima(percorso, intestazioni, righe, mappatura,
                                 tipi, separa_rev)

    anteprima = importa.costruisci_anteprima(
        intestazioni, righe, tipi_ammessi=tipi, separa_revisione=separa_rev,
        pn_esistenti={r["pn"] for r in db.pn_tutti()}, **mappatura)

    cliente_scelto = request.form.get("cliente_id", "")
    cliente_globale = int(cliente_scelto) if cliente_scelto.isdigit() else None

    salvati = 0
    for voce in anteprima.righe:
        cliente_id = cliente_globale
        if voce.cliente:
            cliente_id = db.trova_o_crea_cliente(voce.cliente)
        db.salva_pn(voce.pn, cliente_id, voce.descrizione, "",
                    voce.revisione, voce.unita_misura)
        salvati += 1

    flash(f"Import completato: {salvati} PN salvati in anagrafica "
          f"({anteprima.esclusi} righe escluse).")
    return redirect(url_for("pn_elenco"))


def _mostra_anteprima(percorso: Path, intestazioni: list[str],
                      righe: list[list[str]], mappatura: dict,
                      tipi: tuple[str, ...], separa_rev: bool):
    anteprima = importa.costruisci_anteprima(
        intestazioni, righe, tipi_ammessi=tipi, separa_revisione=separa_rev,
        pn_esistenti={r["pn"] for r in db.pn_tutti()},
        **{k: v for k, v in mappatura.items() if k != "col_pn"},
        col_pn=mappatura.get("col_pn") if mappatura.get("col_pn") is not None else 0)
    return render_template(
        "importa.html", fase="mappatura", percorso=str(percorso),
        intestazioni=intestazioni, mappatura=mappatura, tipi_scelti=tipi,
        separa_revisione=separa_rev, anteprima=anteprima,
        tipi_articolo=importa.TIPI_ARTICOLO, clienti=db.clienti(),
        totale_righe=len(righe))


# --------------------------------------------------------------------------- #
# Storico
# --------------------------------------------------------------------------- #
@app.route("/storico")
@login_richiesto
def storico():
    return render_template("storico.html", righe=db.storico())


if __name__ == "__main__":
    db.inizializza()
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug, host="0.0.0.0", port=port)

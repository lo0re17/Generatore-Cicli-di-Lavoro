"""Generatore Cicli di Lavoro — interfaccia operatore (desktop).

App desktop per CQ / resp. produzione:
  1) scegli il PN e compila i campi variabili
  2) indica quanti cicli (con eventuale ultimo ciclo a pezzi diversi)
  3) (opzionale) compila date di esecuzione e firme per fase
     — per digitalizzare processi gia' eseguiti
  4) genera i file .xlsx pronti da stampare, anche in coda multi-PN

Nessuna dipendenza oltre openpyxl (tkinter e' incluso in Python).
"""

from __future__ import annotations

import os
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import del motore sia da sorgente sia da eseguibile PyInstaller
if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import openpyxl  # noqa: E402

from motore import generatore as G  # noqa: E402
from motore import fasi as F  # noqa: E402
from motore import rileva_campi as R  # noqa: E402
from motore import ordine_interno as OI  # noqa: E402
from motore import riq as RIQ  # noqa: E402
from motore import rileva_riq as RR  # noqa: E402
from motore.anagrafica import Anagrafica, Operatore  # noqa: E402

CARTELLA_TEMPLATES = BASE / "templates"
CARTELLA_OUTPUT_DEFAULT = BASE / "output"
CARTELLA_ANAGRAFICA = BASE / "anagrafica"

COLORE_SFONDO = "#f4f5f7"
COLORE_GRUPPO = "#2c3e50"
OPERATORE_GLOBALE = "= predefinito"


def trova_pn() -> dict[str, Path]:
    """Ritorna {PN: percorso_config.json} per ogni template disponibile."""
    risultato: dict[str, Path] = {}
    if not CARTELLA_TEMPLATES.exists():
        return risultato
    for cartella in sorted(CARTELLA_TEMPLATES.iterdir()):
        cfg = cartella / "config.json"
        if cfg.is_file():
            try:
                config = G.ConfigPN.da_file(cfg)
                risultato[config.pn] = cfg
            except Exception:
                traceback.print_exc()
    return risultato


def trova_config_riq(pn: str) -> RIQ.ConfigRIQ | None:
    """Ritorna la ConfigRIQ per il PN dato, se esiste un template RIQ."""
    cartella = CARTELLA_TEMPLATES / pn
    cfg = RIQ.percorso_config_riq(cartella)
    if not cfg.is_file():
        return None
    try:
        return RIQ.ConfigRIQ.da_file(cfg)
    except Exception:
        traceback.print_exc()
        return None


@dataclass
class VoceRIQ:
    """Dati per generare i RIQ associati 1:1 ai cicli di una VoceCoda."""
    config: RIQ.ConfigRIQ
    valori: dict[str, str]
    righe_misura: list[RIQ.RigaMisura]
    data_esito: str
    stato: str
    cq_operatore: Operatore | None
    firma_cq_modo: str
    rnc: str = ""
    concessione: str = ""


@dataclass
class VoceCoda:
    """Un lavoro accodato: PN + valori + quantita' + opzioni."""
    pn: str
    config: G.ConfigPN
    valori: dict[str, str]
    quantita: int
    pezzi_override: dict[int, str] = field(default_factory=dict)
    interventi: list[G.InterventoFase] = field(default_factory=list)
    riq: VoceRIQ | None = None

    def etichetta(self) -> str:
        oi_campo = self.config.campo_per_tipo("ordine_interno")
        primo = self.valori.get(oi_campo.token, "?") if oi_campo else "?"
        extra = f", firma/date su {len(self.interventi)} fasi" if self.interventi else ""
        riq_txt = " + RIQ" if self.riq else ""
        return f"{self.pn}  ×{self.quantita}  (dal n. {primo}{extra}){riq_txt}"


# =========================================================================== #
# Dialog gestione operatori
# =========================================================================== #
class DialogOperatori(tk.Toplevel):
    def __init__(self, master, anagrafica: Anagrafica):
        super().__init__(master)
        self.title("Operatori e firme")
        self.geometry("680x480")
        self.configure(bg=COLORE_SFONDO)
        self.anagrafica = anagrafica
        self.transient(master)
        self.grab_set()

        sx = tk.Frame(self, bg=COLORE_SFONDO)
        sx.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        ttk.Label(sx, text="Operatori registrati:", background=COLORE_SFONDO
                  ).pack(anchor="w")
        self.lista = tk.Listbox(sx, height=16)
        self.lista.pack(fill="both", expand=True, pady=4)
        self.lista.bind("<<ListboxSelect>>", self._seleziona)

        dx = tk.Frame(self, bg=COLORE_SFONDO)
        dx.pack(side="left", fill="y", padx=(0, 10), pady=10)

        self.var_sigla = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_reparto = tk.StringVar()
        self.var_firma = tk.StringVar()
        self.var_e_cq = tk.BooleanVar(value=False)
        self.var_firma_cq = tk.StringVar()

        for etich, var in (("Sigla *", self.var_sigla), ("Nome", self.var_nome),
                           ("Reparto", self.var_reparto)):
            ttk.Label(dx, text=etich, background=COLORE_SFONDO).pack(anchor="w")
            ttk.Entry(dx, textvariable=var, width=28).pack(anchor="w", pady=(0, 6))

        ttk.Label(dx, text="Firma (immagine .png)", background=COLORE_SFONDO
                  ).pack(anchor="w")
        riga_f = tk.Frame(dx, bg=COLORE_SFONDO)
        riga_f.pack(anchor="w", pady=(0, 6))
        ttk.Entry(riga_f, textvariable=self.var_firma, width=20).pack(side="left")
        ttk.Button(riga_f, text="…", width=3, command=self._scegli_firma
                   ).pack(side="left", padx=4)

        ttk.Separator(dx, orient="horizontal").pack(fill="x", pady=8)
        ttk.Checkbutton(dx, text="È anche Controllo Qualità (CQ)",
                        variable=self.var_e_cq).pack(anchor="w", pady=(0, 4))
        ttk.Label(dx, text="Firma CQ (immagine .png, opzionale)",
                  background=COLORE_SFONDO).pack(anchor="w")
        riga_cq = tk.Frame(dx, bg=COLORE_SFONDO)
        riga_cq.pack(anchor="w", pady=(0, 6))
        ttk.Entry(riga_cq, textvariable=self.var_firma_cq, width=20).pack(side="left")
        ttk.Button(riga_cq, text="…", width=3, command=self._scegli_firma_cq
                   ).pack(side="left", padx=4)
        ttk.Label(dx, text="Se vuota, per il RIQ si usa la firma standard.",
                  style="Aiuto.TLabel", background=COLORE_SFONDO
                  ).pack(anchor="w", pady=(0, 6))

        ttk.Button(dx, text="Salva operatore", command=self._salva
                   ).pack(anchor="w", pady=(10, 4))
        ttk.Button(dx, text="Elimina selezionato", command=self._elimina
                   ).pack(anchor="w")
        ttk.Button(dx, text="Chiudi", command=self.destroy
                   ).pack(anchor="w", pady=(24, 0))

        self._aggiorna_lista()

    def _aggiorna_lista(self) -> None:
        self.lista.delete(0, "end")
        for o in self.anagrafica.operatori:
            firma = " ✓firma" if self.anagrafica.percorso_firma(o) else ""
            cq = " [CQ]" if o.e_cq else ""
            self.lista.insert("end", f"{o.sigla} — {o.nome} [{o.reparto}]{firma}{cq}")

    def _seleziona(self, _event) -> None:
        sel = self.lista.curselection()
        if not sel:
            return
        o = self.anagrafica.operatori[sel[0]]
        self.var_sigla.set(o.sigla)
        self.var_nome.set(o.nome)
        self.var_reparto.set(o.reparto)
        self.var_firma.set(o.firma)
        self.var_e_cq.set(o.e_cq)
        self.var_firma_cq.set(o.firma_cq)

    def _scegli_firma_generica(self, var: tk.StringVar) -> None:
        p = filedialog.askopenfilename(
            parent=self, title="Scegli immagine firma",
            filetypes=[("Immagini PNG", "*.png"), ("Tutte le immagini", "*.png;*.jpg;*.jpeg")])
        if not p:
            return
        src = Path(p)
        dest_dir = self.anagrafica.cartella / "firme"
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / src.name
        if src.resolve() != dest.resolve():
            dest.write_bytes(src.read_bytes())
        var.set(f"firme/{src.name}")

    def _scegli_firma(self) -> None:
        self._scegli_firma_generica(self.var_firma)

    def _scegli_firma_cq(self) -> None:
        self._scegli_firma_generica(self.var_firma_cq)

    def _salva(self) -> None:
        sigla = self.var_sigla.get().strip()
        if not sigla:
            messagebox.showerror("Operatori", "La sigla è obbligatoria.", parent=self)
            return
        self.anagrafica.aggiungi_o_aggiorna(Operatore(
            sigla=sigla,
            nome=self.var_nome.get().strip(),
            reparto=self.var_reparto.get().strip(),
            firma=self.var_firma.get().strip(),
            e_cq=self.var_e_cq.get(),
            firma_cq=self.var_firma_cq.get().strip(),
        ))
        self.anagrafica.salva()
        self._aggiorna_lista()

    def _elimina(self) -> None:
        sel = self.lista.curselection()
        if not sel:
            return
        o = self.anagrafica.operatori[sel[0]]
        if messagebox.askyesno("Operatori", f"Eliminare {o.sigla}?", parent=self):
            self.anagrafica.rimuovi(o.sigla)
            self.anagrafica.salva()
            self._aggiorna_lista()


# =========================================================================== #
# Wizard nuovo PN (creatore di template)
# =========================================================================== #
class DialogNuovoPN(tk.Toplevel):
    """Crea il template di un nuovo PN partendo da un ciclo compilato.

    1. scegli il file .xlsx del ciclo esistente
    2. il riconoscimento automatico propone i campi variabili
    3. rivedi/escludi i campi, conferma PN e descrizione
    4. crea template.xlsx + config.json in templates/<PN>/
    """

    def __init__(self, master, cartella_templates: Path):
        super().__init__(master)
        self.title("Nuovo PN — crea template da un ciclo esistente")
        self.geometry("900x620")
        self.configure(bg=COLORE_SFONDO)
        self.cartella_templates = cartella_templates
        self.proposta: R.Proposta | None = None
        self.creato: Path | None = None
        self.transient(master)
        self.grab_set()

        # --- riga 1: file sorgente
        r1 = tk.Frame(self, bg=COLORE_SFONDO)
        r1.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(r1, text="Ciclo compilato di partenza:",
                  background=COLORE_SFONDO).pack(side="left")
        self.var_file = tk.StringVar()
        ttk.Entry(r1, textvariable=self.var_file).pack(
            side="left", fill="x", expand=True, padx=8)
        ttk.Button(r1, text="Sfoglia…", command=self._scegli_file).pack(side="left")
        ttk.Button(r1, text="🔍 Analizza", command=self._analizza
                   ).pack(side="left", padx=(8, 0))

        # --- riga 2: PN e descrizione
        r2 = tk.Frame(self, bg=COLORE_SFONDO)
        r2.pack(fill="x", padx=12, pady=4)
        ttk.Label(r2, text="PN:", background=COLORE_SFONDO).pack(side="left")
        self.var_pn = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_pn, width=22).pack(side="left", padx=6)
        ttk.Label(r2, text="Descrizione:", background=COLORE_SFONDO
                  ).pack(side="left", padx=(12, 0))
        self.var_descr = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_descr).pack(
            side="left", fill="x", expand=True, padx=6)

        # --- tabella campi proposti
        ttk.Label(self, text="Campi variabili proposti (doppio click per "
                             "includere/escludere):",
                  background=COLORE_SFONDO).pack(anchor="w", padx=12, pady=(8, 2))
        cont = tk.Frame(self, bg=COLORE_SFONDO)
        cont.pack(fill="both", expand=True, padx=12)
        colonne = ("incluso", "token", "etichetta", "tipo", "cella", "valore")
        self.tree = ttk.Treeview(cont, columns=colonne, show="headings",
                                 selectmode="browse")
        intesta = {"incluso": ("✓", 40), "token": ("Token", 200),
                   "etichetta": ("Etichetta", 240), "tipo": ("Tipo", 100),
                   "cella": ("Cella", 60), "valore": ("Valore attuale", 160)}
        for col, (testo, largh) in intesta.items():
            self.tree.heading(col, text=testo)
            self.tree.column(col, width=largh, anchor="w")
        scroll = ttk.Scrollbar(cont, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._toggle_incluso)

        # --- avvisi + pulsanti
        self.lbl_avvisi = ttk.Label(self, text="", style="Aiuto.TLabel",
                                    background=COLORE_SFONDO)
        self.lbl_avvisi.pack(anchor="w", padx=12, pady=(4, 0))
        r3 = tk.Frame(self, bg=COLORE_SFONDO)
        r3.pack(fill="x", padx=12, pady=10)
        self.btn_crea = ttk.Button(r3, text="✔  Crea template", state="disabled",
                                   command=self._crea)
        self.btn_crea.pack(side="left")
        ttk.Button(r3, text="Annulla", command=self.destroy).pack(side="left", padx=8)

    def _scegli_file(self) -> None:
        p = filedialog.askopenfilename(
            parent=self, title="Scegli il ciclo compilato",
            filetypes=[("Cicli Excel", "*.xlsx")])
        if p:
            self.var_file.set(p)
            self._analizza()

    def _analizza(self) -> None:
        percorso = self.var_file.get().strip()
        if not percorso or not Path(percorso).is_file():
            messagebox.showerror("Nuovo PN", "Scegli un file .xlsx valido.",
                                 parent=self)
            return
        try:
            self.proposta = R.analizza(percorso)
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Nuovo PN", f"Analisi fallita:\n{e}", parent=self)
            return
        self.var_pn.set(self.proposta.pn)
        self.var_descr.set(self.proposta.descrizione)
        self.tree.delete(*self.tree.get_children())
        for i, campo in enumerate(self.proposta.campi):
            extra = "" if campo.cella else " (in testo fase)"
            self.tree.insert("", "end", iid=str(i), values=(
                "✓" if campo.incluso else "✗",
                campo.token, campo.etichetta, campo.tipo + extra,
                campo.cella or "—", campo.valore_attuale))
        self.lbl_avvisi.config(text="; ".join(self.proposta.avvisi)
                               or f"{len(self.proposta.campi)} campi rilevati.")
        self.btn_crea.configure(state="normal")

    def _toggle_incluso(self, _event) -> None:
        sel = self.tree.selection()
        if not sel or self.proposta is None:
            return
        idx = int(sel[0])
        campo = self.proposta.campi[idx]
        campo.incluso = not campo.incluso
        vals = list(self.tree.item(sel[0], "values"))
        vals[0] = "✓" if campo.incluso else "✗"
        self.tree.item(sel[0], values=vals)

    def _crea(self) -> None:
        if self.proposta is None:
            return
        self.proposta.pn = self.var_pn.get().strip()
        self.proposta.descrizione = self.var_descr.get().strip()
        if not self.proposta.pn:
            messagebox.showerror("Nuovo PN", "Il PN è obbligatorio.", parent=self)
            return
        dest = self.cartella_templates / self.proposta.pn
        if dest.exists():
            if not messagebox.askyesno(
                    "Nuovo PN", f"Il template per {self.proposta.pn} esiste già.\n"
                    "Sovrascriverlo?", parent=self):
                return
        try:
            self.creato = R.costruisci_template(self.proposta,
                                                self.cartella_templates)
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Nuovo PN", f"Creazione fallita:\n{e}", parent=self)
            return
        messagebox.showinfo("Nuovo PN",
                            f"Template creato per {self.proposta.pn}.\n"
                            "Il PN è ora selezionabile nell'elenco.", parent=self)
        self.destroy()


# =========================================================================== #
# Wizard nuovo template RIQ
# =========================================================================== #
class DialogNuovoRIQ(tk.Toplevel):
    """Crea il template RIQ di un PN partendo da un RIQ compilato.

    Salva template_riq.xlsx + config_riq.json nella stessa cartella
    templates/<PN>/ usata dal template dell'ODL, cosi' i due documenti
    restano associati allo stesso PN.
    """

    def __init__(self, master, cartella_templates: Path):
        super().__init__(master)
        self.title("Nuovo template RIQ — crea da un RIQ esistente")
        self.geometry("880x600")
        self.configure(bg=COLORE_SFONDO)
        self.cartella_templates = cartella_templates
        self.proposta: RR.PropostaRIQ | None = None
        self.creato: Path | None = None
        self.transient(master)
        self.grab_set()

        r1 = tk.Frame(self, bg=COLORE_SFONDO)
        r1.pack(fill="x", padx=12, pady=(12, 4))
        ttk.Label(r1, text="RIQ compilato di partenza:",
                  background=COLORE_SFONDO).pack(side="left")
        self.var_file = tk.StringVar()
        ttk.Entry(r1, textvariable=self.var_file).pack(
            side="left", fill="x", expand=True, padx=8)
        ttk.Button(r1, text="Sfoglia…", command=self._scegli_file).pack(side="left")
        ttk.Button(r1, text="🔍 Analizza", command=self._analizza
                   ).pack(side="left", padx=(8, 0))

        r2 = tk.Frame(self, bg=COLORE_SFONDO)
        r2.pack(fill="x", padx=12, pady=4)
        ttk.Label(r2, text="PN:", background=COLORE_SFONDO).pack(side="left")
        self.var_pn = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_pn, width=22).pack(side="left", padx=6)
        ttk.Label(r2, text="Descrizione:", background=COLORE_SFONDO
                  ).pack(side="left", padx=(12, 0))
        self.var_descr = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_descr).pack(
            side="left", fill="x", expand=True, padx=6)

        ttk.Label(self, text="Campi di testata rilevati (doppio click per "
                             "includere/escludere):",
                  background=COLORE_SFONDO).pack(anchor="w", padx=12, pady=(8, 2))
        cont = tk.Frame(self, bg=COLORE_SFONDO)
        cont.pack(fill="both", expand=True, padx=12)
        colonne = ("incluso", "token", "etichetta", "tipo", "cella", "valore")
        self.tree = ttk.Treeview(cont, columns=colonne, show="headings",
                                 selectmode="browse", height=8)
        intesta = {"incluso": ("✓", 40), "token": ("Token", 200),
                   "etichetta": ("Etichetta", 220), "tipo": ("Tipo", 100),
                   "cella": ("Cella", 60), "valore": ("Valore attuale", 160)}
        for col, (testo, largh) in intesta.items():
            self.tree.heading(col, text=testo)
            self.tree.column(col, width=largh, anchor="w")
        scroll = ttk.Scrollbar(cont, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._toggle_incluso)

        self.lbl_struttura = ttk.Label(self, text="", style="Aiuto.TLabel",
                                       background=COLORE_SFONDO, justify="left")
        self.lbl_struttura.pack(anchor="w", padx=12, pady=(8, 0))
        self.lbl_avvisi = ttk.Label(self, text="", style="Aiuto.TLabel",
                                    background=COLORE_SFONDO)
        self.lbl_avvisi.pack(anchor="w", padx=12, pady=(2, 0))

        r3 = tk.Frame(self, bg=COLORE_SFONDO)
        r3.pack(fill="x", padx=12, pady=10)
        self.btn_crea = ttk.Button(r3, text="✔  Crea template RIQ", state="disabled",
                                   command=self._crea)
        self.btn_crea.pack(side="left")
        ttk.Button(r3, text="Annulla", command=self.destroy).pack(side="left", padx=8)

    def _scegli_file(self) -> None:
        p = filedialog.askopenfilename(
            parent=self, title="Scegli il RIQ compilato",
            filetypes=[("RIQ Excel", "*.xlsx")])
        if p:
            self.var_file.set(p)
            self._analizza()

    def _analizza(self) -> None:
        percorso = self.var_file.get().strip()
        if not percorso or not Path(percorso).is_file():
            messagebox.showerror("Nuovo RIQ", "Scegli un file .xlsx valido.",
                                 parent=self)
            return
        try:
            self.proposta = RR.analizza_riq(percorso)
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Nuovo RIQ", f"Analisi fallita:\n{e}", parent=self)
            return
        self.var_pn.set(self.proposta.pn)
        self.var_descr.set(self.proposta.descrizione)
        self.tree.delete(*self.tree.get_children())
        for i, campo in enumerate(self.proposta.campi):
            self.tree.insert("", "end", iid=str(i), values=(
                "✓" if campo.incluso else "✗",
                campo.token, campo.etichetta, campo.tipo,
                campo.cella or "—", campo.valore_attuale))
        n_misure = max(0, self.proposta.riga_fine_misure
                       - self.proposta.riga_inizio_misure + 1)
        self.lbl_struttura.config(text=(
            f"Tabella misure: righe {self.proposta.riga_inizio_misure}-"
            f"{self.proposta.riga_fine_misure} ({n_misure} caratteristiche)  |  "
            f"Riga esito: {self.proposta.riga_esito}  |  "
            f"Firma CQ: {self.proposta.cella_firma_cq or '—'}  |  "
            f"RNC: {self.proposta.cella_rnc or '—'}"))
        self.lbl_avvisi.config(text="; ".join(self.proposta.avvisi) or "Nessun avviso.")
        self.btn_crea.configure(state="normal")

    def _toggle_incluso(self, _event) -> None:
        sel = self.tree.selection()
        if not sel or self.proposta is None:
            return
        idx = int(sel[0])
        campo = self.proposta.campi[idx]
        campo.incluso = not campo.incluso
        vals = list(self.tree.item(sel[0], "values"))
        vals[0] = "✓" if campo.incluso else "✗"
        self.tree.item(sel[0], values=vals)

    def _crea(self) -> None:
        if self.proposta is None:
            return
        self.proposta.pn = self.var_pn.get().strip()
        self.proposta.descrizione = self.var_descr.get().strip()
        if not self.proposta.pn:
            messagebox.showerror("Nuovo RIQ", "Il PN è obbligatorio.", parent=self)
            return
        dest = self.cartella_templates / self.proposta.pn
        if RIQ.esiste_riq(dest):
            if not messagebox.askyesno(
                    "Nuovo RIQ", f"Il template RIQ per {self.proposta.pn} esiste "
                    "già.\nSovrascriverlo?", parent=self):
                return
        try:
            self.creato = RR.costruisci_template_riq(self.proposta,
                                                      self.cartella_templates)
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Nuovo RIQ", f"Creazione fallita:\n{e}", parent=self)
            return
        messagebox.showinfo("Nuovo RIQ",
                            f"Template RIQ creato per {self.proposta.pn}.",
                            parent=self)
        self.destroy()


# =========================================================================== #
# Strumento inserimento/approvazione misure (calibro)
# =========================================================================== #
class DialogMisureRIQ(tk.Toplevel):
    """Inserisce o approva i valori di misura del RIQ (es. da calibro),
    con possibilita' di generare automaticamente un valore casuale entro
    una varianza data rispetto al nominale."""

    def __init__(self, master, righe: list[RIQ.RigaMisura]):
        super().__init__(master)
        self.title("Misure RIQ — inserimento/approvazione (calibro)")
        self.geometry("820x480")
        self.configure(bg=COLORE_SFONDO)
        self.righe = [RIQ.RigaMisura(**vars(r)) for r in righe]  # copia di lavoro
        self.confermato = False
        self.transient(master)
        self.grab_set()

        intro = ("Le caratteristiche con nominale/tolleranza riconosciuti (es. "
                 "'49 (± 0,3) mm') possono essere compilate automaticamente con "
                 "un valore casuale entro la varianza scelta. I valori restano "
                 "modificabili a mano prima di approvare.")
        ttk.Label(self, text=intro, style="Aiuto.TLabel", background=COLORE_SFONDO,
                  wraplength=780, justify="left").pack(anchor="w", padx=12, pady=(10, 6))

        riga_var = tk.Frame(self, bg=COLORE_SFONDO)
        riga_var.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Label(riga_var, text="Varianza casuale (± mm):", style="Campo.TLabel"
                  ).pack(side="left")
        self.var_varianza = tk.StringVar(value="0.05")
        combo_var = ttk.Combobox(riga_var, textvariable=self.var_varianza,
                                 values=["0.01", "0.02", "0.05"], width=8,
                                 state="readonly")
        combo_var.pack(side="left", padx=6)
        ttk.Button(riga_var, text="🎲 Genera valori casuali",
                   command=self._genera_casuali).pack(side="left", padx=(12, 0))

        cont = tk.Frame(self, bg=COLORE_SFONDO)
        cont.pack(fill="both", expand=True, padx=12)
        colonne = ("pos", "caratteristica", "strumento", "risultato",
                  "accettato", "scartato")
        self.tree = ttk.Treeview(cont, columns=colonne, show="headings",
                                 selectmode="browse")
        intesta = {"pos": ("Pos.", 40), "caratteristica": ("Caratteristica", 260),
                   "strumento": ("Strumento", 140), "risultato": ("Risultato", 100),
                   "accettato": ("Accett.", 60), "scartato": ("Scart.", 60)}
        for col, (testo, largh) in intesta.items():
            self.tree.heading(col, text=testo)
            self.tree.column(col, width=largh, anchor="w")
        scroll = ttk.Scrollbar(cont, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._modifica_riga)

        ttk.Label(self, text="Doppio click su una riga per inserire/correggere "
                             "il risultato a mano.",
                  style="Aiuto.TLabel", background=COLORE_SFONDO
                  ).pack(anchor="w", padx=12, pady=(4, 0))

        r_btn = tk.Frame(self, bg=COLORE_SFONDO)
        r_btn.pack(fill="x", padx=12, pady=10)
        ttk.Button(r_btn, text="✔  Approva e chiudi", style="Genera.TButton",
                   command=self._approva).pack(side="left")
        ttk.Button(r_btn, text="Annulla", command=self.destroy).pack(side="left", padx=8)

        self._popola()

    def _popola(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.righe):
            self.tree.insert("", "end", iid=str(i), values=(
                r.posizione, r.caratteristica, r.strumento,
                r.risultato_rilevato, r.accettato, r.scartato))

    def _genera_casuali(self) -> None:
        try:
            varianza = float(self.var_varianza.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Misure RIQ", "Varianza non valida.", parent=self)
            return
        RIQ.genera_valori_calibro(self.righe, varianza)
        self._popola()

    def _modifica_riga(self, _event) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        r = self.righe[idx]

        dlg = tk.Toplevel(self)
        dlg.title(f"Caratteristica {r.posizione}")
        dlg.configure(bg=COLORE_SFONDO)
        dlg.transient(self)
        dlg.grab_set()
        ttk.Label(dlg, text=r.caratteristica, style="Campo.TLabel",
                  background=COLORE_SFONDO, wraplength=320
                  ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 6))

        ttk.Label(dlg, text="Risultato rilevato:", background=COLORE_SFONDO
                  ).grid(row=1, column=0, sticky="w", padx=10)
        var_ris = tk.StringVar(value=r.risultato_rilevato)
        ttk.Entry(dlg, textvariable=var_ris, width=18).grid(
            row=1, column=1, sticky="w", padx=10, pady=4)

        var_esito = tk.StringVar(value="accettato" if r.accettato else
                                 ("scartato" if r.scartato else ""))
        ttk.Label(dlg, text="Esito:", background=COLORE_SFONDO
                  ).grid(row=2, column=0, sticky="w", padx=10)
        rf = tk.Frame(dlg, bg=COLORE_SFONDO)
        rf.grid(row=2, column=1, sticky="w", padx=10, pady=4)
        ttk.Radiobutton(rf, text="Accettato (V)", value="accettato",
                        variable=var_esito).pack(anchor="w")
        ttk.Radiobutton(rf, text="Scartato (X)", value="scartato",
                        variable=var_esito).pack(anchor="w")

        def conferma() -> None:
            r.risultato_rilevato = var_ris.get().strip()
            r.accettato = "V" if var_esito.get() == "accettato" else ""
            r.scartato = "X" if var_esito.get() == "scartato" else ""
            self._popola()
            dlg.destroy()

        rbtn = tk.Frame(dlg, bg=COLORE_SFONDO)
        rbtn.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(rbtn, text="OK", command=conferma).pack(side="left", padx=6)
        ttk.Button(rbtn, text="Annulla", command=dlg.destroy).pack(side="left")

    def _approva(self) -> None:
        self.confermato = True
        self.destroy()


# =========================================================================== #
# Applicazione principale
# =========================================================================== #
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Generatore Cicli di Lavoro — F&N compositi")
        self.geometry("860x900")
        self.minsize(760, 700)
        self.configure(bg=COLORE_SFONDO)

        self.pn_config: dict[str, Path] = trova_pn()
        self.config_corrente: G.ConfigPN | None = None
        self.campo_vars: dict[str, tk.StringVar] = {}
        self.cartella_output = tk.StringVar(value=str(CARTELLA_OUTPUT_DEFAULT))
        self.anagrafica = Anagrafica(CARTELLA_ANAGRAFICA)

        self.fasi_correnti: list[F.Fase] = []
        self.fase_vars: list[tk.BooleanVar] = []
        self.coda: list[VoceCoda] = []

        # --- stato RIQ
        self.config_riq_corrente: RIQ.ConfigRIQ | None = None
        self.riq_campo_vars: dict[str, tk.StringVar] = {}
        self.righe_misura_riq: list[RIQ.RigaMisura] = []

        self._stili()
        self._costruisci_intestazione()
        self._costruisci_notebook()
        self._costruisci_log()

        if self.pn_config:
            self.pn_var.set(next(iter(self.pn_config)))
            self._cambia_pn()
        else:
            self._log("Nessun template trovato in 'templates/'. "
                      "Aggiungi un PN per iniziare.")

    # ------------------------------------------------------------------ stili
    def _stili(self) -> None:
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("Gruppo.TLabel", font=("Segoe UI", 10, "bold"),
                    foreground=COLORE_GRUPPO, background=COLORE_SFONDO)
        s.configure("Campo.TLabel", background=COLORE_SFONDO)
        s.configure("Aiuto.TLabel", foreground="#7a7a7a",
                    background=COLORE_SFONDO, font=("Segoe UI", 8))
        s.configure("Titolo.TLabel", font=("Segoe UI", 14, "bold"),
                    background=COLORE_SFONDO, foreground=COLORE_GRUPPO)
        s.configure("Genera.TButton", font=("Segoe UI", 11, "bold"))
        s.configure("TNotebook.Tab", font=("Segoe UI", 10))
        s.configure("TCheckbutton", background=COLORE_SFONDO)
        s.configure("TRadiobutton", background=COLORE_SFONDO)

    # ---------------------------------------------------------- intestazione
    def _costruisci_intestazione(self) -> None:
        top = tk.Frame(self, bg=COLORE_SFONDO)
        top.pack(fill="x", padx=16, pady=(12, 4))

        riga_titolo = tk.Frame(top, bg=COLORE_SFONDO)
        riga_titolo.pack(fill="x")
        ttk.Label(riga_titolo, text="Generatore Cicli di Lavoro",
                  style="Titolo.TLabel").pack(side="left")
        ttk.Button(riga_titolo, text="👥 Operatori…",
                   command=self._apri_operatori).pack(side="right")
        ttk.Button(riga_titolo, text="📋 Nuovo template RIQ…",
                   command=self._apri_nuovo_riq).pack(side="right", padx=8)
        ttk.Button(riga_titolo, text="➕ Nuovo PN…",
                   command=self._apri_nuovo_pn).pack(side="right", padx=8)

        riga = tk.Frame(top, bg=COLORE_SFONDO)
        riga.pack(fill="x", pady=(8, 0))
        ttk.Label(riga, text="PN (part number):",
                  style="Campo.TLabel").pack(side="left")
        self.pn_var = tk.StringVar()
        self.combo_pn = ttk.Combobox(riga, textvariable=self.pn_var,
                                     values=list(self.pn_config),
                                     state="readonly", width=28)
        self.combo_pn.pack(side="left", padx=8)
        self.combo_pn.bind("<<ComboboxSelected>>", lambda e: self._cambia_pn())
        self.lbl_descrizione = ttk.Label(riga, text="", style="Aiuto.TLabel")
        self.lbl_descrizione.pack(side="left", padx=8)

    # -------------------------------------------------------------- notebook
    def _costruisci_notebook(self) -> None:
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=6)

        self.tab_campi = tk.Frame(self.nb, bg=COLORE_SFONDO)
        self.tab_produzione = tk.Frame(self.nb, bg=COLORE_SFONDO)
        self.tab_riq = tk.Frame(self.nb, bg=COLORE_SFONDO)
        self.tab_genera = tk.Frame(self.nb, bg=COLORE_SFONDO)
        self.nb.add(self.tab_campi, text="  1. Campi del ciclo  ")
        self.nb.add(self.tab_produzione, text="  2. Date e firme (opz.)  ")
        self.nb.add(self.tab_riq, text="  3. RIQ (opz.)  ")
        self.nb.add(self.tab_genera, text="  4. Genera  ")

        self._costruisci_tab_campi()
        self._costruisci_tab_produzione()
        self._costruisci_tab_riq()
        self._costruisci_tab_genera()

    # ------------------------------------------------------- tab 1: campi
    def _costruisci_tab_campi(self) -> None:
        self.canvas = tk.Canvas(self.tab_campi, bg=COLORE_SFONDO,
                                highlightthickness=0)
        scroll = ttk.Scrollbar(self.tab_campi, orient="vertical",
                               command=self.canvas.yview)
        self.form_frame = tk.Frame(self.canvas, bg=COLORE_SFONDO)
        self.form_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.win_form = self.canvas.create_window((0, 0), window=self.form_frame,
                                                  anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.win_form, width=e.width),
        )
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event) -> None:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _cambia_pn(self) -> None:
        pn = self.pn_var.get()
        cfg = self.pn_config.get(pn)
        if not cfg:
            return
        self.config_corrente = G.ConfigPN.da_file(cfg)
        self.lbl_descrizione.config(text=self.config_corrente.descrizione)
        self._popola_form()
        self._popola_fasi()
        self._popola_form_riq()

    def _popola_form(self) -> None:
        for w in self.form_frame.winfo_children():
            w.destroy()
        self.campo_vars.clear()

        assert self.config_corrente is not None
        gruppo_corrente = None
        r = 0
        self.form_frame.columnconfigure(1, weight=1)
        for campo in self.config_corrente.campi:
            gruppo = campo.gruppo or "Altri campi"
            if gruppo != gruppo_corrente:
                gruppo_corrente = gruppo
                ttk.Label(self.form_frame, text=gruppo, style="Gruppo.TLabel"
                          ).grid(row=r, column=0, columnspan=2, sticky="w",
                                 padx=8, pady=(12, 2))
                r += 1

            testo = campo.etichetta + (" *" if campo.obbligatorio else "")
            ttk.Label(self.form_frame, text=testo, style="Campo.TLabel"
                      ).grid(row=r, column=0, sticky="w", padx=(20, 8), pady=2)
            var = tk.StringVar(value=campo.default)
            ent = ttk.Entry(self.form_frame, textvariable=var)
            ent.grid(row=r, column=1, sticky="ew", padx=(0, 12), pady=2)
            self.campo_vars[campo.token] = var
            if campo.aiuto:
                r += 1
                ttk.Label(self.form_frame, text=campo.aiuto, style="Aiuto.TLabel"
                          ).grid(row=r, column=1, sticky="w", padx=(0, 12))
            r += 1

    # -------------------------------------------------- tab 2: date e firme
    def _costruisci_tab_produzione(self) -> None:
        f = self.tab_produzione
        intro = ("Per digitalizzare un processo già eseguito: compila data di "
                 "esecuzione e firma dell'operatore nelle fasi selezionate.\n"
                 "Lascia disattivato per generare cicli da compilare a mano.")
        ttk.Label(f, text=intro, style="Aiuto.TLabel", background=COLORE_SFONDO
                  ).pack(anchor="w", padx=12, pady=(10, 4))

        self.compila_fasi = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Compila date e firme nelle fasi selezionate",
                        variable=self.compila_fasi).pack(anchor="w", padx=12, pady=4)

        riga1 = tk.Frame(f, bg=COLORE_SFONDO)
        riga1.pack(fill="x", padx=12, pady=4)
        ttk.Label(riga1, text="Data esecuzione:", style="Campo.TLabel").pack(side="left")
        self.var_data_esecuzione = tk.StringVar()
        ttk.Entry(riga1, textvariable=self.var_data_esecuzione, width=14
                  ).pack(side="left", padx=6)
        ttk.Label(riga1, text="gg/mm/aaaa", style="Aiuto.TLabel").pack(side="left")

        ttk.Label(riga1, text="   Operatore predefinito:", style="Campo.TLabel"
                  ).pack(side="left")
        self.var_operatore = tk.StringVar()
        self.combo_operatore = ttk.Combobox(riga1, textvariable=self.var_operatore,
                                            state="readonly", width=30)
        self.combo_operatore.pack(side="left", padx=6)
        ttk.Label(riga1, text="(usato dalle fasi senza operatore specifico)",
                  style="Aiuto.TLabel").pack(side="left")
        self._aggiorna_operatori()

        riga2 = tk.Frame(f, bg=COLORE_SFONDO)
        riga2.pack(fill="x", padx=12, pady=4)
        ttk.Label(riga2, text="Firma come:", style="Campo.TLabel").pack(side="left")
        self.var_tipo_firma = tk.StringVar(value="auto")
        ttk.Radiobutton(riga2, text="automatica (immagine se c'è, altrimenti nome)",
                        value="auto",
                        variable=self.var_tipo_firma).pack(side="left", padx=6)
        ttk.Radiobutton(riga2, text="solo testo (nome/sigla)", value="testo",
                        variable=self.var_tipo_firma).pack(side="left", padx=6)
        ttk.Radiobutton(riga2, text="solo immagine (.png)", value="immagine",
                        variable=self.var_tipo_firma).pack(side="left", padx=6)

        riga3 = tk.Frame(f, bg=COLORE_SFONDO)
        riga3.pack(fill="x", padx=12, pady=(8, 2))
        ttk.Label(riga3, text="Fasi da compilare:", style="Gruppo.TLabel").pack(side="left")
        ttk.Button(riga3, text="tutte", width=6,
                   command=lambda: self._seleziona_fasi(True)).pack(side="left", padx=6)
        ttk.Button(riga3, text="nessuna", width=8,
                   command=lambda: self._seleziona_fasi(False)).pack(side="left")

        cont = tk.Frame(f, bg="white", relief="sunken", bd=1)
        cont.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.canvas_fasi = tk.Canvas(cont, bg="white", highlightthickness=0)
        scroll_f = ttk.Scrollbar(cont, orient="vertical",
                                 command=self.canvas_fasi.yview)
        self.frame_fasi = tk.Frame(self.canvas_fasi, bg="white")
        self.frame_fasi.bind(
            "<Configure>",
            lambda e: self.canvas_fasi.configure(
                scrollregion=self.canvas_fasi.bbox("all")),
        )
        self.canvas_fasi.create_window((0, 0), window=self.frame_fasi, anchor="nw")
        self.canvas_fasi.configure(yscrollcommand=scroll_f.set)
        self.canvas_fasi.pack(side="left", fill="both", expand=True)
        scroll_f.pack(side="right", fill="y")

    def _aggiorna_operatori(self) -> None:
        self.anagrafica.carica()
        etichette = [o.etichetta() for o in self.anagrafica.operatori]
        self.combo_operatore.configure(values=etichette)
        if etichette and not self.var_operatore.get():
            self.var_operatore.set(etichette[0])
        # aggiorna anche le combo per-fase (se il form fasi e' gia' costruito)
        etichette_fase = self._etichette_operatori_fase()
        for combo in getattr(self, "fase_combos", []):
            combo.configure(values=etichette_fase)
        # aggiorna la combo operatore CQ (se il tab RIQ e' gia' costruito)
        combo_cq = getattr(self, "combo_operatore_cq", None)
        if combo_cq is not None:
            etichette_cq = [o.etichetta() for o in self.anagrafica.operatori_cq()]
            combo_cq.configure(values=etichette_cq)
            if etichette_cq and not self.var_operatore_cq.get():
                self.var_operatore_cq.set(etichette_cq[0])

    def _operatore_da_etichetta(self, etich: str) -> Operatore | None:
        for o in self.anagrafica.operatori:
            if o.etichetta() == etich:
                return o
        return None

    def _operatore_selezionato(self) -> Operatore | None:
        return self._operatore_da_etichetta(self.var_operatore.get())

    def _apri_operatori(self) -> None:
        dlg = DialogOperatori(self, self.anagrafica)
        self.wait_window(dlg)
        self._aggiorna_operatori()

    def _apri_nuovo_pn(self) -> None:
        dlg = DialogNuovoPN(self, CARTELLA_TEMPLATES)
        self.wait_window(dlg)
        if dlg.creato is not None:
            self.pn_config = trova_pn()
            self.combo_pn.configure(values=list(self.pn_config))
            nuovo = dlg.creato.name
            if nuovo in self.pn_config:
                self.pn_var.set(nuovo)
                self._cambia_pn()
            self._log(f"Nuovo PN creato: {nuovo}")

    def _popola_fasi(self) -> None:
        for w in self.frame_fasi.winfo_children():
            w.destroy()
        self.fasi_correnti = []
        self.fase_vars = []
        self.fase_op_vars: list[tk.StringVar] = []
        self.fase_data_vars: list[tk.StringVar] = []
        self.fase_combos: list[ttk.Combobox] = []
        if self.config_corrente is None:
            return
        try:
            wb = openpyxl.load_workbook(self.config_corrente.percorso_template())
            self.fasi_correnti = F.rileva_fasi(wb.active)
        except Exception:
            traceback.print_exc()
            return

        # intestazione colonne
        self.frame_fasi.columnconfigure(0, weight=1)
        tk.Label(self.frame_fasi, text="Fase", bg="white",
                 font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=6, pady=(4, 2))
        tk.Label(self.frame_fasi, text="Operatore", bg="white",
                 font=("Segoe UI", 9, "bold")).grid(
            row=0, column=1, sticky="w", padx=6)
        tk.Label(self.frame_fasi, text="Data (se diversa)", bg="white",
                 font=("Segoe UI", 9, "bold")).grid(
            row=0, column=2, sticky="w", padx=6)

        etichette_op = self._etichette_operatori_fase()
        for i, fase in enumerate(self.fasi_correnti, start=1):
            var = tk.BooleanVar(value=True)
            self.fase_vars.append(var)
            tk.Checkbutton(self.frame_fasi, text=fase.etichetta(), variable=var,
                           bg="white", anchor="w", justify="left"
                           ).grid(row=i, column=0, sticky="w", padx=6, pady=1)
            var_op = tk.StringVar(value=OPERATORE_GLOBALE)
            self.fase_op_vars.append(var_op)
            combo = ttk.Combobox(self.frame_fasi, textvariable=var_op,
                                 state="readonly", width=26,
                                 values=etichette_op)
            combo.grid(row=i, column=1, sticky="w", padx=6, pady=1)
            self.fase_combos.append(combo)
            var_data = tk.StringVar()
            self.fase_data_vars.append(var_data)
            ttk.Entry(self.frame_fasi, textvariable=var_data, width=12
                      ).grid(row=i, column=2, sticky="w", padx=6, pady=1)

    def _etichette_operatori_fase(self) -> list[str]:
        return [OPERATORE_GLOBALE] + [o.etichetta() for o in self.anagrafica.operatori]

    def _seleziona_fasi(self, stato: bool) -> None:
        for v in self.fase_vars:
            v.set(stato)

    def _firma_operatore(self, op: Operatore | None) -> tuple[str | None, str | None]:
        """Ritorna (firma_testo, firma_immagine) per l'operatore dato.

        Modalita' "auto": immagine .png se l'operatore ce l'ha, altrimenti
        il nome come testo (mai entrambi).
        """
        if op is None:
            return None, None
        modo = self.var_tipo_firma.get()
        percorso = self.anagrafica.percorso_firma(op)
        if modo == "auto":
            if percorso is not None:
                return None, str(percorso)
            return op.nome or op.sigla, None
        if modo == "immagine":
            if percorso is None:
                raise ValueError(
                    f"L'operatore {op.sigla} non ha un'immagine firma: "
                    "caricala da 'Operatori…' o scegli firma automatica/testo.")
            return None, str(percorso)
        return op.nome or op.sigla, None

    def _costruisci_interventi(self) -> list[G.InterventoFase]:
        """Crea gli interventi data/firma dalle scelte correnti (se attivi).

        Ogni fase puo' avere un operatore e una data propri; in mancanza
        valgono l'operatore predefinito e la data esecuzione globale.
        """
        if not self.compila_fasi.get():
            return []
        data_globale = self.var_data_esecuzione.get().strip() or None
        op_globale = self._operatore_selezionato()

        interventi: list[G.InterventoFase] = []
        for fase, var, var_op, var_data in zip(
                self.fasi_correnti, self.fase_vars,
                self.fase_op_vars, self.fase_data_vars):
            if not var.get():
                continue
            etich_op = var_op.get()
            op = (op_globale if etich_op == OPERATORE_GLOBALE
                  else self._operatore_da_etichetta(etich_op))
            firma_testo, firma_img = self._firma_operatore(op)
            data = var_data.get().strip() or data_globale
            if data is None and firma_testo is None and firma_img is None:
                continue
            interventi.append(G.InterventoFase(
                cella_data=fase.cella_data, data=data,
                cella_firma=fase.cella_firma,
                firma_testo=firma_testo, firma_immagine=firma_img))
        return interventi

    # -------------------------------------------------------------- tab 3: RIQ
    def _costruisci_tab_riq(self) -> None:
        f = self.tab_riq
        intro = ("Genera anche il RIQ (Rapporto Ispezione Qualità) associato "
                 "1:1 a ogni ciclo di questo batch, con lo stesso Ordine Interno. "
                 "Richiede un template RIQ per questo PN (pulsante "
                 "'Nuovo template RIQ…' in alto).")
        ttk.Label(f, text=intro, style="Aiuto.TLabel", background=COLORE_SFONDO,
                  wraplength=780, justify="left").pack(anchor="w", padx=12, pady=(10, 4))

        self.genera_riq_var = tk.BooleanVar(value=False)
        self.chk_genera_riq = ttk.Checkbutton(
            f, text="Genera anche il RIQ per questo PN",
            variable=self.genera_riq_var, command=self._toggle_riq_stato)
        self.chk_genera_riq.pack(anchor="w", padx=12, pady=4)

        self.lbl_stato_riq_template = ttk.Label(f, text="", style="Aiuto.TLabel",
                                                background=COLORE_SFONDO)
        self.lbl_stato_riq_template.pack(anchor="w", padx=12, pady=(0, 8))

        self.frame_riq_corpo = tk.Frame(f, bg=COLORE_SFONDO)
        self.frame_riq_corpo.pack(fill="both", expand=True, padx=0, pady=0)

        # --- campi di testata del RIQ (dinamico, form scrollabile)
        box_campi = tk.LabelFrame(self.frame_riq_corpo, text=" Campi RIQ ",
                                  bg=COLORE_SFONDO, font=("Segoe UI", 9, "bold"),
                                  fg=COLORE_GRUPPO)
        box_campi.pack(fill="x", padx=12, pady=(0, 8))
        self.canvas_riq = tk.Canvas(box_campi, bg=COLORE_SFONDO,
                                    highlightthickness=0, height=160)
        scroll_riq = ttk.Scrollbar(box_campi, orient="vertical",
                                   command=self.canvas_riq.yview)
        self.form_riq_frame = tk.Frame(self.canvas_riq, bg=COLORE_SFONDO)
        self.form_riq_frame.bind(
            "<Configure>",
            lambda e: self.canvas_riq.configure(scrollregion=self.canvas_riq.bbox("all")))
        self.win_form_riq = self.canvas_riq.create_window(
            (0, 0), window=self.form_riq_frame, anchor="nw")
        self.canvas_riq.bind(
            "<Configure>",
            lambda e: self.canvas_riq.itemconfig(self.win_form_riq, width=e.width))
        self.canvas_riq.configure(yscrollcommand=scroll_riq.set)
        self.canvas_riq.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        scroll_riq.pack(side="right", fill="y")

        # --- esito: data, stato, CQ
        box_esito = tk.LabelFrame(self.frame_riq_corpo, text=" Esito e firma CQ ",
                                  bg=COLORE_SFONDO, font=("Segoe UI", 9, "bold"),
                                  fg=COLORE_GRUPPO)
        box_esito.pack(fill="x", padx=12, pady=(0, 8))

        r1 = tk.Frame(box_esito, bg=COLORE_SFONDO)
        r1.pack(fill="x", padx=10, pady=6)
        ttk.Label(r1, text="Data RIQ:", style="Campo.TLabel").pack(side="left")
        self.var_data_riq = tk.StringVar()
        ttk.Entry(r1, textvariable=self.var_data_riq, width=14
                  ).pack(side="left", padx=6)
        self.lbl_data_riq_fonte = ttk.Label(r1, text="", style="Aiuto.TLabel")
        self.lbl_data_riq_fonte.pack(side="left", padx=6)

        ttk.Label(r1, text="   Stato:", style="Campo.TLabel").pack(side="left")
        self.var_stato_riq = tk.StringVar(value=RIQ.CONFORME)
        ttk.Radiobutton(r1, text="Conforme", value=RIQ.CONFORME,
                        variable=self.var_stato_riq,
                        command=self._toggle_non_conformita).pack(side="left", padx=6)
        ttk.Radiobutton(r1, text="Non conforme", value=RIQ.NON_CONFORME,
                        variable=self.var_stato_riq,
                        command=self._toggle_non_conformita).pack(side="left", padx=6)

        self.frame_non_conforme = tk.Frame(box_esito, bg=COLORE_SFONDO)
        self.frame_non_conforme.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(self.frame_non_conforme, text="N. RNC:", style="Campo.TLabel"
                  ).pack(side="left")
        self.var_rnc = tk.StringVar()
        ttk.Entry(self.frame_non_conforme, textvariable=self.var_rnc, width=14
                  ).pack(side="left", padx=6)
        ttk.Label(self.frame_non_conforme, text="N. Richiesta Concessione:",
                  style="Campo.TLabel").pack(side="left", padx=(12, 0))
        self.var_concessione = tk.StringVar()
        ttk.Entry(self.frame_non_conforme, textvariable=self.var_concessione, width=14
                  ).pack(side="left", padx=6)
        self.frame_non_conforme.pack_forget()  # visibile solo se non conforme

        r2 = tk.Frame(box_esito, bg=COLORE_SFONDO)
        r2.pack(fill="x", padx=10, pady=6)
        ttk.Label(r2, text="Operatore CQ:", style="Campo.TLabel").pack(side="left")
        self.var_operatore_cq = tk.StringVar()
        self.combo_operatore_cq = ttk.Combobox(r2, textvariable=self.var_operatore_cq,
                                               state="readonly", width=30)
        self.combo_operatore_cq.pack(side="left", padx=6)

        r3 = tk.Frame(box_esito, bg=COLORE_SFONDO)
        r3.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Label(r3, text="Firma CQ come:", style="Campo.TLabel").pack(side="left")
        self.var_firma_cq_modo = tk.StringVar(value=RIQ.FIRMA_CQ_PNG)
        ttk.Radiobutton(r3, text="solo immagine", value=RIQ.FIRMA_CQ_PNG,
                        variable=self.var_firma_cq_modo).pack(side="left", padx=6)
        ttk.Radiobutton(r3, text="immagine + iniziali", value=RIQ.FIRMA_CQ_PNG_INIZIALI,
                        variable=self.var_firma_cq_modo).pack(side="left", padx=6)
        ttk.Radiobutton(r3, text="solo iniziali/nome", value=RIQ.FIRMA_CQ_INIZIALI,
                        variable=self.var_firma_cq_modo).pack(side="left", padx=6)

        # --- misure
        box_misure = tk.LabelFrame(self.frame_riq_corpo, text=" Misure (calibro) ",
                                   bg=COLORE_SFONDO, font=("Segoe UI", 9, "bold"),
                                   fg=COLORE_GRUPPO)
        box_misure.pack(fill="x", padx=12, pady=(0, 10))
        r4 = tk.Frame(box_misure, bg=COLORE_SFONDO)
        r4.pack(fill="x", padx=10, pady=8)
        ttk.Button(r4, text="🔧 Inserisci/Approva misure…",
                   command=self._apri_misure_riq).pack(side="left")
        self.lbl_stato_misure = ttk.Label(r4, text="Nessuna misura inserita.",
                                          style="Aiuto.TLabel")
        self.lbl_stato_misure.pack(side="left", padx=12)

        self._aggiorna_operatori()
        self._toggle_riq_stato()

    def _toggle_riq_stato(self) -> None:
        stato = "normal" if self.genera_riq_var.get() else "disabled"
        for w in self.frame_riq_corpo.winfo_children():
            self._imposta_stato_ricorsivo(w, stato)

    def _imposta_stato_ricorsivo(self, widget, stato: str) -> None:
        try:
            widget.configure(state=stato)
        except tk.TclError:
            pass
        for figlio in widget.winfo_children():
            self._imposta_stato_ricorsivo(figlio, stato)

    def _toggle_non_conformita(self) -> None:
        if self.var_stato_riq.get() == RIQ.NON_CONFORME:
            self.frame_non_conforme.pack(fill="x", padx=10, pady=(0, 6))
        else:
            self.frame_non_conforme.pack_forget()

    def _popola_form_riq(self) -> None:
        for w in self.form_riq_frame.winfo_children():
            w.destroy()
        self.riq_campo_vars.clear()
        self.righe_misura_riq = []
        self.lbl_stato_misure.config(text="Nessuna misura inserita.")

        pn = self.config_corrente.pn if self.config_corrente else ""
        self.config_riq_corrente = trova_config_riq(pn) if pn else None

        if self.config_riq_corrente is None:
            self.lbl_stato_riq_template.config(
                text=f"Nessun template RIQ trovato per il PN '{pn}'. "
                     "Usa 'Nuovo template RIQ…' per crearlo.")
            self.chk_genera_riq.configure(state="disabled")
            self.genera_riq_var.set(False)
            self._toggle_riq_stato()
            return

        self.chk_genera_riq.configure(state="normal")
        self.lbl_stato_riq_template.config(
            text=f"Template RIQ trovato: {self.config_riq_corrente.template}")

        self.form_riq_frame.columnconfigure(1, weight=1)
        r = 0
        for campo in self.config_riq_corrente.campi:
            testo = campo.etichetta + (" *" if campo.obbligatorio else "")
            ttk.Label(self.form_riq_frame, text=testo, style="Campo.TLabel"
                      ).grid(row=r, column=0, sticky="w", padx=(8, 8), pady=2)
            valore_default = campo.default
            if campo.token == "{{PN}}":
                valore_default = pn
            elif campo.token in self.campo_vars:
                valore_default = self.campo_vars[campo.token].get()
            var = tk.StringVar(value=valore_default)
            ttk.Entry(self.form_riq_frame, textvariable=var).grid(
                row=r, column=1, sticky="ew", padx=(0, 12), pady=2)
            self.riq_campo_vars[campo.token] = var
            r += 1

        # data RIQ: sincronizzata di default con un campo data dell'ODL
        # (preferenza: token/etichetta contenente 'RIQ', altrimenti Data OI)
        fonte = None
        for tok, var in self.campo_vars.items():
            if "RIQ" in tok.upper() and var.get():
                fonte = ("campo ODL " + tok, var.get())
                break
        if fonte is None:
            var_oi = self.campo_vars.get("{{DATA_OI}}")
            if var_oi and var_oi.get():
                fonte = ("campo ODL {{DATA_OI}}", var_oi.get())
        if fonte:
            self.var_data_riq.set(fonte[1])
            self.lbl_data_riq_fonte.config(text=f"(da {fonte[0]})")
        else:
            self.var_data_riq.set("")
            self.lbl_data_riq_fonte.config(
                text="(nessun campo RIQ/Data OI trovato nell'ODL: inserisci a mano)")

        try:
            self.righe_misura_riq = RIQ.leggi_righe_misura(self.config_riq_corrente)
        except Exception:
            traceback.print_exc()
            self.righe_misura_riq = []
        self._aggiorna_stato_misure()
        self._toggle_riq_stato()

    def _aggiorna_stato_misure(self) -> None:
        n_tot = len(self.righe_misura_riq)
        n_compilate = sum(1 for r in self.righe_misura_riq if r.risultato_rilevato)
        self.lbl_stato_misure.config(
            text=f"{n_compilate}/{n_tot} misure compilate." if n_tot
                 else "Nessuna riga misura rilevata nel template.")

    def _apri_misure_riq(self) -> None:
        if not self.righe_misura_riq:
            messagebox.showinfo("Misure RIQ", "Nessuna riga misura disponibile "
                                "per il template RIQ corrente.")
            return
        dlg = DialogMisureRIQ(self, self.righe_misura_riq)
        self.wait_window(dlg)
        if dlg.confermato:
            self.righe_misura_riq = dlg.righe
            self._aggiorna_stato_misure()

    def _apri_nuovo_riq(self) -> None:
        dlg = DialogNuovoRIQ(self, CARTELLA_TEMPLATES)
        self.wait_window(dlg)
        if dlg.creato is not None:
            self._log(f"Nuovo template RIQ creato: {dlg.creato.name}")
            if self.config_corrente and self.config_corrente.pn == dlg.creato.name:
                self._popola_form_riq()

    def _operatore_cq_selezionato(self) -> Operatore | None:
        etich = self.var_operatore_cq.get()
        for o in self.anagrafica.operatori_cq():
            if o.etichetta() == etich:
                return o
        return None

    def _prepara_voce_riq(self) -> VoceRIQ | None:
        """Confeziona i dati RIQ correnti in una VoceRIQ, se richiesto."""
        if not self.genera_riq_var.get():
            return None
        if self.config_riq_corrente is None:
            raise ValueError("Nessun template RIQ disponibile per questo PN.")

        valori_riq = {tok: var.get() for tok, var in self.riq_campo_vars.items()}
        errori = RIQ.valida_riq(self.config_riq_corrente, valori_riq)
        if errori:
            raise ValueError("Dati RIQ non validi:\n" + "\n".join(errori))

        cq = self._operatore_cq_selezionato()
        if cq is None:
            raise ValueError("Seleziona un operatore CQ per il RIQ.")

        return VoceRIQ(
            config=self.config_riq_corrente,
            valori=valori_riq,
            righe_misura=[RIQ.RigaMisura(**vars(r)) for r in self.righe_misura_riq],
            data_esito=self.var_data_riq.get().strip(),
            stato=self.var_stato_riq.get(),
            cq_operatore=cq,
            firma_cq_modo=self.var_firma_cq_modo.get(),
            rnc=self.var_rnc.get().strip(),
            concessione=self.var_concessione.get().strip(),
        )

    # ---------------------------------------------------- tab 4: generazione
    def _costruisci_tab_genera(self) -> None:
        f = self.tab_genera

        box = tk.LabelFrame(f, text=" Batch ", bg=COLORE_SFONDO,
                            font=("Segoe UI", 9, "bold"), fg=COLORE_GRUPPO)
        box.pack(fill="x", padx=12, pady=(10, 4))

        r1 = tk.Frame(box, bg=COLORE_SFONDO)
        r1.pack(fill="x", padx=10, pady=6)
        ttk.Label(r1, text="Numero di cicli:", style="Campo.TLabel").pack(side="left")
        self.spin_quantita = ttk.Spinbox(r1, from_=1, to=999, width=6)
        self.spin_quantita.set(1)
        self.spin_quantita.pack(side="left", padx=8)

        self.ultimo_diverso = tk.BooleanVar(value=False)
        ttk.Checkbutton(r1, text="Ultimo ciclo con n. pezzi diverso:",
                        variable=self.ultimo_diverso,
                        command=self._toggle_ultimo).pack(side="left", padx=(20, 4))
        self.spin_ultimo = ttk.Spinbox(r1, from_=1, to=9999, width=6,
                                       state="disabled")
        self.spin_ultimo.set(1)
        self.spin_ultimo.pack(side="left")

        r2 = tk.Frame(box, bg=COLORE_SFONDO)
        r2.pack(fill="x", padx=10, pady=6)
        ttk.Label(r2, text="Cartella di destinazione:", style="Campo.TLabel"
                  ).pack(side="left")
        ttk.Entry(r2, textvariable=self.cartella_output).pack(
            side="left", fill="x", expand=True, padx=8)
        ttk.Button(r2, text="Sfoglia…", command=self._scegli_cartella
                   ).pack(side="left")

        r3 = tk.Frame(box, bg=COLORE_SFONDO)
        r3.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Button(r3, text="⚙  Genera adesso", style="Genera.TButton",
                   command=self._genera).pack(side="left")
        ttk.Button(r3, text="➕  Aggiungi alla coda",
                   command=self._aggiungi_coda).pack(side="left", padx=8)
        self.btn_apri = ttk.Button(r3, text="📂  Apri cartella",
                                   command=self._apri_cartella, state="disabled")
        self.btn_apri.pack(side="left", padx=8)

        box_coda = tk.LabelFrame(f, text=" Coda multi-PN ", bg=COLORE_SFONDO,
                                 font=("Segoe UI", 9, "bold"), fg=COLORE_GRUPPO)
        box_coda.pack(fill="both", expand=True, padx=12, pady=(6, 10))
        self.lista_coda = tk.Listbox(box_coda, height=6)
        self.lista_coda.pack(fill="both", expand=True, padx=8, pady=(6, 4))
        r4 = tk.Frame(box_coda, bg=COLORE_SFONDO)
        r4.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(r4, text="⚙  Genera tutta la coda", style="Genera.TButton",
                   command=self._genera_coda).pack(side="left")
        ttk.Button(r4, text="Rimuovi selezionato",
                   command=self._rimuovi_da_coda).pack(side="left", padx=8)
        ttk.Button(r4, text="Svuota coda",
                   command=self._svuota_coda).pack(side="left")

    def _toggle_ultimo(self) -> None:
        self.spin_ultimo.configure(
            state="normal" if self.ultimo_diverso.get() else "disabled")

    def _scegli_cartella(self) -> None:
        d = filedialog.askdirectory(initialdir=self.cartella_output.get()
                                    or str(BASE))
        if d:
            self.cartella_output.set(d)

    # -------------------------------------------------------------------- log
    def _costruisci_log(self) -> None:
        box = tk.LabelFrame(self, text=" Esito ", bg=COLORE_SFONDO,
                            font=("Segoe UI", 9, "bold"), fg=COLORE_GRUPPO)
        box.pack(fill="x", padx=16, pady=(4, 12))
        self.txt_log = tk.Text(box, height=6, wrap="word", state="disabled",
                               font=("Consolas", 9), bg="white")
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    def _log(self, testo: str) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", testo + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ----------------------------------------------------------- generazione
    def _raccogli_valori(self) -> dict[str, str]:
        return {tok: var.get() for tok, var in self.campo_vars.items()}

    def _prepara_voce(self) -> VoceCoda | None:
        """Valida l'input corrente e lo confeziona in una VoceCoda."""
        if self.config_corrente is None:
            messagebox.showwarning("Attenzione", "Nessun PN selezionato.")
            return None
        valori = self._raccogli_valori()

        errori = G.valida(self.config_corrente, valori)
        if errori:
            messagebox.showerror("Dati non validi", "\n".join(errori))
            return None

        try:
            quantita = int(self.spin_quantita.get())
            if quantita < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Dati non validi", "Numero di cicli non valido.")
            return None

        pezzi_override: dict[int, str] = {}
        if self.ultimo_diverso.get():
            pezzi_override[quantita - 1] = self.spin_ultimo.get()

        try:
            interventi = self._costruisci_interventi()
        except ValueError as e:
            messagebox.showerror("Firma operatore", str(e))
            return None

        try:
            voce_riq = self._prepara_voce_riq()
        except ValueError as e:
            messagebox.showerror("RIQ", str(e))
            return None

        return VoceCoda(pn=self.config_corrente.pn, config=self.config_corrente,
                        valori=valori, quantita=quantita,
                        pezzi_override=pezzi_override, interventi=interventi,
                        riq=voce_riq)

    def _esegui_voce(self, voce: VoceCoda, cartella: str) -> list[Path]:
        prodotti = G.genera_batch(voce.config, voce.valori, voce.quantita, cartella,
                                  pezzi_override=voce.pezzi_override,
                                  interventi=voce.interventi)
        if voce.riq is not None:
            prodotti += self._genera_riq_batch(voce, cartella)
        return prodotti

    def _genera_riq_batch(self, voce: VoceCoda, cartella: str) -> list[Path]:
        """Genera un RIQ per ciascun ciclo del batch, con lo stesso Ordine
        Interno del ciclo corrispondente (link 1:1 ODL<->RIQ)."""
        riq = voce.riq
        campo_oi = voce.config.campo_per_tipo("ordine_interno")
        base_oi = str(voce.valori.get(campo_oi.token, "")).strip()
        serie_oi = OI.serie(base_oi, voce.quantita)

        cq = riq.cq_operatore
        percorso_cq_png = (self.anagrafica.percorso_firma_cq(cq)
                           or self.anagrafica.percorso_firma(cq)) if cq else None
        nome_cq = cq.etichetta() if cq else ""

        prodotti_riq: list[Path] = []
        for i in range(voce.quantita):
            valori_riq_i = dict(riq.valori)
            oi_token = riq.config.campo_per_tipo("ordine_interno")
            if oi_token:
                valori_riq_i[oi_token.token] = serie_oi[i]

            righe_i = [RIQ.RigaMisura(**vars(r)) for r in riq.righe_misura]
            nome = RIQ.nome_file_riq(riq.config, valori_riq_i)
            out = RIQ.genera_riq(
                riq.config, valori_riq_i, Path(cartella) / nome,
                righe_misura=righe_i,
                data_esito=riq.data_esito or None,
                stato=riq.stato,
                firma_cq_nome=nome_cq,
                firma_cq_png=percorso_cq_png,
                firma_cq_modo=riq.firma_cq_modo,
                rnc=riq.rnc,
                concessione=riq.concessione,
            )
            prodotti_riq.append(out)
        return prodotti_riq

    def _genera(self) -> None:
        voce = self._prepara_voce()
        if voce is None:
            return
        cartella = self.cartella_output.get().strip() or str(CARTELLA_OUTPUT_DEFAULT)
        try:
            prodotti = self._esegui_voce(voce, cartella)
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Errore in generazione", str(e))
            self._log("ERRORE: " + str(e))
            return
        self.btn_apri.configure(state="normal")
        self._log(f"Generati {len(prodotti)} cicli in: {cartella}")
        for p in prodotti:
            self._log("   • " + p.name)
        messagebox.showinfo("Fatto",
                            f"Generati {len(prodotti)} cicli.\nCartella:\n{cartella}")

    # ------------------------------------------------------------------ coda
    def _aggiungi_coda(self) -> None:
        voce = self._prepara_voce()
        if voce is None:
            return
        self.coda.append(voce)
        self.lista_coda.insert("end", voce.etichetta())
        self._log("In coda: " + voce.etichetta())

    def _rimuovi_da_coda(self) -> None:
        sel = self.lista_coda.curselection()
        if not sel:
            return
        self.coda.pop(sel[0])
        self.lista_coda.delete(sel[0])

    def _svuota_coda(self) -> None:
        self.coda.clear()
        self.lista_coda.delete(0, "end")

    def _genera_coda(self) -> None:
        if not self.coda:
            messagebox.showinfo("Coda", "La coda è vuota.")
            return
        cartella = self.cartella_output.get().strip() or str(CARTELLA_OUTPUT_DEFAULT)
        totale = 0
        try:
            for voce in self.coda:
                prodotti = self._esegui_voce(voce, cartella)
                totale += len(prodotti)
                self._log(f"[coda] {voce.pn}: {len(prodotti)} cicli")
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("Errore in generazione", str(e))
            self._log("ERRORE: " + str(e))
            return
        self.btn_apri.configure(state="normal")
        self._svuota_coda()
        messagebox.showinfo("Fatto", f"Coda completata: {totale} cicli generati.")

    def _apri_cartella(self) -> None:
        cartella = self.cartella_output.get().strip() or str(CARTELLA_OUTPUT_DEFAULT)
        try:
            os.startfile(cartella)  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Cartella", cartella)


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()

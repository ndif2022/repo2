# -*- coding: utf-8 -*-
"""
Generatore del servizio di turnazione auto Trinitapoli -> Foggia -> Trinitapoli.

Produce un file Excel con formule automatiche: l'unico foglio da compilare a mano
e' "Assenze"; turni, sostituzioni, doppi turni e recuperi si ricalcolano da soli.

Uso:
    python3 genera_turnazione.py [--inizio AAAA-MM-GG] [--settimane N] [--output FILE.xlsx]
"""

import argparse
import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (CharacterProperties, Paragraph, ParagraphProperties,
                                   RichTextProperties)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

# --------------------------------------------------------------------------- #
# Configurazione
# --------------------------------------------------------------------------- #

PARTECIPANTI = ["Nicola", "Stefano", "Rocco", "Giuseppe", "Savino"]
GIORNI = ["Lunedi", "Martedi", "Mercoledi", "Giovedi", "Venerdi"]
GIORNI_IT = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]

TIPO_LICENZA = "Licenza - settimana intera"
TIPO_ESTERNA = "Attività esterna - settimana intera"
TIPO_ALTRO = "Altro motivo"
TIPI_ASSENZA = [TIPO_LICENZA, TIPO_ESTERNA, TIPO_ALTRO]

TRATTA = "Trinitapoli → Foggia → Trinitapoli"

MAX_ASSENZE = 200          # righe disponibili nel foglio Assenze
ASS_R0 = 5                 # prima riga dati del foglio Assenze
TUR_R0 = 6                 # prima riga dati dei fogli Turnazione / Calcoli
PART_R0 = 5                # prima riga dati del foglio Partecipanti
RIEP_R0 = 6                # prima riga dati del foglio Riepilogo

# --------------------------------------------------------------------------- #
# Stili
# --------------------------------------------------------------------------- #

BLU = "1F3864"
BLU_MED = "2F5597"
AZZURRO = "D9E2F3"
AZZURRO_CHIARO = "EDF2FA"
GRIGIO = "F2F2F2"
GIALLO = "FFF2CC"
VERDE = "E2EFDA"
ARANCIO = "FCE4D6"
ROSSO = "C00000"

F_TITOLO = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
F_SOTTO = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
F_HEAD = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_SEZ = Font(name="Calibri", size=12, bold=True, color=BLU)
F_BASE = Font(name="Calibri", size=10)
F_BOLD = Font(name="Calibri", size=10, bold=True)
F_NOTA = Font(name="Calibri", size=9, italic=True, color="595959")

FILL_TITOLO = PatternFill("solid", fgColor=BLU)
FILL_HEAD = PatternFill("solid", fgColor=BLU_MED)
FILL_BANDA = PatternFill("solid", fgColor=AZZURRO_CHIARO)
FILL_INPUT = PatternFill("solid", fgColor=GIALLO)
FILL_SEZ = PatternFill("solid", fgColor=AZZURRO)
FILL_GRIGIO = PatternFill("solid", fgColor=GRIGIO)

# Palette del grafico: slot categoriali 1-3 (blu, arancio, verde acqua).
# Verificata per daltonismo sulle coppie adiacenti di una barra impilata:
# CVD Delta E 9.2, visione normale 27.6, entrambi sopra le soglie richieste.
SERIE_1 = "2A78D6"
SERIE_2 = "EB6834"
SERIE_3 = "1BAF7A"
SERIE_4 = "EDA100"
INCHIOSTRO = "0B0B0B"
INCHIOSTRO_2 = "52514E"
GRIGLIA = "D8D8D4"

_sottile = Side(style="thin", color="BFBFBF")
BORDO = Border(left=_sottile, right=_sottile, top=_sottile, bottom=_sottile)

C = Alignment(horizontal="center", vertical="center")
L = Alignment(horizontal="left", vertical="center")
LW = Alignment(horizontal="left", vertical="top", wrap_text=True)
CW = Alignment(horizontal="center", vertical="center", wrap_text=True)

FMT_DATA = "DD/MM/YYYY"


def _testo_grafico(dimensione=900, grassetto=False, colore=INCHIOSTRO):
    """Proprieta' tipografiche per titoli, assi e legenda del grafico."""
    carattere = CharacterProperties(sz=dimensione, b=grassetto, solidFill=colore,
                                    latin=None)
    return RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=carattere),
                     endParaRPr=carattere)],
    )


def grafico_carico(ws, prima_riga, ultima_riga, ancora, settimane):
    """Barre impilate orizzontali: composizione dei turni per partecipante.

    La lunghezza della barra e' il totale delle guide, i segmenti ne mostrano
    la composizione. I numeri esatti restano nella tabella sovrastante, che fa
    da vista tabellare del grafico.
    """
    grafico = BarChart()
    grafico.type = "bar"                 # barre orizzontali: i nomi si leggono meglio
    grafico.grouping = "stacked"
    grafico.overlap = 100
    grafico.gapWidth = 60                # barre spesse, spazio contenuto
    grafico.height = 8.5
    grafico.width = 20

    dati = Reference(ws, min_col=2, max_col=4, min_row=prima_riga - 1, max_row=ultima_riga)
    categorie = Reference(ws, min_col=1, min_row=prima_riga, max_row=ultima_riga)
    grafico.add_data(dati, titles_from_data=True)
    grafico.set_categories(categorie)

    # Colore per identita' della serie, in ordine fisso; 2 px di superficie fra
    # i segmenti perche' restino distinti anche stampati in scala di grigi.
    for serie, colore in zip(grafico.series, (SERIE_1, SERIE_2, SERIE_3)):
        serie.graphicalProperties = GraphicalProperties(solidFill=colore)
        serie.graphicalProperties.line.solidFill = "FFFFFF"
        serie.graphicalProperties.line.width = 19050   # 2 px

    grafico.title = "Composizione dei turni per partecipante"
    grafico.title.tx.rich.p[0].pPr = ParagraphProperties(
        defRPr=CharacterProperties(sz=1200, b=True, solidFill=BLU))
    grafico.title.overlay = False

    # x_axis = asse delle categorie (i nomi), y_axis = asse dei valori (i turni)
    grafico.x_axis.title = None
    grafico.y_axis.title = None
    grafico.x_axis.txPr = _testo_grafico(1000, True, INCHIOSTRO)
    grafico.y_axis.txPr = _testo_grafico(900, False, INCHIOSTRO_2)
    grafico.y_axis.numFmt = "0"
    # Fondo scala fissato: con la scala automatica le barre piu' lunghe
    # rischiano di essere tagliate dal bordo dell'area di tracciamento.
    grafico.y_axis.scaling.min = 0
    grafico.y_axis.scaling.max = settimane + 5
    grafico.y_axis.majorUnit = 2
    # Nomi nello stesso ordine della tabella soprastante
    grafico.x_axis.scaling.orientation = "maxMin"
    grafico.x_axis.majorGridlines = None
    grafico.y_axis.majorGridlines.spPr = GraphicalProperties()
    grafico.y_axis.majorGridlines.spPr.line.solidFill = GRIGLIA
    grafico.y_axis.majorGridlines.spPr.line.width = 9525
    for asse in (grafico.x_axis, grafico.y_axis):
        asse.spPr = GraphicalProperties()
        asse.spPr.line.solidFill = GRIGLIA
        asse.majorTickMark = "none"
        asse.minorTickMark = "none"

    grafico.legend.position = "b"
    grafico.legend.overlay = False
    grafico.legend.txPr = _testo_grafico(900, False, INCHIOSTRO)

    grafico.graphical_properties = GraphicalProperties(solidFill="FFFFFF")
    grafico.graphical_properties.line.noFill = True

    ws.add_chart(grafico, ancora)
    return grafico


def grafico_settimane(ws, prima_riga, ultima_riga, ancora):
    """Colonne impilate: composizione dei turni settimana per settimana.

    Ogni colonna vale cinque turni; i segmenti dicono quanti sono ordinari e
    quanti derivano da una variazione. I tre colori del grafico precedente
    conservano qui lo stesso significato.
    """
    grafico = BarChart()
    grafico.type = "col"
    grafico.grouping = "stacked"
    grafico.overlap = 100
    grafico.gapWidth = 50
    grafico.height = 8.5
    grafico.width = 24

    dati = Reference(ws, min_col=3, max_col=6, min_row=prima_riga - 1, max_row=ultima_riga)
    categorie = Reference(ws, min_col=1, min_row=prima_riga, max_row=ultima_riga)
    grafico.add_data(dati, titles_from_data=True)
    grafico.set_categories(categorie)

    for serie, colore in zip(grafico.series, (SERIE_1, SERIE_2, SERIE_3, SERIE_4)):
        serie.graphicalProperties = GraphicalProperties(solidFill=colore)
        serie.graphicalProperties.line.solidFill = "FFFFFF"
        serie.graphicalProperties.line.width = 19050   # 2 px

    grafico.title = "Composizione dei turni settimana per settimana"
    grafico.title.tx.rich.p[0].pPr = ParagraphProperties(
        defRPr=CharacterProperties(sz=1200, b=True, solidFill=BLU))
    grafico.title.overlay = False

    grafico.x_axis.title = None
    grafico.y_axis.title = None
    grafico.x_axis.txPr = _testo_grafico(900, True, INCHIOSTRO)
    grafico.y_axis.txPr = _testo_grafico(900, False, INCHIOSTRO_2)
    grafico.y_axis.numFmt = "0"
    grafico.y_axis.scaling.min = 0
    grafico.y_axis.scaling.max = 5          # cinque turni a settimana, sempre
    grafico.y_axis.majorUnit = 1
    grafico.x_axis.majorGridlines = None
    grafico.y_axis.majorGridlines.spPr = GraphicalProperties()
    grafico.y_axis.majorGridlines.spPr.line.solidFill = GRIGLIA
    grafico.y_axis.majorGridlines.spPr.line.width = 9525
    for asse in (grafico.x_axis, grafico.y_axis):
        asse.spPr = GraphicalProperties()
        asse.spPr.line.solidFill = GRIGLIA
        asse.majorTickMark = "none"
        asse.minorTickMark = "none"

    grafico.legend.position = "b"
    grafico.legend.overlay = False
    grafico.legend.txPr = _testo_grafico(900, False, INCHIOSTRO)

    grafico.graphical_properties = GraphicalProperties(solidFill="FFFFFF")
    grafico.graphical_properties.line.noFill = True

    ws.add_chart(grafico, ancora)
    return grafico


def titolo(ws, testo, sottotitolo, ultima_col):
    """Intestazione grafica comune a tutti i fogli."""
    fine = get_column_letter(ultima_col)
    ws.merge_cells(f"A1:{fine}1")
    ws.merge_cells(f"A2:{fine}2")
    ws["A1"] = testo
    ws["A1"].font = F_TITOLO
    ws["A1"].alignment = L
    ws["A2"] = sottotitolo
    ws["A2"].font = F_SOTTO
    ws["A2"].alignment = L
    for col in range(1, ultima_col + 1):
        ws.cell(row=1, column=col).fill = FILL_TITOLO
        ws.cell(row=2, column=col).fill = FILL_TITOLO
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 17


def intestazioni(ws, riga, etichette, larghezze):
    for i, (testo, larg) in enumerate(zip(etichette, larghezze), start=1):
        cella = ws.cell(row=riga, column=i, value=testo)
        cella.font = F_HEAD
        cella.fill = FILL_HEAD
        cella.alignment = CW
        cella.border = BORDO
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[riga].height = 30


# --------------------------------------------------------------------------- #
# Fogli
# --------------------------------------------------------------------------- #

def foglio_regole(wb, inizio, fine, settimane):
    ws = wb.create_sheet("Regole")
    ws.sheet_properties.tabColor = BLU
    titolo(ws, "SERVIZIO DI TURNAZIONE AUTO",
           f"{TRATTA} · Regolamento e istruzioni d'uso", 3)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 86

    riga = 4

    def sezione(testo):
        nonlocal riga
        riga += 1
        ws.merge_cells(start_row=riga, start_column=1, end_row=riga, end_column=3)
        cella = ws.cell(row=riga, column=1, value=testo)
        cella.font = F_SEZ
        cella.fill = FILL_SEZ
        cella.alignment = L
        ws.row_dimensions[riga].height = 20
        riga += 1

    def voce(chiave, valore):
        nonlocal riga
        a = ws.cell(row=riga, column=2, value=chiave)
        a.font = F_BOLD
        a.alignment = LW
        a.border = BORDO
        b = ws.cell(row=riga, column=3, value=valore)
        b.font = F_BASE
        b.alignment = LW
        b.border = BORDO
        ws.row_dimensions[riga].height = max(15, 13 * (1 + len(valore) // 95))
        riga += 1

    sezione("1. DATI DEL SERVIZIO")
    voce("Tratta", TRATTA)
    voce("Frequenza", "5 giorni a settimana, dal lunedì al venerdì")
    voce("Partecipanti", ", ".join(PARTECIPANTI))
    voce("Periodo coperto", f"{inizio.strftime('%d/%m/%Y')} – {fine.strftime('%d/%m/%Y')} "
                            f"({settimane} settimane)")
    voce("Ordine base di turnazione",
         " · ".join(f"{g}: {p}" for g, p in zip(GIORNI_IT, PARTECIPANTI)))

    sezione("2. REGOLE DI GESTIONE DEI TURNI")
    voce("R1 – Licenza settimana intera",
         "Chi è in licenza per l'intera settimana NON deve recuperare il turno perso. "
         "Il turno viene coperto da un sostituto e si considera chiuso.")
    voce("R2 – Attività esterna settimana intera",
         "Chi è in attività esterna per l'intera settimana NON deve recuperare il turno perso. "
         "Vale la stessa regola della licenza.")
    voce("R3 – Assenza per altro motivo",
         "Qualunque mancata presa della macchina per motivi diversi da R1 e R2 genera un "
         "turno da recuperare nella settimana successiva.")
    voce("R4 – Scelta del sostituto",
         "Il giorno scoperto viene assegnato a chi ha accumulato finora il minor numero di "
         "doppi turni. A parità, si scorre l'ordine base a partire dal collega successivo "
         "all'assente. Così il doppio turno cade sempre su una turnazione diversa dalla prima "
         "e non ricade mai due volte di seguito sulla stessa persona.")
    voce("R5 – Modalità del recupero (solo R3)",
         "Nella settimana successiva chi ha saltato il turno guida due volte: il proprio giorno "
         "più il giorno di chi lo aveva sostituito. Il sostituto, quella settimana, è a riposo "
         "compensativo. Il bilancio dei turni torna in pari.")
    voce("R5-bis – Più recuperi verso la stessa persona",
         "I recuperi dovuti a uno stesso collega si saldano uno per settimana, nell'ordine in cui "
         "sono maturati. Chi sta già guidando un turno di recupero non viene scelto come sostituto "
         "nella stessa settimana.")
    voce("R6 – Comunicazione del giovedì",
         "Entro il giovedì di ogni settimana vanno comunicate a tutti le posizioni della "
         "settimana successiva. Usare il foglio «Comunicazione Giovedì»: contiene il prospetto "
         "e il testo già pronto da inviare.")
    voce("R7 – Slittamento automatico del recupero",
         "Se chi deve recuperare è a sua volta assente nella settimana del recupero, il turno non "
         "si perde: slitta da solo alla prima settimana utile, sempre sul giorno di chi lo aveva "
         "sostituito. Se al termine del periodo il debito è ancora aperto, resta contabilizzato "
         "nel foglio «Riepilogo» alla voce «Recuperi ancora da saldare».")

    sezione("3. COME SI USA IL FILE")
    voce("Foglio da compilare", "Solo «Assenze». Tutti gli altri fogli si aggiornano da soli.")
    voce("Come inserire un'assenza",
         "Una riga per ogni assenza: numero della settimana, nome, tipo di assenza dal menù a "
         "tendina e — solo per «Altro motivo» — la data del giorno saltato.")
    voce("Assenza di settimana intera",
         "Basta indicare settimana, nome e tipo: la data va lasciata vuota, la regola copre "
         "tutti e cinque i giorni.")
    voce("Dove si leggono i risultati",
         "«Turnazione» = calendario completo · «Comunicazione Giovedì» = prospetto da inviare · "
         "«Riepilogo» = conteggi e recuperi per persona.")
    voce("Colonne modificabili a mano",
         "Le celle su sfondo giallo. Tutte le altre contengono formule: non sovrascriverle.")

    sezione("4. LEGENDA DEI TIPI DI TURNO")
    voce("Ordinario", "Il turno si svolge secondo l'ordine base.")
    voce("Sostituzione (doppio turno)",
         "Il titolare è assente: guida un collega, che quella settimana prende la macchina due volte.")
    voce("Recupero",
         "Turno restituito da chi aveva saltato per motivo non giustificato; il collega che lo "
         "aveva sostituito è a riposo compensativo.")
    voce("DA ASSEGNARE",
         "Nessun sostituto disponibile in automatico (troppe assenze contemporanee): "
         "va deciso manualmente.")

    ws.sheet_view.showGridLines = False
    return ws


def foglio_partecipanti(wb):
    ws = wb.create_sheet("Partecipanti")
    ws.sheet_properties.tabColor = BLU_MED
    titolo(ws, "PARTECIPANTI E ORDINE BASE", TRATTA, 5)
    intestazioni(ws, 4, ["N.", "Nome", "Giorno base", "Tratta", "Note"],
                 [6, 22, 16, 34, 40])

    for i, nome in enumerate(PARTECIPANTI):
        r = PART_R0 + i
        valori = [i + 1, nome, GIORNI_IT[i], TRATTA, ""]
        for c, v in enumerate(valori, start=1):
            cella = ws.cell(row=r, column=c, value=v)
            cella.font = F_BOLD if c == 2 else F_BASE
            cella.alignment = L if c in (2, 4, 5) else C
            cella.border = BORDO
            if c == 5:
                cella.fill = FILL_INPUT
        ws.row_dimensions[r].height = 18

    r = PART_R0 + len(PARTECIPANTI) + 1
    ws.cell(row=r, column=2,
            value="Per cambiare un nominativo modificare la cella in colonna B: "
                  "l'intero calendario si aggiorna automaticamente.").font = F_NOTA
    ws.sheet_view.showGridLines = False
    return ws


def foglio_assenze(wb, inizio, settimane):
    ws = wb.create_sheet("Assenze")
    ws.sheet_properties.tabColor = "BF8F00"
    titolo(ws, "ASSENZE E IMPEDIMENTI",
           "Unico foglio da compilare a mano · le celle gialle sono editabili", 7)
    intestazioni(ws, 4,
                 ["Settimana", "Periodo della settimana", "Persona", "Tipo di assenza",
                  "Data (solo per «Altro motivo»)", "Turno da recuperare?", "Note"],
                 [11, 26, 16, 30, 24, 20, 40])

    for i in range(MAX_ASSENZE):
        r = ASS_R0 + i
        for c in range(1, 8):
            cella = ws.cell(row=r, column=c)
            cella.font = F_BASE
            cella.border = BORDO
            cella.alignment = L if c in (4, 7) else C
            if c in (1, 3, 4, 5, 7):
                cella.fill = FILL_INPUT

        # Periodo della settimana (calcolato dal numero di settimana)
        ws.cell(row=r, column=2).value = (
            f'=IF($A{r}="","",TEXT(Calcoli!$AG${TUR_R0}+($A{r}-1)*7,"dd/mm/yyyy")'
            f'&" - "&TEXT(Calcoli!$AG${TUR_R0}+($A{r}-1)*7+4,"dd/mm/yyyy"))'
        )
        # Turno da recuperare?
        ws.cell(row=r, column=6).value = (
            f'=IF($D{r}="","",IF(OR($D{r}="{TIPO_LICENZA}",$D{r}="{TIPO_ESTERNA}"),'
            f'"NO - non recuperabile","SI - settimana successiva"))'
        )
        ws.cell(row=r, column=5).number_format = FMT_DATA
        ws.row_dimensions[r].height = 16

    dv_sett = DataValidation(
        type="list",
        formula1='"' + ",".join(str(n) for n in range(1, settimane + 1)) + '"',
        allow_blank=True, showErrorMessage=True,
    )
    dv_sett.error = "Indicare un numero di settimana compreso nel periodo del servizio."
    ws.add_data_validation(dv_sett)
    dv_sett.add(f"A{ASS_R0}:A{ASS_R0 + MAX_ASSENZE - 1}")

    dv_pers = DataValidation(
        type="list",
        formula1=f"=Partecipanti!$B${PART_R0}:$B${PART_R0 + len(PARTECIPANTI) - 1}",
        allow_blank=True, showErrorMessage=True,
    )
    dv_pers.error = "Selezionare un partecipante dall'elenco."
    ws.add_data_validation(dv_pers)
    dv_pers.add(f"C{ASS_R0}:C{ASS_R0 + MAX_ASSENZE - 1}")

    dv_tipo = DataValidation(
        type="list", formula1='"' + ",".join(TIPI_ASSENZA) + '"',
        allow_blank=True, showErrorMessage=True,
    )
    dv_tipo.error = "Selezionare un tipo di assenza dall'elenco."
    ws.add_data_validation(dv_tipo)
    dv_tipo.add(f"D{ASS_R0}:D{ASS_R0 + MAX_ASSENZE - 1}")

    ws.freeze_panes = f"A{ASS_R0}"
    ws.auto_filter.ref = f"A4:G{ASS_R0 + MAX_ASSENZE - 1}"
    ws.sheet_view.showGridLines = False
    return ws


def foglio_calcoli(wb, inizio, settimane):
    """Foglio di appoggio: risolve assenze, sostituzioni e recuperi.

    I recuperi seguono una coda FIFO per creditore: chi salta un turno per un
    motivo non giustificato resta debitore verso chi lo ha sostituito finche'
    non e' in condizione di restituirlo. Se nella settimana successiva e' a sua
    volta assente, il debito slitta da solo alla prima settimana utile.
    """
    ws = wb.create_sheet("Calcoli")
    ws["A1"] = "Foglio di appoggio - non modificare"
    ws["A1"].font = F_BOLD

    etichette = [
        "Sett.", "Data", "Titolare", "Tipo assenza titolare", "Titolare assente",
        "Recuperabile",
    ] + [f"Disp. {p}" for p in PARTECIPANTI] + [
        "Debitore del recupero", "Recupero attivo",
    ] + [f"Punt. {p}" for p in PARTECIPANTI] + [
        "Sostituto", "Sostituto recuperabile", "Conducente effettivo", "Tipo turno",
        "A riposo", "Motivo variazione", "Nota sintetica",
        "Progressivo debito", "Chiave debito", "Debiti maturati", "Debiti saldati",
        "Riga del debito", "Settimana del debito",
    ]
    for i, testo in enumerate(etichette, start=1):
        cella = ws.cell(row=TUR_R0 - 1, column=i, value=testo)
        cella.font = F_HEAD
        cella.fill = FILL_HEAD
        cella.alignment = CW
        ws.column_dimensions[get_column_letter(i)].width = 18

    n_righe = settimane * 5
    prima = TUR_R0
    plist = f"Partecipanti!$B${PART_R0}:$B${PART_R0 + len(PARTECIPANTI) - 1}"
    pcell = [f"Partecipanti!$B${PART_R0 + i}" for i in range(len(PARTECIPANTI))]
    disp_col = ["G", "H", "I", "J", "K"]
    punt_col = ["N", "O", "P", "Q", "R"]

    for i in range(n_righe):
        r = TUR_R0 + i
        settimana = i // 5 + 1
        giorno = i % 5
        data = inizio + dt.timedelta(weeks=settimana - 1, days=giorno)
        sett_da = TUR_R0 + (settimana - 1) * 5   # lunedi' della settimana corrente
        sett_a = sett_da + 4                     # venerdi' della settimana corrente
        prec_a = sett_da - 1                     # ultima riga delle settimane precedenti

        ws.cell(row=r, column=1, value=settimana)
        cella_data = ws.cell(row=r, column=2, value=data)
        cella_data.number_format = FMT_DATA
        # Titolare del turno secondo l'ordine base
        ws.cell(row=r, column=3, value=f"={pcell[giorno]}")

        # D: tipo di assenza del titolare
        ws.cell(row=r, column=4, value=(
            f'=IF(COUNTIFS(A_PERS,$C{r},A_SETT,$A{r},A_TIPO,"{TIPO_LICENZA}")>0,"{TIPO_LICENZA}",'
            f'IF(COUNTIFS(A_PERS,$C{r},A_SETT,$A{r},A_TIPO,"{TIPO_ESTERNA}")>0,"{TIPO_ESTERNA}",'
            f'IF(COUNTIFS(A_PERS,$C{r},A_DATA,$B{r})>0,'
            f'INDEX(A_TIPO,MATCH(1,INDEX((A_PERS=$C{r})*(A_DATA=$B{r}),0),0)),"")))'
        ))
        # E: titolare assente (1/0)
        ws.cell(row=r, column=5, value=f'=IF($D{r}="",0,1)')
        # F: turno da recuperare (1/0). Se il giorno e' gia' un riposo compensativo
        #    l'assenza del titolare non genera alcun debito.
        ws.cell(row=r, column=6, value=(
            f'=IF($M{r}=1,0,IF($E{r}=0,0,'
            f'IF(OR($D{r}="{TIPO_LICENZA}",$D{r}="{TIPO_ESTERNA}"),0,1)))'
        ))

        # G:K disponibilita' di ciascun partecipante in quella data
        for k, col in enumerate(disp_col):
            ws.cell(row=r, column=ord(col) - 64, value=(
                f'=IF(COUNTIFS(A_PERS,{pcell[k]},A_DATA,$B{r})'
                f'+COUNTIFS(A_PERS,{pcell[k]},A_SETT,$A{r},A_TIPO,"{TIPO_LICENZA}")'
                f'+COUNTIFS(A_PERS,{pcell[k]},A_SETT,$A{r},A_TIPO,"{TIPO_ESTERNA}")>0,0,1)'
            ))

        # AB: debiti maturati verso il titolare di questa riga (settimane precedenti)
        # AC: debiti gia' saldati sul suo giorno
        # AD: posizione, nella coda FIFO, del debito piu' vecchio ancora aperto
        # AE: settimana in cui quel debito e' nato
        # L:  chi deve il recupero
        if settimana == 1:
            ws.cell(row=r, column=28, value=0)
            ws.cell(row=r, column=29, value=0)
            ws.cell(row=r, column=30, value=0)
            ws.cell(row=r, column=31, value="")
            ws.cell(row=r, column=12, value="")
        else:
            ws.cell(row=r, column=28, value=f'=COUNTIF($T${prima}:$T{prec_a},$C{r})')
            ws.cell(row=r, column=29, value=(
                f'=COUNTIFS($C${prima}:$C{prec_a},$C{r},$M${prima}:$M{prec_a},1)'
            ))
            ws.cell(row=r, column=30, value=(
                f'=IF($AB{r}>$AC{r},IFERROR(MATCH($C{r}&"#"&($AC{r}+1),'
                f'$AA${prima}:$AA{prec_a},0),0),0)'
            ))
            ws.cell(row=r, column=31, value=(
                f'=IF($AD{r}=0,"",INDEX($A${prima}:$A{prec_a},$AD{r}))'
            ))
            ws.cell(row=r, column=12, value=(
                f'=IF($AD{r}=0,"",INDEX($C${prima}:$C{prec_a},$AD{r}))'
            ))
        # M: il debitore e' disponibile, quindi il recupero si esegue oggi
        ws.cell(row=r, column=13, value=(
            f'=IF($L{r}="",0,IF(INDEX($G{r}:$K{r},MATCH($L{r},{plist},0))=1,1,0))'
        ))

        # N:R punteggio dei candidati alla sostituzione (piu' basso = scelto).
        # Esclusi gli indisponibili, il titolare stesso e chi in questa settimana
        # sta gia' guidando un turno di recupero.
        for k, col in enumerate(punt_col):
            storico = "0" if r == TUR_R0 else f'COUNTIF($S${TUR_R0}:$S{r - 1},{pcell[k]})'
            impegnato = (f'SUMPRODUCT(($M${sett_da}:$M${sett_a}=1)*'
                         f'($L${sett_da}:$L${sett_a}={pcell[k]}))')
            ws.cell(row=r, column=ord(col) - 64, value=(
                f'=IF(${disp_col[k]}{r}=0,100000,IF({pcell[k]}=$C{r},100000,'
                f'IF({impegnato}>0,100000,'
                f'{storico}*100+MOD({k + 1}-MATCH($C{r},{plist},0)-1,5))))'
            ))

        # S: sostituto scelto
        ws.cell(row=r, column=19, value=(
            f'=IF(OR($E{r}=0,$M{r}=1),"",'
            f'IF(MIN($N{r}:$R{r})>=100000,"DA ASSEGNARE",'
            f'INDEX({plist},MATCH(MIN($N{r}:$R{r}),$N{r}:$R{r},0))))'
        ))
        # T: sostituto di un turno recuperabile: diventa creditore
        ws.cell(row=r, column=20, value=(
            f'=IF(AND($F{r}=1,$S{r}<>"",$S{r}<>"DA ASSEGNARE"),$S{r},"")'
        ))
        # Z, AA: posizione del credito nella coda del creditore e chiave di ricerca
        ws.cell(row=r, column=26, value=f'=IF($T{r}="","",COUNTIF($T${prima}:$T{r},$T{r}))')
        ws.cell(row=r, column=27, value=f'=IF($T{r}="","",$T{r}&"#"&$Z{r})')
        # U: conducente effettivo
        ws.cell(row=r, column=21, value=f'=IF($M{r}=1,$L{r},IF($E{r}=0,$C{r},$S{r}))')
        # V: tipo di turno
        ws.cell(row=r, column=22, value=(
            f'=IF($M{r}=1,"Recupero",IF($E{r}=0,"Ordinario",'
            f'IF($S{r}="DA ASSEGNARE","DA ASSEGNARE","Sostituzione (doppio turno)")))'
        ))
        # W: chi non guida pur essendo di turno
        ws.cell(row=r, column=23, value=f'=IF(OR($M{r}=1,$E{r}=1),$C{r},"")')
        # X: motivo della variazione
        ws.cell(row=r, column=24, value=(
            f'=IF($M{r}=1,'
            f'IF($E{r}=1,$L{r}&" copre il turno recuperando quello saltato nella settimana "'
            f'&$AE{r}&"; "&$C{r}&" assente ("&$D{r}&")",'
            f'$L{r}&" recupera il turno saltato nella settimana "&$AE{r}'
            f'&"; "&$C{r}&" a riposo compensativo"),'
            f'IF($E{r}=1,$C{r}&" assente ("&$D{r}&")"'
            f'&IF($S{r}="DA ASSEGNARE","; nessun sostituto disponibile",'
            f'"; sostituisce "&$S{r}&" con doppio turno")'
            f'&IF($F{r}=1,"; turno da recuperare a partire dalla settimana "&($A{r}+1),'
            f'"; turno non da recuperare"),""))'
        ))
        # Y: nota sintetica, usata nel testo della comunicazione del giovedi'
        ws.cell(row=r, column=25, value=(
            f'=IF($M{r}=1,'
            f'IF($E{r}=1,"recupero del turno saltato nella settimana "&$AE{r}'
            f'&", "&$C{r}&" assente",'
            f'"recupera il turno saltato nella settimana "&$AE{r}&", "&$C{r}&" a riposo"),'
            f'IF($E{r}=1,IF($S{r}="DA ASSEGNARE","turno da assegnare: "&$C{r}&" assente",'
            f'"in sostituzione di "&$C{r}&", doppio turno"),""))'
        ))

    # AG: parametri del servizio, usati dal foglio Assenze
    ws.cell(row=TUR_R0 - 1, column=33, value="Inizio servizio").font = F_BOLD
    cella = ws.cell(row=TUR_R0, column=33, value=inizio)
    cella.number_format = FMT_DATA
    ws.cell(row=TUR_R0 + 1, column=33, value=settimane)
    ws.cell(row=TUR_R0 + 1, column=34, value="numero settimane").font = F_NOTA

    ws.sheet_state = "hidden"
    return ws


def foglio_turnazione(wb, inizio, fine, settimane):
    ws = wb.create_sheet("Turnazione")
    ws.sheet_properties.tabColor = "548235"
    titolo(ws, "CALENDARIO DELLA TURNAZIONE",
           f"{TRATTA} · {inizio.strftime('%d/%m/%Y')} – {fine.strftime('%d/%m/%Y')} · "
           f"calcolo automatico dal foglio «Assenze»", 10)
    ws["A3"] = ("Le colonne da D a I sono calcolate automaticamente. "
                "Compilare a mano solo la colonna «Note».")
    ws["A3"].font = F_NOTA
    intestazioni(ws, 5,
                 ["Sett.", "Data", "Giorno", "Turno previsto", "Stato del titolare",
                  "Conducente effettivo", "Tipo di turno", "Non guida (a riposo)",
                  "Motivo della variazione", "Note"],
                 [7, 12, 12, 16, 26, 20, 26, 20, 62, 30])

    n_righe = settimane * 5
    for i in range(n_righe):
        r = TUR_R0 + i
        settimana = i // 5 + 1
        giorno = i % 5
        banda = (settimana % 2 == 0)

        valori = {
            1: settimana,
            2: f"=Calcoli!$B{r}",
            3: GIORNI_IT[giorno],
            4: f"=Calcoli!$C{r}",
            5: f'=IF(Calcoli!$E{r}=0,"Presente",Calcoli!$D{r})',
            6: f"=Calcoli!$U{r}",
            7: f"=Calcoli!$V{r}",
            8: f"=Calcoli!$W{r}",
            9: f"=Calcoli!$X{r}",
            10: None,
        }
        for c, v in valori.items():
            cella = ws.cell(row=r, column=c, value=v)
            cella.font = F_BOLD if c == 6 else F_BASE
            cella.alignment = LW if c in (9, 10) else C
            cella.border = BORDO
            if banda:
                cella.fill = FILL_BANDA
            if c == 10:
                cella.fill = FILL_INPUT
            if c == 2:
                cella.number_format = FMT_DATA
        ws.row_dimensions[r].height = 16

        if giorno == 4 and i < n_righe - 1:
            for c in range(1, 11):
                bordo = ws.cell(row=r, column=c).border
                ws.cell(row=r, column=c).border = Border(
                    left=bordo.left, right=bordo.right, top=bordo.top,
                    bottom=Side(style="medium", color=BLU_MED),
                )

    ws.freeze_panes = f"D{TUR_R0}"
    ws.auto_filter.ref = f"A5:J{TUR_R0 + n_righe - 1}"
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "5:5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def foglio_comunicazione(wb, settimane):
    ws = wb.create_sheet("Comunicazione Giovedì")
    ws.sheet_properties.tabColor = "C55A11"
    titolo(ws, "COMUNICAZIONE DEL GIOVEDÌ",
           "Prospetto delle posizioni della settimana selezionata, da inviare ai partecipanti", 5)
    for col, larg in zip("ABCDE", [4, 26, 22, 30, 60]):
        ws.column_dimensions[col].width = larg

    ultima = TUR_R0 + settimane * 5 - 1

    ws["B4"] = "Settimana da comunicare (n.)"
    ws["B4"].font = F_BOLD
    ws["C4"] = 1
    ws["C4"].font = Font(name="Calibri", size=12, bold=True, color=BLU)
    ws["C4"].fill = FILL_INPUT
    ws["C4"].alignment = C
    ws["C4"].border = BORDO

    dv = DataValidation(
        type="list", formula1='"' + ",".join(str(n) for n in range(1, settimane + 1)) + '"',
        allow_blank=False, showErrorMessage=True,
    )
    ws.add_data_validation(dv)
    dv.add("C4")

    etichette = [
        ("B5", "Periodo della settimana", "C5",
         f'=TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},($C$4-1)*5+1),"dd/mm/yyyy")&" - "&'
         f'TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},($C$4-1)*5+5),"dd/mm/yyyy")'),
        ("B6", "Comunicazione da inviare entro", "C6",
         f'=TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},($C$4-1)*5+1)-4,"dddd dd/mm/yyyy")'),
    ]
    for cella_e, testo, cella_v, formula in etichette:
        ws[cella_e] = testo
        ws[cella_e].font = F_BOLD
        ws[cella_v] = formula
        ws[cella_v].font = F_BASE
        ws[cella_v].alignment = L

    intestazioni_riga = 8
    for i, testo in enumerate(["", "Giorno", "Data", "Conducente", "Note sul turno"], start=1):
        cella = ws.cell(row=intestazioni_riga, column=i, value=testo)
        cella.font = F_HEAD
        cella.fill = FILL_HEAD
        cella.alignment = CW
        cella.border = BORDO
    ws.row_dimensions[intestazioni_riga].height = 24

    for g in range(5):
        r = intestazioni_riga + 1 + g
        pos = f"($C$4-1)*5+{g + 1}"
        celle = {
            2: f'=INDEX(Turnazione!$C${TUR_R0}:$C${ultima},{pos})',
            3: f'=TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},{pos}),"dd/mm/yyyy")',
            4: f'=INDEX(Turnazione!$F${TUR_R0}:$F${ultima},{pos})',
            5: f'=INDEX(Calcoli!$Y${TUR_R0}:$Y${ultima},{pos})',
        }
        for c in range(1, 6):
            cella = ws.cell(row=r, column=c, value=celle.get(c))
            cella.font = F_BOLD if c == 4 else F_BASE
            cella.alignment = LW if c == 5 else C
            cella.border = BORDO
            if c == 4:
                cella.fill = PatternFill("solid", fgColor=VERDE)
        ws.row_dimensions[r].height = 22

    r = intestazioni_riga + 7
    ws.cell(row=r, column=2, value="TESTO PRONTO DA INVIARE").font = F_SEZ
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws.cell(row=r, column=2).fill = FILL_SEZ

    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 6, end_column=5)
    testo = (
        f'="Servizio auto {TRATTA}"&CHAR(10)&'
        f'"Turnazione settimana n. "&$C$4&" ("&$C$5&")"&CHAR(10)&CHAR(10)&'
    )
    for g in range(5):
        rr = intestazioni_riga + 1 + g
        testo += (f'$B{rr}&" - "&$C{rr}&": "&$D{rr}&'
                  f'IF($E{rr}="",""," ("&$E{rr}&")")&CHAR(10)&')
    testo += '""'
    ws.cell(row=r, column=2, value=testo)
    ws.cell(row=r, column=2).font = F_BASE
    ws.cell(row=r, column=2).alignment = LW
    ws.cell(row=r, column=2).fill = FILL_GRIGIO
    ws.cell(row=r, column=2).border = BORDO

    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "portrait"
    return ws


def foglio_riepilogo(wb, settimane):
    ws = wb.create_sheet("Riepilogo")
    ws.sheet_properties.tabColor = "7030A0"
    titolo(ws, "RIEPILOGO PER PARTECIPANTE",
           f"Conteggi sull'intero periodo del servizio ({settimane} settimane)", 8)
    intestazioni(ws, 5,
                 ["Partecipante", "Turni ordinari", "Doppi turni\n(sostituzioni)",
                  "Recuperi\neffettuati", "Totale guide", "Turni saltati\nnon recuperabili",
                  "Turni saltati\nda recuperare", "Recuperi\nancora da saldare"],
                 [26, 20, 16, 14, 14, 17, 17, 17])

    ultima = TUR_R0 + settimane * 5 - 1
    for i in range(len(PARTECIPANTI)):
        r = RIEP_R0 + i
        p = f"$A{r}"
        valori = {
            1: f"=Partecipanti!$B${PART_R0 + i}",
            2: f'=COUNTIFS(Turnazione!$F${TUR_R0}:$F${ultima},{p},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Ordinario")',
            3: f'=COUNTIFS(Turnazione!$F${TUR_R0}:$F${ultima},{p},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Sostituzione (doppio turno)")',
            4: f'=COUNTIFS(Turnazione!$F${TUR_R0}:$F${ultima},{p},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Recupero")',
            5: f"=SUM($B{r}:$D{r})",
            6: f'=COUNTIFS(Calcoli!$C${TUR_R0}:$C${ultima},{p},'
               f'Calcoli!$E${TUR_R0}:$E${ultima},1,Calcoli!$F${TUR_R0}:$F${ultima},0)',
            7: f'=COUNTIFS(Calcoli!$C${TUR_R0}:$C${ultima},{p},'
               f'Calcoli!$F${TUR_R0}:$F${ultima},1)',
            8: f"=MAX(0,$G{r}-$D{r})",
        }
        for c, v in valori.items():
            cella = ws.cell(row=r, column=c, value=v)
            cella.font = F_BOLD if c in (1, 5) else F_BASE
            cella.alignment = L if c == 1 else C
            cella.border = BORDO
            if c == 5:
                cella.fill = PatternFill("solid", fgColor=VERDE)
            if c == 8:
                cella.fill = PatternFill("solid", fgColor=ARANCIO)
        ws.row_dimensions[r].height = 18

    r = RIEP_R0 + len(PARTECIPANTI)
    ws.cell(row=r, column=1, value="TOTALE").font = F_BOLD
    ws.cell(row=r, column=1).fill = FILL_SEZ
    ws.cell(row=r, column=1).border = BORDO
    for c in range(2, 9):
        col = get_column_letter(c)
        cella = ws.cell(row=r, column=c, value=f"=SUM({col}{RIEP_R0}:{col}{r - 1})")
        cella.font = F_BOLD
        cella.alignment = C
        cella.fill = FILL_SEZ
        cella.border = BORDO

    ultima_dati = r - 1

    # Grafico: composizione dei turni per partecipante
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    intestazione = ws.cell(row=r, column=1, value="DISTRIBUZIONE DEL CARICO")
    intestazione.font = F_SEZ
    intestazione.fill = FILL_SEZ
    intestazione.alignment = L
    grafico_carico(ws, RIEP_R0, ultima_dati, f"A{r + 1}", settimane)

    # Tabella e grafico dell'andamento settimanale
    r += 20
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.row_breaks.append(Break(id=r - 1))
    intestazione = ws.cell(row=r, column=1, value="ANDAMENTO SETTIMANALE")
    intestazione.font = F_SEZ
    intestazione.fill = FILL_SEZ
    intestazione.alignment = L

    r += 2
    riga_intestazioni = r
    etichette = ["Settimana", "Periodo", "Turni ordinari", "Sostituzioni\n(doppi turni)",
                 "Recuperi", "Da assegnare", "Variazioni totali"]
    for c, testo in enumerate(etichette, start=1):
        cella = ws.cell(row=r, column=c, value=testo)
        cella.font = F_HEAD
        cella.fill = FILL_HEAD
        cella.alignment = CW
        cella.border = BORDO
    ws.row_dimensions[r].height = 30

    for w in range(1, settimane + 1):
        r += 1
        pos = f"({w}-1)*5+1"
        valori = {
            1: w,
            2: f'=TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},{pos}),"dd/mm")&" - "&'
               f'TEXT(INDEX(Turnazione!$B${TUR_R0}:$B${ultima},{pos}+4),"dd/mm/yyyy")',
            3: f'=COUNTIFS(Turnazione!$A${TUR_R0}:$A${ultima},$A{r},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Ordinario")',
            4: f'=COUNTIFS(Turnazione!$A${TUR_R0}:$A${ultima},$A{r},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Sostituzione (doppio turno)")',
            5: f'=COUNTIFS(Turnazione!$A${TUR_R0}:$A${ultima},$A{r},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"Recupero")',
            6: f'=COUNTIFS(Turnazione!$A${TUR_R0}:$A${ultima},$A{r},'
               f'Turnazione!$G${TUR_R0}:$G${ultima},"DA ASSEGNARE")',
            7: f"=$D{r}+$E{r}+$F{r}",
        }
        for c, v in valori.items():
            cella = ws.cell(row=r, column=c, value=v)
            cella.font = F_BOLD if c == 7 else F_BASE
            cella.alignment = L if c == 2 else C
            cella.border = BORDO
            if c == 7:
                cella.fill = FILL_GRIGIO
        ws.row_dimensions[r].height = 16

    grafico_settimane(ws, riga_intestazioni + 1, r, f"A{r + 2}")

    r += 20
    ws.cell(row=r, column=1, value="Come leggere il riepilogo").font = F_SEZ
    note = [
        ("Totale guide", "Numero di volte in cui la persona prende effettivamente la macchina."),
        ("Turni saltati non recuperabili",
         "Turni persi per licenza o attività esterna di settimana intera: non vanno recuperati (R1, R2)."),
        ("Turni saltati da recuperare",
         "Turni persi per altri motivi: generano un recupero nella settimana successiva (R3)."),
        ("Recuperi ancora da saldare",
         "Recuperi maturati e non ancora effettuati entro la fine del periodo. Durante il periodo "
         "il recupero slitta da solo alla prima settimana utile (R7): un valore diverso da zero "
         "segnala un debito rimasto aperto, da riportare nel calendario successivo."),
    ]
    for chiave, valore in note:
        r += 1
        etichetta = ws.cell(row=r, column=1, value=chiave)
        etichetta.font = F_BOLD
        etichetta.alignment = LW
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        cella = ws.cell(row=r, column=2, value=valore)
        cella.font = F_NOTA
        cella.alignment = LW
        ws.row_dimensions[r].height = 30

    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0     # adatta solo la larghezza, non comprime l'altezza
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


# --------------------------------------------------------------------------- #
# Composizione della cartella di lavoro
# --------------------------------------------------------------------------- #

def costruisci(inizio, settimane, output):
    fine = inizio + dt.timedelta(weeks=settimane - 1, days=4)

    wb = Workbook()
    wb.remove(wb.active)

    foglio_regole(wb, inizio, fine, settimane)
    foglio_partecipanti(wb)
    foglio_assenze(wb, inizio, settimane)
    foglio_turnazione(wb, inizio, fine, settimane)
    foglio_comunicazione(wb, settimane)
    foglio_riepilogo(wb, settimane)
    foglio_calcoli(wb, inizio, settimane)

    ultima_ass = ASS_R0 + MAX_ASSENZE - 1
    nomi = {
        "A_SETT": f"Assenze!$A${ASS_R0}:$A${ultima_ass}",
        "A_PERS": f"Assenze!$C${ASS_R0}:$C${ultima_ass}",
        "A_TIPO": f"Assenze!$D${ASS_R0}:$D${ultima_ass}",
        "A_DATA": f"Assenze!$E${ASS_R0}:$E${ultima_ass}",
        "PARTECIPANTI":
            f"Partecipanti!$B${PART_R0}:$B${PART_R0 + len(PARTECIPANTI) - 1}",
    }
    for nome, rif in nomi.items():
        wb.defined_names[nome] = DefinedName(nome, attr_text=rif)

    wb.calculation.fullCalcOnLoad = True
    wb.active = wb.sheetnames.index("Turnazione")
    wb.save(output)
    return output, fine


def main():
    ap = argparse.ArgumentParser(description="Genera il file Excel della turnazione auto.")
    ap.add_argument("--inizio", default="2026-09-28",
                    help="Lunedì di inizio servizio (AAAA-MM-GG). Predefinito: 2026-09-28")
    ap.add_argument("--settimane", type=int, default=13,
                    help="Numero di settimane da generare. Predefinito: 13")
    ap.add_argument("--output", default="turnazione-auto-trinitapoli-foggia-20260921.xlsx",
                    help="Nome del file Excel da produrre.")
    args = ap.parse_args()

    inizio = dt.date.fromisoformat(args.inizio)
    if inizio.weekday() != 0:
        raise SystemExit("La data di inizio deve essere un lunedì.")

    percorso, fine = costruisci(inizio, args.settimane, args.output)
    print(f"File generato: {percorso}")
    print(f"Periodo: {inizio.strftime('%d/%m/%Y')} - {fine.strftime('%d/%m/%Y')} "
          f"({args.settimane} settimane)")


if __name__ == "__main__":
    main()

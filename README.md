# Servizio di turnazione auto — Trinitapoli / Foggia

Turnazione della macchina sulla tratta **Trinitapoli → Foggia → Trinitapoli**,
su cinque giorni a settimana, fra cinque partecipanti.

| | |
|---|---|
| Partecipanti | Nicola, Stefano, Rocco, Giuseppe, Savino |
| Frequenza | lunedì – venerdì |
| Ordine base | Lun Nicola · Mar Stefano · Mer Rocco · Gio Giuseppe · Ven Savino |
| Periodo coperto | 28/09/2026 – 25/12/2026 (13 settimane) |

## Contenuto del progetto

| File | Descrizione |
|---|---|
| `turnazione-auto-trinitapoli-foggia-20260921.xlsx` | Il servizio vero e proprio, pronto all'uso |
| `genera_turnazione.py` | Genera il file Excel; serve per estendere il periodo o cambiare la data di avvio |
| `test_turnazione.py` | Verifica automatica delle regole ricalcolando il file con LibreOffice Calc |

## Come si usa il file Excel

Il file contiene sei fogli visibili. **L'unico da compilare a mano è «Assenze»**:
tutto il resto si ricalcola da solo.

| Foglio | A cosa serve |
|---|---|
| `Regole` | Regolamento del servizio, legenda e istruzioni |
| `Partecipanti` | Anagrafica e ordine base; cambiando un nome si aggiorna tutto il calendario |
| `Assenze` | **Foglio di input.** Una riga per assenza: settimana, persona, tipo (menù a tendina) e, solo per «Altro motivo», la data |
| `Turnazione` | Calendario completo: turno previsto, conducente effettivo, doppi turni, recuperi |
| `Comunicazione Giovedì` | Prospetto della settimana scelta e testo già pronto da inviare |
| `Riepilogo` | Turni, doppi turni, recuperi e debiti di turno per ogni partecipante |

Le celle su sfondo giallo sono modificabili; tutte le altre contengono formule.

## Regole implementate

| | Regola | Effetto |
|---|---|---|
| **R1** | Licenza per la settimana intera | Il turno **non** va recuperato |
| **R2** | Attività esterna per la settimana intera | Il turno **non** va recuperato |
| **R3** | Mancata presa della macchina per altri motivi | Turno **da recuperare** nella settimana successiva |
| **R4** | Scelta del sostituto | Va a chi ha finora **meno doppi turni**; a parità si scorre l'ordine base a partire dal collega successivo all'assente, così il doppio turno cade sempre su una turnazione diversa dalla prima |
| **R5** | Modalità del recupero (solo R3) | La settimana dopo chi ha saltato guida due volte — il proprio giorno più quello di chi lo aveva sostituito — e il sostituto è a riposo compensativo |
| **R5-bis** | Più recuperi verso la stessa persona | Si saldano uno per settimana, nell'ordine in cui sono maturati; chi sta già guidando un recupero non viene scelto come sostituto nella stessa settimana |
| **R6** | Comunicazione | Le posizioni della settimana successiva vanno comunicate entro il **giovedì** |
| **R7** | Recupero non effettuabile subito | Se chi deve recuperare è assente anche nella settimana del recupero, il turno **slitta da solo alla prima settimana utile**, sempre sul giorno di chi lo aveva sostituito. Un debito ancora aperto a fine periodo resta in «Riepilogo» → *Recuperi ancora da saldare* |

## Rigenerare o estendere il calendario

```bash
pip install openpyxl
python3 genera_turnazione.py --inizio 2027-01-04 --settimane 26 \
        --output turnazione-auto-trinitapoli-foggia-20270104.xlsx
```

`--inizio` deve essere un lunedì. Il file rigenerato riparte con il foglio
«Assenze» vuoto.

## Verificare il file

```bash
sudo apt-get install -y libreoffice-calc
python3 test_turnazione.py
```

La suite compila il foglio «Assenze» con nove scenari — turnazione base,
licenza di settimana intera, assenza recuperabile con relativo recupero,
rotazione dei doppi turni, slittamento del recupero di una e di più settimane,
catena di assenze recuperabili consecutive, debito aperto a fine periodo,
doppia assenza contemporanea e comunicazione del giovedì — fa ricalcolare il
file da LibreOffice Calc e confronta i risultati con quelli attesi
(52 controlli).

# idea2word

Strumento per **IDEA StatiCa Connection** → **Word**. Parte dal PDF di verifica e prepara le immagini da incollare nella relazione. Ha **due schermate** (tab):

| Schermata | Cosa produce |
|---|---|
| **1 · Viste ortogonali** | Un'unica tavola con le quattro viste del nodo in proiezione ortogonale (3D, Z, Y, X) |
| **2 · Relazione (pagine)** | Le pagine del PDF ritagliate e reimpaginate, tutte dello stesso formato (max 170 × 240 mm), da incollare in Word |

Le due schermate si dividono il lavoro: la 2 **non include** le prime 4 immagini di vista (le gestisce la 1).

## Schermata 1 · Viste ortogonali
- Estrae dal PDF le prime 4 immagini di vista, nell'ordine: 3D, Z, Y, X.
- Le ritaglia sul contenuto (niente margini bianchi né didascalie). Il simbolo degli assi viene eliminato se è separato dall'immagine da spazio bianco, mantenuto se la tocca o si sovrappone (es. parti in calcestruzzo grigio chiaro).
- Le dispone: **3D** in basso a sinistra, **Z** in basso a destra, **Y** in alto a destra, **X** in alto a sinistra.
- Scale iniziali automatiche (Y di riferimento; X e Z adattate alla sua quota/ascissa), correggibili con slider o con **calibrazione assistita**: si cliccano due punti di distanza nota e si inserisce la misura reale in mm.
- Esporta in **PNG** oppure inserisce l'immagine in **Word** nella posizione del cursore, ridimensionata per stare in **max 170 × 240 mm** (vale il vincolo più limitante, proporzioni mantenute).

## Schermata 2 · Relazione (pagine)
1. **Ritaglio**: scarta l'intestazione (Progetto / Progetto n / Autore / logo) e il numero di pagina. Il riquadro orizzontale è comune a tutte le pagine (il più largo tra tutte). Si parte da **"Membrature"**; le 4 viste con le loro didascalie sono escluse.
2. **Reimpaginazione**: gli spazi vuoti a fondo pagina e tra i capitoli vengono tolti (spazio massimo regolabile) e il contenuto viene ridistribuito in pagine tutte identiche.
3. **Scala**: il contenuto è ridotto per stare in **max 170 mm di larghezza × 240 mm di altezza**, vale il vincolo più limitante (larghezza del riquadro oppure altezza del blocco indivisibile più alto, es. un'immagine).
4. **Paragrafi interi**: i tagli di pagina cadono, se possibile, tra un capitolo e l'altro. Immagini, righe di testo e righe di tabella non vengono mai spezzate; una tabella lunga viene tagliata solo tra due righe e l'intestazione ripetuta dal PDF a ogni cambio pagina viene tolta dove la tabella prosegue.
5. **Anteprima** delle pagine (doppio click per ingrandire) e pulsante **Inserisci in Word (cursore)**: le pagine vengono inserite una dopo l'altra, a misura reale in mm, con interruzione di pagina tra l'una e l'altra (disattivabile).
   È possibile anche **Salvare le pagine in PNG**.

Parametri modificabili (poi *Ricalcola* o Invio): larghezza max, altezza max, spazio tra blocchi, risoluzione (dpi).

## Uso
1. Doppio click su `idea2word.pyw` (serve Python 3 installato; al primo avvio installa da solo le librerie).
2. **Apri PDF** in una delle due schermate: passando all'altra tab il PDF già aperto viene caricato automaticamente.
3. Schermata 1: regola scala/offset o usa **Calibra con 2 punti**, poi **Esporta PNG** / **Inserisci in Word**.
   Schermata 2: controlla l'anteprima, poi **Inserisci in Word (cursore)**.

In caso di errori compare una finestra e viene scritto `idea2word_errori.log`.

## Requisiti
Vedi `requirements.txt`: `pymupdf`, `pillow`, `numpy`, `scipy`, `pywin32` (solo Windows, per Word).

## Note
- Le immagini di vista nel PDF sono raster 500×500 px: la qualità della schermata 1 è limitata da IDEA StatiCa. Le pagine della schermata 2 sono invece rese dal PDF alla risoluzione scelta (testo e tabelle nitidi).
- L'impaginazione è pensata per il layout dei PDF di IDEA StatiCa Connection 26.1 (intestazione con «Autore:», numero pagina «n / N», sezione «Membrature»).
- Inserimento in Word: richiede Windows e Word aperto con il cursore nel punto voluto.

## Licenza
MIT

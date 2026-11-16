# idea2word

Compone in **proiezione ortogonale** le quattro viste del nodo contenute nel PDF di verifica di **IDEA StatiCa Connection** (vista 3D, pianta Z, vista Y, vista X).

## Cosa fa
- Estrae dal PDF le prime 4 immagini di vista, nell'ordine: 3D, Z, Y, X.
- Le ritaglia sul contenuto (niente margini bianchi né didascalie). Il simbolo degli assi viene eliminato se è separato dall'immagine da spazio bianco, mantenuto se la tocca o si sovrappone (es. parti in calcestruzzo grigio chiaro).
- Le dispone: **3D** in basso a sinistra, **Z** in basso a destra, **Y** in alto a destra, **X** in alto a sinistra.
- Scale iniziali automatiche (Y di riferimento; X e Z adattate alla sua quota/ascissa), correggibili con slider o con **calibrazione assistita**: si cliccano due punti di distanza nota e si inserisce la misura reale in mm.
- Esporta in **PNG** oppure inserisce l'immagine in **Word** nella posizione del cursore (Windows, Word aperto).

## Uso
1. Doppio click su `idea2word.pyw` (serve Python 3 installato; al primo avvio installa da solo le librerie).
2. **Apri PDF** → anteprima.
3. Regola scala / offset delle viste o usa **Calibra con 2 punti**.
4. **Esporta PNG** oppure **Inserisci in Word (cursore)**.

Per una scala reale calibra tutte e 4 le viste. In caso di errori compare una finestra e viene scritto `ComponiViste_errori.log`.

## Note
- Le immagini nel PDF sono raster 500×500 px: la qualità è limitata da IDEA StatiCa.
- Testato con output di IDEA StatiCa Connection 26.1.

## Licenza
MIT

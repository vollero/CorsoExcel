## **1. Introduzione alle Formule e ai Riferimenti di Cella (30 minuti)**

**Obiettivi:**

- Capire cosa sono le formule e come vengono utilizzate in Excel.
- Comprendere l'importanza dei riferimenti di cella nelle formule.

**Argomenti:**

- Concetto di formula.
- Introduzione ai riferimenti di cella.
- Operatori matematici di base (`+`, `-`, `*`, `/`, `^`).

**Attività ed Esercizi:**

- **Esercizio 1: Calcoli Semplici**

  - **Attività:**
    - Inserire numeri in due celle (ad esempio, `A1` e `B1`).
    - Nella cella `C1`, creare una formula per sommare i due numeri: `=A1+B1`.
    - Ripetere l'esercizio per sottrazione (`=A1-B1`), moltiplicazione (`=A1*B1`), divisione (`=A1/B1`) e potenza (`=A1^B1`).
  - **Obiettivo:**
    - Familiarizzare con l'inserimento di formule e l'uso degli operatori matematici.

- **Esercizio 2: Modifica dei Riferimenti di Cella**

  - **Attività:**
    - Modificare i valori nelle celle `A1` e `B1` e osservare come cambia il risultato in `C1`.
  - **Obiettivo:**
    - Comprendere come i riferimenti di cella rendono le formule dinamiche.

---

## **2. Riferimenti Relativi, Assoluti e Misti (1 ora)**

**Obiettivi:**

- Distinguere tra riferimenti relativi, assoluti e misti.
- Comprendere come i riferimenti influenzano le formule quando vengono copiate.

**Argomenti:**

- **Riferimenti Relativi:** si adattano quando vengono copiati (es. `A1`).
- **Riferimenti Assoluti:** rimangono fissi quando copiati (es. `$A$1`).
- **Riferimenti Misti:** combinazione di assoluti e relativi (es. `$A1`, `A$1`).

**Attività ed Esercizi:**

- **Esercizio 3: Tabella delle Moltiplicazioni**

  - **Attività:**
    - Creare una tabella da `1x1` a `10x10`.
    - Utilizzare riferimenti relativi per riempire la tabella con le formule appropriate.
    - Copiare le formule e osservare come i riferimenti si adattano automaticamente.
  - **Obiettivo:**
    - Comprendere il funzionamento dei riferimenti relativi.

- **Esercizio 4: Uso di Riferimenti Assoluti**

  - **Attività:**
    - Creare un elenco di prodotti con prezzi unitari.
    - Inserire il tasso IVA in una cella fissa (es. `B1`).
    - Calcolare il prezzo totale con IVA per ogni prodotto usando un riferimento assoluto al tasso IVA (`=$B$1`).
  - **Obiettivo:**
    - Apprendere come utilizzare riferimenti assoluti per riferirsi a celle fisse.

- **Esercizio 5: Riferimenti Misti in una Tabella**

  - **Attività:**
    - Creare una tabella di conversione valute con tassi di cambio.
    - Utilizzare riferimenti misti per applicare correttamente i tassi di cambio nelle formule.
  - **Obiettivo:**
    - Comprendere l'uso di riferimenti misti in situazioni pratiche.

---

## **3. Uso delle Funzioni di Base di Excel (1 ora e 30 minuti)**

**Obiettivi:**

- Conoscere e utilizzare le funzioni matematiche e statistiche di base.
- Comprendere la sintassi delle funzioni.
- Applicare funzioni in contesti pratici.

**Argomenti:**

- **Sintassi delle Funzioni:** `=NOME_FUNZIONE(argomenti)`.
- **Funzioni Matematiche:** `SOMMA`, `MEDIA`, `MIN`, `MAX`.
- **Funzioni di Conteggio:** `CONTA.NUMERI`, `CONTA.VALORI`, `CONTA.SE`.
- **Funzioni Logiche:** `SE`, `E`, `O`.
- **Funzioni di Testo:** `CONCATENA`, `SINISTRA`, `DESTRA`, `STRINGA.ESTRAI`.

**Attività ed Esercizi:**

- **Esercizio 6: Calcolo del Totale con SOMMA**

  - **Attività:**
    - Inserire una serie di numeri in un intervallo di celle.
    - Utilizzare `=SOMMA(intervallo)` per calcolare il totale.
  - **Obiettivo:**
    - Apprendere l'uso della funzione `SOMMA`.

- **Esercizio 7: Calcolo della Media dei Voti**

  - **Attività:**
    - Inserire i voti di uno studente in diverse materie.
    - Calcolare la media con `=MEDIA(intervallo)`.
  - **Obiettivo:**
    - Utilizzare la funzione `MEDIA` per analizzare dati.

- **Esercizio 8: Identificare Valori Minimi e Massimi**

  - **Attività:**
    - Inserire dati numerici (es. vendite giornaliere).
    - Utilizzare `=MIN(intervallo)` e `=MAX(intervallo)` per trovare i valori estremi.
  - **Obiettivo:**
    - Comprendere l'uso di `MIN` e `MAX`.

- **Esercizio 9: Conteggio degli Elementi**

  - **Attività:**
    - In una lista mista di numeri e testi, utilizzare `=CONTA.NUMERI(intervallo)` e `=CONTA.VALORI(intervallo)`.
  - **Obiettivo:**
    - Distinguere tra diverse funzioni di conteggio.

- **Esercizio 10: Uso della Funzione SE**

  - **Attività:**
    - Creare una lista di studenti con i loro punteggi.
    - Utilizzare `=SE(punteggio>=60, "Passato", "Riprovato")`.
  - **Obiettivo:**
    - Apprendere l'uso della funzione logica `SE`.

- **Esercizio 11: Funzioni SE Nidificate**

  - **Attività:**
    - Assegnare voti letterali (`A`, `B`, `C`, `D`, `F`) in base ai punteggi utilizzando funzioni `SE` nidificate.
  - **Obiettivo:**
    - Comprendere come nidificare funzioni per creare condizioni multiple.

- **Esercizio 12: Funzioni di Testo**

  - **Attività:**
    - Utilizzare `=CONCATENA(nome, " ", cognome)` per unire testo.
    - Estrarre parti di testo con `=SINISTRA(testo, numero_caratteri)` e `=DESTRA(testo, numero_caratteri)`.
  - **Obiettivo:**
    - Manipolare stringhe di testo con funzioni apposite.

---

## **4. Esercizi di Consolidamento e Applicazioni Pratiche (1 ora)**

**Obiettivi:**

- Applicare tutte le conoscenze acquisite in esercizi più complessi.
- Sviluppare capacità di problem solving utilizzando formule e funzioni.

**Attività ed Esercizi:**

- **Esercizio 13: Gestione di un Budget Familiare**

  - **Attività:**
    - Creare una tabella con le spese mensili suddivise per categorie (affitto, bollette, alimentari, ecc.).
    - Calcolare il totale delle spese.
    - Utilizzare la funzione `SE` per evidenziare se le spese superano un budget predefinito.
  - **Obiettivo:**
    - Integrare formule, funzioni e riferimenti per un caso pratico.

- **Esercizio 14: Calcolo delle Commissioni di Vendita**

  - **Attività:**
    - Creare una lista di venditori con le loro vendite totali.
    - Calcolare le commissioni applicando percentuali diverse in base al volume di vendite utilizzando `SE` e riferimenti assoluti.
  - **Obiettivo:**
    - Applicare funzioni logiche e riferimenti in contesti reali.

- **Esercizio 15: Analisi dei Dati con CONTA.SE e SOMMA.SE**

  - **Attività:**
    - In un elenco di prodotti venduti, utilizzare `=CONTA.SE(intervallo, criterio)` per contare quante volte un prodotto è stato venduto.
    - Utilizzare `=SOMMA.SE(intervallo, criterio, somma_intervallo)` per calcolare il totale delle vendite per un determinato prodotto.
  - **Obiettivo:**
    - Comprendere l'uso di funzioni condizionali per l'analisi dei dati.

- **Esercizio 16: Creazione di un Grafico Semplice**

  - **Attività:**
    - Utilizzare i dati delle vendite per creare un grafico a colonne che mostri le vendite per prodotto.
  - **Obiettivo:**
    - Introdurre la visualizzazione dei dati con grafici (facoltativo se il tempo lo permette).

- **Esercizio 17: Pianificazione di Progetti con Date**

  - **Attività:**
    - Inserire date di inizio e durata in giorni per diverse attività.
    - Calcolare le date di fine utilizzando formule con date.
    - Utilizzare funzioni come `=OGGI()` e `=DATA(anno; mese; giorno)`.
  - **Obiettivo:**
    - Applicare formule con dati di tipo data e ora.

- **Esercizio 18: Valutazione delle Prestazioni con FUNZIONI STATISTICHE**

  - **Attività:**
    - Inserire dati sulle prestazioni (es. tempi di risposta, produttività).
    - Calcolare media, deviazione standard, massimo e minimo.
  - **Obiettivo:**
    - Utilizzare funzioni statistiche per l'analisi avanzata dei dati.

---

### **Esercizi Aggiuntivi per Approfondire**

- **Esercizio 19: Creazione di un Preventivo**

  - **Attività:**
    - Creare un modulo di preventivo con elenco di servizi/prodotti, quantità e prezzi.
    - Calcolare il totale per ogni voce e il totale generale.
    - Applicare sconti in base all'importo totale utilizzando `SE`.
  - **Obiettivo:**
    - Sviluppare un documento utile e applicare le funzioni in un contesto commerciale.

- **Esercizio 20: Registro Ore Lavorate**

  - **Attività:**
    - Registrare le ore di ingresso e uscita dei dipendenti.
    - Calcolare le ore lavorate utilizzando formule con l'ora.
    - Considerare pause e straordinari.
  - **Obiettivo:**
    - Applicare formule con dati di tipo ora e funzioni logiche.

## **1. Introduzione alle Funzioni di Ricerca Avanzate (15 minuti)**

**Obiettivi:**

- Introdurre le funzioni di ricerca e il loro ruolo nell'analisi dei dati.
- Spiegare l'importanza della combinazione di tabelle in Excel.

**Argomenti:**

- Revisione rapida delle formule e funzioni di base.
- Presentazione delle funzioni di ricerca e delle loro applicazioni pratiche.

**Attività:**

- **Discussione Introduttiva:**
  - Chiedere agli studenti quali sfide hanno incontrato nell'utilizzo delle funzioni di base.
  - Presentare esempi di situazioni in cui le funzioni di ricerca sono necessarie.

---

## **2. Funzioni di Ricerca: CERCA.VERT, CERCA.ORIZZ, INDICE, CONFRONTA, CERCA.X (1 ora e 30 minuti)**

**Obiettivi:**

- Comprendere e utilizzare le funzioni di ricerca per trovare e recuperare dati.
- Conoscere le differenze tra le varie funzioni di ricerca e quando utilizzarle.

**Argomenti:**

- **CERCA.VERT:**
  - Sintassi e parametri.
  - Ricerca esatta e approssimata.
  - Limitazioni della funzione.

- **CERCA.ORIZZ:**
  - Utilizzo in tabelle orizzontali.

- **INDICE e CONFRONTA:**
  - Utilizzo combinato per ricerche più flessibili.
  - Vantaggi rispetto a CERCA.VERT.

- **CERCA.X [se disponibile]:**
  - Funzionalità avanzate e sintassi semplificata.

**Attività ed Esercizi:**

### **Esercizio 1: Utilizzo di CERCA.VERT per Recuperare Dati**

- **Attività:**
  - Fornire una tabella con elenco di prodotti (codice, nome, prezzo).
  - Creare un modulo d'ordine dove, inserendo il codice prodotto, si recuperano automaticamente nome e prezzo utilizzando `CERCA.VERT`.

- **Procedura:**
  - Nella cella dove si desidera ottenere il nome del prodotto, inserire la formula:
    ```
    =CERCA.VERT(codice_prodotto; tabella_prodotti; colonna_nome; FALSO)
    ```
  - Spiegare che `FALSO` indica una corrispondenza esatta.

- **Obiettivo:**
  - Comprendere la sintassi di `CERCA.VERT` e come utilizzarla per recuperare informazioni.

### **Esercizio 2: CERCA.VERT con Ricerca Approssimata**

- **Attività:**
  - Fornire una tabella con fasce di reddito e aliquote fiscali.
  - Calcolare l'aliquota fiscale applicabile a diversi redditi utilizzando `CERCA.VERT` con corrispondenza approssimata.

- **Procedura:**
  - Utilizzare la formula:
    ```
    =CERCA.VERT(reddito; tabella_aliquote; colonna_aliquota; VERO)
    ```
  - Spiegare che `VERO` permette la corrispondenza approssimata, utile per intervalli.

- **Obiettivo:**
  - Imparare a utilizzare `CERCA.VERT` per ricerche approssimate.

### **Esercizio 3: Limitazioni di CERCA.VERT e Introduzione a INDICE & CONFRONTA**

- **Attività:**
  - Fornire una tabella dove la chiave di ricerca non è nella prima colonna.
  - Tentare di utilizzare `CERCA.VERT` e discutere le limitazioni.
  - Implementare la ricerca utilizzando `INDICE` e `CONFRONTA`.

- **Procedura:**
  - Spiegare che `CERCA.VERT` non funziona se la chiave non è nella prima colonna.
  - Utilizzare la formula combinata:
    ```
    =INDICE(colonna_ritorno; CONFRONTA(valore_cercato; colonna_chiave; 0))
    ```
  - Spiegare il ruolo di `CONFRONTA` nel trovare la posizione.

- **Obiettivo:**
  - Comprendere le limitazioni di `CERCA.VERT` e come `INDICE` e `CONFRONTA` possono superarle.

### **Esercizio 4: Ricerca Bidirezionale con INDICE & CONFRONTA**

- **Attività:**
  - Fornire una tabella di dati con mesi (colonne) e categorie di spesa (righe).
  - Creare una formula che recupera l'importo speso in una specifica categoria e mese utilizzando `INDICE` e `CONFRONTA` nidificati.

- **Procedura:**
  - Utilizzare la formula:
    ```
    =INDICE(intervallo_dati; CONFRONTA(categoria; intervallo_categorie; 0); CONFRONTA(mese; intervallo_mesi; 0))
    ```
  - Spiegare come `INDICE` utilizza sia il numero di riga che di colonna.

- **Obiettivo:**
  - Utilizzare `INDICE` e `CONFRONTA` per effettuare ricerche bidimensionali.

### **Esercizio 5: Gestione degli Errori nelle Funzioni di Ricerca**

- **Attività:**
  - Modificare le formule precedenti per gestire valori non trovati utilizzando `SE.ERRORE`.

- **Procedura:**
  - Incapsulare la formula con `SE.ERRORE`:
    ```
    =SE.ERRORE(formula; "Valore non trovato")
    ```
  - Spiegare l'importanza di gestire gli errori per migliorare la robustezza del foglio di calcolo.

- **Obiettivo:**
  - Imparare a gestire gli errori nelle formule per evitare risultati indesiderati.

### **Esercizio 6: Utilizzo di CERCA.X [se disponibile]**

- **Attività:**
  - Riscrivere gli esercizi precedenti utilizzando `CERCA.X`.
  - Esplorare le funzionalità aggiuntive, come la ricerca su colonne a sinistra.

- **Procedura:**
  - Sintassi di base:
    ```
    =CERCA.X(valore_cercato; matrice_cercata; matrice_ritorno; [se_non_trovato]; [modo_corrispondenza]; [modo_ricerca])
    ```
  - Dimostrare come `CERCA.X` può sostituire sia `CERCA.VERT` che `INDICE` & `CONFRONTA`.

- **Obiettivo:**
  - Familiarizzare con `CERCA.X` e le sue potenzialità rispetto alle altre funzioni.

---

## **3. Combinazione di Dati da Tabelle Multiple (1 ora e 30 minuti)**

**Obiettivi:**

- Imparare a combinare dati da diverse fonti utilizzando funzioni e strumenti di Excel.
- Comprendere come gestire relazioni tra tabelle.

**Argomenti:**

- **Uso di CERCA.VERT e INDICE/CONFRONTA tra tabelle diverse.**
- **Funzioni Avanzate:**
  - `SOMMA.SE` e `SOMMA.PIÙ.SE`.
  - `CONTA.SE` e `CONTA.PIÙ.SE`.
  - `MEDIA.SE` e `MEDIA.PIÙ.SE`.

- **Introduzione a Power Query (se il tempo lo permette).**

**Attività ed Esercizi:**

### **Esercizio 7: Unire Dati da Due Tabelle con CERCA.VERT**

- **Attività:**
  - Fornire due tabelle: una con informazioni sui clienti e una con ordini effettuati (senza dettagli cliente).
  - Utilizzare `CERCA.VERT` per aggiungere le informazioni del cliente agli ordini.

- **Procedura:**
  - Nella tabella degli ordini, aggiungere colonne per le informazioni del cliente.
  - Utilizzare `CERCA.VERT` per recuperare queste informazioni in base all'ID cliente.

- **Obiettivo:**
  - Applicare `CERCA.VERT` per combinare dati da tabelle diverse.

### **Esercizio 8: Analisi delle Vendite con SOMMA.PIÙ.SE**

- **Attività:**
  - Fornire una tabella di vendite con colonne per data, prodotto, regione e importo.
  - Utilizzare `SOMMA.PIÙ.SE` per calcolare il totale delle vendite per un determinato prodotto e regione.

- **Procedura:**
  - Sintassi:
    ```
    =SOMMA.PIÙ.SE(intervallo_somma; intervallo_criteri1; criterio1; [intervallo_criteri2; criterio2]; ...)
    ```
  - Esempio:
    ```
    =SOMMA.PIÙ.SE(importi; prodotti; "Prodotto A"; regioni; "Nord")
    ```

- **Obiettivo:**
  - Imparare ad aggregare dati basati su criteri multipli.

### **Esercizio 9: Conteggio degli Ordini con CONTA.PIÙ.SE**

- **Attività:**
  - Utilizzare la stessa tabella dell'esercizio precedente.
  - Contare il numero di ordini effettuati in un determinato mese per una specifica regione.

- **Procedura:**
  - Sintassi:
    ```
    =CONTA.PIÙ.SE(intervallo_criteri1; criterio1; [intervallo_criteri2; criterio2]; ...)
    ```
  - Esempio:
    ```
    =CONTA.PIÙ.SE(date_ordini; ">="&DATA(2023;1;1); date_ordini; "<="&DATA(2023;1;31); regioni; "Nord")
    ```

- **Obiettivo:**
  - Utilizzare `CONTA.PIÙ.SE` per l'analisi dei dati.

### **Esercizio 10: Calcolo della Media con MEDIA.PIÙ.SE**

- **Attività:**
  - Calcolare la media delle vendite per prodotto in una specifica regione.

- **Procedura:**
  - Sintassi:
    ```
    =MEDIA.PIÙ.SE(intervallo_media; intervallo_criteri1; criterio1; [intervallo_criteri2; criterio2]; ...)
    ```
  - Esempio:
    ```
    =MEDIA.PIÙ.SE(importi; prodotti; "Prodotto A"; regioni; "Nord")
    ```

- **Obiettivo:**
  - Applicare `MEDIA.PIÙ.SE` per ottenere informazioni dettagliate.

### **Esercizio 11: Combinazione di Tabelle con INDICE & CONFRONTA**

- **Attività:**
  - Fornire tabelle con chiavi di ricerca non nelle prime colonne.
  - Utilizzare `INDICE` e `CONFRONTA` per unire le tabelle correttamente.

- **Procedura:**
  - Seguire la stessa logica dell'Esercizio 3, ma applicata su tabelle diverse.

- **Obiettivo:**
  - Rafforzare l'uso di `INDICE` e `CONFRONTA` in situazioni più complesse.

### **Esercizio 12: Introduzione a Power Query per la Combinazione di Dati**

- **Attività [Facoltativa se il tempo lo permette]:**
  - Importare due tabelle da file esterni.
  - Utilizzare Power Query per unire le tabelle basandosi su una chiave comune.

- **Obiettivo:**
  - Introdurre gli studenti a strumenti avanzati per la gestione dei dati.

---

## **4. Tecniche Avanzate di Ricerca e Tabelle Pivot (45 minuti)**

**Obiettivi:**

- Approfondire tecniche avanzate di ricerca.
- Utilizzare funzioni avanzate e tabelle pivot per l'analisi dei dati.

**Argomenti:**

- **Funzioni Avanzate:**
  - `SCARTO`.
  - `INDIRETTO`.
  - Funzioni Array Dinamiche: `FILTRA`, `UNICI`, `ORDINA`.

- **Introduzione alle Tabelle Pivot.**

**Attività ed Esercizi:**

### **Esercizio 13: Utilizzo di SCARTO per Creare Intervalli Dinamici**

- **Attività:**
  - Creare un grafico che si aggiorna automaticamente quando vengono aggiunti nuovi dati, utilizzando `SCARTO` per definire l'intervallo dati.

- **Procedura:**
  - Definire un nome di intervallo dinamico:
    ```
    =SCARTO(cella_iniziale; 0; 0; CONTA.VALORI(colonna_dati); 1)
    ```
  - Utilizzare questo intervallo nel grafico.

- **Obiettivo:**
  - Comprendere come creare intervalli dinamici per analisi avanzate.

### **Esercizio 14: Creare Riferimenti Indiretti con INDIRETTO**

- **Attività:**
  - Utilizzare `INDIRETTO` per riferirsi a un intervallo il cui nome è specificato in una cella.

- **Procedura:**
  - Se il nome dell'intervallo è in A1:
    ```
    =SOMMA(INDIRETTO(A1))
    ```

- **Obiettivo:**
  - Imparare a utilizzare `INDIRETTO` per creare formule più flessibili.

### **Esercizio 15: Estrazione di Dati con FILTRA [per versioni di Excel che supportano le funzioni dinamiche]**

- **Attività:**
  - Estrarre un elenco di ordini superiori a un certo importo utilizzando `FILTRA`.

- **Procedura:**
  - Sintassi:
    ```
    =FILTRA(intervallo_dati; intervallo_criteri > valore; "Nessun dato")
    ```

- **Obiettivo:**
  - Utilizzare funzioni avanzate per manipolare grandi quantità di dati.

### **Esercizio 16: Identificare Valori Unici con UNICI**

- **Attività:**
  - Ottenere un elenco di clienti unici da una tabella di ordini.

- **Procedura:**
  - Utilizzare:
    ```
    =UNICI(intervallo_clienti)
    ```

- **Obiettivo:**
  - Imparare a isolare valori distinti in un insieme di dati.

### **Esercizio 17: Analisi dei Dati con Tabelle Pivot**

- **Attività:**
  - Creare una tabella pivot per analizzare le vendite per regione e prodotto.
  - Applicare filtri e segmenti per approfondire l'analisi.

- **Procedura:**
  - Guidare gli studenti attraverso la creazione di una tabella pivot:
    - Selezionare i dati.
    - Inserire una tabella pivot.
    - Trascinare campi nelle aree appropriate (Righe, Colonne, Valori, Filtri).

- **Obiettivo:**
  - Introdurre le tabelle pivot come strumento potente per l'analisi dei dati.

---

## **5. Caso di Studio e Consolidamento (30 minuti)**

**Obiettivi:**

- Applicare tutte le competenze acquisite in un caso di studio complesso.
- Sviluppare capacità di problem solving in un contesto realistico.

**Attività:**

- **Caso di Studio: Analisi Completa delle Vendite Aziendali**

  - **Scenario:**
    - Un'azienda ha fornito diverse tabelle: clienti, prodotti, ordini, venditori.
    - Gli studenti devono combinare queste tabelle per rispondere a domande specifiche, ad esempio:
      - Quali sono i prodotti più venduti in ogni regione?
      - Qual è il fatturato totale per venditore?
      - Quali clienti hanno generato più entrate?

  - **Attività:**
    - Unire le tabelle utilizzando funzioni di ricerca avanzate (`CERCA.VERT`, `INDICE`, `CONFRONTA`).
    - Utilizzare `SOMMA.PIÙ.SE`, `CONTA.PIÙ.SE` e tabelle pivot per l'analisi.
    - Presentare i risultati in forma di report o dashboard.

- **Obiettivo:**
  - Mettere in pratica tutte le competenze acquisite durante la lezione in un progetto integrato.

### **1. Funzioni di Ricerca in Tabella (`CERCA.VERT`, `CERCA.X`)**

#### **Dataset: Catalogo Prodotti Esteso**

- **Descrizione del Dataset:**
  - Un catalogo di 10.000 prodotti, ciascuno con:
    - **Codice Prodotto** (es. `P0001`).
    - **Nome Prodotto** (es. `Pro Max Informatica 123`).
    - **Categoria** (es. `Informatica`, `Telefonia`, `Sport`, ecc.).
    - **Prezzo** (valore casuale tra 5,00 € e 2.000,00 €).

#### **Istruzioni per l'Esercizio:**

- **Importa** il dataset `catalogo_prodotti.csv` in Excel.
- **Crea** un modulo d'ordine dove inserire il `Codice Prodotto`.
- **Utilizza** le funzioni `CERCA.VERT` o `CERCA.X` per recuperare automaticamente:
  - Nome Prodotto.
  - Categoria.
  - Prezzo.
- **Gestisci** eventuali errori o codici non trovati utilizzando `SE.ERRORE`.
- **Applica** sconti in base alla categoria o al prezzo utilizzando funzioni condizionali.
- **Calcola** il totale dell'ordine considerando quantità multiple e applicando l'IVA.

---

### **2. Funzioni Aritmetiche con SE**

#### **Dataset: Valutazioni Studenti su Larga Scala**

- **Descrizione del Dataset:**
  - Un elenco di 1.000 studenti con:
    - **Nome** e **Cognome**.
    - **Voti** in 5 materie: Matematica, Italiano, Inglese, Storia, Scienze.
    - Voti generati con una distribuzione realistica (media 70, deviazione standard 15).

#### **Istruzioni per l'Esercizio:**

- **Importa** il dataset `studenti_voti_esteso.csv` in Excel.
- **Calcola** per ogni studente:
  - **Media dei voti**.
  - **Voto minimo e massimo**.
- **Determina** se lo studente è "Promosso" o "Bocciato" utilizzando `SE` (soglia: 60).
- **Assegna** un giudizio utilizzando `PIÙ.SE` o nidificando funzioni `SE`:
  - Media < 60: **Insufficiente**.
  - 60 ≤ Media < 70: **Sufficiente**.
  - 70 ≤ Media < 80: **Buono**.
  - 80 ≤ Media < 90: **Distinto**.
  - Media ≥ 90: **Ottimo**.
- **Analizza** la distribuzione dei voti per ciascuna materia:
  - Calcola **media**, **deviazione standard**, **mediana**.
  - Identifica le materie con migliori e peggiori performance.
- **Crea** grafici per visualizzare la distribuzione delle medie degli studenti.

---

### **3. Manipolazione di Stringhe e Testi**

#### **Dataset: Database Dipendenti Esteso**

- **Descrizione del Dataset:**
  - Un database di 5.000 dipendenti con:
    - **Nome** e **Cognome**.
    - **ID Dipendente** (es. `D0001`).
    - **Reparto** (es. `Vendite`, `Marketing`, `IT`, ecc.).
    - **Telefono** (formato internazionale).
    - **Codice Fiscale** (stringa alfanumerica fittizia).

#### **Istruzioni per l'Esercizio:**

- **Importa** il dataset `dipendenti_esteso.csv` in Excel.
- **Genera** per ogni dipendente:
  - **Indirizzo email aziendale** nel formato `nome.cognome@azienda.com`, tutto in minuscolo.
- **Estrai**:
  - **Iniziali** del dipendente utilizzando le funzioni `SINISTRA`, `DESTRA`, `STRINGA.ESTRAI`.
  - Parti specifiche del **Codice Fiscale** (es. data di nascita, sesso), se il formato lo permette.
- **Crea** un identificatore univoco combinando:
  - **Iniziali**.
  - **ID Dipendente**.
- **Standardizza** il formato dei numeri di telefono, ad esempio:
  - Rimuovere spazi inutili.
  - Aggiungere prefissi se mancanti.
- **Organizza** i dati:
  - Ordina i dipendenti per **reparto** e **cognome**.
  - Crea un **elenco telefonico aziendale**.

---

### **4. Introduzione alle Tabelle Pivot**

#### **Dataset: Storico Vendite su Larga Scala**

- **Descrizione del Dataset:**
  - Un dataset di 50.000 record con informazioni su vendite:
    - **Data Vendita** (ultimi 2 anni).
    - **Prodotto** (100 prodotti differenti).
    - **Regione** (Nord, Sud, Est, Ovest, Centro).
    - **Venditore** (20 venditori).
    - **Quantità** venduta.
    - **Prezzo Unitario**.
    - **Importo Totale**.

#### **Istruzioni per l'Esercizio:**

- **Importa** il dataset `vendite_esteso.csv` in Excel.
- **Crea** diverse tabelle pivot per analizzare:
  - **Vendite per Prodotto**:
    - Identifica i prodotti con maggiori vendite in termini di quantità e valore.
  - **Performance per Regione**:
    - Confronta le vendite tra le diverse regioni.
    - Individua trend regionali.
  - **Risultati dei Venditori**:
    - Valuta le performance individuali dei venditori.
    - Identifica i top performer.
  - **Analisi Temporale**:
    - Analizza le vendite per **mese**, **trimestre** e **anno**.
    - Identifica stagionalità o trend nel tempo.
- **Approfondisci** l'analisi utilizzando:
  - **Filtri** e **segmentazioni** per focalizzarti su specifici prodotti, regioni o periodi.
  - **Campi calcolati** per creare metriche personalizzate (es. Margine di Profitto).
- **Visualizza** i dati:
  - Crea **grafici pivot** per rappresentare graficamente le informazioni chiave.
  - Costruisci una **dashboard interattiva** per presentare i risultati.

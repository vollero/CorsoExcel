## **Introduzione alle Tabelle Pivot**

### **Cosa sono le Tabelle Pivot?**

Le tabelle pivot sono uno strumento potente di Excel che permette di sintetizzare, analizzare ed esplorare grandi quantità di dati in modo interattivo. Consentono di trasformare dati grezzi in informazioni significative, facilitando il processo decisionale.

### **Perché Utilizzare le Tabelle Pivot?**

- **Sintesi dei dati**: Riassumono grandi dataset in tabelle concise.
- **Analisi flessibile**: Permettono di riorganizzare i dati rapidamente.
- **Interattività**: Consentono di filtrare e approfondire i dati.
- **Visualizzazione**: Facilitano la creazione di grafici e report.

---

## **Sezione 1: Creazione di Tabelle Pivot di Base**

### **1. Preparazione dei Dati**

- Assicurarsi che i dati siano organizzati in un formato tabellare, con intestazioni chiare per ogni colonna.
- Verificare che non ci siano righe o colonne vuote all'interno del dataset.
- Rimuovere eventuali duplicati o errori nei dati.

### **2. Creazione di una Tabella Pivot**

**Passaggi:**

1. **Seleziona il Dataset**

   - Clicca su una qualsiasi cella all'interno del tuo dataset.

2. **Inserisci una Tabella Pivot**

   - Vai alla scheda **Inserisci** sulla barra multifunzione.
   - Clicca su **Tabella Pivot**.
   - Nella finestra di dialogo, assicurati che l'intervallo di dati sia corretto.
   - Scegli se posizionare la tabella pivot in un nuovo foglio di lavoro o nello stesso foglio.

3. **Struttura della Tabella Pivot**

   - Una volta creata, vedrai il **Campo Tabella Pivot** sul lato destro.
   - Ci sono quattro aree principali:
     - **Filtri**: Per filtrare l'intera tabella pivot.
     - **Colonne**: Campi da visualizzare come intestazioni di colonna.
     - **Righe**: Campi da visualizzare come righe.
     - **Valori**: Dati numerici da aggregare (somma, conteggio, media, ecc.).

### **3. Popolamento della Tabella Pivot**

**Esempio: Creare una Tabella Pivot che Mostra il Totale delle Vendite per Prodotto**

1. **Aggiungi Campi**

   - Trascina il campo **Nome Prodotto** nell'area **Righe**.
   - Trascina il campo **Ricavo** nell'area **Valori**.

2. **Imposta il Tipo di Calcolo**

   - Assicurati che il campo **Ricavo** sia impostato per eseguire la **Somma**.
   - Se necessario, clicca sulla freccia accanto a **Somma di Ricavo** > **Impostazioni campo valore** > Seleziona **Somma**.

3. **Formato Numerico**

   - Clicca con il tasto destro sul campo **Somma di Ricavo**.
   - Seleziona **Formato numeri...** e scegli il formato **Valuta**.

### **4. Personalizzazione della Tabella Pivot**

- **Ordinamento**

  - Clicca sulla freccia accanto all'intestazione **Nome Prodotto**.
  - Seleziona **Ordina dalla A alla Z** o **Ordina dal più grande al più piccolo** per i valori.

- **Filtraggio**

  - Utilizza il filtro accanto all'intestazione per selezionare specifici prodotti.

- **Stili e Layout**

  - Vai alla scheda **Progettazione** (visibile quando la tabella pivot è selezionata).
  - Scegli uno stile predefinito per migliorare l'aspetto visivo.

---

## **Sezione 2: Operazioni Intermedie con le Tabelle Pivot**

### **1. Filtri e Slicers**

**Utilizzo dei Filtri:**

- Trascina il campo **Categoria** nell'area **Filtri**.
- Nella tabella pivot, utilizza il menu a tendina sopra per selezionare una o più categorie.

**Inserimento di Slicers:**

1. Seleziona la tabella pivot.
2. Vai alla scheda **Analizza** > **Inserisci filtro dati**.
3. Seleziona il campo **Regione Cliente** e clicca su **OK**.
4. Apparirà una finestra interattiva che ti permette di filtrare i dati cliccando sulle regioni.

### **2. Raggruppamento dei Dati**

**Raggruppamento per Data:**

- Trascina il campo **Data Vendita** nell'area **Righe**.
- Clicca con il tasto destro su una data nella tabella pivot.
- Seleziona **Raggruppa**.
- Nella finestra di dialogo, seleziona le opzioni desiderate (es. **Anni**, **Trimestri**, **Mesi**).

**Raggruppamento di Valori Numerici:**

- Se hai un campo numerico (es. **Quantità**), puoi raggrupparlo in intervalli.
- Clicca con il tasto destro su un valore nella tabella pivot.
- Seleziona **Raggruppa**.
- Imposta l'intervallo desiderato (es. da 0 a 100 con passi di 10).

### **3. Campi e Elementi Calcolati**

**Creazione di un Campo Calcolato:**

1. Seleziona la tabella pivot.
2. Vai alla scheda **Analizza** > **Campi, elementi e set** > **Campo calcolato**.
3. Assegna un nome al campo, ad esempio **Margine Percentuale**.
4. Nella formula, inserisci `=Margine / Ricavo`.
5. Clicca su **Aggiungi** e poi su **OK**.
6. Il nuovo campo appare nell'area **Valori**.

**Nota:** Assicurati che i campi **Margine** e **Ricavo** siano presenti nei dati di origine.

### **4. Formattazione Avanzata**

- **Formato Condizionale:**

  - Seleziona le celle nella tabella pivot.
  - Vai alla scheda **Home** > **Formato condizionale**.
  - Applica regole per evidenziare valori al di sopra o al di sotto di una certa soglia.

- **Mostrare Valori come Percentuale:**

  - Clicca sulla freccia accanto al campo nell'area **Valori**.
  - Seleziona **Mostra valori come** > **% del totale generale**.

---

## **Sezione 3: Funzionalità Avanzate delle Tabelle Pivot**

### **1. Utilizzo di Più Tabelle e Relazioni**

**Creazione di un Modello di Dati:**

- Importa le tabelle **Vendite** e **Prodotti** in Excel.
- Vai su **Dati** > **Relazioni**.
- Crea una relazione tra le tabelle basata sul campo **ID Prodotto**.

**Creazione della Tabella Pivot con il Modello di Dati:**

- Inserisci una nuova tabella pivot.
- Nella finestra di dialogo, seleziona **Aggiungi questi dati al modello di dati**.
- Ora puoi utilizzare campi da entrambe le tabelle nella tua tabella pivot.

### **2. Slicers e Timeline Avanzate**

**Inserimento di Slicers Multipli:**

- Seleziona la tabella pivot.
- Vai su **Analizza** > **Inserisci filtro dati**.
- Seleziona più campi, ad esempio **Team Venditore** e **Esperienza Venditore**.
- Posiziona gli slicers sul foglio e utilizza **Opzioni** per personalizzarli.

**Utilizzo della Timeline:**

- Seleziona la tabella pivot.
- Vai su **Analizza** > **Inserisci sequenza temporale**.
- Seleziona il campo **Data Vendita**.
- La timeline permette di filtrare i dati per anni, trimestri, mesi o giorni.

### **3. Aggiornamento Automatico della Tabella Pivot**

- Per aggiornare manualmente, clicca sulla tabella pivot e poi su **Analizza** > **Aggiorna**.
- Per impostare l'aggiornamento automatico:
  - Clicca con il tasto destro sulla tabella pivot.
  - Seleziona **Opzioni tabella pivot**.
  - Nella scheda **Dati**, seleziona **Aggiorna i dati all'apertura del file**.

---

## **Sezione 4: Grafici Pivot**

### **1. Creazione di un Grafico Pivot**

**Passaggi:**

1. Seleziona la tabella pivot.
2. Vai su **Analizza** > **Grafico Pivot**.
3. Scegli il tipo di grafico più appropriato (es. colonne, linee, torta).
4. Clicca su **OK**.

### **2. Personalizzazione del Grafico Pivot**

- **Modifica del Titolo:**

  - Clicca sul titolo del grafico e inserisci un titolo descrittivo.

- **Formattazione degli Assi:**

  - Clicca con il tasto destro sull'asse che vuoi formattare.
  - Seleziona **Formato asse** e modifica le opzioni desiderate.

- **Aggiunta di Linee di Tendenza:**

  - Clicca sulla serie di dati nel grafico.
  - Vai su **Progettazione** > **Aggiungi elemento grafico** > **Linea di tendenza**.

### **3. Interazione tra Tabella Pivot e Grafico Pivot**

- Ogni modifica nella tabella pivot si riflette automaticamente nel grafico pivot.
- Utilizzando gli slicers, puoi filtrare i dati sia nella tabella che nel grafico.

---

## **Esercitazione Pratica: Caso di Studio**

**Scenario:**

Un'azienda vuole analizzare le proprie vendite per ottimizzare le strategie di mercato.

**Attività Guidata:**

1. **Analisi delle Vendite per Categoria e Anno**

   - Crea una tabella pivot con:
     - **Righe**: Categoria
     - **Colonne**: Anno
     - **Valori**: Somma di Ricavo

2. **Identificazione dei Trend Stagionali**

   - Raggruppa le date per **Mese**.
   - Crea un grafico pivot per visualizzare le vendite mensili.

3. **Analisi delle Performance dei Venditori**

   - Crea una tabella pivot con:
     - **Righe**: Nome Venditore
     - **Valori**: Somma di Margine
   - Ordina i venditori dal più performante al meno performante.

4. **Esplorazione delle Correlazioni**

   - Aggiungi il campo **Esperienza Venditore** ai filtri.
   - Utilizza gli slicers per vedere come l'esperienza influisce sulle vendite.

5. **Creazione di un Dashboard Interattivo**

   - Dispone le tabelle pivot e i grafici in un unico foglio.
   - Aggiungi slicers e timeline per permettere un'analisi interattiva.

---

## **Consigli e Best Practice**

- **Pulizia dei Dati:** Prima di creare tabelle pivot, assicurati che i dati siano puliti e coerenti.
- **Nomi Significativi:** Rinomina i campi calcolati e le misure per riflettere chiaramente il loro contenuto.
- **Uso degli Slicers:** Gli slicers rendono i report più interattivi e facili da usare.
- **Formattazione Consistente:** Mantieni uno stile coerente nelle tabelle e nei grafici per una migliore leggibilità.
- **Salvataggio Frequente:** Con dataset grandi, Excel potrebbe rallentare; salva frequentemente per evitare perdite di dati.

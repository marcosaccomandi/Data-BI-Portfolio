<p align="center">
  <img src="./assets/banner.png" alt="Project Banner" width="100%"/>
</p>

<h1 align="center">Pulizia e Reporting Dati in Excel – Costi Fornitori 2024–2025</h1>

<p align="center">
  <b>Autore:</b> Marco Saccomandi  
  <br>
  <b>Progetto:</b> Analisi e Dashboard di Reporting dei Costi Fornitori  
  <br>
  <b>Strumenti:</b> Microsoft Excel · Power Query · Pulizia Dati · Tabelle Pivot  
</p>

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Panoramica
Caso reale di workflow Excel focalizzato su **pulizia, fusione e automazione dei report dei costi fornitori** per un consorzio edilizio.  
Diversi file Excel non strutturati sono stati consolidati in un’unica **Fact Table** che alimenta dashboard pivot dinamiche e report stampabili.

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Anteprima del Progetto
Una breve GIF mostra l’intero flusso di lavoro — dai dati grezzi ai dashboard finali in Excel.

<p align="center">
  <img src="./assets/project_overview.gif" alt="Workflow Preview" width="800"/>
</p>

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Sintesi del Workflow

### Fase 1 – Consolidamento dei Dati Grezzi
I dati provenivano da quattro file Excel indipendenti (Fatture, Fornitori, Costi Extra e Acquisti) senza chiavi comuni.  
Sono stati standardizzati e rinominati per consentire la fusione.

| Immagine | Descrizione |
|-----------|-------------|
| ![01](./assets/01_raw_file_overview.png) | Panoramica della cartella con i file originali Excel. |
| ![02](./assets/02_raw_invoices_2024_2025.png) | Dati fatture 2024–2025 con formati incoerenti. |
| ![03](./assets/03_raw_purchases_summary.png) | Riepilogo fornitori non strutturato. |
| ![04](./assets/04_raw_extra_costs.png) | Foglio dei costi extra (utenze, stipendi, tasse). |
| ![05](./assets/05_raw_secondary_2025.png) | Dataset aggiuntivo 2025 per validazione. |

---

### Fase 2 – Pulizia e Fusione dei Dati
Tutte le sorgenti sono state unite in una **Fact Table** contenente campi normalizzati:  
Anno, Categoria, Sottocategoria e Incidenza %.

| Immagine | Descrizione |
|-----------|-------------|
| ![06](./assets/06_first_merge_preview.png) | Primo test di fusione tra dataset costi e fatture. |
| ![07](./assets/07_fact_table_base.png) | Fact Table di base con campi standardizzati. |
| ![08](./assets/08_fact_table_refined.png) | Versione raffinata con classificazione e logiche di validazione. |

---

### Fase 3 – Reporting e Analisi
Sono state create tabelle pivot dinamiche per riepilogare e analizzare i dati di fornitori e centri di costo.

| Immagine | Descrizione |
|-----------|-------------|
| ![09](./assets/09_cost_center_report.png) | Dashboard di aggregazione per centro di costo. |
| ![10](./assets/10_supplier_detail_pivot.png) | Dettaglio fornitore per documento. |
| ![11](./assets/11_filtered_summary_report.png) | Esempio di report stampabile filtrato finale. |

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Strumenti e Tecniche
- **Microsoft Excel 365** – ambiente principale  
- **Power Query** – fusione e normalizzazione  
- **Tabelle Pivot** – dashboard automatizzate  
- **Formattazione condizionale** – evidenziazione e validazione  
- **Normalizzazione manuale** – pulizia categorie e fornitori  
- **Python (OpenPyXL)** – utilizzato per anonimizzare i nomi dei fornitori nel file finale

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Risultati
✔ 4 file Excel frammentati uniti in una Fact Table unica  
✔ Report pivot completamente automatizzati per fornitori e categorie  
✔ Struttura annuale standardizzata per il monitoraggio dei costi  
✔ Sessione di formazione di **2 ore** per il personale amministrativo  
✔ Maggiore autonomia per utenti non tecnici nella creazione dei report  

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Download del File Demo
Puoi esplorare la versione anonimizzata del file Excel utilizzato nel progetto qui:  
📂 [**Scarica SupplierTabRealCase.xlsx**](./assets/SupplierTabRealCase.xlsx)

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Nota sulla Privacy
Questo progetto è stato sviluppato con **dati operativi reali** di un consorzio edilizio.  
Tutti i nomi dei fornitori e le informazioni identificabili sono stati **anonimizzati** nella versione pubblica  
per garantire la conformità alle normative sulla privacy e la riservatezza dei dati.

---

**Autore:** Marco Saccomandi  
Progetto: *Analisi Costi Fornitori 2024–2025*  
Focus: *Pulizia Dati, Automazione dei Report e Formazione Utente*

---

**Lettura consigliata:**  
Consulta il [Case Summary](CASE_SUMMARY_IT.md) per contesto, azioni e risultati.

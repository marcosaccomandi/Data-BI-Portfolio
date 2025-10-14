<p align="center">
  <img src="./assets/banner.png" alt="Project Banner" width="100%"/>
</p>

<h1 align="center">Excel Data Cleaning & Reporting – Supplier Costs 2024–2025</h1>

<p align="center">
  <b>Author:</b> Marco Saccomandi  
  <br>
  <b>Project:</b> Supplier Cost Analysis and Reporting Dashboard  
  <br>
  <b>Tools:</b> Microsoft Excel · Power Query · Data Cleaning · Pivot Tables  
</p>

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Overview
A real-case Excel workflow focused on **cleaning, merging, and automating supplier cost reports** for a construction consortium.  
Multiple unstructured Excel files were consolidated into a single **Fact Table**, powering dynamic pivot dashboards and printable financial summaries.

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Project GIF Overview
A short GIF showcasing the full workflow — from raw data to final Excel dashboards.

<p align="center">
  <img src="./assets/project_overview.gif" alt="Workflow Preview" width="800"/>
</p>

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Workflow Summary

### Step 1 – Raw Data Consolidation
Raw data came from four independent Excel files (Invoices, Suppliers, Extra Costs, and Purchases) with no shared keys.  
They were standardized and renamed for merge.

| Image | Description |
|--------|--------------|
| ![01](./assets/01_raw_file_overview.png) | Folder overview with original Excel files. |
| ![02](./assets/02_raw_invoices_2024_2025.png) | Raw invoice data (2024–2025) with inconsistent formats. |
| ![03](./assets/03_raw_purchases_summary.png) | Unstructured supplier purchase summary. |
| ![04](./assets/04_raw_extra_costs.png) | Extra cost sheet (utilities, salaries, taxes). |
| ![05](./assets/05_raw_secondary_2025.png) | Additional 2025 dataset for validation. |

---

### Step 2 – Data Cleaning & Merge
All cleaned sources were merged into a single **Fact Table** containing normalized fields:  
Year, Category, Subcategory, and Incidence %.

| Image | Description |
|--------|--------------|
| ![06](./assets/06_first_merge_preview.png) | First merge test between cost and invoice datasets. |
| ![07](./assets/07_fact_table_base.png) | Base Fact Table with standardized fields. |
| ![08](./assets/08_fact_table_refined.png) | Refined version with cost classification and validation logic. |

---

### Step 3 – Reporting & Analysis
Dynamic pivot tables were implemented to summarize and analyze supplier and cost data.

| Image | Description |
|--------|--------------|
| ![09](./assets/09_cost_center_report.png) | Cost center aggregation dashboard. |
| ![10](./assets/10_supplier_detail_pivot.png) | Supplier detail breakdown by document. |
| ![11](./assets/11_filtered_summary_report.png) | Example of final printable summary report. |

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Download the Demo File
You can explore the anonymized Excel version used in this project here:  
📂 [**Download SupplierTabRealCase.xlsx**](./assets/SupplierTabRealCase.xlsx)

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Privacy Note
This project was built using **real operational data** from a construction consortium.  
All supplier names and identifiable information were **anonymized** in the public version  
to ensure full compliance with privacy and data protection requirements.

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Tools & Techniques
- **Microsoft Excel 365** – main environment  
- **Power Query** – merge and normalization  
- **Pivot Tables** – automated dashboards  
- **Conditional Formatting** – highlighting and validation  
- **Manual normalization** – supplier and category cleanup
- **Python (OpenPyXL)** – used to anonymize supplier names in the final Excel file
---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Results
✔ 4 fragmented Excel files merged into 1 unified Fact Table  
✔ Fully automated pivot reports for suppliers and categories  
✔ Standardized year-to-year structure for cost tracking  
✔ 2-hour training session provided to the administrative staff  
✔ Enabled autonomy for non-technical users in Excel report generation  

---

## <img src="./assets/section_icon_color.svg" alt="Project Icon" width="20"/> Author
**Marco Saccomandi**  
Project: *Supplier Cost Analysis 2024–2025*  
Focus: *Data Cleaning, Reporting Automation, and User Training*

---

**Further reading:**  
See the full [Case Summary](CASE_SUMMARY.md) for context, actions, and impact.

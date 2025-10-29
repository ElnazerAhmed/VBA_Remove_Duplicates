# VBA_Remove_Duplicates
excel sheet with duplicate entry for same date
# 📊 Copy and Fill Latest Modified Rows by Date (VBA)

## 🧩 Overview
This VBA macro automates the process of extracting and cleaning data from an Excel worksheet that contains multiple records per date.  
It finds the **latest modified row** for each unique date and intelligently fills in missing cells using values from earlier rows in the same date group.

✅ **Key features:**
- Groups data by the `Date_` column.  
- Picks the row with the **latest `Modified_Date`** per group.  
- Fills empty cells (columns **B–K**) from earlier non-empty rows.  
- Fills empty cells (columns **L–U**) using the most recent non-empty value based on modification date.  
- Outputs the cleaned result into a new sheet named **`Filtered_Latest_Modified`**.

---

## ⚙️ How It Works
1. Reads data from the **source sheet** (first sheet in your workbook).  
2. Groups all rows by the value in column **B** (`Date_`).  
3. Within each group:
   - Finds the record with the most recent value in column **D** (`Modified_Date`).
   - Uses that record as the “base” row.
   - Fills missing cells in columns:
     - **B–K** using the *earliest non-empty value* in the same group.
     - **L–U** using the *last modified non-empty value* in the same group.
4. Creates or clears a worksheet named **`Filtered_Latest_Modified`**.
5. Writes the header row and all processed data efficiently using array output.

---

## 🗂️ Example Use Case
Imagine you track daily updates to medical or operational data, where each date can appear multiple times with different modification times.  
You want to keep only the **latest version per date**, but also preserve important fields that may have been left blank in the final row.  
This macro automates that cleanup in seconds.

---

## 🧠 Core Logic
The macro uses:
- **`Scripting.Dictionary`** to group rows by date.
- **`System.Collections.ArrayList`** to manage grouped indices.
- **In-memory arrays** for fast reading/writing.
- **Fill rules** that distinguish between early (B–K) and late (L–U) columns.

---

## 🚀 How to Use
1. Open your Excel workbook.  
2. Press `Alt + F11` to open the **VBA Editor**.  
3. Insert a new module and paste the macro code.  
4. Modify these settings at the top of the code if needed:
   ```vba
   colStart = 2        ' Column B (start)
   colEnd = 21         ' Column U (end)
   dateCol = 2         ' Date_ column
   modCol = 4          ' Modified_Date column
   specialStart = 12   ' L–U range start
   specialEnd = 21     ' L–U range end


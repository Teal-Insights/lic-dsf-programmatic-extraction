## Input Instructions (Cleaned Markdown)

### 1) Go to the **Input 1 - Basics** sheet
1. **Country and template setup (cell C7):**  
   Select the country name from the dropdown for a new DSA using the previously downloaded template. Ensure country-specific information for CI and realism tools is refreshed.  
   Enter whether the country has an IMF program and IDA operations with IDA status. Country code, HIPC, and MDRI status are reflected automatically.

2. **Current year and first projection year (cells U17–18):**  
   Enter the “current year” and “first year of projection.”  
   These are often the DSA year and DSA vintage year, though the latter may depend on latest data availability. Current year and country name are used to auto-populate country-specific data for Composite Indicator (CI) and country risk scales.

3. **Scale selection (cell C19):**  
   Specify the scale of variables used in **Input 3 - Macro-Debt data**.

4. **Past DSA vintages for realism tools:**  
   Confirm the previous vintage and the vintage from 5 years ago. If missing, the template auto-populates the closest available vintages.

5. **Market-financing pressure inputs:**  
   - EMBI spread in **cells C28–29** is used for the market-financing pressure tool.  
   - If available (e.g., IMF Article IV), enter the real exchange rate overvaluation estimate in **cell C31** to set a stress test larger than the default shock one-time depreciation.

6. **Debt definition (cells C33–34):**  
   Choose external/domestic debt definition (**residency-based** or **currency-based**) from the dropdown.  
   The template then calculates public external and domestic debt based on this selected criterion.

---

### 2) Go to the **Input 2 - Debt Coverage** sheet
1. In **Table A, lines 7–15**, mark “X” when headline public debt data covers the relevant public-sector subsectors.
2. Confirm that **Table B, line 19** debt coverage description is consistent with Table A.
3. Tailored contingent liabilities shock size can be customized in **Table B, lines 21–34** (consistent with Table A).

---

### 3) Go to **Input 3 - Macro-Debt data** sheet
1. **Input data in all six sections:**  
   I. Projected data  
   II. Macroeconomic indicators  
   III. Old (existing) MLT external debt service  
   IV. MLT new external debt disbursements  
   V. Local-currency debt  
   VI. *(Residency-based DSAs only)*

2. Include **locally-issued debt** in sections II–IV of external debt inputs (regardless of currency denomination or creditor residency), since the template classifies it as local debt under the selected Input 1 criterion.

---

### 4) Go to **Input 4 - External Financing** sheet
1. Enter borrowing-term assumptions for:
   - new MLT PPG debt (**lines 9–38**)
   - ST debt (**line 41**)

2. Terms for locally-issued debt treated as external debt (based on selected criterion) are auto-populated using **Input 5 - Local-debt Financing** settings.

---

### 5) Go to **Input 5 - Local-debt Financing** sheet
1. Enter assumptions on borrowing terms for newly issued local-currency debt (**lines 8–22**).
2. Calculate public gross financing needs in **lines 39–63**.
3. Confirm and/or override **line 61–62** assumptions.  
   Debt services associated with IMF disbursements should not create additional borrowing if used for BoP (not budget) support.  
   If used for budget support, **line 62a** should be adjusted.
4. The template can handle up to five domestic debt instruments to fill remaining public GFNs after external financing.  
   If granular information is unavailable, use one representative instrument as an approximation.
5. Residual financing needs are assumed to be met by local currency-denominated short-term debt (**line 112**), which should be minimized and never negative (avoid over-financing).  
   If needed, use **line 61** to eliminate over-financing (deposit accumulation from debt issuance above projected GFNs).

---

### 6) Go to **Input 6 (optional) - Standard Test** sheet
1. Standardized stress tests apply to all countries.  
   Users are encouraged to change default settings, and with strong justification may override them in light yellow-shaded cells (**columns C and G**) in **Inputs D and H**.

---

### 7) Go to **Input 6 - Tailored Tests** sheet
1. Users are encouraged to customize default settings (**columns G and K**) for tailored tests in light yellow-shaded cells (**columns H and L**).

---

### 8) Go to **Input 7 - Residual Financing** sheet
1. Assumptions on marginal public borrowing under stress tests are populated from baseline financing terms.  
   For public DSA, additional financing can be allocated between external borrowing, domestic MLT, and ST borrowing.  
   For external DSA, additional financing needs are assumed to be filled by PPG external debt.
2. Interest rates are specified in:
   - nominal terms for external debt
   - real terms for domestic debt
3. Any customized debt service offset settings should be entered in yellow-shaded cells in **columns D and I**.

---

### 9) Go to **Input 8 - SDR** sheet
1. Enter total SDR allocation and total SDR holdings in **cells B6 and B7**.  
   If SDR allocation exceeds SDR holdings, the member pays interest on net SDR use (given obligation to reconstitute holdings).  
   PV calculations only consider future interest-payment flows.

---

### 10) Go to **Customized Scenarios** sheets *(optional)*
1. Design customized scenarios (for both external and public DSAs) by applying shocks to key economic variables affecting external and total public debt dynamics.  
   Select “Yes” in line 3 of the relevant worksheets to include these scenarios in output charts/tables.
2. Additional instructions are provided in:
   - **Customized scenarios-external**
   - **Customized scenarios-public**

---

### 11) Go to **Output 7 - Risk rating summary** sheet
1. Select final external and overall risk ratings, and enter any applied judgment in yellow cells.  
   Include in the Policy Note/Staff Report:
   - summary table
   - public sector coverage tables

---

## Note
Users will find useful intermediate calculation sheets (green-colored):  
- `Macro-Debt_Data`  
- `Ext_Debt_Data`  
- `PV_Base`  
- `Chart Data`  

These sheets contain many key indicator calculations.

## Reference
Background information:  
- *Staff Guidance Note on the Application of the Joint Fund-Bank Debt Sustainability Framework for Low-Income Countries, 2018*
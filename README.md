# ExSTELLA
Excel VBA macro for use with STELLA data

### **Overview**

This VBA module processes reflectance data collected from a STELLA spectrometer in Excel format.

* Accepts input from a worksheet containing raw\_counts and wavelength\_nm data.
* Prompts the user to select sets of white card and plant sample data.
* Computes average reflectance ratios (plant\_sample / white\_card).
* Outputs a summary table on a new sheet.
* Generates a chart showing reflectance ratios by wavelength.
* Supports processing multiple plant types in one session.

---

### **Global Settings**

These are set as Private variables:
* **`wavelengthMin` / `wavelengthMax`**: Defines the wavelength range to analyze (e.g., 410–940 nm).
* **`numSamples`**: Number of rows per data set (e.g., 18).
* **`hasHeader`**: Boolean to indicate if source sheet includes a header row.

---

### **Procedures & Functions**

#### **Sub `ProcessSTELLAData()`**

Main procedure that runs the workflow:

1. Prompts for a worksheet and plant name.
2. Collects sample ranges for white cards and plant samples.
3. Calculates and outputs average reflectance ratios.
4. Loops to allow adding multiple plant sample sets.
5. Generates a chart of reflectance ratios vs. wavelength.

---

### **Expected Input Sheet Structure**

* **Column H**: `wavelength_nm`
* **Column K**: `raw_counts`
* Data is expected in blocks of 18 rows per sample set.

---

### **Output**

* A new worksheet is added with:

  * Column A: Wavelength
  * Column B+: Reflectance ratios for each plant sample set
* A scatterplot chart is generated on the same sheet.

---

### **Error Handling**

* Prompts and warnings are provided if data is missing or selection criteria are not met.

---

#### **Sub `SetVars()`**

Initializes module-wide settings such as wavelength range and number of samples.

---

#### **Sub `DeleteEmptyRows()`**

Removes completely empty rows from the active worksheet.

---

#### **Function `PromptForWorksheet()` → Worksheet**

Prompts the user to select one worksheet from the open workbook. Returns a `Worksheet` object or `Nothing` if canceled.

---

#### **Function `AverageFromCollection(col As Collection)` → Double**

Calculates the average of numeric values in a `Collection`.

---

#### **Sub `HighlightSets()`**

Highlights rows in the active worksheet where the value in column H equals the minimum wavelength, marking the start of a data block.

---

#### **Function `SortDictionaryKeys(dict As Object)` → Variant**

Returns the keys of a dictionary sorted numerically. (Currently unused.)

---

#### **Function `CollectRanges(setName As String)` → Collection**

Prompts the user to select one or more sample ranges (e.g., White Cards or Plant Samples). Each selected range must have exactly `numSamples` rows.

---

#### **Function `CleanDictionary(dict As Object)` → Object**

Resets all entries in the dictionary by assigning each key a new collection containing two empty sub-collections. This preserves the existing keys (wavelengths) and allows a new white card or plant sample to be added.

A new empty collection is added for:
1. White Card readings
2. Plant Sample readings
   Returns the updated dictionary.

---

#### **Function `AddToWavelengthDict(dict, sampleRange, rangeName, index, ws)` → Object**

Adds raw\_counts data to the dictionary for each wavelength in the selected range.

If a wavelength is not already in the dictionary, it is added with two empty collections.
* `index = 1` for white cards
* `index = 2` for plant samples  

---



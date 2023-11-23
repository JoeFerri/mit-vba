## Script in VBA (Visual Basic for Applications) for inserting alternate blank rows into an Excel worksheet.

**Table of Contents**

[TOCM]

------------
### Example FILE

The *MACRO_RVA_EXAMPLE.xlsm* file consists of 4 sheets:
- **DATA_RAW**: Used to generate sample random data
- **DATA_SET**: used to structure the sample to be processed with the macro
- **RESULT**: Shows the result of executing the macro
- **OPTIONS**: contains the data used by the macro

### Example HOW TO
Copy the data from the **DATA_SET** sheet and paste it into the **RESULT** sheet, change the values in the **OPTIONS** sheet to get different results, then start `MACRO_RVA()`.

------------
### Composition and Content

The Script is made up of two modules including a macro called `MACRO_RVA()` to be executed in the chosen sheet.

The **INITIAL ROW** and **NUMBER OF ROWS** values to be alternated must be entered respectively in cells **D4** and **D6** of a sheet called **OPTIONS**
(*except for modifying the code to use different cells or inserting these values as internal constants of the modules*).

### Execution

To run the macro go to the top menu **bar -> View -> Macro -> View Macro -> MACRO_RVA [Run]**
Or show the develop menu via **File -> Options -> Customize Ribbon -> [Develop]**


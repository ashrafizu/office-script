// THIS IS A SAMPLE OF HOW YOU CAN MASSAGE DATA ONCE YOU HAVE AN EXISTING TABLE INSIDE YOUR EXCEL WORKSHEET
// REFER TO COMMENTS UNDER EACH SECTIONS TO UNDERSTAND FURTHER


// THIS IS THE MAIN FUNCTION IN WHICH SCRIPT WILL RUN
function main(workbook: ExcelScript.Workbook) {
  // Get File Name
  let fileName = workbook.getName();
  // Get the worksheet
  let sheet = workbook.getActiveWorksheet();


  // For Collection Report - assuming the table start at range "A5"
  if (fileName.includes('Collection')) {
    let table = sheet.getTable("Collection");
    // get All Column Names -- THIS IS TO CHANGE ALL THE COLUMN CAPITALIZATION TO UPPERCASE
    let columns = table.getHeaderRowRange().getValues()[0].map(column => column.toString().toUpperCase());

    // THIS IS TO REPLACE THE COLUMN NAMES BY SUBSTITUTING ANY BLANK SPACE EXISTED INSIDE THE COLUMN NAMING WITH AN UNDERSCORE --> THIS SIGN "_"
    for (let col of columns) {
      let attached = col.replace(/ /g, '_'); // I CALLED IT "ATTACHED" BECAUSE IT'S BASICALLY ATTACHING ALL THE COLUMN NAMES TO MAKE THEM STAND LIKE ONE-WORD NAMES
      table.getColumnByName(col).setName(attached);
    };

    // I RE-DECLARED THIS VARIABLE TO CAPTURE NEW COLUMN NAMES WITH UNDERSCORES
    columns = table.getHeaderRowRange().getValues()[0].map(column => column.toString().toUpperCase());

    let caseColumns: string[] = [];
    let tarpColumns: string[] = [];

    // THIS IS TO CREATE AN ARRAY OF COLUMNS BASED ON CATEGORIES - IN THIS CASE "CASE COLUMNS" AND "TARP COLUMNS"
    for (let col of columns) {
        if (col.includes('CASE') && !col.includes('YTD') && !col.includes('SALES')) {
            caseColumns.push(col)
        };
        if (col.includes('TARP') && !col.includes('YTD') && !col.includes('SALES')) {
            tarpColumns.push(col)
        };
    };

    // THIS IS TO ADD ADDITIONAL COLUMNS
    let ytdCase = table.addColumn();
    ytdCase.setName('YTD_CASE');
    let ytdTARP = table.addColumn();
    ytdTARP.setName('YTD_TARP');

    // THIS IS TO ADD IN COLUMNS THAT ARE NEEDED IN AN EXCEL FORMULA - THIS IS A METHOD WHERE I GO THROUGH ONE BY ONE OF THE COLUMNS TO BE IMPOSED IN A SUMMATION FORMULA
    let strCaseFormula: string = "= 0 "; // I STARTED WITH ZERO TO LET AS THE INITIAL STRING BEFORE THE LOOP WITHIN THE ARRAY STARTS
    for (let colCase of caseColumns) {
      strCaseFormula = `${strCaseFormula} + [@[${colCase}]]`
    };

    // THIS IS TO ADD THE FORMULA TO THE INTENDED COLUMN ONCE THE FORMULA HAS BEEN FORMULATED
    table.getColumnByName('YTD_CASE').getRangeBetweenHeaderAndTotal().setFormula(strCaseFormula);

    // THIS IS TO COPY AND PASTE VALUES WITHOUT CAPTURING THE FORMULA -- I DONT WANT TO KEEP THE FORMULA, JUST VALUES
    table.getColumnByName('YTD_CASE').getRangeBetweenHeaderAndTotal().copyFrom(table.getColumnByName('YTD_CASE').getRangeBetweenHeaderAndTotal(), ExcelScript.RangeCopyType.values, false /*to Skip Blanks*/, false /*to Transpose*/);

    // THIS IS TO ADD IN COLUMNS THAT ARE NEEDED IN AN EXCEL FORMULA - THIS IS A METHOD WHERE I GO THROUGH ONE BY ONE OF THE COLUMNS TO BE IMPOSED IN A SUMMATION FORMULA
    let strTarpFormula: string = "= 0 "; // I STARTED WITH ZERO TO LET AS THE INITIAL STRING BEFORE THE LOOP WITHIN THE ARRAY STARTS
    for (let colTarp of tarpColumns) {
      strTarpFormula = `${strTarpFormula} + [@[${colTarp}]]`
    };

    // THIS IS TO ADD THE FORMULA TO THE INTENDED COLUMN ONCE THE FORMULA HAS BEEN FORMULATED
    table.getColumnByName('YTD_TARP').getRangeBetweenHeaderAndTotal().setFormula(strTarpFormula);

    // THIS IS TO COPY AND PASTE VALUES WITHOUT CAPTURING THE FORMULA -- I DONT WANT TO KEEP THE FORMULA, JUST VALUES
    table.getColumnByName('YTD_TARP').getRangeBetweenHeaderAndTotal().copyFrom(table.getColumnByName('YTD_TARP').getRangeBetweenHeaderAndTotal(), ExcelScript.RangeCopyType.values, false /*to Skip Blanks*/, false /*to Transpose*/);

  } else if (fileName.includes('Listing')) {
    let table = sheet.getTable("Listing");
    // get All Column Names
    let columns = table.getHeaderRowRange().getValues()[0].map(column => column.toString().toUpperCase());

    for (let col of columns) {
      // Rename of Certain Columns for Standardization (Applicable for Senheng and VADS)
      if (col.toUpperCase().startsWith('RATE PRORATED TARP')) {
        table.getColumnByName(col).setName('PRORATED_TARP_RATE');
      } else if (col.toUpperCase().startsWith('UPDATE DATE')) {
        table.getColumnByName(col).setName('UPDATE_DT');
      } else {
      let attached = col.replace(/ /g, '_');
      table.getColumnByName(col).setName(attached);
      };
    };

    columns = table.getHeaderRowRange().getValues()[0].map(column => column.toString().toUpperCase());

  };

};

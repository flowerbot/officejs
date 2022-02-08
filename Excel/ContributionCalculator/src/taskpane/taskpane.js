/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
     // const range = context.workbook.getSelectedRange();

      // Read the range address
     // range.load("address");

      // Update the fill color
    //  range.format.fill.color = "yellow";

    //  await context.sync();
    //  console.log(`The range address was ${range.address}.`);


        // Get the worksheet named "Sheet1".
    //const sheet = workbook.getWorksheet(sheetName);
  
    //nb to be run on the 'active' sheet which should be a 'Stage' sheet containing rates
    const rateSheet = context.workbook.getActiveWorksheet();
    const sheetN = rateSheet.getName();
    const rateTable = rateSheet.getTables()[0];
  
    // Get the entire data range.
    //const range = sheet.getUsedRange(true);
  
    const range = rateTable.getRange();
  
    //first clear any filters that may already be there
    let cols = rateTable.getColumns();
  
    for (let i = 0; i < cols.length; i++) {
      cols[i].getFilter().clear();
    }
  
    let visible = rateTable.getRangeBetweenHeaderAndTotal().getVisibleView();
  
    visible.getRows().forEach((r, i) => {
      r.getRange().getColumn(7).clear();
    })
  
  
  
    //then get values from Start Here
    // Get the worksheet named "Start Here ".
    const wsStartHere = workbook.getWorksheet("Start Here");
  
    const DevTypeTable = wsStartHere.getTable("TableDevType");
    const LocalPlanTable = wsStartHere.getTable("Table4");
    const DwellingsTable = wsStartHere.getTable("Table6");
    const CreditsTable = wsStartHere.getTable("Table3");
    const extrasTable = wsStartHere.getTable("Table9");
    const wwwTable = wsStartHere.getTable("Table7");
  
    //console.log(DwellingsTable);
    //console.log(LocalPlanTable);
  
  
    // handle local area plan
  
    //const lapCols: string[] = LocalPlanTable.getColumnByName("Include").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
    const lapCols = LocalPlanTable.getColumnByName("Include").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0]);
  
    //console.log("lapCols:", lapCols);
  
    //const lapPlanCols: string[] = LocalPlanTable.getColumnByName("CP").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
    const lapPlanCols = LocalPlanTable.getColumnByName("CP").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0]);
    /*
      const wwwCols = wwwTable.getColumnByName("Include").getRangeBetweenHeaderAndTotal().getValues().filter((v, i, a) => v[0] as string == "Yes" 
      );
    */
    //const wwwColsArray = wwwTable.getRangeBetweenHeaderAndTotal().getValues().filter((v, i, a) => a[i][1] as string == "Yes"
    const wwwColsArray = wwwTable.getRangeBetweenHeaderAndTotal().getValues().filter((v, i, a) => a[i][1]  == "Yes"
    );
  
    //const wwwCols: string[] = [];
    const wwwCols = [];
    wwwColsArray.forEach(v => wwwCols.push(String(v[0])));
  
    console.log("wwwCols:", wwwCols);
  
    //let bumpCP: string[] = [];
    //let hideCP: string[] = [];

    let bumpCP = [];
    let hideCP = [];
  
    for (let i = 0; i < lapCols.length; i++) {
  
      let planToPush = lapPlanCols[i];
  
      if (lapCols[i] == "Yes") {
  
        bumpCP.push(planToPush);
        //should only ever be 1, if we get one, bail out
        // break;  //cancelled break because also want the 'other' plans
  
      } else {
        hideCP.push(planToPush);
      }
    }
  
  
    //let incWWW: string[] = [];
    let incWWW = [];
    //for(let i=0; i<)
  
    console.log("bumpCP:", bumpCP);
    console.log("hideCP:", hideCP);
  
    if (sheetN.indexOf("Stage") >= 0) {
  
  
  
      //*********************** */
      // make Group filter (works)
  
      //const groups: string[] = rateTable.getColumnByName("Group").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
      const groups = rateTable.getColumnByName("Group").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] );

     // let groupFilterArray: string[] = [];
      let groupFilterArray = [];
  
      //include all except active water/sewer
      groups.forEach(v => {
        if (!(hideCP.includes(String(v))) && !(String(v).indexOf("Water") >= 0) && !(String(v).indexOf("Sewer") >= 0)) {
          groupFilterArray.push(v);
        }
      });
  
      //now add back any water/sewer selected
      wwwCols.forEach(v => groupFilterArray.push(v));
  
      //remove duplicates
      groupFilterArray = groupFilterArray.filter((value, index, self) => {
        return self.indexOf(value) === index;
      }
      );
  
      console.log("groupFilterArray:", groupFilterArray);
  
      const groupCol = rateTable.getColumnByName("Group");
  
      groupCol.getFilter().applyValuesFilter(groupFilterArray);
  
  
  
      // get dwellings and trips
      /* const incDwellings: string[] = DwellingsTable.getColumnByName("Dwellings").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string); */
  
      const totalDwellings = DwellingsTable.getTotalRowRange().getValues();
  
      //1: Dwellings ...totalDwellings[0][1]
      //2: Trips ... totalDwellings[0][2]
      //3:  TotalEts ..totalDwellings[0][3]
  
      console.log(totalDwellings[0][2]);
  
      const totalCredits = CreditsTable.getColumnByName("Number").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0]);
  
      console.log("totalCredits:", totalCredits);
  
      const totalETs = Number(totalDwellings[0][3]) - Number(totalCredits[0]);
      const totalTrips = Number(totalDwellings[0][2]) - Number(totalCredits[1]);
  
      console.log("totalEts:", totalETs);
      console.log("totalTrips:", totalTrips);
  
  
     // let excludeRows: string[] = [];
      let excludeRows = [];
  
      //only do this if on a stage sheet
  
  
      //choose only current items
      const currentColNum = 34;
      const currentCol = rateTable.getColumns()[currentColNum];
      //  console.log(currentCol);
      currentCol.getFilter().applyValuesFilter(["Yes"]);
  
      //handle bumped items
  
      // is this going to cause problems?
      //if (bumpCP.length > 0) {
  
        let myBumpedCP = bumpCP[0];
  
        const bumpedRowRange = rateTable.getRangeBetweenHeaderAndTotal().getValues().forEach((rItem, rIndex) => {
          // console.log(rItem);
          if (rItem[0] == myBumpedCP) {
            excludeRows.push(String(rItem[33]));
          }
        });
  
        console.log("excludeRows:", excludeRows);
  
  
        //************ get TRCP Sector, make filter on prefix  */
  
        const trcpSector = extrasTable.getColumnByName("Value").getRangeBetweenHeaderAndTotal().getValues()[0][0];
  
        const trcpLCA = extrasTable.getColumnByName("Value").getRangeBetweenHeaderAndTotal().getValues()[1][0];
  
        const CP23 = extrasTable.getColumnByName("Value").getRangeBetweenHeaderAndTotal().getValues()[2][0];
  
        console.log("trcpLCA:", trcpLCA);
  
        //if(trcpLCA) trcpSector.push(String(trcpLCA));
  
        console.log("trcpSector:", trcpSector);
  
        const preCol = rateTable.getColumnByName("prefix");
  
       // const prefixFilterArray: string[] = [];
        const prefixFilterArray = [];
  
        console.log("Is it getting here when empty");
  
        preCol.getRangeBetweenHeaderAndTotal().getValues().forEach(v => {
  
          console.log(v, String(v[0]).indexOf("CP04"));
          if ((!(String(v[0]).indexOf("CP04") >= 0) && !(String(v[0]).indexOf("CP23") >= 0))) {
            prefixFilterArray.push(String(v[0]));
          }
        });
        prefixFilterArray.push(String(trcpSector));
        if (trcpLCA) prefixFilterArray.push(String(trcpLCA));
  
        if (CP23) prefixFilterArray.push(String(CP23));
  
        console.log("CP23:", CP23);
  
        console.log(prefixFilterArray);
  
        preCol.getFilter().applyValuesFilter(prefixFilterArray);
  
  
        //************************* */
        // make charge_type Filter
  
        //const chargeTypes: string[] = rateTable.getColumnByName("charge_type").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
        const chargeTypes = rateTable.getColumnByName("charge_type").getRangeBetweenHeaderAndTotal().getValues().map(v => v[0]);
  
       // let ctFilterArray: string[] = [];
        let ctFilterArray = [];
  
        chargeTypes.forEach(v => {
  
          if (!(excludeRows.includes(String(v)))) {
            ctFilterArray.push(v);
          }
        });
  
        /* if(String(trcpSector) != "") {
          ctFilterArray.push(String(trcpSector));
        } */
  
        ctFilterArray = ctFilterArray.filter(
          (value, index, self) => {
            return self.indexOf(value) === index;
          });
  
        console.log("ctFilterArray:", ctFilterArray);
  
  
  
        const ctCol = rateTable.getColumnByName("charge_type");
        ctCol.getFilter().applyValuesFilter(ctFilterArray);
  
        //*****************END OF charge_type filter ******************** */
  
  
  
  
        //**************************ADD NUMBERS TO VISIBLE ETs/Trips/Ha COLUMN ... */
  
  
        const visibleRates = rateTable.getRangeBetweenHeaderAndTotal().getVisibleView();
  
        //console.log('visibleRates:', visibleRates.getRows());
  
        visibleRates.getRows().forEach((r, i) => {
          // charge_comment value
          let cComment = r.getRange().getColumn(6).getValues()[0][0];
          console.log("cComment:", cComment);
          if (cComment == "Per ET") r.getRange().getColumn(7).setValue(totalETs);
          if (cComment == "Trip ends incl admin") r.getRange().getColumn(7).setValue(totalTrips);
        })
  
     // }
  
  
  
  
  
  
  
  
  
      // If the used range is empty, end the script.
      if (!range) {
        console.log(`No data on this sheet.`);
        return;
      }
  
    } else {
  
      console.log("********* NOT A STAGE SHEET ************");
    }
  
    /*
      // Log the address of the used range.
      console.log(`Used range for the worksheet: ${range.getAddress()}`);
    
      // Look through the values in the range for blank rows.
      const values = range.getValues();
      let emptyRows = 0;
      for (let row of values) {
        let emptyRow = true;
    
        // Look at every cell in the row for one with a value.
        for (let cell of row) {
          if (cell.toString().length > 0) {
            emptyRow = false
          }
        }
    
        // If no cell had a value, the row is empty.
        if (emptyRow) {
          emptyRows++;
        }
      }
    
      // Log the number of empty rows.
      console.log(`Total empty rows: ${emptyRows}`);
    
      // Return the number of empty rows for use in a Power Automate flow.
      return emptyRows;
      */



    });
  } catch (error) {
    console.error(error);
  }
}

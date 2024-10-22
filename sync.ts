function main(workbook: ExcelScript.Workbook) {
  

  /*==================*/
  //CHANGE VALUES HERE//
  const planSheetName = "PLAN LIST";
  const allSheetName = "ALL LIST";
  const quarterSheetName = "QUARTER LIST";
  /*let columnsToSync = ["Summary", "Priority", "Status", "Sprint", "Epic Link"];*/
  /*==================*/

  //declaration
  const startTime = new Date();

  //copy latest info from ALL SHEET to PLAN SHEET
  syncSheet(workbook.getWorksheet(allSheetName), workbook.getWorksheet(planSheetName), "PLAN LIST?");

  //copy latest info from PLAN SHEET to QUARTER SHEET
  syncSheet(workbook.getWorksheet(planSheetName), workbook.geatWorksheet(quarterSheetName), "QUARTER LIST?");

    
  

  //set metrics
  const endTime = new Date();
  console.log('TIME (seconds):' + Math.abs(startTime.getTime() - endTime.getTime()) / 1000);

}

function syncSheet(allSheet: ExcelScript.Worksheet, planSheet: ExcelScript.Worksheet, fName: string) {

  //declaration
  const planUsedRange = planSheet.getUsedRange();
  const planUsedColumns = planUsedRange.getColumnCount();
  const planUsedRows = planUsedRange.getRowCount();

  const allUsedRange = allSheet.getUsedRange();
  const allUsedColumns = allUsedRange.getColumnCount();
  const allUsedRows = allUsedRange.getRowCount();

  const columnsToSkip: Array<string> = ["ISSUE KEY", fName];
  fName = fName.toUpperCase();

  //get mapping of plan columns to all and quarter columns
  var columnsToSync: Array<string> = planSheet.getRangeByIndexes(0, 0, 1, planUsedColumns).getValues()[0].map(arr => arr.toString());
  columnsToSync = columnsToSync.map(arr => arr.toUpperCase().replace("Custom field (".toUpperCase(), "").replace(")", ""));
  type mapObj = {
    planColumn: number,
    allColumn: Array<number>
  };
  var columnMap = new Map<string,
    mapObj>();
  var pColumns = planSheet.getRangeByIndexes(0, 0, 1, planUsedColumns).getValues()[0].map(arr => arr.toString());
  var aColumns = allSheet.getRangeByIndexes(0, 0, 1, allUsedColumns).getValues()[0].map(arr => arr.toString());

  aColumns = aColumns.map(arr => arr.toString().toUpperCase().replace("Custom field (".toUpperCase(), "").replace(")", ""));
  pColumns = pColumns.map(arr => arr.toString().toUpperCase().replace("Custom field (".toUpperCase(), "").replace(")", ""));
  const combColumns = aColumns.concat(pColumns).reduce((a:Array<string>, b:string) => {
    if (b && b.length >0 && a.indexOf(b) < 0) a.push(b);
    return a;
  }, []);

  for (var i = 0; i < combColumns.length; i++) {
    if (combColumns[i] && combColumns[i].toString().length > 0) {
      
      columnMap.set(
        combColumns[i].toString(), {
        planColumn: getAllIndexes(pColumns, combColumns[i])[0],
        allColumn: getAllIndexes(aColumns, combColumns[i])
      });

    }
  }

  //validate that ISSUE KEY column exists. Then get all issues from both lists.
  const issueKeyObj = columnMap.get("Issue Key".toUpperCase());
  if (issueKeyObj && issueKeyObj.allColumn.length == 1 && issueKeyObj.planColumn >= 0) {
    console.log('Issue Key column found in both sheets');
  } else {
    console.log('Missing Issue Key column');
    return;
  }

  let allJiraKeys = flattenArray(allSheet.getRangeByIndexes(0, issueKeyObj.allColumn[0], allUsedRows, 1).getValues());
  let planJiraKeys = flattenArray(planSheet.getRangeByIndexes(0, issueKeyObj.planColumn, planUsedRows, 1).getValues());

  //find or create FOUND columns in source sheet
  let foundColumn = columnMap.get(fName);
  if (!foundColumn) {
    columnMap.set(fName, {
      allColumn: [],
      planColumn: -1
    });
    foundColumn = columnMap.get(fName);
  }

  if (foundColumn.allColumn.length == 0) {
    allSheet.getCell(0, allUsedColumns).setValue(fName);
    foundColumn = {
      allColumn: [allUsedColumns],
      planColumn: -1
    };
    columnMap.set(fName, foundColumn);
    console.log('CREATED FOUND COLUMN');
  }

  //reset the FOUND column
  allSheet.getRangeByIndexes(1, foundColumn.allColumn[0], allUsedRows, 1).setValue(null);

  //remove conditional formatting from the entire sheet
  planSheet.getRange().clearAllConditionalFormats();
  allSheet.getRange().clearAllConditionalFormats();  

  //go through ALL ITEMS and update PLAN ITEMS & QUARTER LIST
  console.log("go through ALL ITEMS and update PLAN ITEMS & QUARTER LIST");
  for (var allRowIndex = 1; allRowIndex < allJiraKeys.length; allRowIndex++) {

    let jiraKey = allJiraKeys[allRowIndex];
    if (!jiraKey.toString().startsWith('ABC')){
      continue;
    }

    //----locate item in plan list, and update values if needed
    let planRowIndex = planJiraKeys.indexOf(jiraKey);
    let color = '';
    if (planRowIndex > 0) {

      //----set FOUND/NOTFOUND. And color of background in ALL LIST.
      allSheet.getCell(allRowIndex, foundColumn.allColumn[0]).setValue('FOUND');
      //planSheet.getCell(planRowIndex, foundColumn.planColumn).setValue('FOUND');

      //----update PLAN LIST row for the colums we want to sync
      for (var j = 0; j < columnsToSync.length; j++) {
        let syncColumn = columnMap.get(columnsToSync[j]);
        if (syncColumn) {
          let allColumnIndex = columnMap.get(columnsToSync[j]).allColumn;
          let planColumnIndex = columnMap.get(columnsToSync[j]).planColumn;

          //----copy data
          if (columnsToSkip.indexOf(columnsToSync[j]) == -1 && allColumnIndex.length > 0 && planColumnIndex >= 0) {
            let aValues: Array<string> = new Array();
            for (let k = 0; k < allColumnIndex.length; k++) {
              aValues.push(allSheet.getCell(allRowIndex, allColumnIndex[k]).getValue().toString());
            }

            let valueToCopy = aValues.filter(n => n).join(", ");
            if (valueToCopy !== planSheet.getCell(planRowIndex, planColumnIndex).getValue()) {
              planSheet.getCell(planRowIndex, planColumnIndex).setValue(valueToCopy);
            }
          }
        } else {
          //console.log("Can't find column " + columnsToSync[j]);
        }

      }
    } else {
      allSheet.getCell(allRowIndex, foundColumn.allColumn[0]).setValue('NOT FOUND');
    }
  }

  //go through plan ITEMS and do final formatting/data operations
  console.log("go through plan ITEMS and do final formatting/data operations");
  //-- set URL if it's missing
  var column = columnMap.get("URL".toUpperCase());
  if (column && column.planColumn >= 0) {
    for (let i = 0; i < planJiraKeys.length; i++) {
      if (planJiraKeys[i].toString().startsWith('ABC')) {
        let link = `https://jira.yourdomain.com/browse/${planJiraKeys[i].toString()}`;
        planSheet.getCell(i, column.planColumn).clear();
        planSheet.getCell(i, column.planColumn).setValue(link);
      }
    }
  }

  //--strikethrough resolved and closed items
  console.log("strikethrough resolved and closed items");
  var column:mapObj = columnMap.get("STATUS".toUpperCase());
  if (column.planColumn >= 0) {
    const statusesToMatch = ["Closed", "Resolved", "Ready for Test"].map(arr => arr.toString().toUpperCase());
    for (let i = 1; i < planUsedRows; i++) {
      var status = planSheet.getCell(i, column.planColumn).getValue().toString().toUpperCase();
      if (statusesToMatch.indexOf(status) >= 0) {
        planSheet.getCell(i, column.planColumn).getEntireRow().getFormat().getFont().setStrikethrough(true);
      }
    }
  }

  //--set TO DO on .PLAN if it has no value
  console.log("set TO DO on .PLAN if it has no value");
  var column = columnMap.get(".PLAN".toUpperCase());
  if (column && column.planColumn >= 0) {
    for (let i = 0; i < planJiraKeys.length; i++) {
      if (planJiraKeys[i].toString().startsWith('ABC') && !planSheet.getCell(i, column.planColumn).getValue()) {
        planSheet.getCell(i, column.planColumn).clear();
        planSheet.getCell(i, column.planColumn).setValue('TO DO');
      }
    }
  }

  //--set TO DO on .REQS if it has no value
  console.log("set TO DO on .REQS if it has no value");
  var column = columnMap.get(".REQS".toUpperCase());
  if (column && column.planColumn >= 0) {
    for (let i = 0; i < planJiraKeys.length; i++) {
      if (planJiraKeys[i].toString().startsWith('ABC') && !planSheet.getCell(i, column.planColumn).getValue()) {
        planSheet.getCell(i, column.planColumn).clear();
        planSheet.getCell(i, column.planColumn).setValue('TO DO');
      }
    }
  }
  
  //create conditional formatting on certain columns
  //--issue key (search for duplicates)
  var column = columnMap.get("Issue Key".toUpperCase());
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    addDuplicateCondition(columnRange);

  }

  //--status
  var column = columnMap.get("Status".toUpperCase());
  if (column){
    let columnRange: Array<ExcelScript.Range>  = [];
    if (column.planColumn >= 0){
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0){
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    addColorCondition(columnRange, "Open", "#C6EFCE", "#006100");
    addColorCondition(columnRange, "In Progress", "#FFEB9C", "#9c0006");
    addColorCondition(columnRange, "Closed", "#ffc7ce", "#9c0006");
    addColorCondition(columnRange, "Resolved", "#ffc7ce", "#9c0006");
    addColorCondition(columnRange, "Ready for Test", "#ffc7ce", "#9c0006");

  }

  //--priority
  var column = columnMap.get("Priority".toUpperCase());
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0) {
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    addColorCondition(columnRange, "Critical", "#ffc7ce", "#9c0006");
    addColorCondition(columnRange, "Major", "#FFEB9C", "#9C5700");
    addColorCondition(columnRange, "Minor", "#D9E1F2", "#002060");
    addColorCondition(columnRange, "Trivial", "#D9E1F2", "#002060");
  }

  //--issue type
  var column = columnMap.get("Issue Type".toUpperCase());
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0) {
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    addColorCondition(columnRange, "Epic", "#ffc7ce", "#9c0006");
  }

  //--found / not found
  var column = foundColumn;
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0) {
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    addColorCondition(columnRange, "NOT FOUND", "#ffc7ce", "#9c0006");
    addColorCondition(columnRange, "FOUND", "#C6EFCE", "#006100");
  }

  //--requirements
  var column = columnMap.get(".Reqs".toUpperCase());
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0) {
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }
    
    removeConditionsInColumns(columnRange);
    addColorCondition(columnRange, "TO DO", "#ffc7ce", "#9c0006"); 
    addColorCondition(columnRange, "READY", "#C6EFCE", "#006100");
  }

  //--quarter
  var column = columnMap.get(".PLAN".toUpperCase());
  if (column) {
    let columnRange: Array<ExcelScript.Range> = [];
    if (column.planColumn >= 0) {
      columnRange.push(planSheet.getCell(1, column.planColumn).getEntireColumn());
    }
    if (column.allColumn.length > 0) {
      columnRange.push(allSheet.getCell(1, column.allColumn[0]).getEntireColumn());
    }

    removeConditionsInColumns(columnRange);
    let currentQuarter = '23-Q4';
    let nextQuarter = '24-Q1';
    let nextNextQuarter = '24-Q2';

    addColorCondition(columnRange, currentQuarter, "#FFEB9C", "#9C5700");
    addColorCondition(columnRange, nextQuarter, "#C6EFCE", "#006100");
    addColorCondition(columnRange, nextNextQuarter, "#bfb6f2", "#403394");
    addColorCondition(columnRange, "LATER", "#EDEDED", "#595959");
    addColorCondition(columnRange, "N/A", "#EDEDED", "#595959");
    addColorCondition(columnRange, "TO DO", "#ffc7ce", "#9c0006");
  }

  //set standard font for the entire sheet
  planSheet.getRange().getFormat().getFont().setName("Arial");
  allSheet.getRange().getFormat().getFont().setName("Arial");

}

function flattenArray(arr: Array<Array<string | number | boolean>>): Array<string | number | boolean> {
  let arrFlat: Array<string | number | boolean> = new Array();
  for (var i = 0; i < arr.length; i++) {
    arrFlat.push(arr[i][0]);
  }
  return arrFlat;
}

function getAllIndexes(arr: Array<string | number | boolean>, val: string): Array<number> {
  let indexes: Array<number> = new Array();
  for (var i = 0; i < arr.length; i++)
    if (arr[i] && arr[i].toString().length > 0 && arr[i].toString() === val.toString()) {
      indexes.push(i);
    }
  return indexes;
}

function removeConditionsInColumns(ranges: Array<ExcelScript.Range>) {
  for (var i = 0; i < ranges.length; i++){
    ranges[i].clearAllConditionalFormats();
  }
}

function addColorCondition(ranges: Array<ExcelScript.Range>, text: string, bgColor: string, fontColor: string) {
  let conditionalFormatting: ExcelScript.ConditionalFormat;
  for (var i = 0; i < ranges.length; i++) {
    conditionalFormatting = ranges[i].addConditionalFormat(ExcelScript.ConditionalFormatType.containsText);
    conditionalFormatting.getTextComparison().setRule({
      operator: ExcelScript.ConditionalTextOperator.contains,
      text: text
    });
    conditionalFormatting.getTextComparison().getFormat().getFill().setColor(bgColor);
    conditionalFormatting.getTextComparison().getFormat().getFont().setColor(fontColor);
  }
  
}

function addDuplicateCondition(ranges: Array<ExcelScript.Range>){
  let conditionalFormatting: ExcelScript.ConditionalFormat;
  for (var i = 0; i < ranges.length; i++) {
    conditionalFormatting = ranges[i].addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria);
    conditionalFormatting.getPreset().setRule({ criterion: ExcelScript.ConditionalFormatPresetCriterion.duplicateValues });
    conditionalFormatting.getPreset().getFormat().getFill().setColor("#ffc7ce");
    conditionalFormatting.getPreset().getFormat().getFont().setColor("#9c0006");
    conditionalFormatting.setStopIfTrue(false);
    conditionalFormatting.setPriority(0);
  }
}

Office.initialize = () => {};

Office.onReady((info) => {
  console.log("Office is Ready!");

  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("setup").onclick = setup;
  }
});

async function setup() {
  try {
      await Excel.run(async (context) => {
      context.workbook.worksheets.getItemOrNullObject("Sample").delete();
      const sheet = context.workbook.worksheets.add("Sample");
  
      let expensesTable = sheet.tables.add("A1:E1", true);
      expensesTable.name = "SalesTable";
      expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];
  
      expensesTable.rows.add(null, [
          ["Frames", 5000, 7000, 6544, 4377],
          ["Saddles", 400, 323, 276, 651],
          ["Brake levers", 12000, 8766, 8456, 9812],
          ["Chains", 1550, 1088, 692, 853],
          ["Mirrors", 225, 600, 923, 544],
          ["Spokes", 6005, 7634, 4589, 8765]
      ]);
  
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
  
      sheet.activate();
      await context.sync();
      });
  } catch (error) {
    console.error(error);
  }
}
  
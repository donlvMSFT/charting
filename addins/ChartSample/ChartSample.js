Office.initialize = () => {};

Office.onReady((info) => {
  console.log("Office is Ready!");

  if (info.host === Office.HostType.Excel) {
    document.getElementById("setup").onclick = setup;
    document.getElementById("test").onclick = test;
    document.getElementById("error").onclick = error;
    $("#dl_setup").on("click", () => tryCatch(dl_setup));
    $("#dl_shape").on("click", () => tryCatch(dl_shape));
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

async function test() {
  try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        let range = sheet.getRange("A1");
        range.format.fill.color = "yellow";

        let rangeAreas = sheet.getRanges("A2:E2,A7:E7");
        rangeAreas.clear();
        await context.sync();

        rangeAreas.select();

        let range1 = sheet.getRange("B1");
        range1.format.fill.color = "green";
    
        await context.sync();
      });
  } catch (error) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange("A10");
        range.values = [["Error"]];
        let range1 = sheet.getRange("A11");
        range1.values = [[error.message]];
        await context.sync();
      });
    } catch (innerError) {
      console.error("Failed to log error to Excel:", innerError);
    }
  }
}

async function error() {
  try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
  
        let expensesTable = sheet.tables.add("A1:E1", true);
        expensesTable.name = "SalesTable";
        expensesTable.getHeaderRowRange().values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"]];

        await context.sync();
      });
  } catch (error) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange("A10");
        range.values = [["Error"]];
        let range1 = sheet.getRange("A11");
        range1.values = [[error.message]];
        await context.sync();
      });
    } catch (innerError) {
      console.error("Failed to log error to Excel:", innerError);
    }
  }
}

async function dl_setup() {
  await Excel.run(async (context) => {
    // Get first chart on the sheet.
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    await context.sync();
    // Load points in the data series.
    let series = chart.series.getItemAt(0);
    // Create a new data label at point 1 in series.
    series.points.getItemAt(1).hasDataLabel = true;

    await context.sync();
  });
}

async function dl_shape() {
  await Excel.run(async (context) => {
    // Get first chart on the sheet.
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let chart = sheet.charts.getItemAt(0);
    await context.sync();
    // Load points in the data series.
    let series = chart.series.getItemAt(0);
    let label = series.points.getItemAt(1).dataLabel;

    // Set the new label properties.
    label.set({
      geometricShapeType: Excel.GeometricShapeType.triangle
    });

    await context.sync();
  });
}
  
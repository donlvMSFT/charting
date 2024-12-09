Office.initialize = () => {};

Office.onReady((info) => {
  console.log("Office is Ready!");

  if (info.host === Office.HostType.Excel) {
    document.getElementById("setup").onclick = setup;
    document.getElementById("test").onclick = test;
    document.getElementById("error").onclick = error;
    document.getElementById("dl_setup").onclick = dl_setup;
    document.getElementById("dl_shape").onclick = dl_shape;

    document.getElementById("set_datalabel_size_multiple").onclick = set_datalabel_size_multiple;
    document.getElementById("set_anchor_top").onclick = set_anchor_top;
    document.getElementById("setdatalabel_newapi").onclick = setdatalabel_newapi;
    document.getElementById("addshape").onclick = addshape;
    document.getElementById("getactiveshape").onclick = getactiveshape;

    document.getElementById("testLeaderLinesAPI").onclick = testLeaderLinesAPI;
    document.getElementById("substring").onclick = substring;
    document.getElementById("testTextRuns").onclick = testTextRuns;
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

async function substring() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add();
    sheet.activate();
    const range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, range);

    let label = chart.series.getItemAt(0).points.getItemAt(1).dataLabel;
    label.text = "Test substring APIs";
    await context.sync();

    let text = label.getSubstring(0, 4);
    text.font.color = "red";
    await context.sync();
  });
}

async function testLeaderLinesAPI() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    await context.sync();
    if (!sheet.id) {
      sheet = context.workbook.worksheets.add(sheetName);
    }

    sheet.activate();

    const count = sheet.charts.getCount();
    await context.sync();

    if (count.value > 0) {
      const chart = sheet.charts.getItemAt(0);
      chart.delete();
    }

    let range = sheet.getRange("A1:C4");
    range.values = [
      ["Type", "Product A", "Product B"],
      ["Q1", 15, 20],
      ["Q2", 22, 15],
      ["Q3", 33, 47]
    ];
    let chart = sheet.charts.add(Excel.ChartType.line, range);
    chart.title.text = "Sales Quantity";
    await context.sync();

    await addLabels(context);

    await changeFormat(context);
  });
}

async function changeFormat(context) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.getItemAt(0);
  let series = chart.series.getItemAt(0);
  let seriesDataLabels = series.dataLabels;
  let lineformat = seriesDataLabels.leaderLines.format;

  lineformat.line.color = "blue";
  lineformat.line.weight = 2;
  lineformat.line.lineStyle = Excel.ChartLineStyle.dot;

  console.log("changes leaderlines format");
}

async function addLabels(context) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let chart = sheet.charts.getItemAt(0);
  let series = chart.series.getItemAt(0);
  series.hasDataLabels = true;
  series.points.load("items");
  await context.sync();
  series.points.items.forEach((point) => point.dataLabel.load("top"));
  await context.sync();

  series.points.items[1].dataLabel.top = series.points.items[1].dataLabel.top - 50;
  series.points.items[2].dataLabel.top = series.points.items[2].dataLabel.top + 50;
  series.dataLabels.geometricShapeType = Excel.GeometricShapeType.rectangle;
  series.dataLabels.showCategoryName = true;
  series.dataLabels.format.border.weight = 1;
  await context.sync();
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

async function addshape() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    // sheet.activate();
    let a = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.hexagon);
    await context.sync();
  });
}

async function getactiveshape() {
  await Excel.run(async (context) => {
    const shape = context.workbook.getActiveShape();
    const shapenull = context.workbook.getActiveShapeOrNullObject();

    shape.load("name");
    await context.sync();
    console.log(shape);

    context.workbook.worksheets.getActiveWorksheet().getRange("D1").values = [[shape.name]];

    if (!shapenull.isNullObject) {
      shapenull.load("name");
      await context.sync();
      console.log(shapenull);
    } else {
      console.log("shape.isNullObject=true");
    }

    await context.sync();
  });
}

async function testTextRuns() {
  //const sheetName: string = "TestLeaderLinesAPI";
  await Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const cellSrc = sheet.getRange("A1:A2");
    const cellSrcTextRun1 = {
      textRuns: [
        {
          text: "Sample",
          font: {
            bold: true,
            color: "#00B0F0",
            size: 14,
            italic: true,
            name: "Abadi",
            underline: "Single",
            strikethrough: true
          }
        },
        {
          text: "1",
          font: { subscript: true }
        },
        {
          text: "String"
        },
        {
          text: "2",
          font: {
            superscript: true,
            color: "black",
            tintAndShade: 0.5
          }
        }
      ]
    };
    const cellSrcTextRun2 = {
      textRuns: [
        {
          text: "",
          font: { color: "#00B0F0" }
        }
      ]
    };
    cellSrc.clear(Excel.ClearApplyTo.all);
    cellSrc.setCellProperties([[cellSrcTextRun1], [cellSrcTextRun2]]);
    sheet.getUsedRange().format.autofitColumns();
    await ctx.sync();

    const cellDest = sheet.getRange("A1:A2");
    const textRunCellProperty = cellDest.getCellProperties({
      textRuns: true
    });

    await ctx.sync();

    const cellTextRuns1 = textRunCellProperty.value[0][0].textRuns;
    const cellTextRuns2 = textRunCellProperty.value[1][0].textRuns;

    console.log(JSON.stringify(cellTextRuns1, undefined, "  "));

    sheet.getRange("D2").values = [[JSON.stringify(cellTextRuns1, undefined, "  ")]];
  });
}

async function set_datalabel_size_multiple() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add();
    sheet.activate();
    const range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, range);

    const datalabel0 = chart.series.getItemAt(0).points.getItemAt(0).dataLabel;
    datalabel0.geometricShapeType = "Rectangle";
    datalabel0.setWidth(10);
    datalabel0.setHeight(10);

    const datalabel1 = chart.series.getItemAt(0).points.getItemAt(1).dataLabel;
    datalabel1.geometricShapeType = "Rectangle";
    datalabel1.setWidth(15);
    datalabel1.setHeight(15);

    const datalabel2 = chart.series.getItemAt(0).points.getItemAt(2).dataLabel;
    datalabel2.geometricShapeType = "Rectangle";
    datalabel2.setWidth(20);
    datalabel2.setHeight(20);
    await context.sync();
  });
}

async function set_datalabel_width() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    var label = chart.series.getItemAt(0).points.getItemAt(get_datalabel_size_index()).dataLabel;
    if (test_size_with_existing_api) {
      label.left = get_datalabel_size_width();
    } else {
      label.setWidth(get_datalabel_size_width());
    }
    await context.sync();
  });
}
async function set_datalabel_height() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    var label = chart.series.getItemAt(0).points.getItemAt(get_datalabel_size_index()).dataLabel;
    if (test_size_with_existing_api) {
      label.top = get_datalabel_size_height();
    } else {
      label.setHeight(get_datalabel_size_height());
    }
    await context.sync();
  });
}
async function get_datalabel_top() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const points = chart.series.getItemAt(0).points;
    const datalabel = points.getItemAt(0).dataLabel;
    datalabel.load("top, left, width, height");
    await context.sync();
    console.log(
      `[0]: top=${datalabel.top}, left=${datalabel.left}, width=${datalabel.width}, height=${datalabel.height}`
    );
    const count = points.getCount();
    await context.sync();
    if (count.value > 1) {
      const datalabel1 = points.getItemAt(1).dataLabel;
      datalabel1.load("top, left, width, height");
      await context.sync();
      console.log(
        `[1]: top=${datalabel1.top}, left=${datalabel1.left}, width=${datalabel1.width}, height=${datalabel1.height}`
      );
    }
    if (count.value > 2) {
      const datalabel2 = points.getItemAt(2).dataLabel;
      datalabel2.load("top, left, width, height");
      await context.sync();
      console.log(
        `[2]: top=${datalabel2.top}, left=${datalabel2.left}, width=${datalabel2.width}, height=${datalabel2.height}`
      );
    }
    await context.sync();
  });
}
async function set_datalabel_top() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const datalabel0 = chart.series.getItemAt(0).points.getItemAt(0).dataLabel;
    datalabel0.top = 10;
    datalabel0.left = 0;
    const datalabel1 = chart.series.getItemAt(0).points.getItemAt(1).dataLabel;
    datalabel1.top = 100;
    datalabel1.left = 200;
    const datalabel2 = chart.series.getItemAt(0).points.getItemAt(2).dataLabel;
    datalabel2.top = 50;
    datalabel2.left = 100;
    await context.sync();
  });
}
async function setdatalabels_location_series() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    chart.series.getItemAt(0).dataLabels.position = Excel.ChartDataLabelPosition.center;

    await context.sync();
  });
}

async function setdatalabels_location() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    chart.dataLabels.position = Excel.ChartDataLabelPosition.insideBase;

    await context.sync();
  });
}

async function setdatalabel_location() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    chart.series.getItemAt(0).points.getItemAt(0).dataLabel.position = Excel.ChartDataLabelPosition.insideEnd;
    chart.series.getItemAt(0).points.getItemAt(1).dataLabel.position = Excel.ChartDataLabelPosition.insideBase;

    await context.sync();
  });
}

async function getdatalabels_location() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    const label = series.points.getItemAt(0).dataLabel;
    const labels = chart.dataLabels;
    const labels_series = series.dataLabels;
    label.load("position");
    labels.load("position");
    labels_series.load("position");
    await context.sync();

    console.log("datalabel position: " + label.position);
    console.log("chart.datalabels position: " + labels.position);
    console.log("series.datalabels position: " + labels_series.position);

    await context.sync();
  });
}

async function setup1() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    // const range = sheet.getRange("A1:A2");
    // range.values = [[3], [22]];
    //const range = sheet.getRange("A1");
    //range.values = [[11]];
    sheet.charts.add(Excel.ChartType.columnClustered, range);

    await context.sync();
  });
}

async function showdatalabel() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const count = sheet.charts.getCount();
    await context.sync();

    const chart = sheet.charts.getItemAt(count.value - 1);
    const datalabel = chart.series
      .getItemAt(0)
      .points.getItemAt(1)
      .dataLabel.load("showLegendKey,geometricShapeType");
    await context.sync();

    console.log(
      //datalabel.showValue,
      datalabel.showLegendKey,
      datalabel.geometricShapeType
    );

    await context.sync();
  });
}

async function hidedatalabel() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    chart.dataLabels.showValue = false;

    await context.sync();
  });
}

async function setdatalabel() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    chart.dataLabels.showValue = true;

    await context.sync();
  });
}

async function setdatalabelformat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    chart.dataLabels.showLegendKey = true;
    chart.dataLabels.showValue = true;
    chart.dataLabels.separator = ".";
    chart.dataLabels.horizontalAlignment = Excel.HorizontalAlignment.distributed;
    chart.dataLabels.format.fill.setSolidColor("pink");
    chart.dataLabels.format.font.bold = true;
    chart.dataLabels.numberFormat = "0.00";

    await context.sync();
  });
}

async function getdatalabelformat() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const showValue = chart.dataLabels.load("showValue, showLegendKey, separator, horizontalAlignment, numberFormat");
    const solidColor = chart.dataLabels.format.fill.getSolidColor();
    const bold = chart.dataLabels.format.font.load("bold");
    await context.sync();

    console.log("horizontalAlignment: " + showValue.horizontalAlignment);
    console.log("seperator: " + showValue.separator);
    console.log("showLegendKey: " + showValue.showLegendKey);
    console.log("showValue: " + showValue.showValue);
    console.log("solidColor: " + solidColor.value);
    console.log("bold: " + bold.bold);
    console.log("numberFormat: " + showValue.numberFormat);

    await context.sync();
  });
}

var iterShapeType = 81; // Excel.GeometricShapeType.lineInverse
const failingShapes = [1, 37, 125, 184, 185, 186];
const shapeTypeCallout = [107, 108, 110, 111, 116, 117, 118];
var iterFailingShapes = 0;
var iterCalloutShapes = 0;
async function setdatalabels_newapi() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const label = chart.series.getItemAt(0).points.getItemAt(0).dataLabel;
    label.load("geometricShapeType");
    await context.sync();
    console.log(label.geometricShapeType);

    const iter_type = get_selected_iter_datalabels();
    switch (iter_type) {
      case "failed":
        // test all failing shapes
        const shape = failingShapes[iterFailingShapes % failingShapes.length];
        console.log("shape: " + shape);
        iterFailingShapes += 1;
        chart.dataLabels.geometricShapeType = shape;
        //sheet.shapes.addGeometricShape(shape);
        break;
      case "callout":
        // test all callout shapes
        const shape1 = shapeTypeCallout[iterCalloutShapes % shapeTypeCallout.length];
        console.log("shape: " + shape1);
        iterCalloutShapes += 1;
        chart.dataLabels.geometricShapeType = shape1;
        break;
      default:
        // test all shapes
        console.log("iter: " + iterShapeType);
        chart.dataLabels.geometricShapeType = iterShapeType;
        iterShapeType += 1;
    }

    await context.sync();
  });
}

async function setdatalabel_newapi() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add();
    sheet.activate();
    const range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, range);

    const labels = chart.series.getItemAt(0).dataLabels;
    const label1 = chart.series.getItemAt(0).points.getItemAt(1).dataLabel;

    labels.geometricShapeType = Excel.GeometricShapeType.rectangle;
    label1.geometricShapeType = Excel.GeometricShapeType.wedgeRectCallout;

    labels.load("geometricShapeType, showAsStickyCallout");
    label1.load("geometricShapeType, showAsStickyCallout");
    await context.sync();

    sheet.getRange("D1").values = [["chart.datalabels geometricShapeType: " + labels.geometricShapeType + ", callout: " + labels.showAsStickyCallout]];
    sheet.getRange("D2").values = [["points[0].label  geometricShapeType: " + label1.geometricShapeType + ", callout: " + label1.showAsStickyCallout]];
    console.log(
      "chart.datalabels geometricShapeType: " + labels.geometricShapeType + ", callout: " + labels.showAsStickyCallout
    );
    console.log(
      "points[0].label  geometricShapeType: " + label1.geometricShapeType + ", callout: " + label1.showAsStickyCallout
    );

    await context.sync();
  });
}

async function getdatalabel_newapi() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    const label1 = chart.series.getItemAt(0).points.getItemAt(0).dataLabel;
    const label2 = chart.series.getItemAt(0).points.getItemAt(1).dataLabel;
    const label3 = chart.series.getItemAt(0).points.getItemAt(2).dataLabel;

    label1.load("geometricShapeType");
    label2.load("geometricShapeType");
    label3.load("geometricShapeType");
    await context.sync();

    console.log("points[0].label: " + label1.geometricShapeType);
    console.log("points[1].label: " + label2.geometricShapeType);
    console.log("points[2].label: " + label3.geometricShapeType);

    await context.sync();
  });
}

async function setdatalabels_newapi_property() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);

    chart.dataLabels.geometricShapeType = Excel.GeometricShapeType.cloudCallout;
    //chart.dataLabels.geometricShapeType = 0;
    //chart.dataLabels.geometricShapeType = null;
    await context.sync();

    const datalabels = chart.dataLabels.load("geometricShapeType");
    await context.sync();

    console.log("chart.datalabels: " + datalabels.geometricShapeType);
  });
}

async function setdatalabels_newapi_property_get() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const datalabels = chart.dataLabels.load("geometricShapeType");
    await context.sync();

    console.log("chart.datalabels: " + datalabels.geometricShapeType);
  });
}

async function setdatalabels_series() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);

    series.dataLabels.geometricShapeType = Excel.GeometricShapeType.heart;
    await context.sync();

    const datalabels = series.dataLabels.load("geometricShapeType");
    await context.sync();

    console.log("series.datalabels: " + datalabels.geometricShapeType);
  });
}

async function getdatalabels_series() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    const datalabels = series.dataLabels.load("geometricShapeType");
    await context.sync();

    console.log("series.datalabels: " + datalabels.geometricShapeType);
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    let chart = sheet.charts.getItemAt(0);
    let series = chart.series.getItemAt(0);
    let title = "Label Title";
    let content = "This is a data label.";
    let label = series.points.getItemAt(0).dataLabel;
    label.text = title + "\n" + content;
    label.load("geometricShapeType, ShowAsStickyCallout");
    await context.sync();

    console.log("Label shape type: " + label.geometricShapeType);
    console.log("Label data calllout: " + label.showAsStickyCallout);

    if (label.showAsStickyCallout == false) {
      label.geometricShapeType = Excel.geometricShapeType.callout1;
    }
    await context.sync();
  });
}

async function callout() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    const label = series.points.getItemAt(0).dataLabel;
    const labels = chart.dataLabels;
    const labels_series = series.dataLabels;
    label.load("showAsStickyCallout");
    labels.load("showAsStickyCallout");
    labels_series.load("showAsStickyCallout");
    await context.sync();

    console.log("datalabel callout: " + label.showAsStickyCallout);
    console.log("chart.datalabels callout: " + labels.showAsStickyCallout);
    console.log("series.datalabels callout: " + labels_series.showAsStickyCallout);

    await context.sync();
  });
}

async function boolget() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    const label = series.points.getItemAt(0).dataLabel;
    const labels = chart.dataLabels;
    label.load("showCategoryName, showValue, autoText");
    labels.load("showCategoryName, showValue, autoText");
    await context.sync();

    console.log("showValue: " + label.showValue + ", " + labels.showValue);
    console.log("showCategoryName: " + label.showCategoryName + ", " + labels.showCategoryName);
    console.log("autoText: " + label.autoText + ", " + labels.autoText);

    // set when null
    //label.autoText = true;

    await context.sync();
  });
}

const supportedCharts = [
  Excel.ChartType._3DColumn,
  Excel.ChartType._3DLine,
  Excel.ChartType.pieOfPie,
  Excel.ChartType._3DBarStacked100,
  Excel.ChartType.areaStacked,
  Excel.ChartType.xyscatterLines,
  Excel.ChartType.radarMarkers,
  Excel.ChartType.stockHLC,
  Excel.ChartType.stockOHLC,
  Excel.ChartType.stockVHLC,
  Excel.ChartType.stockVOHLC
];
const unsupportedCharts = [
  Excel.ChartType.regionMap,
  Excel.ChartType.surface,
  Excel.ChartType.treemap,
  Excel.ChartType.sunburst,
  Excel.ChartType.histogram,
  Excel.ChartType.boxwhisker,
  Excel.ChartType.waterfall,
  Excel.ChartType.funnel
];
var iterTestCharts = 0; // test all
// var iterTestCharts = 11; // test unsupported
async function setcharts() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const count = sheet.charts.getCount();
    await context.sync();

    if (count.value > 0) {
      // console.log("delete previous chart");
      const chart = sheet.charts.getItemAt(0);
      chart.delete();
    }

    const set2drange = function(context) {
      range = sheet.getRange("A1:B5");
      range.values = [
        ["Type", "Num"],
        ["a", 3],
        ["b", 22],
        ["c", 111],
        ["d", 5]
      ];
      return range;
    };

    const set3drange = function(context) {
      range = sheet.getRange("A1:C4");
      range.values = [
        ["Type", "Num1", "Num2"],
        ["a", 3, 101],
        ["b", 22, 23],
        ["c", 111, 4]
      ];
      return range;
    };

    const setstock3drange = function(context) {
      range = sheet.getRange("A1:D4");
      range.values = [
        ["Name", "High Price", "Low Price", "Closing Price"],
        ["lake", 150, 99, 101],
        ["mountain", 250, 149, 199],
        ["sea", 99, 101, 100]
      ];
      return range;
    };

    const setstock4drange = function(context) {
      range = sheet.getRange("A1:E5");
      range.values = [
        ["Name", "Opening Price/ Volumn traded", "High Price", "Low Price", "Closing Price"],
        ["lake", 133, 150, 99, 101],
        ["mountain", 150, 250, 149, 199],
        ["sea", 99, 99, 101, 100],
        ["sky", 50, 150, 20, 100]
      ];
      return range;
    };

    const setstock5drange = function(context) {
      range = sheet.getRange("A1:F6");
      range.values = [
        ["Name", "Volumn traded", "Opening Price", "High Price", "Low Price", "Closing Price"],
        ["lake", 10, 133, 150, 99, 101],
        ["mountain", 9, 150, 250, 149, 199],
        ["sea", 25, 99, 99, 101, 100],
        ["sky", 66, 50, 150, 20, 100],
        ["river", 15, 77, 100, 66, 71]
      ];
      return range;
    };

    if (iterTestCharts < supportedCharts.length) {
      console.log("Set up supported chart.");
    } else if (iterTestCharts < supportedCharts.length + unsupportedCharts.length) {
      console.log("Set up unsupported chart.");
    } else {
      console.log("Test done.");
      return;
    }
    const chartType = supportedCharts.concat(unsupportedCharts)[iterTestCharts];

    switch (chartType) {
      case Excel.ChartType._3DColumn:
      case Excel.ChartType.pieOfPie:
      case Excel.ChartType.regionMap:
      case Excel.ChartType.treemap:
      case Excel.ChartType.sunburst:
      case Excel.ChartType.boxwhisker:
      case Excel.ChartType.waterfall:
      case Excel.ChartType.funnel:
        range = set2drange(context);
        break;
      case Excel.ChartType._3DLine:
      case Excel.ChartType._3DBarStacked100:
      case Excel.ChartType.areaStacked:
      case Excel.ChartType.xyscatterLines:
      case Excel.ChartType.radarMarkers:
      case Excel.ChartType.surface:
      case Excel.ChartType.histogram:
        range = set3drange(context);
        break;
      case Excel.ChartType.stockHLC:
        range = setstock3drange(context);
        break;
      case Excel.ChartType.stockOHLC:
      case Excel.ChartType.stockVHLC:
        range = setstock4drange(context);
        break;
      case Excel.ChartType.stockVOHLC:
        range = setstock5drange(context);
        break;
    }
    sheet.charts.add(chartType, range);
    iterTestCharts++;

    await context.sync();
  });
}

function get_selected_iter_datalabels() {
  var e = document.getElementById("iter_datalabels");
  var value = e.options[e.selectedIndex].value;
  return value;
}

async function get_anchor_top() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    var label = series.points.getItemAt(get_datalabel_index()).dataLabel;

    var anchor = label.getTailAnchor();
    anchor.load("left, top");
    await context.sync();

    console.log(`Data label edge anchor position: (${anchor.left}, ${anchor.top})`);

    await context.sync();
  });
}

async function set_anchor_top() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add();
    sheet.activate();
    const range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, range);

    const series = chart.series.getItemAt(0);
    var label = series.points.getItemAt(1).dataLabel;
    label.geometricShapeType = Excel.GeometricShapeType.wedgeRectCallout;

    label.load("width, height, top, left");
    await context.sync();

    var anchor = label.getTailAnchor();
    anchor.top = -label.height;
    anchor.left = -label.width;

    anchor.load("left, top");
    await context.sync();

    console.log(`Data label edge anchor position: (${anchor.left}, ${anchor.top})`);
  });
}

async function set_anchor_left() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const chart = sheet.charts.getItemAt(0);
    const series = chart.series.getItemAt(0);
    var label = series.points.getItemAt(get_datalabel_index()).dataLabel;
    var anchor = label.getTailAnchor();

    label.load("width, height, top, left");
    await context.sync();

    // anchor.left = label.width;
    // anchor.left = 200;

    var selected_anchor_pos = get_selected_anchor_position();
    switch (selected_anchor_pos) {
      case "left-top-datalabel":
        anchor.left = -label.width;
        break;
      case "right-below-datalabel":
        anchor.left = 2 * label.width;
        break;
      case "left-edge-datalabel":
        anchor.left = 0;
        break;
      case "top-edge-datalabel":
        anchor.left = 0.25 * label.width;
        break;
      case "inside-datalabel":
        anchor.left = 0.5 * label.width;
        break;
      case "edge-chart":
        anchor.left = -label.left;
        break;
      case "outside-chart":
        anchor.left = -label.left - 10;
        break;
      case "set-by-get":
        anchor.load("left, top");
        await context.sync();
        anchor.left = anchor.left;
        break;
    }
    await context.sync();
  });
}

function get_selected_anchor_position() {
  var e = document.getElementById("anchor_pos_select");
  var value = e.options[e.selectedIndex].value;
  return value;
}

function get_datalabel_index() {
  var index = parseInt(document.getElementById("datalabel_index").value);
  if (isNaN(index)) {
    index = 1; // default 1
  }
  // console.log(`index: ${index}`);
  return index;
}

const pieChart = [
  Excel.ChartType.pieOfPie,
  Excel.ChartType.barOfPie
  //Excel.ChartType.pie, // correct
  //Excel.ChartType.pieExploded // correct
];
var iterTestChartsPie = 0;
async function setcharts_pie() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("A1:A3");
    range.values = [[3], [22], [111]];
    const count = sheet.charts.getCount();
    await context.sync();

    if (count.value > 0) {
      // console.log("delete previous chart");
      const chart = sheet.charts.getItemAt(0);
      chart.delete();
    }

    // correct
    const set2drange2num = function(context) {
      range = sheet.getRange("A1:B3");
      range.values = [
        ["Type", "Num"],
        ["a", 3],
        ["b", 22]
      ];
      return range;
    };

    // incorrect
    const set2drange3num = function(context) {
      range = sheet.getRange("A1:B4");
      range.values = [
        ["Type", "Num"],
        ["a", 3],
        ["b", 22],
        ["c", 111]
      ];
      return range;
    };

    // correct
    const set2drange3num2 = function(context) {
      range = sheet.getRange("A1:B4");
      range.values = [
        ["Type", "Num"],
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ];
      return range;
    };

    // incorrect
    const set2drange4num = function(context) {
      range = sheet.getRange("A1:B5");
      range.values = [
        ["Type", "Num"],
        ["a", 3],
        ["b", 22],
        ["c", 111],
        ["d", 5]
      ];
      return range;
    };

    // incorrect
    const set2drange5num = function(context) {
      range = sheet.getRange("A1:B6");
      range.values = [
        ["Type", "Num"],
        ["a", 3],
        ["b", 22],
        ["c", 111],
        ["d", 5],
        ["e", 66]
      ];
      return range;
    };

    if (iterTestChartsPie < pieChart.length) {
      console.log("Set up pie chart.");
    } else {
      console.log("Test done.");
      return;
    }
    const chartType = pieChart[iterTestChartsPie];

    range = set2drange3num2(context);
    sheet.charts.add(chartType, range);
    iterTestChartsPie++;

    await context.sync();
  });
}
function get_datalabel_size_index() {
  var index = parseInt(document.getElementById("datalabel_size_index").value);
  if (isNaN(index)) {
    index = 1; // default 1
  }
  // console.log(`index: ${index}`);
  return index;
}
function get_datalabel_size_width() {
  var width = parseFloat(document.getElementById("datalabel_size_width").value);
  if (isNaN(width)) {
    width = 25.5; // default 25
  }
  //console.log(`width: ${width}`);
  return width;
}
function get_datalabel_size_height() {
  var height = parseFloat(document.getElementById("datalabel_size_height").value);
  if (isNaN(height)) {
    height = 25.5; // default 25
  }
  //console.log(`height: ${height}`);
  return height;
}
function set_datalabel_size(width, height) {
  document.getElementById("datalabel_size_width").value = width;
  document.getElementById("datalabel_size_height").value = height;
}

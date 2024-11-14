// $("#run").on("click", () => tryCatch(run));

Office.onReady(function () {
  // Office is ready.
  $(document).ready(function () {
    // The document is ready.
    $("#run").on("click", () => tryCatch(run));
  });
});

const lineSeperator = navigator.platform.startsWith("Win") ? "\r\n" : "\n";

async function run() {
  await Excel.run(async (context) => {
    // Truy xuat api url
    const apiUrl = $("#apiUrl").val();
    if (!apiUrl || typeof apiUrl !== "string") {
      throw new Error("Can't detect API URL. Please fill the API URL textbox.");
    }

    // Truy xuat vung dang duoc chon
    let selectedRange: Excel.Range = context.workbook.getSelectedRange();
    selectedRange = selectedRange.load("values");
    selectedRange = selectedRange.load("columnIndex");
    selectedRange = selectedRange.load("rowIndex");
    selectedRange = await context.sync(selectedRange);

    // Doc sql tu gia tri cua vung dang duoc chon
    const [firstRowIndex, sql] = readSQL(selectedRange.values);

    // Goi API parse file SQL
    const response = await fetch(apiUrl, {
      method: "POST",
      body: sql
    });

    // Truy xuat cau SQL ket qua
    const sqlResult = await response.text();

    // Tach cau SQL ket qua
    const sqlResultLines = sqlResult.split(lineSeperator);

    // Truy xuat cot cuoi cung cua vung dang duoc chon
    let lastColumn: Excel.Range = selectedRange.getLastColumn();
    lastColumn = lastColumn.load("columnIndex");
    lastColumn = await context.sync(lastColumn);

    // Truy xuat vi tri ghi tiep theo
    const rowIndex = selectedRange.rowIndex + firstRowIndex;
    const columnIndex = lastColumn.columnIndex + 2;

    // Truy xuat vung ghi
    const writeRange: Excel.Range = context.workbook.worksheets
      .getActiveWorksheet()
      .getRangeByIndexes(rowIndex, columnIndex, sqlResultLines.length, 1);

    // Ghi ket qua SQL
    writeRange.values = sqlResultLines.map((line) => [line]);

    // Xuat thong bao
    notification("Converted successfully!", "INFO");

    await context.sync();
  });
}

/**
 * Ham doc SQL tu vung dang chon
 */
function readSQL(values: any[][]): [number, string] {
  let sql: string = "";

  let firstRow = undefined;

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];

    let rowEmpty: boolean = true;

    for (const value of row) {
      if (!value) {
        continue;
      }

      if (rowEmpty) {
        rowEmpty = false;
      }

      if (firstRow === undefined) {
        firstRow = rowIndex;
      }

      sql += value + lineSeperator;
    }

    if (rowEmpty && firstRow !== undefined) {
      sql += lineSeperator;
    }
  }

  // for (const row of values) {
  //   for (const value of row) {
  //     if (!value) {
  //       continue;
  //     }

  //     sql += value + "\n";
  //   }
  // }

  return [firstRow, sql];
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    notification(error, "ERROR");
  }
}

/**
 * Ham xuat thong bao o giao dien
 */
function notification(message: string, messageType: "ERROR" | "INFO") {
  const notification = $("#notification");

  if (messageType === "ERROR") {
    notification.css("color", "red");
  } else {
    notification.css("color", "blue");
  }

  notification.text(message);
}

import * as React from "react";

const App = () => {
  const [columnRange, setColumnRange] = React.useState<string | null>("");
  const [rowRange, setRowRange] = React.useState<string | null>("");
  const [errors, setErrors] = React.useState<string[]>([]);

  const handleSelectColumns = async () => {
    try {
      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        return range.address;
      });

      setColumnRange(result);
    } catch (error) {
      console.error(error);
    }
  };

  const handleSelectRows = async () => {
    try {
      const result = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        return range.address;
      });

      setRowRange(result);
    } catch (error) {
      console.error(error);
    }
  };

  const validateRanges = (columnRange, rowRange) => {
    let newErrors = [];

    // clean the ranges to remove sheet names if present to not interfere with validation
    const cleanColumnRange = columnRange.includes("!") ? columnRange.split("!")[1] : columnRange;
    const cleanRowRange = rowRange.includes("!") ? rowRange.split("!")[1] : rowRange;

    const columnRangeMatch = cleanColumnRange.match(/[A-Za-z]+(?=\d|\s*$)/g);
    const rowRangeMatch = cleanRowRange.match(/[A-Za-z]+(?=\d|\s*$)/g);

    if (!columnRangeMatch || !rowRangeMatch) {
      newErrors.push("Range format is incorrect");
      setErrors(newErrors);
      return false;
    }

    const columnStart = columnRangeMatch[0];
    const columnEnd = columnRangeMatch[columnRangeMatch.length - 1];
    const rowRangeStart = rowRangeMatch[0];
    const rowRangeEnd = rowRangeMatch[rowRangeMatch.length - 1];

    // make sure column range is 1 row
    const columnRowNumbers = cleanColumnRange.match(/\d+/g);
    if (columnRowNumbers && new Set(columnRowNumbers).size !== 1) {
      newErrors.push("Column range must be a single row");
    }

    // make sure row range is within column bounds
    if (rowRangeStart !== columnStart || rowRangeEnd !== columnEnd) {
      newErrors.push("Row range must start and end within the column bounds");
    }

    if (newErrors.length > 0) {
      setErrors(newErrors);
      return false;
    }

    return true;
  };

  const handleExtractData = async () => {
    try {
      setErrors([]);

      if (!validateRanges(columnRange, rowRange)) {
        return;
      }

      const data = await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        if (columnRange.includes("!")) {
          const sheetName = columnRange.split("!")[0];
          sheet = context.workbook.worksheets.getItem(sheetName);
        }

        // define the ranges for columns and rows
        const columnRangeAddress = columnRange.includes("!") ? columnRange.split("!")[1] : columnRange;
        const rowRangeAddress = rowRange.includes("!") ? rowRange.split("!")[1] : rowRange;

        // get the range for the column headers
        const headerRange = sheet.getRange(columnRangeAddress);
        headerRange.load("values");
        // get the range for the rows
        const dataRange = sheet.getRange(rowRangeAddress);
        dataRange.load("values");

        await context.sync();

        // convert 2D array into JSON objects
        const headers = headerRange.values[0];
        const rows = dataRange.values;
        const jsonData = rows.map((row) => {
          let obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });

        return jsonData;
      });

      sendDataToServer(data);
    } catch (error) {
      console.error("Error extracting data: ", error);
      setErrors([...errors, "Error extracting data"]);
    }
  };

  const sendDataToServer = async (jsonData) => {
    console.log(jsonData);
    // TODO: send json data to server for processing, create new sheet and display processed data in nice format
    // const endpoint = "http://localhost:3001/process-data";
    // try {
    //   const response = await fetch(endpoint, {
    //     method: "POST",
    //     body: JSON.stringify(jsonData),
    //     headers: {
    //       "Content-Type": "application/json",
    //     },
    //   });
    //   if (!response.ok) {
    //     throw new Error(`Server responded with ${response.status}`);
    //   }
    //   const result = await response.json();
    // } catch (error) {
    //   console.error("Failed to send data to the server:", error);
    // }
  };

  return (
    <div className="min-h-screen w-full bg-slate-100 p-4 text-slate-700">
      <div className="w-full">
        <div>
          <h2 className="text-slate-600 font-black text-xl border-b border-slate-300 pb-2">Instructions</h2>
          <ul className="list list-decimal list-inside font-medium mt-6">
            <li>Set the range for your columns</li>
            <li>Set the range for your rows based on your set columns</li>
            <li>After column and row ranges have been set, you may process your data</li>
          </ul>
        </div>

        <h2 className="text-slate-600 font-black text-xl border-b border-slate-300 pb-2 mt-6">Set Ranges</h2>
        <div className="w-full flex flex-col mt-6">
          <button
            className="rounded-t w-full bg-blue-600 text-white hover:bg-blue-700 px-4 py-2 h-11 font-bold"
            onClick={handleSelectColumns}
          >
            Set Selected Columns
          </button>
          <input
            disabled
            id="set-columns-input"
            className="rounded-b w-full bg-white px-4 py-4 h-11 font-bold border border-solid border-slate-400"
            onChange={(e) => {
              setColumnRange(e.target.value);
            }}
            value={columnRange}
          />
        </div>
      </div>

      <div className="w-full mt-6">
        <div className="w-full flex flex-col">
          <button
            className="rounded-t w-full bg-blue-600 text-white hover:bg-blue-700 px-4 py-2 h-11 font-bold"
            onClick={handleSelectRows}
          >
            Set Selected Rows
          </button>
          <input
            disabled
            id="set-rows-input"
            className="rounded-b w-full bg-white px-4 py-2 h-11 font-bold border border-solid border-slate-400"
            onChange={(e) => {
              setRowRange(e.target.value);
            }}
            value={rowRange}
          />
        </div>
      </div>

      {columnRange && rowRange && (
        <div>
          <h2 className="font-black text-slate-600 text-xl border-b border-slate-300 pb-2 mt-6">
            Send Data for Processing
          </h2>
          {errors.length ? (
            <div className="p-4 mt-6 bg-red-50 border border-red-300 rounded">
              <p className="text-red-500 font-black text-xl">Errors</p>
              <ul className="mt-2 list list-disc list-inside font-medium text-red-500">
                {errors.map((error, i) => {
                  return <li key={`${error}_${i}`}>{error}</li>;
                })}
              </ul>
            </div>
          ) : null}
          <div className="mt-6">
            <button
              className={`text-white rounded px-4 py-2 h-11 font-bold w-full bg-green-500 hover:bg-green-600`}
              onClick={handleExtractData}
            >
              Process Data
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;

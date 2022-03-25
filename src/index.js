const express = require("express");
const reader = require("xlsx");
const fs = require("fs");

const app = express();

app.use("/excel", async (req, res) => {
  fs.readFile(__dirname + "/filestoExcel/output.json", (err, data) => {
    if (err) {
      console.error(err);
      return;
    }
    const results = JSON.parse(data);
    const resultsJSON = results.results;
    const woorkbook = reader.utils.book_new();
    const woorkSheet = reader.utils.json_to_sheet(resultsJSON);
    reader.utils.book_append_sheet(woorkbook, woorkSheet, "results");
    const exportFileName = __dirname + `/filestoExcel/results.xlsx`;
    reader.writeFile(woorkbook, exportFileName);
    return res.send("Generado");
  });
});

app.use("/", async (req, res) => {
  const file = reader.readFile(
    __dirname + "/files/TIENDAS SUBIDAS CON ZOONA.xlsx"
  );
  const sheets = file.SheetNames;

  const data = await Promise.all(
    sheets.map(async (sheet) => {
      return new Promise((resolve) =>
        resolve({
          name: "Tiendas",
          data: reader.utils.sheet_to_json(file.Sheets[sheet]),
        })
      );
    })
  );

  fs.writeFile(
    __dirname + "/files/tiendas.json",
    JSON.stringify(data),
    (err) => {
      if (err) {
        console.log("An error occured while writing JSON Object to File.");
        return res.send(err);
      }

      console.log("JSON file has been saved.");
    }
  );

  res.send("Leido");
});

app.listen(3000, () => console.log("Server on port", 3000));

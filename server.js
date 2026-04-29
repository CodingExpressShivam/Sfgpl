const express = require("express");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
app.use(express.json());
app.use(express.static("public"));

const FILE_PATH = "users.xlsx";

app.post("/save", async (req, res) => {
  const data = req.body;

  let workbook = new ExcelJS.Workbook();
  let worksheet;

  try {
    if (fs.existsSync(FILE_PATH)) {
      await workbook.xlsx.readFile(FILE_PATH);
      worksheet = workbook.getWorksheet("Sheet1");
    } else {
      worksheet = workbook.addWorksheet("Sheet1");

      worksheet.columns = [
        { header: "Date", key: "date" },
        { header: "Time", key: "time" },
        { header: "Challan", key: "challan" },
        { header: "Item", key: "item" },
        { header: "Quantity", key: "quantity" },
        { header: "Amount", key: "amount" },
        { header: "Unit", key: "unit" },
        { header: "Vehicle", key: "vehicle" },
        { header: "Driver", key: "driver" },
        { header: "Location", key: "location" }
      ];
    }

    worksheet.addRow(data);
    await workbook.xlsx.writeFile(FILE_PATH);

    res.send("Saved");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error saving data");
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

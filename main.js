const ExcelJS = require("exceljs");
const PptxGenJS = require("pptxgenjs");

// Ruta del archivo Excel
const excelPath =
  "assetsCONTROL 678_2024-1 - EVALUACIÓN - STP - CAMIÓN ALJIBE - LYLG-45.xlsx";

// Crear una nueva presentación
const pres = new PptxGenJS();

// Leer el archivo Excel
const workbook = new ExcelJS.Workbook();
workbook.xlsx
  .readFile(excelPath)
  .then(() => {
    const worksheet = workbook.getWorksheet(1); // Obtener la primera hoja
    const rows = [];

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      rows.push(row.getCell(1).value); // Obtener el valor de la primera columna
    });

    for (let i = 0; i < rows.length; i += 2) {
      const slide = pres.addSlide();

      // Añadir la primera actividad (arriba izquierda)
      if (rows[i]) {
        slide.addText(rows[i], {
          x: 0.5,
          y: 0.5,
          fontSize: 24,
          align: "left",
        });
      }

      // Añadir la segunda actividad (centro izquierda)
      if (rows[i + 1]) {
        slide.addText(rows[i + 1], {
          x: 0.5,
          y: 2.5,
          fontSize: 24,
          align: "left",
        });
      }
    }

    // Guardar la presentación en PowerPoint
    const pptPath = "ruta/al/archivo.pptx";
    pres
      .writeFile({ fileName: pptPath })
      .then(() => {
        console.log(`Presentación guardada en: ${pptPath}`);
      })
      .catch((err) => {
        console.error("Error al guardar la presentación:", err);
      });
  })
  .catch((err) => {
    console.error("Error al leer el archivo Excel:", err);
  });

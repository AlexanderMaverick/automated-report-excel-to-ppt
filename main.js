const ExcelJS = require("exceljs");
const PptxGenJS = require("pptxgenjs");
const path = require("path");

// Ruta del archivo Excel
const excelPath =
  "assets/CONTROL 678_2024-1 - EVALUACIÓN - STP - CAMIÓN ALJIBE - LYLG-45.xlsx";

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
      rows.push(row.getCell(6).value); // Obtener el valor de la primera columna
    });

    for (let i = 9; i < rows.length; i += 2) {
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

    // Determinar la ruta del archivo PowerPoint
    const excelDir = path.dirname(excelPath); // Obtener el directorio del archivo Excel
    const pptPath = path.join(excelDir, "archivo.pptx"); // Construir la ruta del archivo PowerPoint

    // Guardar la presentación en PowerPoint
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

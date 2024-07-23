const ExcelJS = require("exceljs");
const PptxGenJS = require("pptxgenjs");
const path = require("path");

// Ruta del archivo Excel
const excelPath =
  "assets/CONTROL 678_2024-1 - EVALUACIÓN - STP - CAMIÓN ALJIBE - LYLG-45.xlsx";

// Crear una nueva presentación
const pres = new PptxGenJS();
pres.defineLayout({ name: "Carta", width: 8.5, height: 11 });
pres.layout = "Carta";

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

    for (let i = 7; i < rows.length; i += 2) {
      const slide = pres.addSlide();

      // Establecer el color de fondo de la diapositiva
      slide.background = { color: "616161" }; // Cambia 'FFFFFF' por el color que desees

      // Añadir la primera actividad (arriba izquierda)
      if (rows[i]) {
        slide.addText(
          [
            { text: `${i + 1}. `, options: { color: "#F39200", fontSize: 14 } }, // Color del número
            { text: rows[i], options: { color: "FFFFFF", fontSize: 14 } }, // Color del texto
          ],
          { x: 0.5, y: 0.5, align: "left" }
        );
      }

      // Añadir la segunda actividad (centro izquierda)
      if (rows[i + 1]) {
        slide.addText(
          [
            { text: `${i + 2}. `, options: { color: "#F39200", fontSize: 14 } }, // Color del número
            { text: rows[i + 1], options: { color: "FFFFFF", fontSize: 14 } }, // Color del texto
          ],
          { x: 0.5, y: 5.5, align: "left" }
        );
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

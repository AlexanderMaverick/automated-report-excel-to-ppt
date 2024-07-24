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

    const totalPages = Math.ceil((rows.length - 7) / 2);

    // MAIN SLIDE

    // SLIDES DESING

    for (let i = 7; i < rows.length; i += 2) {
      const slide = pres.addSlide();

      // Establecer el color de fondo de la diapositiva
      slide.background = { color: "616161" }; // Cambia 'FFFFFF' por el color que desees

      // Añadir la primera actividad (arriba izquierda)
      if (rows[i]) {
        slide.addText(
          [
            { text: `${i - 6}. `, options: { color: "#F39200", fontSize: 14 } }, // Color del número
            { text: rows[i], options: { color: "FFFFFF", fontSize: 14 } }, // Color del texto
          ],
          { x: 0.3, y: 2.5, align: "left" }
        );
      }

      // Añadir la segunda actividad (centro izquierda)
      if (rows[i + 1]) {
        slide.addText(
          [
            { text: `${i - 5}. `, options: { color: "#F39200", fontSize: 14 } }, // Color del número
            { text: rows[i + 1], options: { color: "FFFFFF", fontSize: 14 } }, // Color del texto
          ],
          { x: 0.3, y: 5.5, align: "left" }
        );
      }

      // Agregar una franja naranja al top de la diapositiva
      slide.addShape(pres.ShapeType.rect, {
        x: 0.5,
        y: 0.8,
        w: 7.4,
        h: 0.05,
        fill: { color: "F39200" }, // Color naranja
      });

      // Agregar una franja naranja al pie de la diapositiva
      slide.addShape(pres.ShapeType.rect, {
        x: 0,
        y: 10.6,
        w: 8.5,
        h: 0.23,
        fill: { color: "F39200" }, // Color naranja
      });

      // Calcular el número de página
      const pageNumber = Math.ceil((i - 5) / 2);

      // Agregar el número de página en la franja naranja
      slide.addText(`Página ${pageNumber} de ${totalPages}`, {
        x: 1.4,
        y: 10.71,
        fontSize: 14,
        color: "FFFFFF", // Color del texto (blanco)
        align: "right",
      });
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

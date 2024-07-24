const ExcelJS = require("exceljs");
const PptxGenJS = require("pptxgenjs");
const path = require("path");

// excel path
const excelPath =
  "assets/CONTROL 678_2024-1 - EVALUACIÓN - STP - CAMIÓN ALJIBE - LYLG-45.xlsx";

// new pptx
const pres = new PptxGenJS();
pres.defineLayout({ name: "Carta", width: 8.5, height: 11 });
pres.layout = "Carta";

// reading xlsx
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

      // background slide color
      slide.background = { color: "616161" };

      // adding activitie n°1
      if (rows[i]) {
        slide.addText(
          [
            { text: `${i - 6}. `, options: { color: "#F39200", fontSize: 14 } },
            { text: rows[i], options: { color: "FFFFFF", fontSize: 14 } },
          ],
          { x: 0.3, y: 2.5, w: 8, h: 0.5, align: "left" }
        );
      }

      // adding activitie n°2
      if (rows[i + 1]) {
        slide.addText(
          [
            { text: `${i - 5}. `, options: { color: "#F39200", fontSize: 14 } },
            { text: rows[i + 1], options: { color: "FFFFFF", fontSize: 14 } },
          ],
          { x: 0.3, y: 6.5, w: 8, h: 0.5, align: "left" }
        );
      }

      // adding separator orange top
      slide.addShape(pres.ShapeType.rect, {
        x: 0.5,
        y: 0.8,
        w: 7.4,
        h: 0.05,
        fill: { color: "F39200" },
      });

      // adding separator dark gary above activitie n°1
      slide.addShape(pres.ShapeType.rect, {
        x: 0.3,
        y: 2.3,
        w: 7.9,
        h: 0.05,
        fill: { color: "404040" },
      });

      // adding separator dark gary above activitie n°2
      slide.addShape(pres.ShapeType.rect, {
        x: 0.3,
        y: 6.3,
        w: 7.9,
        h: 0.05,
        fill: { color: "404040" },
      });

      // adding separator orange bottom
      slide.addShape(pres.ShapeType.rect, {
        x: 0,
        y: 10.6,
        w: 8.5,
        h: 0.23,
        fill: { color: "F39200" },
      });

      // page number footer
      const pageNumber = Math.ceil((i - 5) / 2);

      // page number footer desing
      slide.addText(`Página ${pageNumber} de ${totalPages}`, {
        x: 1.4,
        y: 10.71,
        fontSize: 14,
        color: "FFFFFF",
        align: "right",
      });
    }
    // pptx patht
    const excelDir = path.dirname(excelPath);
    const pptPath = path.join(excelDir, "archivo.pptx");

    // save pptx
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

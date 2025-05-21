/* global Office */
/* global Word, Excel, PowerPoint, performance, console */

async function addFooter(message) {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getFooter(Word.HeaderFooterType.primary)
      .insertParagraph(`From Agent: ${message}`, "End");

    await context.sync();
  });
}

async function fillColor(cell, color) {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  });
}

async function addTextToSlide(text) {
  await PowerPoint.run(async (context) => {
    context.presentation.slides.getItemAt(0).shapes.addTextBox(text, {
      left: Math.random() * 200,
      top: Math.random() * 200,
      height: 150,
      width: 150,
    });
    await context.sync();
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("addfooter", async (message) => {
      const start = performance.now();
      const { Footer: footer } = JSON.parse(message);
      await addFooter(footer);
      const duration = performance.now() - start;
      const result = `Demo add-in: Footer added! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });
  } else if (info.host === Office.HostType.Excel) {
    Office.actions.associate("fillcolor", async (message) => {
      const start = performance.now();
      const { Cell: cell, Color: color } = JSON.parse(message);
      await fillColor(cell, color);
      const duration = performance.now() - start;
      const result = `Demo add-in: Action completed! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });
  } else if (info.host === Office.HostType.PowerPoint) {
    Office.actions.associate("addtexttoslide", async (message) => {
      const start = performance.now();
      const { Text: text } = JSON.parse(message);
      await addTextToSlide(text);
      const duration = performance.now() - start;
      const result = `Demo add-in: text added to slide! completed in ${duration.toFixed(0)} ms.`;
      console.log(`Returning result: "${result}"`);
      return result;
    });
  }
});

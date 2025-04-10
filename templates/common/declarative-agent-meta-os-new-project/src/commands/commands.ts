/* global Office */
/* global Excel, performance, console */

async function fillColor(cell, color) {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  });
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
Office.onReady((info) => {
  Office.actions.associate("fillcolor", async (message) => {
    const start = performance.now();
    const { Cell: cell, Color: color } = JSON.parse(message);
    await fillColor(cell, color);
    const duration = performance.now() - start;
    const result = `Demo add-in: Action completed! completed in ${duration.toFixed(0)} ms.`;
    console.log(`Returning result: "${result}"`);
    return result;
  });
});

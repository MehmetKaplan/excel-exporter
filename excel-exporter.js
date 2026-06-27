#!/usr/bin/env node
/**
 * excel-exporter.js
 *
 * Exports a single sheet from an Excel workbook into a new Excel file.
 * The output file is named after the exported sheet (sheet name + .xlsx).
 *
 * Modes:
 *   --break-links  (default)  keep same-sheet formulas; replace any formula
 *                             referencing another sheet with its last value.
 *   --values-only             replace EVERY formula with its computed value.
 *
 * Usage:
 *   node excel-exporter.js <input.xlsx> "<Sheet Name>" [outputDir] [--break-links | --values-only]
 *
 * Requires: exceljs  (npm install exceljs)
 */

const path = require('path');
const ExcelJS = require('exceljs');

function fail(msg) {
  console.error('Error: ' + msg);
  process.exit(1);
}

async function main() {
  const argv = process.argv.slice(2);
  const flags = argv.filter(a => a.startsWith('--'));
  const positional = argv.filter(a => !a.startsWith('--'));
  const [inputFile, sheetName, outputDir = '.'] = positional;

  const usage =
    'Usage: node excel-exporter.js <input.xlsx> "<Sheet Name>" [outputDir] [mode]\n' +
    '  mode (one of, default --break-links):\n' +
    '    --break-links   break cross-sheet references, keep same-sheet formulas\n' +
    '    --values-only   replace every formula with its computed value';

  if (!inputFile || !sheetName) fail(usage);

  const known = ['--break-links', '--values-only'];
  const unknown = flags.filter(f => !known.includes(f));
  if (unknown.length) fail(`Unknown option(s): ${unknown.join(', ')}\n${usage}`);
  if (flags.includes('--break-links') && flags.includes('--values-only')) {
    fail('Choose only one of --break-links or --values-only.');
  }

  const mode = flags.includes('--values-only') ? 'values' : 'links';

  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(inputFile);

  const srcSheet = srcWb.getWorksheet(sheetName);
  if (!srcSheet) {
    fail(`Sheet "${sheetName}" not found. Available: ${srcWb.worksheets.map(w => w.name).join(', ')}`);
  }

  // Build a matcher that detects references to OTHER sheets.
  // A same-sheet reference may appear unqualified (A1) or qualified with this
  // sheet's own name (Sheet1!A1 or 'Sheet 1'!A1). Anything qualified with a
  // different sheet name is a cross-sheet reference.
  const escaped = sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // Matches  Name!  or  'Name'!  preceding a reference.
  const sheetQualifier = /(?:'([^']+)'|([A-Za-z_\u00A0-\uFFFF][A-Za-z0-9_.\u00A0-\uFFFF ]*))!/g;

  function referencesOtherSheet(formula) {
    sheetQualifier.lastIndex = 0;
    let m;
    while ((m = sheetQualifier.exec(formula)) !== null) {
      const ref = (m[1] !== undefined ? m[1] : m[2]).trim();
      if (ref !== sheetName) return true;
    }
    return false;
  }

  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet(sheetName, {
    properties: srcSheet.properties,
    views: srcSheet.views,
  });

  // Copy column widths/styles.
  srcSheet.columns.forEach((col, i) => {
    if (col && col.width) outSheet.getColumn(i + 1).width = col.width;
  });

  // Copy merged cells.
  if (srcSheet.model && srcSheet.model.merges) {
    srcSheet.model.merges.forEach(range => outSheet.mergeCells(range));
  }

  srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const outRow = outSheet.getRow(rowNumber);
    if (row.height) outRow.height = row.height;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const outCell = outRow.getCell(colNumber);
      const v = cell.value;

      if (v && typeof v === 'object' && v.formula !== undefined) {
        // It's a formula cell.
        if (mode === 'values') {
          // Strip all formulas: write only the computed value.
          outCell.value = v.result !== undefined ? v.result : null;
        } else if (referencesOtherSheet(v.formula)) {
          // Break the reference: keep the last computed result.
          outCell.value = v.result !== undefined ? v.result : null;
        } else if (v.shareType === 'array') {
          // Array (CSE) formula: preserve the array context so Excel does NOT
          // insert an implicit-intersection "@" before range references.
          outCell.value = {
            formula: v.formula,
            result: v.result,
            shareType: 'array',
            ref: v.ref || cell.address,
          };
        } else {
          // Same-sheet formula: keep it.
          outCell.value = { formula: v.formula, result: v.result };
        }
      } else if (v && typeof v === 'object' && v.sharedFormula !== undefined) {
        // Shared formula resolved by exceljs still exposes .formula above in
        // most cases; fall back to its result if only sharedFormula is present.
        outCell.value = v.result !== undefined ? v.result : null;
      } else {
        outCell.value = v;
      }

      // Preserve style. Clone it — exceljs shares style objects between cells,
      // and assigning the shared reference can drop fills/fonts on write.
      if (cell.style) outCell.style = JSON.parse(JSON.stringify(cell.style));
    });
    outRow.commit();
  });

  // Copy data validations (dropdown "list of values", numeric/date limits, etc.).
  // These live at sheet level, keyed by cell address, not on the cell itself.
  if (srcSheet.dataValidations && srcSheet.dataValidations.model) {
    outSheet.dataValidations.model = JSON.parse(
      JSON.stringify(srcSheet.dataValidations.model)
    );
  }

  // Copy conditional formatting (color scales, "highlight cells" rules, etc.).
  // Rule-based coloring is NOT stored on the cell, so it must be copied here.
  if (Array.isArray(srcSheet.conditionalFormattings)) {
    srcSheet.conditionalFormattings.forEach(cf => {
      outSheet.addConditionalFormatting(JSON.parse(JSON.stringify(cf)));
    });
  }

  const outName = `${sheetName}.xlsx`;
  const outPath = path.join(outputDir, outName);
  await outWb.xlsx.writeFile(outPath);
  console.log(`Exported "${sheetName}" -> ${outPath}`);
}

main().catch(err => fail(err.message));

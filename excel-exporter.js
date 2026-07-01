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
const JSZip = require('jszip');
const fs = require('fs');

function fail(msg) {
  console.error('Error: ' + msg);
  process.exit(1);
}

// --- Conditional-formatting restoration (raw XML) -------------------------
// ExcelJS loses dxf borders that use theme/indexed colors, which is exactly
// what Excel writes for "format the border via a formula" rules. To preserve
// them, we graft the source's <dxfs> table and the source sheet's
// <conditionalFormatting> blocks into the freshly written output file.

async function sheetPathForName(zip, name) {
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetMatch = [...wbXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)]
    .find(x => x[1] === name);
  if (!sheetMatch) return null;
  const rel = [...relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)]
    .find(x => x[1] === sheetMatch[2]);
  if (!rel) return null;
  const target = rel[2].replace(/^\/?xl\//, '').replace(/^\.\//, '');
  return 'xl/' + target;
}

function extractDxfs(stylesXml) {
  const m = stylesXml.match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/);
  if (!m) return { count: 0, inner: '' };
  return { count: (m[1].match(/<dxf\b/g) || []).length, inner: m[1] };
}

async function restoreConditionalFormatting(srcPath, outPath, sheetName, opts = {}) {
  const srcZip = await JSZip.loadAsync(fs.readFileSync(srcPath));
  const outZip = await JSZip.loadAsync(fs.readFileSync(outPath));

  const srcSheetPath = await sheetPathForName(srcZip, sheetName);
  const outSheetPath = await sheetPathForName(outZip, sheetName);
  if (!srcSheetPath || !outSheetPath) return false;

  const srcSheetXml = await srcZip.file(srcSheetPath).async('string');
  const srcCfBlocks =
    srcSheetXml.match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g) || [];
  const hasSrcDv = /<dataValidations\b/.test(srcSheetXml);
  // Nothing to restore (no classic CF and no data validations) → leave as-is.
  // (Note: x14/extLst CF is not handled here.)
  if (!srcCfBlocks.length && !hasSrcDv) return false;

  const srcStyles = await srcZip.file('xl/styles.xml').async('string');
  const outStyles = await outZip.file('xl/styles.xml').async('string');
  const srcDxfs = extractDxfs(srcStyles);
  const outDxfs = extractDxfs(outStyles);
  const offset = outDxfs.count;

  const mergedCount = outDxfs.count + srcDxfs.count;
  const mergedInner = outDxfs.inner + srcDxfs.inner;
  const dxfsEl = `<dxfs count="${mergedCount}">${mergedInner}</dxfs>`;
  let newStyles;
  if (/<dxfs\b[^>]*\/>/.test(outStyles)) {
    newStyles = outStyles.replace(/<dxfs\b[^>]*\/>/, dxfsEl);
  } else if (/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/.test(outStyles)) {
    newStyles = outStyles.replace(/<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/, dxfsEl);
  } else if (mergedCount > 0) {
    newStyles = /<tableStyles\b/.test(outStyles)
      ? outStyles.replace(/<tableStyles\b/, dxfsEl + '<tableStyles')
      : outStyles.replace(/<\/styleSheet>/, dxfsEl + '</styleSheet>');
  } else {
    newStyles = outStyles;
  }

  // Remap dxfId by the offset; optionally drop rules whose formula references
  // another sheet (so broken links don't leave dangling cross-sheet rules).
  const remapped = srcCfBlocks.map(block => {
    let b = block.replace(/dxfId="(\d+)"/g, (_, n) => `dxfId="${parseInt(n, 10) + offset}"`);
    if (opts.breakLinks) {
      b = b.replace(/<cfRule\b[\s\S]*?(?:\/>|<\/cfRule>)/g, rule => {
        const formulas = [...rule.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map(x => x[1]);
        const refsOther = formulas.some(f => {
          const decoded = f.replace(/&apos;/g, "'");
          const qualifiers = [...decoded.matchAll(/(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!/g)];
          return qualifiers.some(q => (q[1] || q[2]) !== sheetName);
        });
        return refsOther ? '' : rule;
      });
      if (!/<cfRule\b/.test(b)) return '';
    }
    return b;
  }).filter(Boolean);

  let outSheetXml = await outZip.file(outSheetPath).async('string');

  // --- Data validations: replace ExcelJS's (duplicated/overlapping) block with
  // the source's verbatim block. Strip xr:uid attributes, which rely on a
  // namespace the output may not declare. Under breakLinks, drop validations
  // whose formula references another sheet.
  let srcDvBlock = (srcSheetXml.match(/<dataValidations\b[\s\S]*?<\/dataValidations>/) || [])[0];
  if (srcDvBlock) {
    srcDvBlock = srcDvBlock.replace(/\s+xr:uid="[^"]*"/g, '');
    if (opts.breakLinks) {
      srcDvBlock = srcDvBlock.replace(
        /<dataValidation\b[\s\S]*?(?:\/>|<\/dataValidation>)/g, dv => {
          const formulas = [...dv.matchAll(/<formula\d?>([\s\S]*?)<\/formula\d?>/g)].map(x => x[1]);
          const refsOther = formulas.some(f => {
            const decoded = f.replace(/&apos;/g, "'").replace(/&quot;/g, '"');
            const qualifiers = [...decoded.matchAll(/(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!/g)];
            return qualifiers.some(q => (q[1] || q[2]) !== sheetName);
          });
          return refsOther ? '' : dv;
        });
      // Recount after possible drops.
      const remaining = (srcDvBlock.match(/<dataValidation\b/g) || []).length;
      srcDvBlock = srcDvBlock.replace(/(<dataValidations\b[^>]*\bcount=")\d+(")/, `$1${remaining}$2`);
      if (remaining === 0) srcDvBlock = '';
    }
  }
  // Remove whatever data validations ExcelJS wrote.
  outSheetXml = outSheetXml.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/g, '');

  // Remove any CF ExcelJS wrote (we re-insert our faithful copy).
  outSheetXml = outSheetXml.replace(
    /<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, ''
  );

  // Re-insert in schema order: conditionalFormatting, then dataValidations,
  // both after sheetData and before pageMargins.
  const cfInsert = remapped.join('');
  const anchor = /<pageMargins\b/.test(outSheetXml) ? '<pageMargins'
               : /<\/worksheet>/.test(outSheetXml) ? '</worksheet>' : null;
  const payload = cfInsert + (srcDvBlock || '');
  if (payload && anchor) {
    outSheetXml = outSheetXml.replace(
      anchor === '<pageMargins' ? /<pageMargins\b/ : /<\/worksheet>/,
      payload + anchor
    );
  }

  outZip.file('xl/styles.xml', newStyles);
  outZip.file(outSheetPath, outSheetXml);
  fs.writeFileSync(outPath, await outZip.generateAsync({ type: 'nodebuffer' }));
  return true;
}
// -------------------------------------------------------------------------


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

  // Discover which tables and defined names would still RESOLVE in a
  // single-sheet output. Anything that wouldn't (a table living on another
  // sheet, a defined name scoped to another sheet, or any workbook-scoped
  // defined name — since we drop the workbook's defined names entirely) must be
  // frozen to its last computed value, exactly like a cross-sheet cell
  // reference. Structured references such as RolesAndCosts[Code] carry NO
  // "Sheet!" qualifier, so the sheetQualifier scan below cannot catch them;
  // without this step they survive as live formulas pointing at a table that no
  // longer exists, and Excel reports #NAME?/#REF! (the "broken output").
  const localTableNames = new Set();   // tables physically on the exported sheet
  const localDefinedNames = new Set(); // names still valid in the output
  try {
    const probeZip = await JSZip.loadAsync(fs.readFileSync(inputFile));
    const srcSheetPath = await sheetPathForName(probeZip, sheetName);

    // Tables that belong to the exported sheet (via its worksheet rels).
    if (srcSheetPath) {
      const relName =
        'xl/worksheets/_rels/' + srcSheetPath.split('/').pop() + '.rels';
      const relFile = probeZip.file(relName);
      if (relFile) {
        const relXml = await relFile.async('string');
        const targets = [...relXml.matchAll(/Target="([^"]*tables\/table\d+\.xml)"/g)]
          .map(x => x[1].replace(/^\/?xl\//, '').replace(/^\.\.\//, '').replace(/^\.\//, ''));
        for (let t of targets) {
          const tPath = t.startsWith('xl/') ? t : 'xl/' + t.replace(/^worksheets\/\.\.\//, '');
          const tFile = probeZip.file(tPath) || probeZip.file('xl/' + t);
          if (tFile) {
            const tXml = await tFile.async('string');
            const nm = tXml.match(/<table\b[^>]*\bname="([^"]+)"/);
            if (nm) localTableNames.add(nm[1]);
          }
        }
      }
    }

    // Defined names: keep only those whose scope (localSheetId) is the exported
    // sheet. Workbook-scoped names are dropped from the output, so treat them as
    // external too. (The output carries no <definedNames> at all today; this is
    // future-proofing for when a same-sheet name genuinely resolves.)
    const wbXml = await probeZip.file('xl/workbook.xml').async('string');
    const sheetOrder = [...wbXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)].map(x => x[1]);
    const exportedSheetIndex = sheetOrder.indexOf(sheetName);
    for (const m of wbXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
      const attrs = m[1];
      const nameM = attrs.match(/\bname="([^"]+)"/);
      if (!nameM) continue;
      const localM = attrs.match(/\blocalSheetId="(\d+)"/);
      // Skip built-in names like _xlnm.Print_Area; they aren't referenced by
      // user formulas as bare identifiers.
      if (nameM[1].startsWith('_xlnm.')) continue;
      if (localM && parseInt(localM[1], 10) === exportedSheetIndex) {
        localDefinedNames.add(nameM[1]);
      }
    }
  } catch (e) {
    // If probing fails, fall back to the conservative behaviour: treat every
    // structured/defined-name reference as external (frozen), which is always
    // safe for a single-sheet export.
  }

  // Build a matcher that detects references to OTHER sheets.
  // A same-sheet reference may appear unqualified (A1) or qualified with this
  // sheet's own name (Sheet1!A1 or 'Sheet 1'!A1). Anything qualified with a
  // different sheet name is a cross-sheet reference.
  const escaped = sheetName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // Matches  Name!  or  'Name'!  preceding a reference.
  const sheetQualifier = /(?:'([^']+)'|([A-Za-z_\u00A0-\uFFFF][A-Za-z0-9_.\u00A0-\uFFFF ]*))!/g;

  // Excel/OOXML built-in function names that ExcelJS exposes with an _xlfn.
  // (and sometimes _xlws.) prefix. These are NOT table/defined-name references
  // and must never be treated as external.
  const FUNCTION_PREFIX = /^_xl(fn|ws)\./;

  // A structured table reference looks like  Name[...]  possibly with the
  // _xlfn. junk stripped. Capture the identifier that precedes a "[".
  function referencesMissingTableOrName(formula) {
    // Structured table references:  TableName[Column]  /  TableName[[#Data],...]
    const tableRefs = [...formula.matchAll(/([A-Za-z_\u00A0-\uFFFF][A-Za-z0-9_.\u00A0-\uFFFF]*)\s*\[/g)];
    for (const t of tableRefs) {
      let ident = t[1];
      // A leading _xlfn. is a function-call artefact (e.g. _xlfn.SUMIFS is not a
      // table). Strip a single known function prefix before judging.
      const bare = ident.replace(FUNCTION_PREFIX, '');
      // If what precedes "[" is a recognised worksheet function, it's a call,
      // not a table. We can't enumerate every function, but a real structured
      // reference's identifier is followed immediately by "[" with no "(" — and
      // functions are followed by "(". Since we matched "[", treat any
      // identifier that is NOT a local table as an unresolvable table ref.
      if (bare && !localTableNames.has(bare)) {
        // Guard: skip if this "[" is actually an array/reference construct that
        // happens to follow a function name with no space (rare). Structured
        // references never resolve without their table, so freezing is safe.
        return true;
      }
    }
    // Bare defined-name references (e.g.  Project_Start ,  task_start ). Match
    // identifiers that are not immediately followed by "(" (function call),
    // not preceded by "!" (already sheet-qualified, handled elsewhere), and
    // not part of a cell address (letters+digits like A1, V33).
    const identRe = /(?<![!:.\w])([A-Za-z_\u00A0-\uFFFF][A-Za-z0-9_.\u00A0-\uFFFF]*)/g;
    let im;
    while ((im = identRe.exec(formula)) !== null) {
      const ident = im[1];
      const after = formula[identRe.lastIndex];
      if (after === '(' ) continue;              // function call
      if (after === '[') continue;               // handled as table ref above
      if (FUNCTION_PREFIX.test(ident)) continue; // function artefact
      if (/^[A-Za-z]{1,3}\$?\d+$/.test(ident)) continue; // cell address A1/$A$1-ish
      if (/^(TRUE|FALSE)$/i.test(ident)) continue;
      // A pure column/row like "A", "AZ" isn't a defined name; skip single/double
      // letter tokens with no digits only when they look like columns.
      if (/^[A-Za-z]{1,3}$/.test(ident)) continue;
      if (localDefinedNames.has(ident)) continue; // resolves locally
      // Only flag identifiers that we KNOW are defined names elsewhere in the
      // workbook; otherwise we'd freeze ordinary text. We conservatively flag
      // only when the identifier is a known non-local defined name.
      if (knownDefinedNames.has(ident)) return true;
    }
    return false;
  }

  // Set of every user defined name in the workbook (any scope), used to decide
  // whether a bare identifier in a formula is a defined-name reference.
  const knownDefinedNames = new Set();
  try {
    const probeZip2 = await JSZip.loadAsync(fs.readFileSync(inputFile));
    const wbXml2 = await probeZip2.file('xl/workbook.xml').async('string');
    for (const m of wbXml2.matchAll(/<definedName\b[^>]*\bname="([^"]+)"/g)) {
      if (!m[1].startsWith('_xlnm.')) knownDefinedNames.add(m[1]);
    }
  } catch (e) { /* best effort */ }

  function referencesOtherSheet(formula) {
    sheetQualifier.lastIndex = 0;
    let m;
    while ((m = sheetQualifier.exec(formula)) !== null) {
      const ref = (m[1] !== undefined ? m[1] : m[2]).trim();
      if (ref !== sheetName) return true;
    }
    return false;
  }

  // A formula must be frozen to its last value if it references another sheet
  // OR a table / defined name that won't exist in the single-sheet output.
  function referencesExternal(formula) {
    return referencesOtherSheet(formula) || referencesMissingTableOrName(formula);
  }

  const outWb = new ExcelJS.Workbook();
  const outSheet = outWb.addWorksheet(sheetName, {
    properties: srcSheet.properties,
    views: srcSheet.views,
  });

  // Copy column widths and grouping (collapse/expand outline).
  srcSheet.columns.forEach((col, i) => {
    if (!col) return;
    const outCol = outSheet.getColumn(i + 1);
    if (col.width) outCol.width = col.width;
    if (col.outlineLevel) outCol.outlineLevel = col.outlineLevel;
    if (col.collapsed) outCol.collapsed = true;
    if (col.hidden) outCol.hidden = true;
  });

  // Copy merged cells.
  if (srcSheet.model && srcSheet.model.merges) {
    srcSheet.model.merges.forEach(range => outSheet.mergeCells(range));
  }

  srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const outRow = outSheet.getRow(rowNumber);
    if (row.height) outRow.height = row.height;
    // Row grouping (collapse/expand outline).
    if (row.outlineLevel) outRow.outlineLevel = row.outlineLevel;
    if (row.collapsed) outRow.collapsed = true;
    if (row.hidden) outRow.hidden = true;

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const outCell = outRow.getCell(colNumber);
      const v = cell.value;

      if (v && typeof v === 'object' && v.formula !== undefined) {
        // It's a formula cell.
        if (mode === 'values') {
          // Strip all formulas: write only the computed value.
          outCell.value = v.result !== undefined ? v.result : null;
        } else if (referencesExternal(v.formula)) {
          // Break the reference (cross-sheet, or a table / defined name that
          // won't exist in the single-sheet output): keep the last computed
          // result so the exported cell shows a value instead of #NAME?/#REF!.
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

  // Outline/summary properties. ExcelJS does not reliably populate
  // outlineLevelRow/Col on read, so compute them from the actual levels; the
  // group brackets won't render unless these are set. summaryBelow/Right are
  // copied when explicitly present (they default to true when undefined).
  let maxRowLevel = 0;
  outSheet.eachRow({ includeEmpty: true }, r => {
    if (r.outlineLevel > maxRowLevel) maxRowLevel = r.outlineLevel;
  });
  let maxColLevel = 0;
  outSheet.columns.forEach(c => {
    if (c && c.outlineLevel > maxColLevel) maxColLevel = c.outlineLevel;
  });
  if (maxRowLevel) outSheet.properties.outlineLevelRow = maxRowLevel;
  if (maxColLevel) outSheet.properties.outlineLevelCol = maxColLevel;

  const sp = srcSheet.properties || {};
  if (sp.summaryBelow !== undefined) outSheet.properties.summaryBelow = sp.summaryBelow;
  if (sp.summaryRight !== undefined) outSheet.properties.summaryRight = sp.summaryRight;

  // NOTE: Data validations are intentionally NOT copied through the ExcelJS
  // object model. ExcelJS re-emits them as duplicated, overlapping sqref ranges
  // (e.g. AU6:BL54 plus AU10:BL54), which Excel rejects with a repair warning.
  // They are restored verbatim from the source XML in the post-write pass.

  // NOTE: Conditional formatting is intentionally NOT copied through the
  // ExcelJS object model here. ExcelJS drops dxf borders that use theme/indexed
  // colors (common in Excel-authored "format border by formula" rules). Instead
  // it is restored faithfully from the source XML in a post-write pass below
  // (restoreConditionalFormatting), which preserves fills, borders, fonts,
  // color scales, data bars and icon sets exactly as Excel wrote them.

  const outName = `${sheetName}.xlsx`;
  const outPath = path.join(outputDir, outName);
  fs.mkdirSync(outputDir, { recursive: true });
  await outWb.xlsx.writeFile(outPath);

  // Restore conditional formatting from source XML (preserves dxf borders).
  await restoreConditionalFormatting(inputFile, outPath, sheetName, {
    breakLinks: mode === 'links',
  });

  console.log(`Exported "${sheetName}" -> ${outPath}`);
}

main().catch(err => fail(err.message));

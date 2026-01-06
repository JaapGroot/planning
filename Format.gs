function readMappedTextFormats_(sheet, startRow, numRows) {
  const rA = sheet.getRange(startRow, 1, numRows, 1);   // A
  const rHK = sheet.getRange(startRow, 8, numRows, 4);  // H-K
  const rM = sheet.getRange(startRow, 13, numRows, 1);  // M

  function merge6_(a1, hk4, m1) {
    const out = new Array(numRows);
    for (let i = 0; i < numRows; i++) out[i] = [a1[i][0], hk4[i][0], hk4[i][1], hk4[i][2], hk4[i][3], m1[i][0]];
    return out;
  }

  return {
    num: merge6_(rA.getNumberFormats(), rHK.getNumberFormats(), rM.getNumberFormats()),
    family: merge6_(rA.getFontFamilies(), rHK.getFontFamilies(), rM.getFontFamilies()),
    size: merge6_(rA.getFontSizes(), rHK.getFontSizes(), rM.getFontSizes()),
    weight: merge6_(rA.getFontWeights(), rHK.getFontWeights(), rM.getFontWeights()),
    style: merge6_(rA.getFontStyles(), rHK.getFontStyles(), rM.getFontStyles()),
    color: merge6_(rA.getFontColors(), rHK.getFontColors(), rM.getFontColors()),
    hAlign: merge6_(rA.getHorizontalAlignments(), rHK.getHorizontalAlignments(), rM.getHorizontalAlignments()),
    vAlign: merge6_(rA.getVerticalAlignments(), rHK.getVerticalAlignments(), rM.getVerticalAlignments()),
    wrap: merge6_(rA.getWrapStrategies(), rHK.getWrapStrategies(), rM.getWrapStrategies()),
  };
}

function readSingleColTextFormats_(sheet, startRow, numRows, col) {
  const r = sheet.getRange(startRow, col, numRows, 1);
  return {
    num: r.getNumberFormats(),
    family: r.getFontFamilies(),
    size: r.getFontSizes(),
    weight: r.getFontWeights(),
    style: r.getFontStyles(),
    color: r.getFontColors(),
    hAlign: r.getHorizontalAlignments(),
    vAlign: r.getVerticalAlignments(),
    wrap: r.getWrapStrategies(),
  };
}

(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.TouTemplate = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  function monthIndexFromLabel(value) {
    const text = String(value || "").trim();
    const match = text.match(/(?:^|[-年])(\d{1,2})月?$/) || text.match(/^(\d{1,2})月$/);
    if (!match) return -1;
    const monthIndex = Number(match[1]) - 1;
    return monthIndex >= 0 && monthIndex < 12 ? monthIndex : -1;
  }

  function normalizePackage(value) {
    return String(value || "").includes("带现货") ? "spot" : "term";
  }

  function toFiniteNumber(value) {
    if (value === null || value === undefined || String(value).trim() === "") return NaN;
    return Number(value);
  }

  function findSheet(workbook, name) {
    return workbook.sheets.find((sheet) => sheet.name === name);
  }

  function findHeaderRow(rows, requiredLabels) {
    return rows.findIndex((row) =>
      requiredLabels.every((label) => row.some((value) => String(value || "").trim() === label)),
    );
  }

  function findAdjacentValue(sheet, label) {
    for (const row of sheet?.rows || []) {
      const index = row.findIndex((value) => String(value || "").trim() === label);
      if (index < 0) continue;
      for (let column = index + 1; column < row.length; column += 1) {
        if (row[column] !== null && row[column] !== undefined && row[column] !== "") return row[column];
      }
    }
    return undefined;
  }

  function findBelowValue(sheet, label) {
    for (let rowIndex = 0; rowIndex < (sheet?.rows || []).length; rowIndex += 1) {
      const column = sheet.rows[rowIndex].findIndex((value) => String(value || "").trim() === label);
      if (column < 0) continue;
      for (let next = rowIndex + 1; next < sheet.rows.length; next += 1) {
        const value = sheet.rows[next][column];
        if (value !== null && value !== undefined && value !== "") return value;
      }
    }
    return undefined;
  }

  function parseBandCode(value) {
    const text = String(value || "").trim();
    if (text.includes("尖峰")) return "S";
    if (text.includes("高峰")) return "H";
    if (text.includes("低谷")) return "V";
    if (text.includes("平段")) return "P";
    return "";
  }

  function parseHourlyMonthMatrix(sheet) {
    const headerIndex = findHeaderRow(sheet.rows, ["时段", "1月"]);
    if (headerIndex < 0) throw new Error(`${sheet.name} 未找到“时段/月份”表头。`);
    const header = sheet.rows[headerIndex];
    const monthColumns = header
      .map((value, index) => ({ index, monthIndex: monthIndexFromLabel(value) }))
      .filter((item) => item.monthIndex >= 0);
    const hourRows = sheet.rows
      .slice(headerIndex + 1)
      .filter((row) => /^\d{2}:\d{2}-\d{2}:\d{2}$/.test(String(row[0] || "").trim()))
      .slice(0, 24);
    if (hourRows.length !== 24 || !monthColumns.length) {
      throw new Error(`${sheet.name} 未识别到完整的 24 小时数据。`);
    }
    return { monthColumns, hourRows };
  }

  function parsePriceWorkbook(workbook, fileName, currentData) {
    const periodSheet = findSheet(workbook, "国网峰平谷时段") || findSheet(workbook, "峰平谷时段");
    const termSheet = findSheet(workbook, "旬及以上基础价");
    const spotSheet = findSheet(workbook, "带现货基础价");
    const gridSheet = findSheet(workbook, "国网分时电价");
    if (!periodSheet || !termSheet || !spotSheet || !gridSheet) {
      throw new Error("模板需要包含国网峰平谷时段（兼容原名“峰平谷时段”）、旬及以上基础价、带现货基础价、国网分时电价四个工作表。");
    }

    const period = parseHourlyMonthMatrix(periodSheet);
    const titleText = periodSheet.rows.flat().map((value) => String(value || "")).join(" ");
    const yearMatch = titleText.match(/(?:19|20)\d{2}/);
    const analysisYear = yearMatch ? Number(yearMatch[0]) : currentData.analysisYear;
    const bandCodes = Array(12).fill(null);
    period.monthColumns.forEach(({ index, monthIndex }) => {
      const codes = period.hourRows.map((row) => parseBandCode(row[index]));
      if (codes.every(Boolean)) bandCodes[monthIndex] = codes.join("");
    });
    const parseBaseSheet = (sheet) => {
      const parsed = parseHourlyMonthMatrix(sheet);
      const matrix = Array(12).fill(null);
      parsed.monthColumns.forEach(({ index, monthIndex }) => {
        const values = parsed.hourRows.map((row) => toFiniteNumber(row[index]));
        if (values.every(Number.isFinite)) matrix[monthIndex] = values;
      });
      return matrix;
    };
    const termPrices = parseBaseSheet(termSheet);
    const spotPrices = parseBaseSheet(spotSheet);

    const gridHeaderIndex = findHeaderRow(gridSheet.rows, ["月份", "平段基价", "折价电费"]);
    if (gridHeaderIndex < 0) throw new Error("国网分时电价未找到月份、平段基价和折价电费。");
    const gridFormulaText = gridSheet.rows
      .flat()
      .map((value) => String(value || "").trim())
      .find((value) => value.includes("计算公式")) || currentData.gridFormulaText;
    const gridHeader = gridSheet.rows[gridHeaderIndex].map((value) => String(value || "").trim());
    const monthColumn = gridHeader.indexOf("月份");
    const baseColumn = gridHeader.indexOf("平段基价");
    const adjustmentColumn = gridHeader.indexOf("折价电费");
    const levelColumns = {
      S: gridHeader.indexOf("尖峰电价"),
      H: gridHeader.indexOf("高峰电价"),
      P: gridHeader.indexOf("平段电价"),
      V: gridHeader.indexOf("低谷电价"),
    };
    const gridInputs = Array(12).fill(null);
    const gridLevels = Array(12).fill(null);
    gridSheet.rows.slice(gridHeaderIndex + 1).forEach((row) => {
      const monthIndex = monthIndexFromLabel(row[monthColumn]);
      const base = toFiniteNumber(row[baseColumn]);
      const adjustment = toFiniteNumber(row[adjustmentColumn]);
      if (monthIndex < 0 || !Number.isFinite(base) || !Number.isFinite(adjustment)) return;
      gridInputs[monthIndex] = [base, adjustment];
      const calculated = {
        S: base * 1.92 + adjustment,
        H: base * 1.6 + adjustment,
        P: base + adjustment,
        V: base * 0.45 + adjustment,
      };
      gridLevels[monthIndex] = Object.fromEntries(
        Object.entries(levelColumns).map(([code, column]) => {
          const parsed = column >= 0 ? toFiniteNumber(row[column]) : NaN;
          return [code, Number.isFinite(parsed) ? parsed : calculated[code]];
        }),
      );
    });

    const comparisonSheet = findSheet(workbook, "套餐电价对比");
    const savingsSheet = findSheet(workbook, "电费节省分析");
    const packageText = findBelowValue(comparisonSheet, "当前选择")
      ?? findBelowValue(savingsSheet, "当前选择/数值");
    const markupValue = findAdjacentValue(comparisonSheet, "统一加价(元/MWh)")
      ?? findAdjacentValue(savingsSheet, "统一加价(元/MWh)");
    const priceMonthIndexes = Array.from({ length: 12 }, (_, index) => index).filter(
      (index) => termPrices[index] && spotPrices[index] && gridInputs[index] && bandCodes[index],
    );
    if (!priceMonthIndexes.length) throw new Error("模板中没有可联动计算的完整月份。");

    return {
      ...currentData,
      sourceName: fileName,
      analysisYear,
      gridFormulaText,
      bandCodes,
      termPrices,
      spotPrices,
      gridInputs,
      gridLevels,
      priceMonthIndexes,
      coveredMonths: priceMonthIndexes.length,
      defaultPackage: normalizePackage(packageText),
      defaultMarkup: Number.isFinite(Number(markupValue)) ? Number(markupValue) : 20,
    };
  }

  return { parsePriceWorkbook };
});

const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const statusMessage = document.querySelector("#status-message");
const fileList = document.querySelector("#file-list");
const fileListItems = document.querySelector("#file-list-items");
const clearFilesButton = document.querySelector("#clear-files");
const configPanel = document.querySelector("#config-panel");
const sourceSummary = document.querySelector("#source-summary");
const sheetField = document.querySelector("#sheet-field");
const sheetSelect = document.querySelector("#sheet-select");
const companyColumn = document.querySelector("#company-column");
const dateColumn = document.querySelector("#date-column");
const timeColumn = document.querySelector("#time-column");
const valueColumn = document.querySelector("#value-column");
const intervalMode = document.querySelector("#interval-mode");
const processButton = document.querySelector("#process-button");
const resultPanel = document.querySelector("#result-panel");
const resultSummary = document.querySelector("#result-summary");
const qualityMessage = document.querySelector("#quality-message");
const resultTable = document.querySelector("#result-table");
const previewNote = document.querySelector("#preview-note");
const downloadCsvButton = document.querySelector("#download-csv");
const downloadXlsxButton = document.querySelector("#download-xlsx");
const filterCompany = document.querySelector("#filter-company");
const filterMonthStart = document.querySelector("#filter-month-start");
const filterMonthEnd = document.querySelector("#filter-month-end");
const resetFiltersButton = document.querySelector("#reset-filters");
const curvePanel = document.querySelector("#curve-panel");
const curveCompany = document.querySelector("#curve-company");
const curveType = document.querySelector("#curve-type");
const curveMonthField = document.querySelector("#curve-month-field");
const curveStartMonthField = document.querySelector("#curve-start-month-field");
const curveEndMonthField = document.querySelector("#curve-end-month-field");
const curveMonth = document.querySelector("#curve-month");
const curveStartMonth = document.querySelector("#curve-start-month");
const curveEndMonth = document.querySelector("#curve-end-month");
const curveSummary = document.querySelector("#curve-summary");
const curveTotal = document.querySelector("#curve-total");
const curvePeak = document.querySelector("#curve-peak");
const curveAverage = document.querySelector("#curve-average");
const curveChart = document.querySelector("#curve-chart");
const downloadCurveButton = document.querySelector("#download-curve");

if ("serviceWorker" in navigator && ["http:", "https:"].includes(window.location.protocol)) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch((error) =>
      console.warn("离线缓存未启用", error),
    );
  });
}

const mappingSelects = {
  company: companyColumn,
  date: dateColumn,
  time: timeColumn,
  value: valueColumn,
};

let sourceFiles = [];
let parsedSources = [];
let fileEntries = [];
let currentSheet = null;
let currentAnalysis = null;
let currentResult = null;

fileInput.addEventListener("change", () => {
  const files = [...fileInput.files];
  if (files.length) loadFiles(files);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("is-dragging");
  });
});

dropZone.addEventListener("drop", (event) => {
  const files = [...event.dataTransfer.files];
  if (files.length) loadFiles(files);
});

clearFilesButton.addEventListener("click", clearFiles);
sheetSelect.addEventListener("change", () => selectSingleSheet(Number(sheetSelect.value)));
processButton.addEventListener("click", processCurrentSheet);
intervalMode.addEventListener("change", () => currentSheet && processCurrentSheet());
downloadCsvButton.addEventListener("click", downloadCsv);
downloadXlsxButton.addEventListener("click", downloadXlsx);
[filterCompany, filterMonthStart, filterMonthEnd].forEach((control) => {
  control.addEventListener("change", () => currentSheet && processCurrentSheet());
});
resetFiltersButton.addEventListener("click", () => {
  filterCompany.value = "";
  filterMonthStart.value = "";
  filterMonthEnd.value = "";
  processCurrentSheet();
});
[curveType, curveMonth, curveStartMonth, curveEndMonth].forEach(
  (control) => control.addEventListener("change", updateCurve),
);
downloadCurveButton.addEventListener("click", downloadCurvePng);

async function loadFiles(files) {
  const supportedExtensions = new Set(["xls", "xlsx", "csv", "tsv"]);
  sourceFiles = files;
  parsedSources = [];
  currentSheet = null;
  currentResult = null;
  configPanel.hidden = true;
  resultPanel.hidden = true;
  curvePanel.hidden = true;
  document.querySelector("#tou-panel").hidden = true;

  fileEntries = files.map((file) => {
    const extension = file.name.split(".").pop().toLowerCase();
    return {
      file,
      extension,
      status: supportedExtensions.has(extension) ? "等待读取" : "不支持此格式",
      isError: !supportedExtensions.has(extension),
    };
  });
  renderFileList();

  const readableEntries = fileEntries.filter((entry) => !entry.isError);
  if (!readableEntries.length) {
    showStatus("请选择 .xls、.xlsx、.csv 或 .tsv 文件。", true);
    return;
  }

  for (let index = 0; index < readableEntries.length; index += 1) {
    const entry = readableEntries[index];
    entry.status = "正在读取";
    renderFileList();
    showStatus(`正在读取第 ${index + 1}/${readableEntries.length} 个文件：${entry.file.name}`);

    try {
      const workbook = await parseFile(entry.file, entry.extension);
      parsedSources.push({ file: entry.file, workbook });
      const firstSheetRows = workbook.sheets[0]?.rows.length || 0;
      entry.status = `${workbook.sheets.length} 个工作表，约 ${Math.max(0, firstSheetRows - 1).toLocaleString("zh-CN")} 行`;
    } catch (error) {
      console.error(error);
      entry.status = error.message || "读取失败";
      entry.isError = true;
    }
    renderFileList();
  }

  if (!parsedSources.length) {
    showStatus("没有成功读取的文件，请检查文件格式。", true);
    return;
  }

  prepareSources();
  const failedCount = fileEntries.filter((entry) => entry.isError).length;
  showStatus(
    failedCount
      ? `成功读取 ${parsedSources.length} 个文件，${failedCount} 个文件未处理。`
      : `已合并读取 ${parsedSources.length} 个文件。`,
    failedCount > 0,
  );
  configPanel.hidden = false;
  configPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function parseFile(file, extension) {
  if (extension === "xls" || extension === "xlsx") {
    return HourlyEngine.parseXlsx(await file.arrayBuffer(), JSZip);
  }
  return HourlyEngine.parseCsv(decodeText(await file.arrayBuffer()));
}

function decodeText(arrayBuffer) {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(arrayBuffer);
  } catch {
    try {
      return new TextDecoder("gb18030").decode(arrayBuffer);
    } catch {
      return new TextDecoder().decode(arrayBuffer);
    }
  }
}

function renderFileList() {
  fileList.hidden = fileEntries.length === 0;
  const fragment = document.createDocumentFragment();

  fileEntries.forEach((entry) => {
    const item = document.createElement("li");
    const name = document.createElement("span");
    const status = document.createElement("span");
    name.textContent = entry.file.name;
    name.title = entry.file.name;
    status.textContent = entry.status;
    status.classList.toggle("file-error", entry.isError);
    item.append(name, status);
    fragment.append(item);
  });

  fileListItems.replaceChildren(fragment);
}

function clearFiles() {
  sourceFiles = [];
  parsedSources = [];
  fileEntries = [];
  currentSheet = null;
  currentAnalysis = null;
  currentResult = null;
  fileInput.value = "";
  fileList.hidden = true;
  fileListItems.replaceChildren();
  configPanel.hidden = true;
  resultPanel.hidden = true;
  curvePanel.hidden = true;
  document.querySelector("#tou-panel").hidden = true;
  showStatus("已清空，可以重新选择多个文件。", false);
}

function prepareSources() {
  if (parsedSources.length === 1) {
    populateSheetSelector(parsedSources[0].workbook);
    selectSingleSheet(0);
    return;
  }

  sheetField.hidden = true;
  currentSheet = HourlyEngine.combineSheets(
    parsedSources.map((source) => {
      const sheet = HourlyEngine.normalizeSourceSheet(source.workbook.sheets[0]);
      return {
        name: source.file.name,
        sheetName: sheet.name,
        rows: sheet.rows,
      };
    }),
  );
  configureCurrentSheet(
    `${parsedSources.length} 个文件 · ${Math.max(0, currentSheet.rows.length - 1).toLocaleString("zh-CN")} 行合并数据 · 每个文件的第一个工作表`,
  );
}

function populateSheetSelector(workbook) {
  sheetSelect.replaceChildren();
  workbook.sheets.forEach((sheet, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = sheet.name;
    sheetSelect.append(option);
  });
  sheetField.hidden = workbook.sheets.length <= 1;
}

function selectSingleSheet(index) {
  const source = parsedSources[0];
  const originalSheet = source.workbook.sheets[index];
  currentSheet = HourlyEngine.normalizeSourceSheet(originalSheet);
  const layoutLabel = currentSheet.sourceLayout === "wide"
    ? `横向分时格式 · ${currentSheet.sourceTimeColumnCount} 个时点列`
    : "逐时点明细格式";
  configureCurrentSheet(
    `${source.file.name} · ${originalSheet.name} · ${layoutLabel} · ${Math.max(0, originalSheet.rows.length - 1).toLocaleString("zh-CN")} 行源数据`,
  );
}

function configureCurrentSheet(summary) {
  currentAnalysis = HourlyEngine.analyzeSheet(currentSheet.rows);
  sourceSummary.textContent = summary;
  populateMappingSelectors(currentAnalysis.headers);
  filterCompany.value = "";
  filterMonthStart.value = "";
  filterMonthEnd.value = "";

  const mapping = getMapping();
  if (Object.values(mapping).every((columnIndex) => columnIndex >= 0)) {
    processCurrentSheet();
  } else {
    resultPanel.hidden = true;
    showStatus("文件已读取，请确认企业、日期、时点和电量字段。", false);
  }
}

function populateMappingSelectors(headers) {
  const guesses = HourlyEngine.guessMapping(headers);

  Object.entries(mappingSelects).forEach(([key, select]) => {
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "-1";
    placeholder.textContent = "请选择";
    select.append(placeholder);

    headers.forEach((header, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = header;
      select.append(option);
    });

    select.value = String(guesses[key]);
  });
}

function getMapping() {
  return Object.fromEntries(
    Object.entries(mappingSelects).map(([key, select]) => [key, Number(select.value)]),
  );
}

function processCurrentSheet() {
  const mapping = getMapping();
  if (Object.values(mapping).some((columnIndex) => columnIndex < 0)) {
    showStatus("请先选择企业／对象名称、日期、时点和电量数值列。", true);
    return;
  }

  if (filterMonthStart.value && filterMonthEnd.value && filterMonthStart.value > filterMonthEnd.value) {
    showStatus("开始月份不能晚于结束月份。", true);
    return;
  }

  try {
    const discoveryResult = HourlyEngine.processData(currentSheet.rows, mapping, {
      headerIndex: currentAnalysis.headerIndex,
      outputMode: "daily",
      intervalMode: intervalMode.value,
      dateStart: filterMonthStart.value ? `${filterMonthStart.value}-01` : "",
      dateEnd: filterMonthEnd.value ? getMonthEnd(filterMonthEnd.value) : "",
    });
    const selectedCompany = filterCompany.value || discoveryResult.meta.availableCompanies[0] || "";
    const monthlyResult = HourlyEngine.buildMonthlySummary(currentSheet.rows, mapping, {
      headerIndex: currentAnalysis.headerIndex,
      intervalMode: intervalMode.value,
      companyFilter: selectedCompany,
      dateStart: filterMonthStart.value ? `${filterMonthStart.value}-01` : "",
      dateEnd: filterMonthEnd.value ? getMonthEnd(filterMonthEnd.value) : "",
    });
    currentResult = HourlyEngine.transposeMonthlySummary(monthlyResult);
    renderResult(currentResult);
    showStatus("月度 24 小时数据已生成，可以检查预览并导出。", false);
  } catch (error) {
    console.error(error);
    showStatus(`处理失败：${error.message || "请检查字段选择"}`, true);
  }
}

function renderResult(result) {
  const { meta } = result;
  resultPanel.hidden = false;
  document.querySelector("#file-count").textContent = parsedSources.length.toLocaleString("zh-CN");
  document.querySelector("#valid-rows").textContent = meta.validRows.toLocaleString("zh-CN");
  document.querySelector("#company-count").textContent = meta.companyCount.toLocaleString("zh-CN");
  document.querySelector("#month-count").textContent = meta.monthCount.toLocaleString("zh-CN");
  document.querySelector("#output-count").textContent = result.rows.length.toLocaleString("zh-CN");

  const monthRange = meta.monthStart
    ? meta.monthStart === meta.monthEnd
      ? meta.monthStart
      : `${meta.monthStart} 至 ${meta.monthEnd}`
    : "未识别月份";
  resultSummary.textContent = `${meta.selectedCompany || "未选择企业"} · 各月份按列展示 · ${monthRange}`;
  populateFilterControls(meta);
  renderQuality(meta);
  renderTable(result);
  prepareCurveControls(meta);
  window.dispatchEvent(new CustomEvent("hourly-data-updated", {
    detail: {
      sheet: currentSheet,
      analysis: currentAnalysis,
      mapping: getMapping(),
      intervalMode: intervalMode.value,
      company: filterCompany.value || meta.selectedCompany,
    },
  }));
}

function populateFilterControls(meta) {
  const previousCompany = filterCompany.value;
  filterCompany.replaceChildren();
  meta.availableCompanies.forEach((company) => {
    const option = document.createElement("option");
    option.value = company;
    option.textContent = company;
    filterCompany.append(option);
  });
  filterCompany.value = meta.availableCompanies.includes(previousCompany)
    ? previousCompany
    : meta.availableCompanies.includes(meta.selectedCompany)
      ? meta.selectedCompany
      : meta.availableCompanies[0] || "";

  const minimumMonth = meta.availableMonths[0] || "";
  const maximumMonth = meta.availableMonths[meta.availableMonths.length - 1] || "";
  [filterMonthStart, filterMonthEnd].forEach((input) => {
    input.min = minimumMonth;
    input.max = maximumMonth;
    if (input.value && (input.value < minimumMonth || input.value > maximumMonth)) input.value = "";
  });
}

function prepareCurveControls(meta) {
  const previousMonth = curveMonth.value;
  const previousStartMonth = curveStartMonth.value;
  const previousEndMonth = curveEndMonth.value;
  const selectedCompany = filterCompany.value || meta.selectedCompany || "";
  const companyDatesResult = selectedCompany
    ? HourlyEngine.processData(currentSheet.rows, getMapping(), {
        headerIndex: currentAnalysis.headerIndex,
        outputMode: "daily",
        intervalMode: intervalMode.value,
        companyFilter: selectedCompany,
      })
    : null;
  const companyMonths = [...new Set(
    (companyDatesResult?.meta.selectedDates || []).map((date) => date.slice(0, 7)),
  )];

  setSelectValues(curveMonth, companyMonths);
  setSelectValues(curveStartMonth, companyMonths);
  setSelectValues(curveEndMonth, companyMonths);

  curveCompany.textContent = selectedCompany || "—";
  curveMonth.value = companyMonths.includes(previousMonth)
    ? previousMonth
    : companyMonths[companyMonths.length - 1] || "";
  curveStartMonth.value = companyMonths.includes(previousStartMonth)
    ? previousStartMonth
    : companyMonths[0] || "";
  curveEndMonth.value = companyMonths.includes(previousEndMonth)
    ? previousEndMonth
    : companyMonths[companyMonths.length - 1] || "";

  curvePanel.hidden = !selectedCompany;
  if (!curvePanel.hidden) updateCurve();
}

function setSelectValues(select, values) {
  const fragment = document.createDocumentFragment();
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    fragment.append(option);
  });
  select.replaceChildren(fragment);
}

function updateCurve() {
  const selectedCompany = filterCompany.value;
  if (!currentSheet || !selectedCompany) return;
  const companyDatesResult = HourlyEngine.processData(currentSheet.rows, getMapping(), {
    headerIndex: currentAnalysis.headerIndex,
    outputMode: "daily",
    intervalMode: intervalMode.value,
    companyFilter: selectedCompany,
  });
  const companyDates = companyDatesResult.meta.selectedDates;
  const companyMonths = [...new Set(companyDates.map((date) => date.slice(0, 7)))];
  preserveSelectValue(curveMonth, companyMonths, companyMonths[companyMonths.length - 1] || "");
  preserveSelectValue(curveStartMonth, companyMonths, companyMonths[0] || "");
  preserveSelectValue(curveEndMonth, companyMonths, companyMonths[companyMonths.length - 1] || "");

  const type = curveType.value;
  curveMonthField.hidden = type !== "daily";
  curveStartMonthField.hidden = type !== "multi";
  curveEndMonthField.hidden = type !== "multi";

  let label = "";
  let series = [];
  let chartOptions = { dense: false, showLegend: false };

  if (type === "daily") {
    const datesInMonth = companyDates.filter((date) => date.startsWith(`${curveMonth.value}-`));
    series = datesInMonth
      .map((date, index) => {
        const dailyResult = HourlyEngine.processData(currentSheet.rows, getMapping(), {
          headerIndex: currentAnalysis.headerIndex,
          outputMode: "daily",
          intervalMode: intervalMode.value,
          companyFilter: selectedCompany,
          dateStart: date,
          dateEnd: date,
        });
        return {
          label: date,
          values: dailyResult.rows[0]?.slice(2, 26) || [],
          color: getCurveColor(index, "daily"),
        };
      })
      .filter((item) => item.values.some((value) => typeof value === "number"));
    label = `${selectedCompany} · ${curveMonth.value} · ${series.length} 天日曲线叠加`;
    chartOptions = { dense: true, showLegend: false };
  } else if (type === "monthly") {
    series = companyMonths
      .map((month, index) => {
        const monthlyResult = HourlyEngine.processData(currentSheet.rows, getMapping(), {
          headerIndex: currentAnalysis.headerIndex,
          outputMode: "sum",
          intervalMode: intervalMode.value,
          companyFilter: selectedCompany,
          dateStart: `${month}-01`,
          dateEnd: getMonthEnd(month),
        });
        return {
          label: month,
          values: monthlyResult.rows[0]?.slice(1, 25) || [],
          color: getCurveColor(index, "monthly"),
        };
      })
      .filter((item) => item.values.some((value) => typeof value === "number"));
    label = `${selectedCompany} · ${series.length} 个月月度曲线对比`;
    chartOptions = { dense: series.length > 12, showLegend: true };
  } else {
    if (curveStartMonth.value > curveEndMonth.value) {
      curveEndMonth.value = curveStartMonth.value;
    }
    label = `${selectedCompany} · ${curveStartMonth.value} 至 ${curveEndMonth.value} 多月总曲线`;
    const curveResult = HourlyEngine.processData(currentSheet.rows, getMapping(), {
      headerIndex: currentAnalysis.headerIndex,
      outputMode: "sum",
      intervalMode: intervalMode.value,
      companyFilter: selectedCompany,
      dateStart: `${curveStartMonth.value}-01`,
      dateEnd: getMonthEnd(curveEndMonth.value),
    });
    series = [
      {
        label,
        values: curveResult.rows[0]?.slice(1, 25) || [],
        color: "#245c46",
      },
    ];
  }

  const points = series.flatMap((item) =>
    item.values.flatMap((value, hourIndex) =>
      typeof value === "number" ? [{ value, hourIndex }] : [],
    ),
  );
  const total = points.reduce((sum, point) => sum + point.value, 0);
  const peakPoint = points.reduce(
    (highest, point) => (!highest || point.value > highest.value ? point : highest),
    null,
  );

  curveSummary.textContent = label;
  curveTotal.textContent = formatNumber(total);
  curvePeak.textContent = peakPoint
    ? `${formatNumber(peakPoint.value)} MWh（第 ${peakPoint.hourIndex + 1} 小时）`
    : "无数据";
  curveAverage.textContent = formatNumber(points.length ? total / points.length : 0);
  renderCurveChart(series, label, chartOptions);
}

function preserveSelectValue(select, values, fallback) {
  const previousValue = select.value;
  setSelectValues(select, values);
  select.value = values.includes(previousValue) ? previousValue : fallback;
}

function getMonthEnd(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  const finalDay = new Date(year, monthNumber, 0).getDate();
  return `${month}-${String(finalDay).padStart(2, "0")}`;
}

function createSvgElement(name, attributes = {}, text = "") {
  const element = document.createElementNS("http://www.w3.org/2000/svg", name);
  Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, String(value)));
  if (text !== "") element.textContent = text;
  return element;
}

function getCurveColor(index, type) {
  const monthlyPalette = [
    "#245c46",
    "#d97706",
    "#2563eb",
    "#be123c",
    "#7c3aed",
    "#0891b2",
    "#65a30d",
    "#c2410c",
    "#4f46e5",
    "#a21caf",
    "#0f766e",
    "#ca8a04",
  ];
  if (type === "monthly" && index < monthlyPalette.length) return monthlyPalette[index];
  const hue = Math.round((index * 137.508 + (type === "daily" ? 12 : 205)) % 360);
  return `hsl(${hue} 68% 44%)`;
}

function renderCurveChart(series, label, options = {}) {
  curveChart.replaceChildren();
  curveChart.dataset.hasData = "false";
  curveChart.setAttribute("aria-label", label);
  const width = 960;
  const legendColumns = 6;
  const legendRows = options.showLegend ? Math.ceil(series.length / legendColumns) : 0;
  const legendHeight = legendRows ? legendRows * 24 + 12 : 0;
  const height = 380 + legendHeight;
  const margin = { top: 28 + legendHeight, right: 34, bottom: 54, left: 76 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const numericValues = series.flatMap((item) =>
    item.values.filter((value) => typeof value === "number"),
  );

  curveChart.setAttribute("viewBox", `0 0 ${width} ${height}`);
  curveChart.dataset.chartWidth = String(width);
  curveChart.dataset.chartHeight = String(height);

  if (!numericValues.length) {
    curveChart.append(
      createSvgElement(
        "text",
        { x: width / 2, y: height / 2, "text-anchor": "middle", fill: "#68776f", "font-size": 15 },
        "当前范围没有可绘制的数据",
      ),
    );
    return;
  }

  curveChart.dataset.hasData = "true";
  let minimum = Math.min(0, ...numericValues);
  let maximum = Math.max(0, ...numericValues);
  if (minimum === maximum) maximum = minimum + 1;
  const padding = (maximum - minimum) * 0.08;
  maximum += padding;
  if (minimum < 0) minimum -= padding;

  const x = (index) => margin.left + (index / 23) * plotWidth;
  const y = (value) => margin.top + ((maximum - value) / (maximum - minimum)) * plotHeight;

  for (let tick = 0; tick <= 5; tick += 1) {
    const value = minimum + ((maximum - minimum) * tick) / 5;
    const yPosition = y(value);
    curveChart.append(
      createSvgElement("line", {
        x1: margin.left,
        y1: yPosition,
        x2: width - margin.right,
        y2: yPosition,
        stroke: "#dce5df",
        "stroke-width": 1,
      }),
      createSvgElement(
        "text",
        {
          x: margin.left - 12,
          y: yPosition + 4,
          "text-anchor": "end",
          fill: "#68776f",
          "font-size": 12,
        },
        formatNumber(value),
      ),
    );
  }

  for (let index = 0; index < 24; index += 1) {
    if (index % 2 !== 0 && index !== 23) continue;
    curveChart.append(
      createSvgElement(
        "text",
        {
          x: x(index),
          y: height - 22,
          "text-anchor": "middle",
          fill: "#68776f",
          "font-size": 12,
        },
        `${index + 1}时`,
      ),
    );
  }

  series.forEach((item) => {
    let pathData = "";
    let segmentOpen = false;
    item.values.forEach((value, index) => {
      if (typeof value !== "number") {
        segmentOpen = false;
        return;
      }
      pathData += `${segmentOpen ? " L" : " M"} ${x(index)} ${y(value)}`;
      segmentOpen = true;
    });
    if (!pathData) return;

    const path = createSvgElement("path", {
      d: pathData.trim(),
      fill: "none",
      stroke: item.color,
      "stroke-width": options.dense ? 1.6 : series.length === 1 ? 3 : 2.2,
      "stroke-opacity": options.dense ? 0.58 : 0.92,
      "stroke-linecap": "round",
      "stroke-linejoin": "round",
    });
    if (!options.dense) path.append(createSvgElement("title", {}, item.label));
    curveChart.append(path);

    if (series.length !== 1) return;
    item.values.forEach((value, index) => {
      if (typeof value !== "number") return;
      const point = createSvgElement("circle", {
        cx: x(index),
        cy: y(value),
        r: 4,
        fill: "#ffffff",
        stroke: item.color,
        "stroke-width": 2,
      });
      point.append(
        createSvgElement("title", {}, `第 ${index + 1} 小时：${formatNumber(value)} MWh`),
      );
      curveChart.append(point);
    });
  });

  if (options.showLegend) {
    series.forEach((item, index) => {
      const column = index % legendColumns;
      const row = Math.floor(index / legendColumns);
      const itemWidth = plotWidth / legendColumns;
      const legendX = margin.left + column * itemWidth;
      const legendY = 22 + row * 24;
      curveChart.append(
        createSvgElement("line", {
          x1: legendX,
          y1: legendY,
          x2: legendX + 20,
          y2: legendY,
          stroke: item.color,
          "stroke-width": 3,
          "stroke-linecap": "round",
        }),
        createSvgElement(
          "text",
          { x: legendX + 27, y: legendY + 4, fill: "#42534a", "font-size": 12 },
          item.label,
        ),
      );
    });
  }

  curveChart.append(
    createSvgElement(
      "text",
      { x: 18, y: margin.top - 8, fill: "#68776f", "font-size": 12 },
      "MWh",
    ),
  );
}

function renderQuality(meta) {
  const messages = [];
  let level = "ok";
  const failedFiles = fileEntries.filter((entry) => entry.isError).length;
  const incompatibleFiles = (currentSheet.files || []).filter(
    (file) => file.missingHeaders.length > 0,
  );

  if (failedFiles) {
    messages.push(`${failedFiles} 个文件读取失败或格式不受支持`);
    level = "error";
  }

  if (incompatibleFiles.length) {
    messages.push(`${incompatibleFiles.length} 个文件存在缺失表头，缺失列数据已留空`);
    level = "error";
  }

  if (meta.invalidRows) {
    messages.push(`${meta.invalidRows} 行因企业、日期、时点或电量为空/无效而未参与计算`);
    level = "error";
  }

  if (meta.incompleteGroups) {
    messages.push(`${meta.incompleteGroups} 个企业日期组的时点少于本批次识别出的 ${meta.expectedSlots} 个时点`);
    if (level !== "error") level = "warning";
  }

  if (meta.difference != null && meta.difference > 0.00001) {
    messages.push(`输入与输出合计相差 ${meta.difference} MWh`);
    level = "error";
  }

  if (!messages.length) {
    messages.push(
      `校验通过：${parsedSources.length} 个文件已合并，识别为${meta.intervalMode === "end" ? "区间结束" : "区间开始"}时点，输入与输出电量合计一致`,
    );
  }

  qualityMessage.textContent = `${messages.join("；")}。`;
  qualityMessage.classList.toggle("has-warning", level === "warning");
  qualityMessage.classList.toggle("has-error", level === "error");
}

function renderTable(result) {
  const headRow = document.createElement("tr");
  result.headers.forEach((header) => {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = header;
    headRow.append(cell);
  });
  resultTable.tHead.replaceChildren(headRow);

  const previewRows = result.rows.slice(0, 100);
  const bodyFragment = document.createDocumentFragment();
  previewRows.forEach((row) => {
    const tableRow = document.createElement("tr");
    row.forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = typeof value === "number" ? formatNumber(value) : value ?? "";
      tableRow.append(cell);
    });
    bodyFragment.append(tableRow);
  });
  resultTable.tBodies[0].replaceChildren(bodyFragment);

  previewNote.textContent =
    result.rows.length > previewRows.length
      ? `页面预览前 ${previewRows.length} 行，下载文件包含全部 ${result.rows.length} 行。`
      : `共 ${result.rows.length} 行，页面已显示全部结果。`;
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 6,
  }).format(value);
}

function showStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("is-error", isError);
}

function getSelectedCompanyDownloadResult() {
  const company = filterCompany.value;
  if (!company) {
    showStatus("请先在“选择企业／对象”中单选一项，再下载月度数据。", true);
    filterCompany.focus();
    return null;
  }

  const monthlyResult = HourlyEngine.buildMonthlySummary(currentSheet.rows, getMapping(), {
    headerIndex: currentAnalysis.headerIndex,
    intervalMode: intervalMode.value,
    companyFilter: company,
  });
  if (!monthlyResult.rows.length) {
    showStatus("所选企业没有可导出的月度数据。", true);
    return null;
  }
  return HourlyEngine.transposeMonthlySummary(monthlyResult);
}

function makeMonthlyOutputName(extension) {
  const safeCompany = filterCompany.value.replace(/[\\/:*?"<>|]/g, "_");
  return `${safeCompany}_月度24小时分时数据.${extension}`;
}

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadCsv() {
  if (!currentResult) return;
  const monthlyResult = getSelectedCompanyDownloadResult();
  if (!monthlyResult) return;
  const csv = `\uFEFF${HourlyEngine.toCsv(monthlyResult)}`;
  triggerDownload(
    new Blob([csv], { type: "text/csv;charset=utf-8" }),
    makeMonthlyOutputName("csv"),
  );
  showStatus(`已导出 ${filterCompany.value} 的 ${monthlyResult.meta.monthCount} 个月月度数据。`, false);
}

async function downloadXlsx() {
  if (!currentResult) return;
  downloadXlsxButton.disabled = true;
  downloadXlsxButton.textContent = "正在生成…";

  try {
    const monthlyResult = getSelectedCompanyDownloadResult();
    if (!monthlyResult) return;
    const bytes = await HourlyEngine.toXlsx(monthlyResult, JSZip);
    triggerDownload(
      new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      makeMonthlyOutputName("xlsx"),
    );
    showStatus(`已导出 ${filterCompany.value} 的 ${monthlyResult.meta.monthCount} 个月月度数据。`, false);
  } catch (error) {
    console.error(error);
    showStatus(`Excel 导出失败：${error.message}`, true);
  } finally {
    downloadXlsxButton.disabled = false;
    downloadXlsxButton.textContent = "下载所选企业 Excel";
  }
}

async function downloadCurvePng() {
  if (curveChart.dataset.hasData !== "true") {
    showStatus("当前曲线没有可下载的数据。", true);
    return;
  }

  const clonedSvg = curveChart.cloneNode(true);
  const chartWidth = Number(curveChart.dataset.chartWidth) || 960;
  const chartHeight = Number(curveChart.dataset.chartHeight) || 380;
  const exportScale = 2;
  clonedSvg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  clonedSvg.setAttribute("width", String(chartWidth * exportScale));
  clonedSvg.setAttribute("height", String(chartHeight * exportScale));
  const background = createSvgElement("rect", {
    x: 0,
    y: 0,
    width: chartWidth,
    height: chartHeight,
    fill: "#ffffff",
  });
  clonedSvg.insertBefore(background, clonedSvg.firstChild);

  const svgText = new XMLSerializer().serializeToString(clonedSvg);
  const svgUrl = URL.createObjectURL(new Blob([svgText], { type: "image/svg+xml;charset=utf-8" }));

  try {
    const image = new Image();
    await new Promise((resolve, reject) => {
      image.onload = resolve;
      image.onerror = reject;
      image.src = svgUrl;
    });

    const canvas = document.createElement("canvas");
    canvas.width = chartWidth * exportScale;
    canvas.height = chartHeight * exportScale;
    const context = canvas.getContext("2d");
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    const safeCompany = (filterCompany.value || "企业").replace(/[\\/:*?"<>|]/g, "_");
    const typeLabel = { daily: "整月日曲线", monthly: "全部月份月曲线", multi: "多月总曲线" }[
      curveType.value
    ];
    triggerDownload(pngBlob, `${safeCompany}_${typeLabel}.png`);
  } catch (error) {
    console.error(error);
    showStatus("曲线图片生成失败，请重试。", true);
  } finally {
    URL.revokeObjectURL(svgUrl);
  }
}

(function () {
  "use strict";

  let data = globalThis.TouPriceData;
  const engine = globalThis.HourlyEngine;
  const templateParser = globalThis.TouTemplate;
  if (!data || !engine || !templateParser) return;

  const panel = document.querySelector("#tou-panel");
  const companyLabel = document.querySelector("#tou-company");
  const templateInput = document.querySelector("#tou-template-input");
  const templateButton = document.querySelector("#tou-template-button");
  const templateName = document.querySelector("#tou-template-name");
  const packageSelect = document.querySelector("#tou-package");
  const markupInput = document.querySelector("#tou-markup");
  const gridFormulaText = document.querySelector("#tou-grid-formula-text");
  const note = document.querySelector("#tou-note");
  const tabs = [...document.querySelectorAll("[data-tou-tab]")];
  const views = [...document.querySelectorAll("[data-tou-panel]")];
  let templateHandle = null;
  let templateSignature = "";
  let templateTimer = null;
  let templateLoading = false;

  const state = {
    monthly: null,
    company: "",
    sourceMonthIndexes: [],
    availableMonthIndexes: [],
  };
  const allMonthIndexes = Array.from({ length: 12 }, (_, index) => index);

  const numberFormat = new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 3,
  });
  const moneyFormat = new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  function number(value) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function monthIndexFromKey(month) {
    const match = String(month || "").match(/-(\d{2})$/);
    return match ? Number(match[1]) - 1 : -1;
  }

  function getFileSignature(file) {
    return `${file.name}|${file.size}|${file.lastModified}`;
  }

  async function loadTemplateFile(file, options = {}) {
    if (templateLoading) return false;
    templateLoading = true;
    templateButton.disabled = true;
    templateButton.textContent = "正在读取…";
    try {
      const workbook = await engine.parseXlsx(await file.arrayBuffer(), JSZip);
      data = templateParser.parsePriceWorkbook(workbook, file.name, data);
      packageSelect.value = data.defaultPackage;
      markupInput.value = String(data.defaultMarkup);
      templateSignature = getFileSignature(file);
      templateName.textContent = options.watching
        ? `${file.name} · 自动监测中`
        : `${file.name} · 已载入（更新后请重新选择）`;
      renderBaseTable("#tou-term-table", data.termPrices);
      renderBaseTable("#tou-spot-table", data.spotPrices);
      refreshAll();
      globalThis.showStatus?.(
        options.auto
          ? `检测到模板已更新，已自动重新计算：${file.name}。`
          : `已读取最新电价分析模板：${file.name}。`,
        false,
      );
      return true;
    } catch (error) {
      console.error(error);
      templateName.textContent = "读取失败，继续使用上一次成功的数据";
      globalThis.showStatus?.(`电价分析模板读取失败：${error.message}`, true);
      return false;
    } finally {
      templateLoading = false;
      templateButton.disabled = false;
      templateButton.textContent = "选择或刷新 Excel";
      templateInput.value = "";
    }
  }

  function stopTemplateWatch() {
    if (templateTimer) clearInterval(templateTimer);
    templateTimer = null;
  }

  function startTemplateWatch() {
    stopTemplateWatch();
    if (!templateHandle) return;
    templateTimer = setInterval(async () => {
      if (templateLoading) return;
      try {
        const file = await templateHandle.getFile();
        if (getFileSignature(file) !== templateSignature) {
          await loadTemplateFile(file, { auto: true, watching: true });
        }
      } catch (error) {
        console.warn("模板自动监测已停止", error);
        stopTemplateWatch();
        templateName.textContent = "自动监测已停止，可重新选择 Excel";
      }
    }, 4000);
  }

  async function chooseTemplateFile() {
    if (typeof globalThis.showOpenFilePicker !== "function") {
      templateInput.click();
      return;
    }
    try {
      const [handle] = await globalThis.showOpenFilePicker({
        multiple: false,
        types: [{
          description: "Excel 工作簿",
          accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] },
        }],
      });
      const file = await handle.getFile();
      const loaded = await loadTemplateFile(file, { watching: true });
      if (loaded) {
        templateHandle = handle;
        startTemplateWatch();
      }
    } catch (error) {
      if (error?.name !== "AbortError") {
        console.error(error);
        globalThis.showStatus?.(`无法打开电价分析模板：${error.message}`, true);
      }
    }
  }

  function getPackageName() {
    return packageSelect.value === "spot" ? "带现货套餐" : "旬及以上套餐";
  }

  function getMarkup() {
    return Number.isFinite(Number(markupInput.value)) ? Number(markupInput.value) : 0;
  }

  function getBasePrices(monthIndex) {
    const source = packageSelect.value === "spot" ? data.spotPrices : data.termPrices;
    return source[monthIndex] || null;
  }

  function getPackagePrices(monthIndex) {
    const base = getBasePrices(monthIndex);
    if (!base) return null;
    const markup = getMarkup();
    return base.map((value) => value + markup);
  }

  function getGridLevels(monthIndex) {
    if (data.gridLevels?.[monthIndex]) return data.gridLevels[monthIndex];
    const input = data.gridInputs[monthIndex];
    if (!input) return null;
    const [base, adjustment] = input;
    return {
      S: base * 1.92 + adjustment,
      H: base * 1.6 + adjustment,
      P: base + adjustment,
      V: base * 0.45 + adjustment,
    };
  }

  function getGridPrices(monthIndex) {
    const levels = getGridLevels(monthIndex);
    const codes = data.bandCodes[monthIndex];
    if (!levels || !codes) return null;
    return [...codes].map((code) => levels[code]);
  }

  function getSourceEnergyByMonth() {
    const result = new Map();
    for (const row of state.monthly?.rows || []) {
      const monthIndex = monthIndexFromKey(row[1]);
      if (monthIndex < 0) continue;
      const values = row.slice(2, 26).map((value) => number(value));
      if (!result.has(monthIndex)) result.set(monthIndex, Array(24).fill(0));
      const totals = result.get(monthIndex);
      values.forEach((value, hourIndex) => { totals[hourIndex] += value; });
    }
    return result;
  }

  function getEnergyByMonth() {
    return new Map(
      [...getSourceEnergyByMonth()].filter(([monthIndex]) => data.priceMonthIndexes.includes(monthIndex)),
    );
  }

  function makeCell(tag, value, className = "") {
    const cell = document.createElement(tag);
    cell.textContent = value == null ? "—" : String(value);
    if (className) cell.className = className;
    return cell;
  }

  function renderTable(table, headers, rows, options = {}) {
    const head = document.createElement("tr");
    headers.forEach((header) => head.append(makeCell("th", header)));
    table.tHead.replaceChildren(head);
    const fragment = document.createDocumentFragment();
    rows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      row.forEach((value, columnIndex) => {
        const formatted = typeof value === "number"
          ? (options.moneyColumns?.includes(columnIndex) ? moneyFormat.format(value) : numberFormat.format(value))
          : value;
        const td = makeCell("td", formatted);
        if (value === null || value === undefined) td.classList.add("tou-placeholder");
        if (options.cellClass) {
          const className = options.cellClass(value, rowIndex, columnIndex);
          if (className) td.classList.add(className);
        }
        tr.append(td);
      });
      fragment.append(tr);
    });
    table.tBodies[0].replaceChildren(fragment);
  }

  function timeRows(matrix, monthIndexes) {
    return data.hours.map((hour, hourIndex) => [
      hour,
      ...monthIndexes.map((monthIndex) => matrix[monthIndex]?.[hourIndex] ?? null),
    ]);
  }

  function bandClass(value) {
    if (value === "S" || value === "尖峰") return "tou-band-super";
    if (value === "H" || value === "高峰") return "tou-band-peak";
    if (value === "V" || value === "低谷") return "tou-band-valley";
    if (value === "P" || value === "平段") return "tou-band-flat";
    return "";
  }

  function renderBaseTable(tableId, matrix) {
    renderTable(
      document.querySelector(tableId),
      ["时段", ...data.months],
      timeRows(matrix, allMonthIndexes),
    );
  }

  function renderEnergyTable() {
    const energyByMonth = getSourceEnergyByMonth();
    const rows = data.hours.map((hour, hourIndex) => {
      const availableValues = [...energyByMonth.values()].map((values) => values[hourIndex] || 0);
      const values = allMonthIndexes.map((monthIndex) => energyByMonth.get(monthIndex)?.[hourIndex] ?? null);
      const average = availableValues.length
        ? availableValues.reduce((sum, value) => sum + value, 0) / availableValues.length
        : null;
      return [hour, ...values, average];
    });
    renderTable(
      document.querySelector("#tou-energy-table"),
      ["时段", ...data.months, "月份平均值"],
      rows,
    );
  }

  function renderGridTables() {
    gridFormulaText.textContent = data.gridFormulaText || "暂无计算方式说明";
    const inputs = allMonthIndexes.map((monthIndex) => {
      const input = data.gridInputs[monthIndex];
      const levels = getGridLevels(monthIndex);
      if (!input || !levels) return [data.months[monthIndex], null, null, null, null, null, null];
      const [base, adjustment] = input;
      return [data.months[monthIndex], base, adjustment, levels.S, levels.H, levels.P, levels.V];
    });
    renderTable(
      document.querySelector("#tou-grid-input-table"),
      ["月份", "平段基价", "折价电费", "尖峰电价", "高峰电价", "平段电价", "低谷电价"],
      inputs,
    );
    const matrix = Array(12).fill(null);
    data.priceMonthIndexes.forEach((index) => { matrix[index] = getGridPrices(index); });
    renderTable(
      document.querySelector("#tou-grid-table"),
      ["时段", ...data.months],
      timeRows(matrix, allMonthIndexes),
      {
        cellClass(value, rowIndex, columnIndex) {
          if (!columnIndex || typeof value !== "number") return "";
          const monthIndex = allMonthIndexes[columnIndex - 1];
          return bandClass(data.bandCodes[monthIndex]?.[rowIndex]);
        },
      },
    );
  }

  function renderComparison() {
    const monthIndexes = data.priceMonthIndexes;
    const differenceMatrix = Array(12).fill(null);
    monthIndexes.forEach((monthIndex) => {
      const grid = getGridPrices(monthIndex);
      const packagePrices = getPackagePrices(monthIndex);
      differenceMatrix[monthIndex] = grid.map((value, hourIndex) => value - packagePrices[hourIndex]);
    });
    renderTable(
      document.querySelector("#tou-comparison-table"),
      ["时段", ...data.months],
      timeRows(differenceMatrix, allMonthIndexes),
      {
        cellClass(value, rowIndex, columnIndex) {
          if (!columnIndex || typeof value !== "number") return "";
          return value >= 0 ? "tou-positive" : "tou-negative";
        },
      },
    );
    const summary = allMonthIndexes.map((monthIndex) => {
      const values = differenceMatrix[monthIndex];
      if (!values) return [data.months[monthIndex], "待更新", "待更新", "待更新"];
      return [
        data.months[monthIndex],
        values.filter((value) => value >= 0).length,
        values.filter((value) => value < 0).length,
        values.reduce((sum, value) => sum + value, 0) / values.length,
      ];
    });
    renderTable(
      document.querySelector("#tou-comparison-summary-table"),
      ["月份", "正价差小时数", "负价差小时数", "24 小时平均价差"],
      summary,
    );
  }

  function drawLineChart(svg, series, label, unit) {
    svg.replaceChildren();
    svg.setAttribute("aria-label", label);
    const width = 960;
    const height = 400;
    const margin = { top: 62, right: 34, bottom: 58, left: 74 };
    const values = series.flatMap((item) => item.values.filter(Number.isFinite));
    if (!values.length) {
      svg.append(createSvg("text", { x: 480, y: 205, "text-anchor": "middle", fill: "#68776f" }, "当前月份没有可绘制的数据"));
      return;
    }
    let minimum = Math.min(0, ...values);
    let maximum = Math.max(...values);
    if (minimum === maximum) maximum = minimum + 1;
    const padding = (maximum - minimum) * 0.08;
    maximum += padding;
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;
    const x = (index) => margin.left + (index / 23) * plotWidth;
    const y = (value) => margin.top + ((maximum - value) / (maximum - minimum)) * plotHeight;

    for (let tick = 0; tick <= 5; tick += 1) {
      const value = minimum + ((maximum - minimum) * tick) / 5;
      const position = y(value);
      svg.append(
        createSvg("line", { x1: margin.left, y1: position, x2: width - margin.right, y2: position, stroke: "#dce5df", "stroke-width": 1 }),
        createSvg("text", { x: margin.left - 10, y: position + 4, "text-anchor": "end", fill: "#68776f", "font-size": 12 }, numberFormat.format(value)),
      );
    }
    for (let index = 0; index < 24; index += 1) {
      if (index % 2 && index !== 23) continue;
      svg.append(createSvg("text", { x: x(index), y: height - 22, "text-anchor": "middle", fill: "#68776f", "font-size": 12 }, `${index + 1}时`));
    }
    series.forEach((item) => {
      const path = item.values.map((value, index) => `${index ? "L" : "M"}${x(index)},${y(value)}`).join(" ");
      svg.append(createSvg("path", { d: path, fill: "none", stroke: item.color, "stroke-width": 3, "stroke-linejoin": "round", "stroke-linecap": "round" }));
    });
    series.forEach((item, index) => {
      const startX = margin.left + index * 180;
      svg.append(
        createSvg("line", { x1: startX, y1: 25, x2: startX + 28, y2: 25, stroke: item.color, "stroke-width": 4 }),
        createSvg("text", { x: startX + 36, y: 30, fill: "#42534a", "font-size": 13 }, item.label),
      );
    });
    svg.append(createSvg("text", { x: 16, y: margin.top - 8, fill: "#68776f", "font-size": 12 }, unit));
  }

  function createSvg(name, attributes = {}, text = "") {
    const element = document.createElementNS("http://www.w3.org/2000/svg", name);
    Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, String(value)));
    if (text) element.textContent = text;
    return element;
  }

  function renderPriceChart() {
    const monthIndex = state.availableMonthIndexes.at(-1) ?? data.priceMonthIndexes.at(-1) ?? 0;
    const grid = getGridPrices(monthIndex);
    const packagePrices = getPackagePrices(monthIndex);
    const title = `${data.months[monthIndex] || "所选月"}国网与${getPackageName()}24小时电价曲线`;
    document.querySelector("#tou-price-chart-title").textContent = title;
    drawLineChart(
      document.querySelector("#tou-price-chart"),
      grid && packagePrices
        ? [
            { label: "国网电价", values: grid, color: "#2563eb" },
            { label: getPackageName(), values: packagePrices, color: "#d97706" },
          ]
        : [],
      title,
      "元/MWh",
    );
  }

  function renderSavings() {
    const energyByMonth = getEnergyByMonth();
    const monthIndexes = [...energyByMonth.keys()].sort((a, b) => a - b);
    const rows = monthIndexes.map((monthIndex) => {
      const energy = energyByMonth.get(monthIndex);
      const grid = getGridPrices(monthIndex);
      const packagePrices = getPackagePrices(monthIndex);
      const totalEnergy = energy.reduce((sum, value) => sum + value, 0);
      const gridCost = energy.reduce((sum, value, hourIndex) => sum + value * grid[hourIndex], 0);
      const packageCost = energy.reduce((sum, value, hourIndex) => sum + value * packagePrices[hourIndex], 0);
      const saving = gridCost - packageCost;
      return [data.months[monthIndex], totalEnergy, gridCost, packageCost, saving, totalEnergy ? saving / totalEnergy : 0];
    });
    const totals = rows.reduce(
      (sum, row) => sum.map((value, index) => value + (index ? number(row[index]) : 0)),
      [0, 0, 0, 0, 0, 0],
    );
    totals[0] = "合计";
    totals[5] = totals[1] ? totals[4] / totals[1] : 0;
    renderTable(
      document.querySelector("#tou-savings-table"),
      ["月份", "企业月电量(MWh)", "国网月费用(元)", "套餐月费用(元)", "当月节省(元)", "度电节省(元/MWh)"],
      [...rows, totals],
      { moneyColumns: [2, 3, 4] },
    );

    document.querySelector("#tou-covered-months").textContent = monthIndexes.length;
    document.querySelector("#tou-energy-total").textContent = `${numberFormat.format(totals[1])} MWh`;
    document.querySelector("#tou-grid-cost").textContent = `${moneyFormat.format(totals[2])} 元`;
    document.querySelector("#tou-package-cost").textContent = `${moneyFormat.format(totals[3])} 元`;
    document.querySelector("#tou-saving-total").textContent = `${moneyFormat.format(totals[4])} 元`;
    const recommendation = document.querySelector("#tou-recommendation");
    recommendation.textContent = monthIndexes.length
      ? totals[4] > 0
        ? `测算结果：${getPackageName()}预计节省 ${moneyFormat.format(totals[4])} 元，按当前参数可优先考虑。`
        : `测算结果：${getPackageName()}预计增加 ${moneyFormat.format(Math.abs(totals[4]))} 元，按当前参数暂不建议选择。`
      : "当前企业没有落在模板覆盖月份内的月度电量，暂不能计算节省。";
    recommendation.classList.toggle("is-saving", totals[4] > 0 && monthIndexes.length > 0);
    recommendation.classList.toggle("is-costlier", totals[4] <= 0 && monthIndexes.length > 0);
  }

  function getAnalysisSnapshot() {
    const energyByMonth = getEnergyByMonth();
    const monthIndexes = [...energyByMonth.keys()].sort((a, b) => a - b);
    const months = monthIndexes.map((monthIndex) => {
      const energy = energyByMonth.get(monthIndex);
      const grid = getGridPrices(monthIndex);
      const packagePrices = getPackagePrices(monthIndex);
      const totalEnergy = energy.reduce((sum, value) => sum + value, 0);
      const gridCost = energy.reduce((sum, value, hourIndex) => sum + value * grid[hourIndex], 0);
      const packageCost = energy.reduce((sum, value, hourIndex) => sum + value * packagePrices[hourIndex], 0);
      return {
        monthIndex,
        month: data.months[monthIndex],
        energy,
        grid,
        packagePrices,
        totalEnergy,
        gridCost,
        packageCost,
        saving: gridCost - packageCost,
      };
    });
    return {
      sourceName: data.sourceName,
      analysisYear: data.analysisYear,
      company: state.company,
      packageType: packageSelect.value,
      packageName: getPackageName(),
      markup: getMarkup(),
      priceMonthIndexes: [...data.priceMonthIndexes],
      months,
    };
  }

  function refreshAll() {
    if (!state.monthly) return;
    companyLabel.textContent = state.company || "—";
    renderEnergyTable();
    renderGridTables();
    renderComparison();
    renderPriceChart();
    renderSavings();
    const priceMonths = data.priceMonthIndexes.map((index) => data.months[index]).join("、");
    note.textContent = state.sourceMonthIndexes.length
      ? `已按月份读取 ${state.company} 的 ${state.sourceMonthIndexes.length} 个月电量；不限制数据年份，当前模板可计算 ${priceMonths}。所有表格均保留 1—12 月位置。`
      : `当前企业没有可匹配的月度电量。模板已按 1—12 月保留位置，当前具备完整价格的月份为 ${priceMonths}。`;
  }

  function updateFromSource(detail) {
    const company = detail?.company;
    if (!company) return;
    state.monthly = engine.buildMonthlySummary(detail.sheet.rows, detail.mapping, {
      headerIndex: detail.analysis.headerIndex,
      intervalMode: detail.intervalMode,
      companyFilter: company,
    });
    state.company = company;
    state.sourceMonthIndexes = [...new Set(state.monthly.rows
      .map((row) => monthIndexFromKey(row[1]))
      .filter((index) => index >= 0))].sort((a, b) => a - b);
    state.availableMonthIndexes = state.sourceMonthIndexes
      .filter((index) => data.priceMonthIndexes.includes(index));
    panel.hidden = false;
    refreshAll();
  }

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((item) => {
        const active = item === tab;
        item.classList.toggle("is-active", active);
        item.setAttribute("aria-selected", String(active));
      });
      views.forEach((view) => {
        view.hidden = view.dataset.touPanel !== tab.dataset.touTab;
      });
      if (tab.dataset.touTab === "curves") renderPriceChart();
    });
  });
  packageSelect.addEventListener("change", refreshAll);
  markupInput.addEventListener("input", refreshAll);
  templateButton.addEventListener("click", chooseTemplateFile);
  templateInput.addEventListener("change", () => {
    const [file] = templateInput.files;
    if (file) {
      templateHandle = null;
      stopTemplateWatch();
      loadTemplateFile(file);
    }
  });

  renderBaseTable("#tou-term-table", data.termPrices);
  renderBaseTable("#tou-spot-table", data.spotPrices);
  packageSelect.value = data.defaultPackage;
  markupInput.value = String(data.defaultMarkup);
  templateName.textContent = `当前使用内置参考数据：${data.sourceName}`;
  window.addEventListener("hourly-data-updated", (event) => updateFromSource(event.detail));
  globalThis.TouAnalysis = {
    parsePriceWorkbook: (workbook, fileName) => templateParser.parsePriceWorkbook(workbook, fileName, data),
    getAnalysisSnapshot,
  };
})();

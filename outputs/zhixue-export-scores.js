/*
 * 智学网“校级报告 -> 学生成绩”导出脚本
 *
 * 用法：
 * 1. 打开智学网当前“学生成绩”页面，并保持已登录。
 * 2. 按 F12 / 右键检查，打开 Console 控制台。
 * 3. 粘贴本文件全部内容并回车。
 * 4. 等待脚本自动读取“全部班级”、翻页，并按班级自动拆分 sheet。
 */
(async () => {
  const CONFIG = {
    fileName: "智学网学生成绩导出.xlsx",
    waitMs: 900,
    maxPages: 80,
    sheetJsUrl: "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
  };

  const FALLBACK_HEADERS = [
    "序号",
    "准考证号",
    "姓名",
    "班级",
    "总分",
    "校次",
    "校次进退步",
    "班次",
    "班次进退步",
    "语文",
    "数学",
    "英语",
    "科学",
  ];

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const normalize = (text) => String(text || "").replace(/\s+/g, " ").trim();

  const isVisible = (el) => {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const style = window.getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
  };

  const findClickableByText = (text) => {
    const candidates = Array.from(document.querySelectorAll("a, button, li, span, div"));
    return candidates
      .filter(isVisible)
      .filter((el) => normalize(el.innerText || el.textContent) === text)
      .sort((a, b) => a.getBoundingClientRect().width - b.getBoundingClientRect().width)[0];
  };

  const waitForTableChange = async (previousSignature) => {
    for (let i = 0; i < 25; i += 1) {
      await sleep(200);
      const currentSignature = getPageSignature();
      if (currentSignature && currentSignature !== previousSignature) return;
    }
  };

  const clickText = async (text) => {
    const el = findClickableByText(text);
    if (!el) throw new Error(`没有找到可点击项：${text}`);
    const previousSignature = getPageSignature();
    el.click();
    await sleep(CONFIG.waitMs);
    await waitForTableChange(previousSignature);
  };

  const cellTexts = (row) =>
    Array.from(row.querySelectorAll("th, td"))
      .filter(isVisible)
      .map((cell) => normalize(cell.innerText || cell.textContent))
      .filter(Boolean);

  const looksLikeStudentRow = (cells) =>
    cells.length >= 4 && /^\d+$/.test(cells[0]) && /^\d{6,}$/.test(cells[1]) && /班$/.test(cells[3]);

  const getVisibleRows = () =>
    Array.from(document.querySelectorAll("table tr"))
      .filter(isVisible)
      .map(cellTexts)
      .filter((cells) => cells.length > 0);

  const getScoreHeaders = () => {
    const rows = getVisibleRows();
    const rankHeaders = rows.find(
      (cells) => cells[0] === "总分" && cells.includes("校次") && cells.includes("班次"),
    );
    const subjectHeaders = rows.find(
      (cells) =>
        cells[0] === "总分" &&
        cells.length >= 2 &&
        !cells.includes("校次") &&
        !cells.includes("班次") &&
        cells.slice(1).some(Boolean),
    );

    if (!rankHeaders) return FALLBACK_HEADERS.slice(4);
    return [...rankHeaders, ...(subjectHeaders ? subjectHeaders.slice(1) : [])];
  };

  const getHeaders = () => [...FALLBACK_HEADERS.slice(0, 4), ...getScoreHeaders()];

  const looksLikeScoreRow = (cells) =>
    cells.length >= 5 &&
    /^-?\d+(\.\d+)?$/.test(cells[0]) &&
    /^-?\d+$/.test(cells[1]) &&
    !/^\d{6,}$/.test(cells[1]);

  function readCurrentPageRows() {
    const rows = getVisibleRows();
    const studentRows = rows.filter(looksLikeStudentRow);
    const scoreRows = rows.filter(looksLikeScoreRow);
    const scoreColumnCount = getScoreHeaders().length;

    if (!studentRows.length || !scoreRows.length) {
      console.warn("当前页未识别到完整表格，已跳过。", { studentRows, scoreRows });
      return [];
    }

    const count = Math.min(studentRows.length, scoreRows.length);
    const merged = [];

    for (let i = 0; i < count; i += 1) {
      const left = studentRows[i].slice(0, 4);
      const right = scoreRows[i].slice(0, scoreColumnCount);
      merged.push([...left, ...right]);
    }

    return merged;
  }

  function getPageSignature() {
    return readCurrentPageRows()
      .slice(0, 5)
      .map((row) => row.join("|"))
      .join(";");
  }

  const getNextButton = () => {
    const candidates = Array.from(document.querySelectorAll("a, button, li, span"))
      .filter(isVisible)
      .filter((el) => normalize(el.innerText || el.textContent) === "下一页");

    return candidates.find((el) => {
      const className = String(el.className || "").toLowerCase();
      const ariaDisabled = el.getAttribute("aria-disabled") === "true";
      const disabled = el.disabled || ariaDisabled || className.includes("disabled");
      return !disabled;
    });
  };

  const dedupeRows = (rows) => {
    const seen = new Set();
    return rows.filter((row) => {
      const key = row.slice(1).join("|");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const collectActiveTab = async () => {
    const rows = [];

    for (let page = 1; page <= CONFIG.maxPages; page += 1) {
      const currentRows = readCurrentPageRows();
      rows.push(...currentRows);
      console.log(`已读取第 ${page} 页：${currentRows.length} 行`);

      const next = getNextButton();
      if (!next) break;

      const previousSignature = getPageSignature();
      next.click();
      await sleep(CONFIG.waitMs);
      await waitForTableChange(previousSignature);

      if (getPageSignature() === previousSignature) break;
    }

    return dedupeRows(rows);
  };

  const clickAllClasses = async () => {
    const allClasses = findClickableByText("全部班级");
    if (!allClasses) {
      console.warn("没有找到“全部班级”筛选项，将直接读取当前列表。");
      return;
    }

    const previousSignature = getPageSignature();
    allClasses.click();
    await sleep(CONFIG.waitMs);
    await waitForTableChange(previousSignature);
  };

  const groupRowsByClass = (rows) =>
    rows.reduce((grouped, row) => {
      const className = row[3] || "未识别班级";
      if (!grouped[className]) grouped[className] = [];

      const copied = [...row];
      copied[0] = String(grouped[className].length + 1);
      grouped[className].push(copied);

      return grouped;
    }, {});

  const loadSheetJs = async () => {
    if (window.XLSX) return;

    await new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = CONFIG.sheetJsUrl;
      script.onload = resolve;
      script.onerror = () => reject(new Error("SheetJS 加载失败，请检查网络或 CDN 是否被拦截。"));
      document.head.appendChild(script);
    });
  };

  const safeSheetName = (name) => name.replace(/[\\/?*[\]:]/g, "_").slice(0, 31);

  const classSortKey = (name) => {
    const gradeOrder = {
      一: 1,
      二: 2,
      三: 3,
      四: 4,
      五: 5,
      六: 6,
      七: 7,
      八: 8,
      九: 9,
    };
    const match = String(name).match(/^([一二三四五六七八九]|\d+)(\d+)班$/);
    if (!match) return [999, 999, String(name)];

    const grade = /^\d+$/.test(match[1]) ? Number(match[1]) : gradeOrder[match[1]];
    return [grade, Number(match[2]), String(name)];
  };

  const compareClasses = ([classA], [classB]) => {
    const keyA = classSortKey(classA);
    const keyB = classSortKey(classB);

    for (let i = 0; i < keyA.length; i += 1) {
      if (keyA[i] < keyB[i]) return -1;
      if (keyA[i] > keyB[i]) return 1;
    }

    return 0;
  };

  const applyColumnWidths = (worksheet, headers) => {
    worksheet["!cols"] = headers.map((header) => {
      const text = String(header);
      if (text.includes("准考证号")) return { wch: 14 };
      if (text.includes("进退步")) return { wch: 12 };
      if (text === "姓名") return { wch: 10 };
      if (text === "班级") return { wch: 8 };
      return { wch: Math.max(7, text.length + 2) };
    });
  };

  const exportWorkbook = (allRows, dataByClass) => {
    const workbook = XLSX.utils.book_new();
    const headers = getHeaders();

    const allClassSheet = XLSX.utils.aoa_to_sheet([headers, ...allRows]);
    applyColumnWidths(allClassSheet, headers);
    XLSX.utils.book_append_sheet(workbook, allClassSheet, "全部班级");

    Object.entries(dataByClass).sort(compareClasses).forEach(([className, rows]) => {
      const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      applyColumnWidths(worksheet, headers);
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(className));
    });

    const summaryRows = allRows.map((row) => [row[3] || "未识别班级", ...row]);
    const allSheet = XLSX.utils.aoa_to_sheet([["来源班级", ...headers], ...summaryRows]);
    applyColumnWidths(allSheet, ["来源班级", ...headers]);
    XLSX.utils.book_append_sheet(workbook, allSheet, "汇总");
    XLSX.writeFile(workbook, CONFIG.fileName);
  };

  console.clear();
  console.log("开始加载 Excel 导出组件...");
  await loadSheetJs();

  console.log("切换到全部班级并开始读取...");
  await clickAllClasses();

  const allRows = await collectActiveTab();
  const result = groupRowsByClass(allRows);
  Object.entries(result).forEach(([className, rows]) => console.log(`${className}：${rows.length} 行`));

  exportWorkbook(allRows, result);
  console.log("导出完成：", CONFIG.fileName, result);
})();

var etApp = window.Application;
const collegeNameList = ["机械工程与自动化学院", "汽车与交通工程学院", "材料科学与工程学院", "化学与环境工程学院", "电气工程学院", "电子与信息工程学院", "经济管理学院", "文化传媒与艺术设计学院", "土木建筑工程学院", "外国语学院", "理学院", "交叉科学学院"];
const summaryTitleList = ["学号", "姓名", "学院", "年级", "班级", "参与开始时间", "参与结束时间", "思想成长类-学分", "创新创业-学分", "志愿公益类-学分", "实践实习类-学分", "文体活动-学分", "工作履历-学分", "技能特长-学分", "学分总和"];
const detailTitleList = ["姓名", "学号", "院系", "年级", "班级", "学分类型", "单位", "一级分类", "二级分类", "学分数值", "实际获得学分", "转换后学分", "获取原因", "发送人", "参与时长（分）", "参与开始时间", "参与结束时间", "获取时间", "来源", "状态", "备注"];

// initialize user interface  (IIFE)
(function () {
    try {
        initializeUserInterface_summaryBookList();
        initializeUserInterface_buttonsForSummaryBookList();
        initializeUserInterface_detailBookList();
        initializeUserInterface_buttonsForDetailBookList();
        initializeUserInterface_summaryDataFilters();
        document.getElementById("start-button").addEventListener("click", function () { main(); });
        updateUserInterface("unlock");
    } catch (error) {
        alert(error.stack);
        window.close();
    }
})();

function initializeUserInterface_summaryBookList() {

    const bookListNode = document.getElementById("summary-book-list");

    for (let bookNum = 1; bookNum <= etApp.Workbooks.Count; bookNum++) {

        let book = etApp.Workbooks.Item(bookNum);

        const bookNode = document.createElement("div");
        const checkboxNode = document.createElement("input");
        const labelNode = document.createElement("label");

        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `summary-book-${bookNum}`);
        checkboxNode.setAttribute("name", "summary-book");
        checkboxNode.setAttribute("value", book.Name);

        labelNode.setAttribute("for", `summary-book-${bookNum}`);
        labelNode.textContent = book.Name;

        let checkResult = checkSummaryBook(book);
        if (checkResult !== "pass") {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("title", checkResult);
            labelNode.setAttribute("style", "color: gray;");
        }

        bookNode.appendChild(checkboxNode);
        bookNode.appendChild(labelNode);
        bookListNode.appendChild(bookNode);
    }
}

function initializeUserInterface_buttonsForSummaryBookList() {

    const selectAll = document.getElementById("select-all-summary-files");
    const deselectAll = document.getElementById("deselect-all-summary-files");
    const hideInvalid = document.getElementById("hide-invalid-summary-files");

    /** @type {NodeListOf<HTMLInputElement>} */
    const validList = document.querySelectorAll("input[name='summary-book']:not([disabled])");
    /** @type {NodeListOf<HTMLInputElement>} */
    const invalidList = document.querySelectorAll("input[name='summary-book'][disabled]");

    if (validList.length === 0) {
        selectAll.setAttribute("disabled", "");
        deselectAll.setAttribute("disabled", "");
    }

    selectAll.addEventListener("click", function () {
        for (const checkbox of validList) {
            checkbox.checked = true;
        }
    });

    deselectAll.addEventListener("click", function () {
        for (const checkbox of validList) {
            checkbox.checked = false;
        }
    });

    hideInvalid.addEventListener("click", function () {
        if (validList.length === 0) {
            hideInvalid.parentElement.setAttribute("hidden", "");
            document.getElementById("summary-book-list").setAttribute("hidden", "");
            document.getElementById("no-valid-summary-files").removeAttribute("hidden");
        }
        else {
            hideInvalid.setAttribute("hidden", "");
            for (const checkbox of invalidList) {
                checkbox.parentElement.setAttribute("hidden", "");
            }
        }
    });
}

function initializeUserInterface_detailBookList() {

    const bookListNode = document.getElementById("detail-book-list");

    for (let bookNum = 1; bookNum <= etApp.Workbooks.Count; bookNum++) {

        let book = etApp.Workbooks.Item(bookNum);

        const bookNode = document.createElement("div");
        const checkboxNode = document.createElement("input");
        const labelNode = document.createElement("label");

        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `detail-book-${bookNum}`);
        checkboxNode.setAttribute("name", "detail-book");
        checkboxNode.setAttribute("value", book.Name);

        labelNode.setAttribute("for", `detail-book-${bookNum}`);
        labelNode.textContent = book.Name;

        let checkResult = checkDetailBook(book);
        if (checkResult !== "pass") {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("title", checkResult);
            labelNode.setAttribute("style", "color: gray;");
        }

        bookNode.appendChild(checkboxNode);
        bookNode.appendChild(labelNode);
        bookListNode.appendChild(bookNode);
    }
}

function initializeUserInterface_buttonsForDetailBookList() {

    const selectAll = document.getElementById("select-all-detail-files");
    const deselectAll = document.getElementById("deselect-all-detail-files");
    const hideInvalid = document.getElementById("hide-invalid-detail-files");

    /** @type {NodeListOf<HTMLInputElement>} */
    const validList = document.querySelectorAll("input[name='detail-book']:not([disabled])");
    /** @type {NodeListOf<HTMLInputElement>} */
    const invalidList = document.querySelectorAll("input[name='detail-book'][disabled]");

    if (validList.length === 0) {
        selectAll.setAttribute("disabled", "");
        deselectAll.setAttribute("disabled", "");
    }

    selectAll.addEventListener("click", function () {
        for (const checkbox of validList) {
            checkbox.checked = true;
        }
    });

    deselectAll.addEventListener("click", function () {
        for (const checkbox of validList) {
            checkbox.checked = false;
        }
    });

    hideInvalid.addEventListener("click", function () {
        if (validList.length === 0) {
            hideInvalid.parentElement.setAttribute("hidden", "");
            document.getElementById("detail-book-list").setAttribute("hidden", "");
            document.getElementById("no-valid-detail-files").removeAttribute("hidden");
        }
        else {
            hideInvalid.setAttribute("hidden", "");
            for (const checkbox of invalidList) {
                checkbox.parentElement.setAttribute("hidden", "");
            }
        }
    });
}

function initializeUserInterface_summaryDataFilters() {

    const useCollegeFilter = document.getElementById("use-college-filter");
    const collegeFilterInput = document.getElementById("college-filter-input");

    const useGradeFilter = document.getElementById("use-grade-filter");
    const gradeFilterOperator = document.getElementById("grade-filter-operator");
    const gradeFilterInput = document.getElementById("grade-filter-input");

    const useClassFilter = document.getElementById("use-class-filter");
    const classFilterOperator = document.getElementById("class-filter-operator");
    const classFilterInput = document.getElementById("class-filter-input");

    for (const collegeName of collegeNameList) {
        const collegeNameNode = document.createElement("option");
        collegeNameNode.setAttribute("value", collegeName);
        collegeNameNode.textContent = collegeName;
        collegeFilterInput.appendChild(collegeNameNode);
    }

    useCollegeFilter.addEventListener("change", function () {
        if (useCollegeFilter.checked) {
            collegeFilterInput.removeAttribute("disabled");
        }
        else {
            collegeFilterInput.setAttribute("disabled", "");
        }
    });

    useGradeFilter.addEventListener("change", function () {
        if (useGradeFilter.checked) {
            gradeFilterOperator.removeAttribute("disabled");
            gradeFilterInput.removeAttribute("disabled");
        }
        else {
            gradeFilterOperator.setAttribute("disabled", "");
            gradeFilterInput.setAttribute("disabled", "");
        }
    });

    useClassFilter.addEventListener("change", function () {
        if (useClassFilter.checked) {
            classFilterOperator.removeAttribute("disabled");
            classFilterInput.removeAttribute("disabled");
        }
        else {
            classFilterOperator.setAttribute("disabled", "");
            classFilterInput.setAttribute("disabled", "");
        }
    });
}

function updateUserInterface(operation) {

    const fieldsetList = document.getElementsByTagName("fieldset");
    const startButton = document.getElementById("start-button");

    switch (operation) {
        case "unlock":
            for (const fieldset of fieldsetList) {
                fieldset.removeAttribute("disabled");
            }
            startButton.removeAttribute("disabled");
            break;
        case "lock":
            for (const fieldset of fieldsetList) {
                fieldset.setAttribute("disabled", "");
            }
            startButton.setAttribute("disabled", "");
            break;
    }
}

/**
 * 检查学分汇总工作簿
 * @param {Et.Workbook} summaryBook 学分汇总工作簿
 * @returns {string} 检查结果
 */
function checkSummaryBook(summaryBook) {

    if (summaryBook.Worksheets.Count !== 1) {
        return "工作表数量异常";
    }

    /** 学分汇总工作表 @type {Et.Worksheet} */
    let summarySheet = summaryBook.Worksheets.Item(1);

    if (summarySheet.Name !== "学分汇总") {
        return "工作表名称异常";
    }

    let dataTitle = (function () {
        let columnNumMax = summarySheet.UsedRange.Columns.Count;
        let startCell = summarySheet.Cells.Item(1, 1);
        let endCell = summarySheet.Cells.Item(1, columnNumMax);
        /** @type {any[]} */
        let titleArr = summarySheet.Range(startCell, endCell).Value2[0];
        return titleArr.map((ele) => String(ele));
    })();

    if (new Set(dataTitle).size !== dataTitle.length) {
        return "存在重复的数据标题";
    }

    for (const summaryTitle of summaryTitleList) {
        if (dataTitle.includes(summaryTitle) === false) {
            return `未找到【${summaryTitle}】数据列`;
        }
    }

    const rowNumMax = summarySheet.UsedRange.Rows.Count;

    if (rowNumMax < 2) {
        return "未找到数据行";
    }

    let formulaStr;

    formulaStr = (function () {
        let columnNum = summarySheet.Rows.Item(1).Find("学号", undefined, -4163, 1).Column;
        let startCell = summarySheet.Cells.Item(2, columnNum);
        let endCell = summarySheet.Cells.Item(rowNumMax, columnNum);
        let rangeAddr = summarySheet.Range(startCell, endCell).Address(false, false);
        return `COUNTA(UNIQUE('[${summaryBook.Name}]${summarySheet.Name}'!${rangeAddr}))`;
    })();
    if (etApp.Evaluate(formulaStr) !== rowNumMax - 1) {
        return "存在【学号】相同的数据行";
    }

    formulaStr = (function () {
        let columnNum = summarySheet.Rows.Item(1).Find("参与开始时间", undefined, -4163, 1).Column;
        let startCell = summarySheet.Cells.Item(2, columnNum);
        let endCell = summarySheet.Cells.Item(rowNumMax, columnNum);
        let rangeAddr = summarySheet.Range(startCell, endCell).Address(false, false);
        return `COUNTA(UNIQUE('[${summaryBook.Name}]${summarySheet.Name}'!${rangeAddr}))`;
    })();
    if (etApp.Evaluate(formulaStr) !== 1) {
        return "存在【参与开始时间】不同的数据行";
    }

    formulaStr = (function () {
        let columnNum = summarySheet.Rows.Item(1).Find("参与结束时间", undefined, -4163, 1).Column;
        let startCell = summarySheet.Cells.Item(2, columnNum);
        let endCell = summarySheet.Cells.Item(rowNumMax, columnNum);
        let rangeAddr = summarySheet.Range(startCell, endCell).Address(false, false);
        return `COUNTA(UNIQUE('[${summaryBook.Name}]${summarySheet.Name}'!${rangeAddr}))`;
    })();
    if (etApp.Evaluate(formulaStr) !== 1) {
        return "存在【参与结束时间】不同的数据行";
    }

    return "pass";
}

/**
 * 检查学分明细工作簿
 * @param {Et.Workbook} detailBook 学分明细工作簿
 * @returns {string} 检查结果
 */
function checkDetailBook(detailBook) {

    if (detailBook.Worksheets.Count !== 1) {
        return "工作表数量异常";
    }

    /** 学分明细工作表 @type {Et.Worksheet} */
    let detailSheet = detailBook.Worksheets.Item(1);

    if (detailSheet.Name !== "学分明细") {
        return "工作表名称异常";
    }

    let dataTitle = (function () {
        let startCell = detailSheet.Cells.Item(1, 1);
        let endCell = detailSheet.Cells.Item(1, detailTitleList.length);
        /** @type {any[]} */
        let titleArr = detailSheet.Range(startCell, endCell).Value2[0];
        return titleArr.map((ele) => String(ele));
    })();

    for (let i = 0; i < dataTitle.length; i++) {
        if (dataTitle[i] !== detailTitleList[i]) {
            return `匹配【${detailTitleList[i]}】数据列失败`;
        }
    }

    const rowNumMax = detailSheet.UsedRange.Rows.Count;

    if (rowNumMax < 2) {
        return "未找到数据行";
    }

    let formulaStr = (function () {
        let columnNum, startCell, endCell, rangeAddr, criteria_range;
        const bookName = detailBook.Name;
        const sheetName = detailSheet.Name;

        columnNum = detailTitleList.indexOf("学分类型") + 1;
        startCell = detailSheet.Cells.Item(2, columnNum);
        endCell = detailSheet.Cells.Item(rowNumMax, columnNum);
        rangeAddr = detailSheet.Range(startCell, endCell).Address(false, false);
        criteria_range = `'[${bookName}]${sheetName}'!${rangeAddr},"实践实习类",`;

        columnNum = detailTitleList.indexOf("单位") + 1;
        startCell = detailSheet.Cells.Item(2, columnNum);
        endCell = detailSheet.Cells.Item(rowNumMax, columnNum);
        rangeAddr = detailSheet.Range(startCell, endCell).Address(false, false);
        criteria_range += `'[${bookName}]${sheetName}'!${rangeAddr},"学分",`;

        columnNum = detailTitleList.indexOf("状态") + 1;
        startCell = detailSheet.Cells.Item(2, columnNum);
        endCell = detailSheet.Cells.Item(rowNumMax, columnNum);
        rangeAddr = detailSheet.Range(startCell, endCell).Address(false, false);
        criteria_range += `'[${bookName}]${sheetName}'!${rangeAddr},"有效"`;

        return `COUNTIFS(${criteria_range})`;
    })();
    if (etApp.Evaluate(formulaStr) < 1) {
        return "未找到有效的【实践实习类-学分】数据行";
    }

    return "pass";
}

function main() {
    try {
        updateUserInterface("lock");

        if (document.querySelectorAll("input[name='summary-book']:checked").length === 0) {
            alert("未选择学分汇总文件");
            updateUserInterface("unlock");
            return;
        }
        if (document.querySelectorAll("input[name='detail-book']:checked").length === 0) {
            alert("未选择学分明细文件");
            updateUserInterface("unlock");
            return;
        }

        // 获取学分明细数据
        let summaryData = getSummaryData();
        // 筛选学分明细数据，仅保留【符合全部筛选条件】的数据行。筛没了就退出
        summaryData = filterSummaryData(summaryData);
        if (summaryData.length === 0) {
            alert("未找到符合筛选条件的学分汇总数据行");
            updateUserInterface("unlock");
            return;
        }
        // 检查学分汇总数据，保证学号不重复、时间一致。有问题就退出
        let checkResult = checkSummaryData(summaryData);
        if (checkResult !== "pass") {
            alert(checkResult);
            updateUserInterface("unlock");
            return;
        }


        // 获取学分汇总数据
        let detailData = getDetailData();
        // 筛选学分汇总数据，仅保留【实践实习类-学分-有效】的数据行，提高性能
        detailData = filterDetailData(detailData);

        // 生成输出工作簿
        let outputBook = createOutputBook(summaryData, detailData);

        // 处理完成，弹窗提示并退出
        alert(`处理完成。输出工作簿：${outputBook.Name}`);
        window.close();
    } catch (error) {
        alert(error.stack);
    }
}

/**
 * 获取学分汇总数据
 * @returns {any[]} 学分汇总数据
 */
function getSummaryData() {

    let tempBook = etApp.Workbooks.Add();

    let tempSheetNameList = [];

    for (const checkbox of document.querySelectorAll("input[name='summary-book']:checked")) {
        /** 学分汇总工作表 @type {Et.Worksheet} */
        let summarySheet = etApp.Workbooks.Item(checkbox.value).Worksheets.Item(1);
        summarySheet.Copy(tempBook.Worksheets.Item(1), null);
        tempSheetNameList.push(etApp.ActiveSheet.Name);
    }

    /** 学分汇总数据 @type {any[]} */
    let summaryData = [];

    for (const tempSheetName of tempSheetNameList) {

        /** 临时工作表（学分汇总） @type {Et.Worksheet} */
        let tempSheet = tempBook.Worksheets.Item(tempSheetName);

        // 格式化学分汇总工作表
        for (let columnNum = tempSheet.UsedRange.Columns.Count; columnNum >= 1; columnNum--) {
            const dataTitle = String(tempSheet.Cells.Item(1, columnNum).Value2);
            if (summaryTitleList.includes(dataTitle) === false) {
                tempSheet.Columns.Item(columnNum).Delete(-4159);
            }
        }
        for (let i = summaryTitleList.length - 1; i >= 0; i--) {
            tempSheet.Columns.Item(1).Insert(-4161);
            const columnNum = tempSheet.Rows.Item(1).Find(summaryTitleList[i], undefined, -4163, 1).Column;
            tempSheet.Columns.Item(columnNum).Cut(tempSheet.Columns.Item(1));
        }

        // 确定目标区域
        const targetRange = (function () {
            const rowNumMax = tempSheet.UsedRange.Rows.Count;
            const columnNumMax = summaryTitleList.length;
            const startCell = tempSheet.Cells.Item(2, 1);
            const endCell = tempSheet.Cells.Item(rowNumMax, columnNumMax);
            return tempSheet.Range(startCell, endCell);
        })();

        // 将数据添加到数组
        summaryData = summaryData.concat(targetRange.Value2);
    }

    tempBook.Close(false);
    return summaryData;
}

/**
 * 获取学分明细数据
 * @returns {any[]} 学分明细数据
 */
function getDetailData() {

    let tempBook = etApp.Workbooks.Add();

    let tempSheetNameList = [];

    for (const checkbox of document.querySelectorAll("input[name='detail-book']:checked")) {
        /** 学分明细工作表 @type {Et.Worksheet} */
        let detailSheet = etApp.Workbooks.Item(checkbox.value).Worksheets.Item(1);
        detailSheet.Copy(tempBook.Worksheets.Item(1), null);
        tempSheetNameList.push(etApp.ActiveSheet.Name);
    }

    /** 学分明细数据 @type {any[]} */
    let detailData = [];

    for (const tempSheetName of tempSheetNameList) {

        /** 临时工作表（学分明细） @type {Et.Worksheet} */
        let tempSheet = tempBook.Worksheets.Item(tempSheetName);

        // 确定目标区域
        const targetRange = (function () {
            const rowNumMax = tempSheet.UsedRange.Rows.Count;
            const columnNumMax = detailTitleList.length;
            const startCell = tempSheet.Cells.Item(2, 1);
            const endCell = tempSheet.Cells.Item(rowNumMax, columnNumMax);
            return tempSheet.Range(startCell, endCell);
        })();

        // 将数据添加到数组
        detailData = detailData.concat(targetRange.Value2);
    }

    tempBook.Close(false);
    return detailData;
}

/**
 * 筛选学分汇总数据
 * @param {any[]} summaryData 学分汇总数据
 * @returns {any[]} 筛选后的学分汇总数据
 */
function filterSummaryData(summaryData) {

    let filteredData = summaryData;

    // 如果没设置筛选器，直接返回原始数组
    if (document.querySelectorAll("input[name='summary-data-filter']:checked") === 0) {
        return filteredData;
    }

    // 筛选学院
    if (document.getElementById("use-college-filter").checked) {
        const positionIndex = summaryTitleList.indexOf("学院");
        const filterValue = String(document.getElementById("college-filter-input").value);
        filteredData = filteredData.filter((ele) => String(ele[positionIndex]) === filterValue);
    }

    // 筛选年级
    if (document.getElementById("use-grade-filter").checked) {
        const positionIndex = summaryTitleList.indexOf("年级");
        const filterValue = Number(document.getElementById("grade-filter-input").value);
        switch (document.getElementById("grade-filter-operator").value) {
            case "greater":
                filteredData = filteredData.filter((ele) => Number(ele[positionIndex]) > filterValue);
                break;
            case "less":
                filteredData = filteredData.filter((ele) => Number(ele[positionIndex]) < filterValue);
                break;
            case "equal":
                filteredData = filteredData.filter((ele) => Number(ele[positionIndex]) === filterValue);
                break;
            default:
                alert("unknown grade filter operator");
                return [];
        }
    }

    // 筛选班级
    if (document.getElementById("use-class-filter").checked) {
        const positionIndex = summaryTitleList.indexOf("班级");
        const filterValue = String(document.getElementById("class-filter-input").value);
        switch (document.getElementById("class-filter-operator").value) {
            case "exact-match":
                filteredData = filteredData.filter((ele) => String(ele[positionIndex]) === filterValue);
                break;
            case "partial-match":
                filteredData = filteredData.filter((ele) => String(ele[positionIndex]).includes(filterValue));
                break;
            default:
                alert("unknown class filter operator");
                return [];
        }
    }

    return filteredData;
}

/**
 * 筛选学分明细数据
 * @param {any[]} detailData 学分明细数据
 * @returns {any[]} 筛选后的学分明细数据
 */
function filterDetailData(detailData) {

    let filteredData = detailData;

    let positionIndex = [detailTitleList.indexOf("学分类型"), detailTitleList.indexOf("单位"), detailTitleList.indexOf("状态")];
    let filterValue = ["实践实习类", "学分", "有效"];
    filteredData = filteredData.filter((ele) =>
        ele[positionIndex[0]] === filterValue[0] &&
        ele[positionIndex[1]] === filterValue[1] &&
        ele[positionIndex[2]] === filterValue[2]
    );

    return filteredData;
}

/**
 * 检查学分汇总数据
 * @param {any[]} summaryData 学分汇总数据
 * @returns {string} 检查结果
 */
function checkSummaryData(summaryData) {

    let positionIndex;

    // 检查学号
    positionIndex = summaryTitleList.indexOf("学号");
    let studentNumberList = summaryData.map((ele) => String(ele[positionIndex]));
    if (new Set(studentNumberList).size !== studentNumberList.length) {
        return "合并并筛选后的学分汇总数据中，存在重复的学号";
    }

    // 检查参与开始时间
    positionIndex = summaryTitleList.indexOf("参与开始时间");
    let startTimeList = summaryData.map((ele) => String(ele[positionIndex]));
    if (new Set(startTimeList).size !== 1) {
        return "合并并筛选后的学分汇总数据中，参与开始时间不统一";
    }
    // 检查参与结束时间
    positionIndex = summaryTitleList.indexOf("参与结束时间");
    let endTimeList = summaryData.map((ele) => String(ele[positionIndex]));
    if (new Set(endTimeList).size !== 1) {
        return "合并并筛选后的学分汇总数据中，参与结束时间不统一";
    }

    return "pass";
}

/**
 * 生成输出工作簿
 * @param {any[]} summaryData 学分汇总数据
 * @param {any[]} detailData 学分明细数据
 * @returns {Et.Workbook} 输出工作簿
 */
function createOutputBook(summaryData, detailData) {

    let outputBook = etApp.Workbooks.Add();

    /** 输出工作表类型（表格拆分方式） @type {string} */
    const outputSheetType = document.getElementById("output-sheet-type").value;

    /** 输出工作表名称 @type {string} */
    let outputSheetName;
    /** 输出工作表名称清单 @type {string[]} */
    let outputSheetNameList = [];
    /** 输出工作表 @type {Et.Worksheet} */
    let outputSheet;
    /** 输出行号 @type {number} */
    let outputRowNum;

    for (let i = 0; i < summaryData.length; i++) {

        // 根据设置的表格拆分方式，获取输出工作表名称
        switch (outputSheetType) {
            case "sheet-type-college":
                outputSheetName = String(summaryData[i][summaryTitleList.indexOf("学院")]);
                break;
            case "sheet-type-grade":
                outputSheetName = String(summaryData[i][summaryTitleList.indexOf("年级")]);
                break;
            case "sheet-type-class":
                outputSheetName = String(summaryData[i][summaryTitleList.indexOf("班级")]);
                break;
            case "sheet-type-all-in-one":
                outputSheetName = "output-sheet";
                break;
            default:
                alert("unknown output sheet type");
        }

        // 获取输出工作表，获取输出位置（行号）
        if (outputSheetNameList.includes(outputSheetName)) {
            outputSheet = outputBook.Worksheets.Item(outputSheetName);
            outputRowNum = outputSheet.UsedRange.Rows.Count + 1;
        }
        else {
            outputSheet = initializeOutputSheet(outputBook, outputSheetName);
            outputSheetNameList.push(outputSheetName);
            outputRowNum = 2;
        }

        // 向输出工作表写入学分汇总数据
        const targetRange = (function () {
            const columnNumMax = summaryTitleList.length;
            const startCell = outputSheet.Cells.Item(outputRowNum, 1);
            const endCell = outputSheet.Cells.Item(outputRowNum, columnNumMax);
            return outputSheet.Range(startCell, endCell);
        })();
        targetRange.Value2 = summaryData[i];

        // 向输出工作表写入学分明细数据
        let detailFound = detailData.find((ele) => {
            const studentNumber_summary = String(summaryData[i][summaryTitleList.indexOf("学号")]);
            const studentNumber_detail = String(ele[detailTitleList.indexOf("学号")]);
            const obtainReason_detail = String(ele[detailTitleList.indexOf("获取原因")]);
            return studentNumber_summary === studentNumber_detail && (obtainReason_detail.includes("三下乡") || obtainReason_detail.includes("返家乡"));
        });
        const targetCell = outputSheet.Cells.Item(outputRowNum, summaryTitleList.length + 1);
        if (detailFound !== undefined) {
            targetCell.Value2 = detailFound[detailTitleList.indexOf("获取原因")];
        }
        else {
            targetCell.Value2 = "-";
        }
    }

    // 处理输出工作簿：格式化输出工作表，删除空表
    for (let outputSheetNum = 1; outputSheetNum <= outputBook.Worksheets.Count; outputSheetNum++) {

        outputSheet = outputBook.Worksheets.Item(outputSheetNum);

        if (outputSheetNameList.includes(outputSheet.Name)) {
            calculateOutputSheet(outputSheet);
            formatOutputSheet(outputSheet);
            if (document.getElementById("colorful-mark-output-sheet").checked) {
                colorfulMarkOutputSheet(outputSheet);
            }
        }
        else {
            outputSheet.Delete();
        }
    }

    return outputBook;
}

/**
 * 初始化输出工作表
 * @param {Et.Workbook} outputBook 输出工作簿
 * @param {string} outputSheetName 输出工作表名称
 * @returns 空的输出工作表
 */
function initializeOutputSheet(outputBook, outputSheetName) {

    // 在输出工作簿中，新建工作表作为输出工作表
    let outputSheet = outputBook.Worksheets.Add();

    // 重命名输出工作表
    outputSheet.Name = outputSheetName;

    const targetRange = (function () {
        const columnNumMax = summaryTitleList.length;
        const startCell = outputSheet.Cells.Item(1, 1);
        const endCell = outputSheet.Cells.Item(1, columnNumMax);
        return outputSheet.Range(startCell, endCell);
    })();
    targetRange.Value2 = summaryTitleList;

    return outputSheet;
}

/**
 * 计算输出工作表
 * @param {Et.Worksheet} sheet 输出工作表
 */
function calculateOutputSheet(sheet) {

    // 插入空列
    const columnNumList = [16, 15, 15, 11, 11, 6, 6];
    for (let i = 0; i < columnNumList.length; i++) {
        sheet.Columns.Item(columnNumList[i]).Insert(-4161);
    }

    const rowNumMax = sheet.UsedRange.Rows.Count;

    // 根据设置的学分标准版本，写入公式
    let startCell, endCell;
    if (document.getElementById("credit-standred-version").value === "auto") {
        // 根据学号前两位，计算入学年份
        startCell = sheet.Cells.Item(2, 6);
        endCell = sheet.Cells.Item(rowNumMax, 6);
        startCell.Formula = '=VALUE("20"&LEFT(A2,2))';
        startCell.AutoFill(sheet.Range(startCell, endCell), 1);
        // 根据入学年份，计算标准版本
        startCell = sheet.Cells.Item(2, 7);
        endCell = sheet.Cells.Item(rowNumMax, 7);
        startCell.Formula = '=IFS(F2>=2023,"2023版",F2<2023,"2019版")';
        startCell.AutoFill(sheet.Range(startCell, endCell), 1);
    }
    else {
        // 由于手动选择了标准版本，入学年份列写入横杠占位
        startCell = sheet.Cells.Item(2, 6);
        endCell = sheet.Cells.Item(rowNumMax, 6);
        startCell.Value2 = "-";
        startCell.AutoFill(sheet.Range(startCell, endCell), 1);
        // 由于手动选择了标准版本，标准版本列直接写入
        startCell = sheet.Cells.Item(2, 7);
        endCell = sheet.Cells.Item(rowNumMax, 7);
        startCell.Value2 = document.getElementById("credit-standred-version").value;
        startCell.AutoFill(sheet.Range(startCell, endCell), 1);
    }

    // 将列号码和公式，按键值对存储到map
    const columnFormulaMap = new Map();
    // 总必修学分
    columnFormulaMap.set(13, '=SUM(J2:L2)');
    // 必修有效学分
    columnFormulaMap.set(14, '=SUM(MIN(J2,2),MIN(K2,2),MIN(L2,1))');
    // 总选修学分
    columnFormulaMap.set(19, '=SUM(O2:R2)');
    // 选修有效学分
    columnFormulaMap.set(20, '=IFS(G2="2023版",MIN(O2+SUM(P2:R2),O2+2,3),G2="2019版",MIN(S2,3))');
    // 如果设置了【重新计算学分总和】，重新计算【总学分】
    if (document.getElementById("recalculate-total-credits").checked) {
        columnFormulaMap.set(21, '=SUM(M2,S2)');
    }
    // 总有效学分
    columnFormulaMap.set(22, '=SUM(N2,T2)');
    // 毕业资格
    columnFormulaMap.set(24, '=IFS(G2="2023版",IF(AND(V2>=8,OR(ISNUMBER(FIND("三下乡",W2)),ISNUMBER(FIND("返家乡",W2)))),"OK","NG"),G2="2019版",IF(V2>=8,"OK","NG"))');

    for (const [columnNum, formulaStr] of columnFormulaMap) {
        startCell = sheet.Cells.Item(2, columnNum);
        endCell = sheet.Cells.Item(rowNumMax, columnNum);
        startCell.Formula = formulaStr;
        startCell.AutoFill(sheet.Range(startCell, endCell), 1);
    }
}

/**
 * 格式化输出工作表
 * @param {Et.Worksheet} sheet 输出工作表
 */
function formatOutputSheet(sheet) {

    // 如果设置了【输出数据转换为值】，复制并粘贴为值
    if (document.getElementById("convert-output-data-into-values").checked) {
        sheet.UsedRange.Copy();
        sheet.Range("A1").PasteSpecial(-4163, -4142, false, false);
    }

    // 配置第1行标题
    sheet.Range("H1:U1").Clear();
    sheet.Range("F1:X1").Value2 = ["入学年份", "标准版本", "统计开始时间", "统计结束时间", "必修学分", "", "", "", "", "选修学分", "", "", "", "", "", "总学分", "总有效学分", "实践实习记录", "毕业资格"];
    // 配置第2行标题
    sheet.Rows.Item(2).Insert(-4121);
    sheet.Range("J2:T2").Value2 = ["思想成长", "创新创业", "志愿公益", "总必修学分", "有效学分", "实践实习", "文体活动", "工作履历", "技能特长", "总选修学分", "有效学分"];
    const rangeAddrList = ["A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:F2", "G1:G2", "H1:H2", "I1:I2", "J1:N1", "O1:T1", "U1:U2", "V1:V2", "W1:W2", "X1:X2"];
    for (const rangeAddr of rangeAddrList) {
        sheet.Range(rangeAddr).Merge();
    }

    // 设置行高、列宽
    sheet.UsedRange.RowHeight = 20;
    const columnWidthList = [10, 8, 12, 5, 9, 8, 8, 12, 12, 8, 8, 8, 10, 8, 8, 8, 8, 8, 10, 8, 8, 10, 12, 8];
    for (let i = 0; i < columnWidthList.length; i++) {
        sheet.Columns.Item(i + 1).ColumnWidth = columnWidthList[i];
    }

    // 设置对齐方式
    sheet.UsedRange.HorizontalAlignment = -4108;
    sheet.Range("B:C").HorizontalAlignment = -4131;
    sheet.Range("E:E").HorizontalAlignment = -4131;
    sheet.Range("W:W").HorizontalAlignment = -4131;
    sheet.Range("1:1").HorizontalAlignment = -4108;

    // 设置数据格式
    sheet.Range("J:V").NumberFormatLocal = "0.00";
}

/**
 * 
 * @param {Et.Worksheet} sheet 
 */
function colorfulMarkOutputSheet(sheet) {

    const colorMap = new Map();
    colorMap.set("red", 12498932);
    colorMap.set("green", 11855049);
    colorMap.set("blue", 15386293);
    colorMap.set("yellow", 9889790);
    colorMap.set("gray", 14277081);

    let targetRange;
    const rowNumMax = sheet.UsedRange.Rows.Count;

    // 预填充颜色：标题灰色，信息部分红色，学分部分黄色
    sheet.Range("A1:X2").Interior.Color = colorMap.get("gray");

    targetRange = (function () {
        const startCell = sheet.Cells.Item(3, 1);
        const endCell = sheet.Cells.Item(rowNumMax, 9);
        return sheet.Range(startCell, endCell);
    })();
    targetRange.Interior.Color = colorMap.get("red");
    targetRange = (function () {
        const startCell = sheet.Cells.Item(3, 10);
        const endCell = sheet.Cells.Item(rowNumMax, 24);
        return sheet.Range(startCell, endCell);
    })();
    targetRange.Interior.Color = colorMap.get("yellow");

    // 读取数据，存储到数组
    const data = (function () {
        const startCell = sheet.Cells.Item(3, 1);
        const endCell = sheet.Cells.Item(rowNumMax, 24);
        return sheet.Range(startCell, endCell).Value2;
    })();


    for (let i = 0; i < data.length; i++) {

        const rowNum = i + 3;

        if (String(data[i][23]) === "OK") {

            // 信息部分设置为蓝色
            targetRange = (function () {
                const startCell = sheet.Cells.Item(rowNum, 1);
                const endCell = sheet.Cells.Item(rowNum, 9);
                return sheet.Range(startCell, endCell);
            })();
            targetRange.Interior.Color = colorMap.get("blue");

            // 学分部分设置为绿色
            targetRange = (function () {
                const startCell = sheet.Cells.Item(rowNum, 10);
                const endCell = sheet.Cells.Item(rowNum, 24);
                return sheet.Range(startCell, endCell);
            })();
            targetRange.Interior.Color = colorMap.get("green");
        }
        else if (String(data[i][23]) === "NG") {

            // 学分部分

            // 总有效学分=8
            if (Number(data[i][21]) === 8) {
                targetRange = (function () {
                    const startCell = sheet.Cells.Item(rowNum, 10);
                    const endCell = sheet.Cells.Item(rowNum, 22);
                    return sheet.Range(startCell, endCell);
                })();
                targetRange.Interior.Color = colorMap.get("green");
            }
            else {
                // 必修有效学分=5
                if (Number(data[i][13]) === 5) {
                    targetRange = (function () {
                        const startCell = sheet.Cells.Item(rowNum, 10);
                        const endCell = sheet.Cells.Item(rowNum, 14);
                        return sheet.Range(startCell, endCell);
                    })();
                    targetRange.Interior.Color = colorMap.get("green");
                }
                else {
                    if (Number(data[i][9]) >= 2) {
                        sheet.Cells.Item(rowNum, 10).Interior.Color = colorMap.get("green");
                    }
                    if (Number(data[i][10]) >= 2) {
                        sheet.Cells.Item(rowNum, 11).Interior.Color = colorMap.get("green");
                    }
                    if (Number(data[i][11]) >= 1) {
                        sheet.Cells.Item(rowNum, 12).Interior.Color = colorMap.get("green");
                    }
                }

                // 选修有效学分=3
                if (Number(data[i][19]) === 3) {
                    targetRange = (function () {
                        const startCell = sheet.Cells.Item(rowNum, 15);
                        const endCell = sheet.Cells.Item(rowNum, 20);
                        return sheet.Range(startCell, endCell);
                    })();
                    targetRange.Interior.Color = colorMap.get("green");
                }
                else if (String(data[i][6]) === "2023版") {
                    if (Number(data[i][14]) >= 1) {
                        sheet.Cells.Item(rowNum, 15).Interior.Color = colorMap.get("green");
                    }
                    if (Number(data[i][15]) + Number(data[i][16]) + Number(data[i][17]) >= 2) {
                        targetRange = (function () {
                            const startCell = sheet.Cells.Item(rowNum, 16);
                            const endCell = sheet.Cells.Item(rowNum, 18);
                            return sheet.Range(startCell, endCell);
                        })();
                        targetRange.Interior.Color = colorMap.get("green");
                    }
                }
            }

            // 实践实习记录
            if (String(data[i][6]) === "2019版" || String(data[i][6]) === "2023版" && (String(data[i][22]).includes("三下乡") || String(data[i][22]).includes("返家乡"))) {
                sheet.Cells.Item(rowNum, 23).Interior.Color = colorMap.get("green");
            }

        }
    }

    // 添加边框
    sheet.UsedRange.Borders.LineStyle = 1;
}
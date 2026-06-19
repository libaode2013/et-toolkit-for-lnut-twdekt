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

        // 获取学分汇总数据
        let summaryData = getSummaryData();
        // 筛选学分汇总数据
        if (document.querySelectorAll("input[name='summary-data-filter']:checked") !== 0) {
            summaryData = filterSummaryData(summaryData);
        }
        if (summaryData.length === 0) {
            alert("未找到符合筛选条件的学分汇总数据行");
            updateUserInterface("unlock");
            return;
        }
        // 检查学分汇总数据
        let checkResult = checkSummaryData(summaryData);
        if (checkResult !== "pass") {
            alert(checkResult);
            updateUserInterface("unlock");
            return;
        }

        // 获取学分明细数据
        let detailData = getDetailData();
        // 整理学分明细数据
        detailData = filterDetailData(detailData);

        // 生成输出工作簿
        let outputBook = createOutputBook(summaryData, detailData);

        // 弹窗提示完成
        alert(`处理完成，输出工作簿：${outputBook.Name}`);
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

        // 删除多余的数据列
        for (let columnNum = tempSheet.UsedRange.Columns.Count; columnNum > 0; columnNum--) {
            let dataTitle = String(tempSheet.Cells.Item(1, columnNum).Value2);
            if (summaryTitleList.includes(dataTitle) === false) {
                tempSheet.Columns.Item(columnNum).Delete(-4159);
            }
        }

        // 调整数据列位置
        for (let i = summaryTitleList.length - 1; i >= 0; i--) {
            tempSheet.Columns.Item(1).Insert(-4161);
            let columnNum = tempSheet.Rows.Item(1).Find(summaryTitleList[i], undefined, -4163, 1).Column;
            tempSheet.Columns.Item(columnNum).Cut(tempSheet.Columns.Item(1));
        }

        // 确定数据所在的目标区域
        let targetRange = (function () {
            let rowNumMax = tempSheet.UsedRange.Rows.Count;
            let columnNumMax = summaryTitleList.length;
            let startCell = tempSheet.Cells.Item(2, 1);
            let endCell = tempSheet.Cells.Item(rowNumMax, columnNumMax);
            return tempSheet.Range(startCell, endCell);
        })();

        // 将数据添加到数组
        summaryData = summaryData.concat(targetRange.Value2);
    }

    // 关闭临时工作簿
    tempBook.Close(false);

    return summaryData;
}

/**
 * 筛选学分汇总数据
 * @param {any[]} summaryData 学分汇总数据
 * @returns {any[]} 筛选后的学分汇总数据
 */
function filterSummaryData(summaryData) {

    let filteredData = summaryData;

    // 筛选学院
    if (document.getElementById("use-college-filter").checked) {
        let positionIndex = summaryTitleList.indexOf("学院");
        let filterValue = String(document.getElementById("college-filter-input").value);
        filteredData = filteredData.filter((ele) => String(ele[positionIndex]) === filterValue);
    }

    // 筛选年级
    if (document.getElementById("use-grade-filter").checked) {
        let positionIndex = summaryTitleList.indexOf("年级");
        let filterValue = Number(document.getElementById("grade-filter-input").value);
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
        let positionIndex = summaryTitleList.indexOf("班级");
        let filterValue = String(document.getElementById("class-filter-input").value);
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

        let targetRange = (function () {
            let rowNumMax = tempSheet.UsedRange.Rows.Count;
            let columnNumMax = detailTitleList.length;
            let startCell = tempSheet.Cells.Item(2, 1);
            let endCell = tempSheet.Cells.Item(rowNumMax, columnNumMax);
            return tempSheet.Range(startCell, endCell);
        })();

        detailData = detailData.concat(targetRange.Value2);
    }

    tempBook.Close(false);

    return detailData;
}

/**
 * 处理学分明细数据
 * @param {any[]} detailData 学分明细数据
 * @returns {any[]} 处理后的学分明细数据
 */
function filterDetailData(detailData) {

    let filteredData = detailData;

    let positionIndex = [detailTitleList.indexOf("学分类型"), detailTitleList.indexOf("单位"), detailTitleList.indexOf("状态")];
    let filterValue = ["实践实习类", "学分", "有效"];

    filteredData = filteredData.filter((ele) =>
        String(ele[positionIndex[0]]) === filterValue[0] &&
        String(ele[positionIndex[1]]) === filterValue[1] &&
        String(ele[positionIndex[2]]) === filterValue[2]
    );

    return filteredData;
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

        // get output sheet name
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

        // get output sheet and output row num
        if (outputSheetNameList.includes(outputSheetName)) {
            outputSheet = outputBook.Worksheets.Item(outputSheetName);
            outputRowNum = outputSheet.UsedRange.Rows.Count + 1;
        }
        else {
            outputSheet = outputBook.Worksheets.Add();
            outputSheet.Name = outputSheetName;
            let targetRange = (function () {
                let columnNumMax = summaryTitleList.length;
                let startCell = outputSheet.Cells.Item(1, 1);
                let endCell = outputSheet.Cells.Item(1, columnNumMax);
                return outputSheet.Range(startCell, endCell);
            })();
            targetRange.Value2 = summaryTitleList;
            outputSheetNameList.push(outputSheetName);
            outputRowNum = 2;
        }

        // write summary data
        let targetRange = (function () {
            let columnNumMax = summaryTitleList.length;
            let startCell = outputSheet.Cells.Item(outputRowNum, 1);
            let endCell = outputSheet.Cells.Item(outputRowNum, columnNumMax);
            return outputSheet.Range(startCell, endCell);
        })();
        targetRange.Value2 = summaryData[i];

        // find detail data
        let found = detailData.find((ele) => {
            const summaryStuNum = String(summaryData[i][summaryTitleList.indexOf("学号")]);
            const detailStuNum = String(ele[detailTitleList.indexOf("学号")]);
            const detailReason = String(ele[detailTitleList.indexOf("获取原因")]);
            return summaryStuNum === detailStuNum && (detailReason.includes("三下乡") || detailReason.includes("返家乡"));
        });
        targetRange = outputSheet.Cells.Item(outputRowNum, summaryTitleList.length + 1);
        if (found !== undefined) {
            targetRange.Value2 = found[detailTitleList.indexOf("获取原因")];
        }
        else {
            targetRange.Value2 = "-";
        }
    }

    for (let outputSheetNum = 1; outputSheetNum <= outputBook.Worksheets.Count; outputSheetNum++) {

        outputSheet = outputBook.Worksheets.Item(outputSheetNum);

        if (outputSheetNameList.includes(outputSheet.Name)) {
            calculateOutputSheet(outputSheet);
            organizeOutputSheet(outputSheet);
        }
        else {
            outputSheet.Delete();
        }
    }

    return outputBook;
}

/**
 * 
 * @param {Et.Worksheet} outputSheet 
 */
function calculateOutputSheet(outputSheet) {
    // old code

    let outputRowNumMax = outputSheet.UsedRange.Rows.Count;

    let targetColumnNumList = [16, 15, 15, 11, 11, 6, 6];
    for (const targetColumnNum of targetColumnNumList) {
        outputSheet.Columns.Item(targetColumnNum).Insert(-4161);
    }

    let targetStartCell, targetEndCell;

    switch (document.getElementById("credit-standred-version").value) {
        case "standred-version-auto":
            targetStartCell = outputSheet.Cells.Item(2, 6);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 6);
            targetStartCell.Formula = `=VALUE("20"&LEFT(A2,2))`;
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);

            targetStartCell = outputSheet.Cells.Item(2, 7);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 7);
            targetStartCell.Formula = `=IFS(F2>=2023,"2023版",F2<2023,"2019版")`;
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);
            break;
        case "standred-version-2019":
            targetStartCell = outputSheet.Cells.Item(2, 6);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 6);
            targetStartCell.Value2 = "-";
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);

            targetStartCell = outputSheet.Cells.Item(2, 7);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 7);
            targetStartCell.Value2 = "2019版";
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);
            break;
        case "standred-version-2023":
            targetStartCell = outputSheet.Cells.Item(2, 6);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 6);
            targetStartCell.Value2 = "-";
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);

            targetStartCell = outputSheet.Cells.Item(2, 7);
            targetEndCell = outputSheet.Cells.Item(outputRowNumMax, 7);
            targetStartCell.Value2 = "2023版";
            targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);
            break;
        default:
            alert("unknown credit standred version");
            return;
    }

    let targetColumnNumMap = new Map();
    targetColumnNumMap.set(13, `=SUM(J2:L2)`);
    targetColumnNumMap.set(14, `=SUM(MIN(J2,2),MIN(K2,2),MIN(L2,1))`);
    targetColumnNumMap.set(19, `=SUM(O2:R2)`);
    targetColumnNumMap.set(20, `=IFS(G2="2023版",MIN(O2+SUM(P2:R2),O2+2,3),G2="2019版",MIN(S2,3))`);
    if (document.getElementById("recalculate-total-credits").checked) {
        targetColumnNumMap.set(21, `=SUM(M2,S2)`);
    }
    targetColumnNumMap.set(22, `=SUM(N2,T2)`);
    targetColumnNumMap.set(24, `=IFS(G2="2023版",IF(AND(V2>=8,OR(ISNUMBER(FIND("三下乡",W2)),ISNUMBER(FIND("返家乡",W2)))),"OK","NG"),G2="2019版",IF(V2>=8,"OK","NG"))`);

    for (const [targetColumnNum, formulaStr] of targetColumnNumMap) {

        targetStartCell = outputSheet.Cells.Item(2, targetColumnNum);
        targetEndCell = outputSheet.Cells.Item(outputRowNumMax, targetColumnNum);
        targetStartCell.Formula = formulaStr;
        targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);
    }
}

/**
 * 
 * @param {Et.Worksheet} outputSheet 
 */
function organizeOutputSheet(outputSheet) {
    // old code

    if (document.getElementById("convert-output-data-into-values").checked) {
        outputSheet.UsedRange.Copy();
        outputSheet.Range("A1").PasteSpecial(-4163, -4142, false, false);
    }

    outputSheet.Range("H1:U1").Clear();
    outputSheet.Range("F1:X1").Value2 = ["入学年份", "标准版本", "统计开始时间", "统计结束时间", "思想成长", "创新创业", "志愿公益", "总必修学分", "有效学分", "实践实习", "文体活动", "工作履历", "技能特长", "总选修学分", "有效学分", "总学分", "总有效学分", "实践实习记录", "毕业资格"];

    let columnWidthList = [10, 8, 12, 5, 9, 8, 8, 12, 12, 8, 8, 8, 10, 8, 8, 8, 8, 8, 10, 8, 8, 10, 12, 8];
    for (let i = columnWidthList.length; i >= 1; i--) {
        outputSheet.Columns.Item(i).ColumnWidth = columnWidthList[i - 1];
    }

    outputSheet.UsedRange.RowHeight = 20;

    outputSheet.UsedRange.HorizontalAlignment = -4108;
    outputSheet.Range("B:C").HorizontalAlignment = -4131;
    outputSheet.Range("E:E").HorizontalAlignment = -4131;
    outputSheet.Range("W:W").HorizontalAlignment = -4131;
    outputSheet.Range("1:1").HorizontalAlignment = -4108;

    outputSheet.Range("J:V").NumberFormatLocal = "0.00";
}
/**
 * @author LiBaode <libaode2013@163.com>
 * @version 20260126A
 */
var etApp = window.Application;
const collegeNameList = ["机械工程与自动化学院", "汽车与交通工程学院", "材料科学与工程学院", "化学与环境工程学院", "电气工程学院", "电子与信息工程学院", "经济管理学院", "文化传媒与艺术设计学院", "土木建筑工程学院", "外国语学院", "理学院", "交叉科学学院"];
const summaryTitleList = ["学号", "姓名", "学院", "年级", "班级", "参与开始时间", "参与结束时间", "思想成长类-学分", "创新创业-学分", "志愿公益类-学分", "实践实习类-学分", "文体活动-学分", "工作履历-学分", "技能特长-学分", "学分总和"];
const detailsTitleList = ["姓名", "学号", "院系", "年级", "班级", "学分类型", "单位", "一级分类", "二级分类", "数量", "转换后学分", "获取原因", "发送人", "参与时长（分）", "参与开始时间", "参与结束时间", "获取时间", "来源", "是否有效"];

initUserInterface();

function initUserInterface() {
    try {
        initUserInterface_summaryBookList();
        initUserInterface_detialsBookList();
        initUserInterface_summaryDataFilterSettings();
        initUserInterface_otherHtmlElements();
        // 解锁用户界面
        updateUserInterface("unlock-user-interface");
    } catch (error) {
        let errInfoStr = "";
        errInfoStr += "初始化用户界面时发生错误，错误信息：\n\n";
        errInfoStr += String(error.stack);
        errInfoStr += "\n\n可联系开发者获取更多支持。libaode2013@163.com";
        alert(errInfoStr);
        window.close();
    }
}

function initUserInterface_summaryBookList() {

    const summaryBookListNode = document.getElementById("summary-book-list");

    let summaryBook, summaryBookNode, checkboxNode, labelNode;

    for (let bookNum = etApp.Workbooks.Count; bookNum >= 1; bookNum--) {

        summaryBook = etApp.Workbooks.Item(bookNum);

        summaryBookNode = document.createElement("div");
        checkboxNode = document.createElement("input");
        labelNode = document.createElement("label");

        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `summary-book-${bookNum}`);
        checkboxNode.setAttribute("name", "summary-book");
        checkboxNode.setAttribute("value", summaryBook.Name);

        labelNode.setAttribute("for", `summary-book-${bookNum}`);
        labelNode.textContent = summaryBook.Name;

        if (checkSummaryBook(summaryBook) !== true) {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("style", "color: gray;");
        }

        summaryBookNode.appendChild(checkboxNode);
        summaryBookNode.appendChild(labelNode);
        summaryBookListNode.appendChild(summaryBookNode);
    }
}

function initUserInterface_detialsBookList() {

    const detailsBookListNode = document.getElementById("details-book-list");

    let detailsBook, detailsBookNode, checkboxNode, labelNode;

    for (let bookNum = etApp.Workbooks.Count; bookNum >= 1; bookNum--) {

        detailsBook = etApp.Workbooks.Item(bookNum);

        detailsBookNode = document.createElement("div");
        checkboxNode = document.createElement("input");
        labelNode = document.createElement("label");

        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `details-book-${bookNum}`);
        checkboxNode.setAttribute("name", "details-book");
        checkboxNode.setAttribute("value", detailsBook.Name);

        labelNode.setAttribute("for", `details-book-${bookNum}`);
        labelNode.textContent = detailsBook.Name;

        if (checkDetailsBook(detailsBook) !== true) {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("style", "color: gray;");
        }

        detailsBookNode.appendChild(checkboxNode);
        detailsBookNode.appendChild(labelNode);
        detailsBookListNode.appendChild(detailsBookNode);
    }
}

function initUserInterface_summaryDataFilterSettings() {

    const useCollegeFilter = document.getElementById("use-college-filter");
    const collegeFilterInput = document.getElementById("college-filter-input");

    const useGradeFilter = document.getElementById("use-grade-filter");
    const gradeFilterOperator = document.getElementById("grade-filter-operator");
    const gradeFilterInput = document.getElementById("grade-filter-input");

    const useClassFilter = document.getElementById("use-class-filter");
    const classFilterOperator = document.getElementById("class-filter-operator");
    const classFilterInput = document.getElementById("class-filter-input");

    // 初始化学院筛选器中的学院列表
    let collegeNameNode;
    for (const collegeName of collegeNameList) {
        collegeNameNode = document.createElement("option");
        collegeNameNode.setAttribute("value", collegeName);
        collegeNameNode.textContent = collegeName;
        collegeFilterInput.appendChild(collegeNameNode);
    }

    // 勾选使用筛选器时，解锁对应选项和输入框，反之则锁定
    useCollegeFilter.addEventListener("change", () => {
        if (useCollegeFilter.checked) {
            collegeFilterInput.removeAttribute("disabled");
        }
        else {
            collegeFilterInput.setAttribute("disabled", "");
        }
    });

    useGradeFilter.addEventListener("change", () => {
        if (useGradeFilter.checked) {
            gradeFilterOperator.removeAttribute("disabled");
            gradeFilterInput.removeAttribute("disabled");
        }
        else {
            gradeFilterOperator.setAttribute("disabled", "");
            gradeFilterInput.setAttribute("disabled", "");
        }
    });

    useClassFilter.addEventListener("change", () => {
        if (useClassFilter.checked) {
            classFilterInput.removeAttribute("disabled");
            classFilterOperator.removeAttribute("disabled");
        }
        else {
            classFilterInput.setAttribute("disabled", "");
            classFilterOperator.setAttribute("disabled", "");
        }
    });
}

function initUserInterface_otherHtmlElements() {

    const hideInvalidSummaryFiles = document.getElementById("hide-invalid-summary-files");
    const noValidSummaryFiles = document.getElementById("no-valid-summary-files");

    hideInvalidSummaryFiles.addEventListener("click", () => {
        hideInvalidSummaryFiles.setAttribute("hidden", "");
        for (const checkbox of document.querySelectorAll(`input[name="summary-book"][disabled]`)) {
            checkbox.parentElement.setAttribute("hidden", "");
        }
        if (document.querySelectorAll(`input[name="summary-book"]:not([disabled])`).length === 0) {
            noValidSummaryFiles.removeAttribute("hidden");
        }
    });

    const hideInvalidDetailsFiles = document.getElementById("hide-invalid-details-files");
    const noValidDetailsFiles = document.getElementById("no-valid-details-files");

    hideInvalidDetailsFiles.addEventListener("click", () => {
        hideInvalidDetailsFiles.setAttribute("hidden", "");
        for (const checkbox of document.querySelectorAll(`input[name="details-book"][disabled]`)) {
            checkbox.parentElement.setAttribute("hidden", "");
        }
        if (document.querySelectorAll(`input[name="details-book"]:not([disabled])`).length === 0) {
            noValidDetailsFiles.removeAttribute("hidden");
        }
    });

    const startButton = document.getElementById("start-button");

    startButton.addEventListener("click", () => {
        main();
    });
}


/**
 * 更新用户界面
 * @param {string} operation 要执行的操作
 */
function updateUserInterface(operation) {

    switch (operation) {
        // 锁定用户界面
        case "lock-user-interface":
            for (const fieldset of document.getElementsByName("option-group")) {
                fieldset.setAttribute("disabled", "");
            }
            document.getElementById("start-button").setAttribute("disabled", "");
            break;
        // 解锁用户界面
        case "unlock-user-interface":
            for (const fieldset of document.getElementsByName("option-group")) {
                fieldset.removeAttribute("disabled");
            }
            document.getElementById("start-button").removeAttribute("disabled");
            break;
        default:
            alert("unknown operation parameter for update user interface");
    }
}

/**
 * 
 * @param {Et.Workbook} summaryBook 
 * @returns 
 */
function checkSummaryBook(summaryBook) {

    if (summaryBook.Worksheets.Count !== 1) {
        return false;
    }

    let summarySheet = summaryBook.Worksheets.Item(1);

    if (summarySheet.Name !== "学分汇总") {
        return false;
    }

    for (const summaryTitle of summaryTitleList) {
        if (summarySheet.Rows.Item(1).Find(summaryTitle, undefined, -4163, 1) === null) {
            return false;
        }
    }

    let summaryRowNumMax = summarySheet.UsedRange.Rows.Count;

    if (summaryRowNumMax < 2) {
        return false;
    }

    if (checkSummaryData(summarySheet) !== "pass") {
        return false;
    }

    return true;
}

/**
 * 
 * @param {Et.Worksheet} summarySheet 
 */
function checkSummaryData(summarySheet) {

    let summaryBookName = summarySheet.Parent.Name;

    let summarySheetName = summarySheet.Name;

    let summaryRowNumMax = summarySheet.UsedRange.Rows.Count;

    let targetColumnNum, targetStartCell, targetEndCell, targetRangeAddr, formulaStr;

    let checkResult;

    targetColumnNum = summarySheet.Rows.Item(1).Find("学号", undefined, -4163, 1).Column;
    targetStartCell = summarySheet.Cells.Item(2, targetColumnNum);
    targetEndCell = summarySheet.Cells.Item(summaryRowNumMax, targetColumnNum);
    targetRangeAddr = summarySheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${summaryBookName}]${summarySheetName}'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== summaryRowNumMax - 1) {
        checkResult = "汇总数据中，存在学号相同的重复数据行。";
        return checkResult;
    }

    targetColumnNum = summarySheet.Rows.Item(1).Find("参与开始时间", undefined, -4163, 1).Column;
    targetStartCell = summarySheet.Cells.Item(2, targetColumnNum);
    targetEndCell = summarySheet.Cells.Item(summaryRowNumMax, targetColumnNum);
    targetRangeAddr = summarySheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${summaryBookName}]${summarySheetName}'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        checkResult = "汇总数据中，参与开始时间不唯一。";
        return checkResult;
    }

    targetColumnNum = summarySheet.Rows.Item(1).Find("参与结束时间", undefined, -4163, 1).Column;
    targetStartCell = summarySheet.Cells.Item(2, targetColumnNum);
    targetEndCell = summarySheet.Cells.Item(summaryRowNumMax, targetColumnNum);
    targetRangeAddr = summarySheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${summaryBookName}]${summarySheetName}'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        checkResult = "汇总数据中，参与结束时间不唯一。";
        return checkResult;
    }

    checkResult = "pass";
    return checkResult;
}

/**
 * 
 * @param {Et.Workbook} detailsBook 
 * @returns 
 */
function checkDetailsBook(detailsBook) {

    if (detailsBook.Worksheets.Count !== 1) {
        return false;
    }

    let detailsSheet = detailsBook.Worksheets.Item(1);

    if (detailsSheet.Name !== "学分明细") {
        return false;
    }

    for (let i = 1; i <= detailsTitleList.length; i++) {
        if (detailsSheet.Cells.Item(1, i).Value2 !== detailsTitleList[i - 1]) {
            return false;
        }
    }

    if (detailsSheet.Rows.Count < 2) {
        return false;
    }

    if (checkDetailsData(detailsSheet) !== "pass") {
        return false;
    }

    return true;
}

/**
 * 
 * @param {Et.Worksheet} inputSheet 
 * @returns 
 */
function checkDetailsData(inputSheet) {

    let inputBookName = inputSheet.Parent.Name;

    let inputSheetName = inputSheet.Name;

    let inputRowNumMax = inputSheet.UsedRange.Rows.Count;

    let targetColumnNum, targetStartCell, targetEndCell, targetRangeAddr, formulaStr;

    let checkResult;

    targetColumnNum = inputSheet.Rows.Item(1).Find("学分类型", undefined, -4163, 1).Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTIFS('[${inputBookName}]${inputSheetName}'!${targetRangeAddr},"实践实习类",`;
    targetColumnNum = inputSheet.Rows.Item(1).Find("单位", undefined, -4163, 1).Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr += `'[${inputBookName}]${inputSheetName}'!${targetRangeAddr},"学分",`;
    targetColumnNum = inputSheet.Rows.Item(1).Find("是否有效", undefined, -4163, 1).Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr += `'[${inputBookName}]${inputSheetName}'!${targetRangeAddr},"是")`;

    if (etApp.Evaluate(formulaStr) < 1) {
        checkResult = "明细数据中，未找到有效的实践实习学分列。";
        return checkResult;
    }

    checkResult = "pass";
    return checkResult;
}

function main() {
    try {
        // 锁定用户界面
        updateUserInterface("lock-user-interface");

        if (document.querySelectorAll(`input[name="summary-book"]:checked`).length === 0) {
            alert("未选择学分汇总文件。");
            updateUserInterface("unlock-user-interface");
            return;
        }
        if (document.querySelectorAll(`input[name="details-book"]:checked`).length === 0) {
            alert("未选择学分明细文件。");
            updateUserInterface("unlock-user-interface");
            return;
        }

        let tempSheet, checkResult;

        tempSheet = createTempSummarySheet();

        checkResult = checkSummaryData(tempSheet);
        if (checkResult !== "pass") {
            closeBookSilently(tempSheet.Parent);
            alert(checkResult);
            updateUserInterface("unlock-user-interface");
            return;
        }

        let filteredSummaryData = getFilteredSummaryData(tempSheet);

        // 临时工作簿用完了，关闭临时工作簿
        // closeBookSilently(tempSheet.Parent);

        // 如果没找到符合筛选条件的输入数据，（检查过程中）弹窗提示，解锁用户界面，终止执行
        if (filteredSummaryData.length === 0) {
            return;
        }

        tempSheet = createTempDetailsSheet();

        checkResult = checkDetailsData(tempSheet);
        if (checkResult !== "pass") {
            closeBookSilently(tempSheet.Parent);
            alert(checkResult);
            updateUserInterface("unlock-user-interface");
            return;
        }

        let filteredDetailsData = getFilteredDetailsData(tempSheet);

        // closeBookSilently(tempSheet.Parent);

        if (filteredDetailsData.length === 0) {
            return;
        }

        let outputBook = createOutputSummaryBook(filteredSummaryData, filteredDetailsData);

        alert(`处理完成，输出数据在【${outputBook.Name}】。`);

        window.close();

    } catch (error) {
        let errInfoStr = "";
        errInfoStr += "运行主函数过程时发生错误，错误信息：\n\n";
        errInfoStr += String(error.stack);
        errInfoStr += "\n\n可联系开发者获取更多支持。libaode2013@163.com";
        alert(errInfoStr);
    }
}

function createTempSummarySheet() {

    let tempBook = etApp.Workbooks.Add();

    let inputSheet, inputSheetNameList = [];

    for (const checkbox of document.querySelectorAll(`input[name="summary-book"]:checked`)) {
        inputSheet = etApp.Workbooks.Item(checkbox.value).Worksheets.Item(1);
        inputSheet.Copy(tempBook.Worksheets.Item(1), null);
        inputSheetNameList.push(etApp.ActiveSheet.Name);
    }

    let tempSheet = createOutputSummarySheet(tempBook, "temp-summary-sheet");

    let inputRowNumMax, tempRowNum, targetStartCell, targetEndCell;

    for (const inputSheetName of inputSheetNameList) {
        inputSheet = tempBook.Worksheets.Item(inputSheetName);

        organizeInputSummarySheet(inputSheet);

        inputRowNumMax = inputSheet.UsedRange.Rows.Count;
        tempRowNum = tempSheet.UsedRange.Rows.Count + 1;

        targetStartCell = inputSheet.Cells.Item(2, 1);
        targetEndCell = inputSheet.Cells.Item(inputRowNumMax, summaryTitleList.length);

        inputSheet.Range(targetStartCell, targetEndCell).Copy(tempSheet.Cells.Item(tempRowNum, 1));
    }

    return tempSheet;
}

/**
 * 创建输出工作表
 * @param {Et.Workbook} outputBook 输出工作簿
 * @param {string} outputSheetName 输出工作表名称
 * @returns 创建的输出工作表
 */
function createOutputSummarySheet(outputBook, outputSheetName) {

    // 在输出工作簿中，新建工作表作为输出工作表
    let outputSheet = outputBook.Worksheets.Add();

    // 重命名输出工作表
    outputSheet.Name = outputSheetName;

    let targetStartCell = outputSheet.Cells.Item(1, 1);
    let targetEndCell = outputSheet.Cells.Item(1, summaryTitleList.length);
    // 向输出工作表的首行，填入数据标题
    outputSheet.Range(targetStartCell, targetEndCell).Value2 = summaryTitleList;

    return outputSheet;
}

/**
 * 整理输入工作表
 * @param {Et.Worksheet} inputSheet 要整理的输入工作表
 */
function organizeInputSummarySheet(inputSheet) {

    let inputDataTitle;

    for (let inputColumnNum = inputSheet.UsedRange.Columns.Count; inputColumnNum >= 1; inputColumnNum--) {

        inputDataTitle = String(inputSheet.Cells.Item(1, inputColumnNum).Value2);

        if (summaryTitleList.includes(inputDataTitle) === false) {
            inputSheet.Columns.Item(inputColumnNum).Delete();
        }
    }

    let targetColumnNum;

    for (let i = summaryTitleList.length - 1; i >= 0; i--) {

        inputSheet.Columns.Item(1).Insert(-4161);

        targetColumnNum = inputSheet.Rows.Item(1).Find(summaryTitleList[i], undefined, -4163, 1).Column;

        inputSheet.Columns.Item(targetColumnNum).Cut(inputSheet.Columns.Item(1));
    }
}

/**
 * 
 * @param {Et.Worksheet} tempSheet 
 */
function getFilteredSummaryData(tempSheet) {

    let tempRowNumMax = tempSheet.UsedRange.Rows.Count;

    let targetStartCell = tempSheet.Cells.Item(2, 1);
    let targetEndCell = tempSheet.Cells.Item(tempRowNumMax, summaryTitleList.length);

    let originalSummaryData = tempSheet.Range(targetStartCell, targetEndCell).Value2;

    if (document.querySelectorAll(`input[name="summary-data-filter"]:checked`).length === 0) {
        return originalSummaryData;
    }

    let positionIndex, filterValue;

    let filteredSummaryData = originalSummaryData;

    if (document.getElementById("use-college-filter").checked) {
        positionIndex = summaryTitleList.indexOf("学院");
        filterValue = String(document.getElementById("college-filter-input").value);
        filteredSummaryData = filteredSummaryData.filter((inputDataRow) => String(inputDataRow[positionIndex] === filterValue));
    }

    // 根据设置的数据筛选器，筛选【年级】符合条件的数据
    if (document.getElementById("use-grade-filter").checked) {
        positionIndex = summaryTitleList.indexOf("年级");
        filterValue = Number(document.getElementById("grade-filter-input").value);
        switch (document.getElementById("grade-filter-operator").value) {
            case "greater":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) > filterValue);
                break;
            case "less":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) < filterValue);
                break;
            case "equal":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) === filterValue);
                break;
            case "greater-or-equal":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) >= filterValue);
                break;
            case "less-or-equal":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) <= filterValue);
                break;
            case "not-equal":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) !== filterValue);
                break;
            default:
                alert("unknown grade filter operator");
                return [];
        }
    }

    // 根据设置的数据筛选器，筛选【班级】符合条件的数据
    if (document.getElementById("use-class-filter").checked) {
        positionIndex = summaryTitleList.indexOf("班级");
        filterValue = String(document.getElementById("class-filter-input").value);
        switch (document.getElementById("class-filter-operator").value) {
            case "exact-match":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => String(inputDataRow[positionIndex]) === filterValue);
                break;
            case "partial-match":
                filteredSummaryData = filteredSummaryData.filter((inputDataRow) => String(inputDataRow[positionIndex]).includes(filterValue));
                break;
            default:
                alert("unknown class filter operator");
                return [];
        }
    }


    // 如果筛选后的数据为空，数组长度为0，弹窗提示，解锁用户界面，返回两次终止主函数
    if (filteredSummaryData.length === 0) {
        alert("没有符合筛选条件的输入数据。");
        updateUserInterface("unlock-user-interface");
        return [];
    }

    // 返回筛选后的输入数据
    return filteredSummaryData;
}

function createTempDetailsSheet() {

    let tempBook = etApp.Workbooks.Add();

    let inputSheet, inputSheetNameList = [];

    for (const checkbox of document.querySelectorAll(`input[name="details-book"]:checked`)) {
        inputSheet = etApp.Workbooks.Item(checkbox.value).Worksheets.Item(1);
        inputSheet.Copy(tempBook.Worksheets.Item(1), null);
        inputSheetNameList.push(etApp.ActiveSheet.Name);
    }

    let tempSheet = tempBook.Worksheets.Add();
    tempSheet.Name = "temp-details-sheet";

    let targetStartCell, targetEndCell;

    targetStartCell = tempSheet.Cells.Item(1, 1);
    targetEndCell = tempSheet.Cells.Item(1, detailsTitleList.length);
    tempSheet.Range(targetStartCell, targetEndCell).Value2 = detailsTitleList;

    let inputRowNumMax, tempRowNum;

    for (const inputSheetName of inputSheetNameList) {
        inputSheet = tempBook.Worksheets.Item(inputSheetName);

        inputRowNumMax = inputSheet.UsedRange.Rows.Count;
        tempRowNum = tempSheet.UsedRange.Rows.Count + 1;

        targetStartCell = inputSheet.Cells.Item(2, 1);
        targetEndCell = inputSheet.Cells.Item(inputRowNumMax, detailsTitleList.length);

        inputSheet.Range(targetStartCell, targetEndCell).Copy(tempSheet.Cells.Item(tempRowNum, 1));
    }

    return tempSheet;
}

/**
 * 
 * @param {Et.Worksheet} tempSheet 
 */
function getFilteredDetailsData(tempSheet) {

    let tempRowNumMax = tempSheet.UsedRange.Rows.Count;

    let targetStartCell = tempSheet.Cells.Item(2, 1);
    let targetEndCell = tempSheet.Cells.Item(tempRowNumMax, detailsTitleList.length);

    let filteredDetailsData = tempSheet.Range(targetStartCell, targetEndCell).Value2;

    let positionIndex, filterValue;

    positionIndex = detailsTitleList.indexOf("学分类型");
    filterValue = "实践实习类";
    filteredDetailsData = filteredDetailsData.filter((detailsDataRow) => String(detailsDataRow[positionIndex]) === filterValue);

    positionIndex = detailsTitleList.indexOf("单位");
    filterValue = "学分";
    filteredDetailsData = filteredDetailsData.filter((detailsDataRow) => String(detailsDataRow[positionIndex]) === filterValue);

    positionIndex = detailsTitleList.indexOf("是否有效");
    filterValue = "是";
    filteredDetailsData = filteredDetailsData.filter((detailsDataRow) => String(detailsDataRow[positionIndex]) === filterValue);

    // 如果筛选后的数据为空，数组长度为0，弹窗提示，解锁用户界面，返回两次终止主函数
    if (filteredDetailsData.length === 0) {
        alert("没有符合筛选条件的输入数据。");
        updateUserInterface("unlock-user-interface");
        return [];
    }

    // 返回筛选后的输入数据
    return filteredDetailsData;
}

/**
 * 
 * @param {Array} filteredSummaryData 
 * @param {Array} filteredDetailsData 
 * @returns 
 */
function createOutputSummaryBook(filteredSummaryData, filteredDetailsData) {

    let outputBook = etApp.Workbooks.Add();

    let outputSheetType = document.getElementById("output-sheet-type").value;

    let outputSheetName, outputSheetNameList = [], outputSheet, outputRowNum, targetStartCell, targetEndCell;

    let summaryStunumIndex = summaryTitleList.indexOf("学号");
    let detailsStunumIndex = detailsTitleList.indexOf("学号");
    let detailsReasonIndex = detailsTitleList.indexOf("获取原因");
    let detailsDataRow, detailsDataValue;

    for (let i = 0; i < filteredSummaryData.length; i++) {

        switch (outputSheetType) {
            case "sheet-type-college":
                outputSheetName = filteredSummaryData[i][summaryTitleList.indexOf("学院")];
                break;
            case "sheet-type-grade":
                outputSheetName = filteredSummaryData[i][summaryTitleList.indexOf("年级")];
                break;
            case "sheet-type-class":
                outputSheetName = filteredSummaryData[i][summaryTitleList.indexOf("班级")];
                break;
            case "sheet-type-all-in-one":
                outputSheetName = "output-sheet";
                break;
            default:
                alert("unknown output sheet type");
        }

        // 重定向输出工作表，如果没有则新建，获取输出位置（行的号码）
        if (outputSheetNameList.includes(outputSheetName)) {
            outputSheet = outputBook.Worksheets.Item(outputSheet);
            outputRowNum = outputSheet.UsedRange.Rows.Count + 1;
        }
        else {
            outputSheet = createOutputSummarySheet(outputBook, outputSheetName);
            outputSheetNameList.push(outputSheetName);
            outputRowNum = 2;
        }

        targetStartCell = outputSheet.Cells.Item(outputRowNum, 1);
        targetEndCell = outputSheet.Cells.Item(outputRowNum, summaryTitleList.length);

        outputSheet.Range(targetStartCell, targetEndCell).Value2 = filteredSummaryData[i];

        // 写入明细数据
        targetStartCell = outputSheet.Cells.Item(outputRowNum, summaryTitleList.length + 1);
        detailsDataRow = filteredDetailsData.find((detailsDataRow) => {
            const summaryStunum = String(filteredSummaryData[i][summaryStunumIndex]);
            const detailsStunum = String(detailsDataRow[detailsStunumIndex]);
            const detailsReason = String(detailsDataRow[detailsReasonIndex]);
            return detailsStunum === summaryStunum && (detailsReason.includes("三下乡") || detailsReason.includes("返家乡"));
        });

        if (detailsDataRow !== undefined) {
            detailsDataValue = detailsDataRow[detailsReasonIndex];
            targetStartCell.Value2 = detailsDataValue;
        }
        else {
            targetStartCell.Value2 = "-";
        }
    }

    for (let outputSheetNum = outputBook.Worksheets.Count; outputSheetNum >= 1; outputSheetNum--) {

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
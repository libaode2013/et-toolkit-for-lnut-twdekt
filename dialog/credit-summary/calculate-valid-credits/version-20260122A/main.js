/**
 * @author LiBaode <libaode2013@163.com>
 * @version 20260122A
 */
var etApp = window.Application;
const collegeNameList = ["机械工程与自动化学院", "汽车与交通工程学院", "材料科学与工程学院", "化学与环境工程学院", "电气工程学院", "电子与信息工程学院", "经济管理学院", "文化传媒与艺术设计学院", "土木建筑工程学院", "外国语学院", "理学院", "交叉科学学院"];
const inputDataTitleList = ["学号", "姓名", "学院", "年级", "班级", "参与开始时间", "参与结束时间", "思想成长类-学分", "创新创业-学分", "志愿公益类-学分", "实践实习类-学分", "文体活动-学分", "工作履历-学分", "技能特长-学分", "学分总和"];

// 初始化用户界面
initUserInterface();

function initUserInterface() {
    try {
        initUserInterface_inputBookList();
        initUserInterface_inputDataFilterSettings();
        initUserInterface_outputSheetSettings();
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

function initUserInterface_inputBookList() {

    const inputBookListNode = document.getElementById("input-book-list");

    let inputBook, inputBookNode, checkboxNode, labelNode;

    for (let inputBookNum = etApp.Workbooks.Count; inputBookNum >= 1; inputBookNum--) {

        inputBook = etApp.Workbooks.Item(inputBookNum);

        inputBookNode = document.createElement("div");
        checkboxNode = document.createElement("input");
        labelNode = document.createElement("label");

        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `input-book-${inputBookNum}`);
        checkboxNode.setAttribute("name", "input-book");
        checkboxNode.setAttribute("value", inputBook.Name);

        labelNode.setAttribute("for", `input-book-${inputBookNum}`);
        labelNode.textContent = inputBook.Name;

        // 如果检查不通过，禁用复选框，标签文字显示为灰色，提示此文件不可选中
        if (checkInputBook(inputBook) === false) {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("style", "color: gray;");
        }

        inputBookNode.appendChild(checkboxNode);
        inputBookNode.appendChild(labelNode);
        inputBookListNode.appendChild(inputBookNode);
    }
}

/**
 * 检查输入工作簿
 * @param {Et.Workbook} inputBook 要检查的输入工作簿
 * @returns 检查结果，不通过返回false
 */
function checkInputBook(inputBook) {

    // 检查工作簿中的工作表，应有且仅有1个，名为“学分汇总”的工作表
    if (inputBook.Worksheets.Count !== 1) {
        return false;
    }

    let inputSheet = inputBook.Worksheets.Item(1);

    if (inputSheet.Name !== "学分汇总") {
        return false;
    }

    // 检查工作表的数据标题，应有所有数组元素
    for (const inputDataTitle of inputDataTitleList) {
        if (inputSheet.Rows.Item(1).Find(inputDataTitle) === null) {
            return false;
        }
    }

    // 检查数据行数，应至少有1行数据
    let inputRowNumMax = inputSheet.UsedRange.Rows.Count;

    if (inputRowNumMax < 2) {
        return false;
    }

    let targetColumnNum, targetStartCell, targetEndCell, targetRangeAddr, formulaStr;

    // 检查此工作表中，是否存在重复的学号
    targetColumnNum = inputSheet.Rows.Item(1).Find("学号").Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${inputBook.Name}]学分汇总'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== inputRowNumMax - 1) {
        return false;
    }

    // 检查此工作表中，参与开始时间是否唯一
    targetColumnNum = inputSheet.Rows.Item(1).Find("参与开始时间").Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${inputBook.Name}]学分汇总'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        return false;
    }

    // 检查此工作表中，参与结束时间是否唯一
    targetColumnNum = inputSheet.Rows.Item(1).Find("参与结束时间").Column;
    targetStartCell = inputSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = inputSheet.Cells.Item(inputRowNumMax, targetColumnNum);
    targetRangeAddr = inputSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${inputBook.Name}]学分汇总'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        return false;
    }

    return true;
}

function initUserInterface_inputDataFilterSettings() {

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

function initUserInterface_outputSheetSettings() {

    // 删除参评资格为NG的数据行，首先要计算参评资格
    // 添加事件监听器：勾选【删除NG数据行】时，自动勾选【计算年均有效学分】；取消勾选【计算年均有效学分】时，自动取消勾选【删除NG数据行】
    const calculateAnnualAverageValidCredits = document.getElementById("calculate-annual-average-valid-credits");
    const deleteNotGoodResultRows = document.getElementById("delete-not-good-result-rows");

    calculateAnnualAverageValidCredits.addEventListener("change", () => {
        if (calculateAnnualAverageValidCredits.checked === false) {
            deleteNotGoodResultRows.checked = false;
        }
    });
    deleteNotGoodResultRows.addEventListener("change", () => {
        if (deleteNotGoodResultRows.checked) {
            calculateAnnualAverageValidCredits.checked = true;
        }
    });

    // 由于部分数据为计算结果，以来原始数据完整性，当勾选【任意“删除列数据”】时，自动勾选【输出数据转换为值】，先全部转换为值，取消依赖，再删除
    // 所以当取消勾选【输出数据转换为值】时，自动取消勾选【全部“删除列数据”】

    // 当指定【学分标准版本】时，版本与【入学时间】无关，故【入学时间】列数据不再计算，用“-”占位，且此列数据需要删除

    // 故当指定【学分标准版本】时，自动勾选【删除入学时间列数据】；又因为勾选了【任意“删除列数据”】，自动勾选【输出数据转换为值】
    // 当取消勾选【删除入学时间列数据】时，自动将【学分标准版本】调整为“自动匹配”

    // 而当取消勾选【输出数据转换为值】时，自动取消勾选【全部“删除列数据”】，包括【删除入学时间列数据】，此时自动将【学分标准版本】调整为“自动匹配”

    // 用户界面交互逻辑如上，添加的事件监听器如下
    const creditStandredVersion = document.getElementById("credit-standred-version");
    const convertOutputDataIntoValues = document.getElementById("convert-output-data-into-values");
    const deleteEnrollmentYearColumn = document.getElementById("delete-enrollment-year-column");
    const deleteColumnsInOutputSheets = document.getElementsByName("delete-columns-in-output-sheets");

    creditStandredVersion.addEventListener("change", () => {
        switch (creditStandredVersion.value) {
            case "standred-version-2019":
            case "standred-version-2023":
                deleteEnrollmentYearColumn.checked = true;
                deleteEnrollmentYearColumn.dispatchEvent(new Event("change"));
                break;
            case "standred-version-auto":
                break;
            default:
                alert("invalid credit standred version");
                window.close();
        }
    });

    convertOutputDataIntoValues.addEventListener("change", () => {
        if (convertOutputDataIntoValues.checked === false) {
            for (const checkbox of deleteColumnsInOutputSheets) {
                checkbox.checked = false;
            }
            deleteEnrollmentYearColumn.dispatchEvent(new Event("change"));
        }
    });

    deleteEnrollmentYearColumn.addEventListener("change", () => {
        if (deleteEnrollmentYearColumn.checked === false) {
            creditStandredVersion.value = "standred-version-auto";
        }
    });

    for (const checkbox of deleteColumnsInOutputSheets) {
        checkbox.addEventListener("change", () => {
            if (checkbox.checked) {
                convertOutputDataIntoValues.checked = true;
            }
        });
    }

    // 添加事件监听器：点击【显示高级选项】按钮时，隐藏按钮，并取消隐藏选项
    const showAdvancedOptions = document.getElementById("show-advanced-options-for-output-sheets");
    const advancedOptions = document.getElementById("advanced-options-for-output-sheets");
    showAdvancedOptions.addEventListener("click", () => {
        showAdvancedOptions.setAttribute("hidden", "");
        advancedOptions.removeAttribute("hidden");
    });
}

function initUserInterface_otherHtmlElements() {

    // 添加事件监听器：当点击【隐藏无效的输入文件】按钮时，将【被禁用的复选框的父元素】隐藏
    // 如果没有【没被禁用的复选框】，取消隐藏标签，提示未找到有效的输入文件
    const hideInvalidInputFiles = document.getElementById("hide-invalid-input-files");
    const noValidInputFile = document.getElementById("no-valid-input-file");

    hideInvalidInputFiles.addEventListener("click", () => {
        hideInvalidInputFiles.setAttribute("hidden", "");
        for (const checkbox of document.querySelectorAll(`input[name="input-book"][disabled]`)) {
            checkbox.parentElement.setAttribute("hidden", "");
        }
        if (document.querySelectorAll(`input[name="input-book"]:not([disabled])`).length === 0) {
            noValidInputFile.removeAttribute("hidden");
        }
    });

    // 开始按钮，点击运行main函数
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
            alert("unknown operation param for update UI");
    }
}

function main() {
    try {

        // 锁定用户界面
        updateUserInterface("lock-user-interface");

        // 如果没勾选任何输入工作簿，弹窗提示，解锁用户界面，终止执行
        if (document.querySelectorAll(`input[name="input-book"]:checked`).length === 0) {
            alert("未选择输入工作簿。");
            updateUserInterface("unlock-user-interface");
            return;
        }

        let tempSheet = createTempSheet();

        // 检查合并后的输入数据，失败则弹窗提示，解锁用户界面，终止执行
        if (checkInputData(tempSheet) === false) {
            updateUserInterface("unlock-user-interface");
            return;
        }

        // 获取临时工作表中的输入数据
        let filteredInputData = getFilteredInputData(tempSheet);

        // 临时工作簿用完了，关闭临时工作簿
        closeBookSilently(tempSheet.Parent);

        // 如果没找到符合筛选条件的输入数据，（检查过程中）弹窗提示，解锁用户界面，终止执行
        if (filteredInputData.length === 0) {
            return;
        }

        // 生成输出工作簿
        let outputBook = createOutputBook(filteredInputData);

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

/**
 * 创建临时工作表
 * @returns 存有所有输入数据的临时工作表
 */
function createTempSheet() {

    //  新建工作簿作为临时工作簿
    let tempBook = etApp.Workbooks.Add();

    let inputSheet, inputSheetNameList = [];

    // 将选择的输入工作簿，中的输入工作表，复制到临时工作簿，并存储新工作表名到数组
    for (const checkbox of document.querySelectorAll(`input[name="input-book"]:checked`)) {
        inputSheet = etApp.Workbooks.Item(checkbox.value).Worksheets.Item(1);
        inputSheet.Copy(tempBook.Worksheets.Item(1), null);
        inputSheetNameList.push(etApp.ActiveSheet.Name);
    }

    // 新建工作表作为临时工作表，重命名并填入数据标题
    let tempSheet = createOutputSheet(tempBook, "temp-sheet");

    let inputRowNumMax, tempRowNum, targetStartCell, targetEndCell;

    // 逐个处理临时工作簿中的输入工作表副本
    for (const inputSheetName of inputSheetNameList) {
        // 重定向输入工作表
        inputSheet = tempBook.Worksheets.Item(inputSheetName);
        // 整理输入工作表
        organizeInputSheet(inputSheet);

        inputRowNumMax = inputSheet.UsedRange.Rows.Count;
        tempRowNum = tempSheet.UsedRange.Rows.Count + 1;

        targetStartCell = inputSheet.Cells.Item(2, 1);
        targetEndCell = inputSheet.Cells.Item(inputRowNumMax, inputDataTitleList.length);
        // 复制其中的数据到临时工作表
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
function createOutputSheet(outputBook, outputSheetName) {

    // 在输出工作簿中，新建工作表作为输出工作表
    let outputSheet = outputBook.Worksheets.Add();

    // 重命名输出工作表
    outputSheet.Name = outputSheetName;

    let targetStartCell = outputSheet.Cells.Item(1, 1);
    let targetEndCell = outputSheet.Cells.Item(1, inputDataTitleList.length);
    // 向输出工作表的首行，填入数据标题
    outputSheet.Range(targetStartCell, targetEndCell).Value2 = inputDataTitleList;

    return outputSheet;
}

/**
 * 整理输入工作表
 * @param {Et.Worksheet} inputSheet 要整理的输入工作表
 */
function organizeInputSheet(inputSheet) {

    let inputDataTitle;

    // 删除数据标题不在数组中的数据列
    for (let inputColumnNum = inputSheet.UsedRange.Columns.Count; inputColumnNum >= 1; inputColumnNum--) {

        inputDataTitle = String(inputSheet.Cells.Item(1, inputColumnNum).Value2);

        if (inputDataTitleList.includes(inputDataTitle) === false) {

            inputSheet.Columns.Item(inputColumnNum).Delete();
        }
    }

    let targetColumnNum;

    // 按数组中的标题顺序排列数据
    for (let i = inputDataTitleList.length - 1; i >= 0; i--) {

        inputSheet.Columns.Item(1).Insert(-4161);

        targetColumnNum = inputSheet.Rows.Item(1).Find(inputDataTitleList[i]).Column;

        inputSheet.Columns.Item(targetColumnNum).Cut(inputSheet.Columns.Item(1));
    }
}

/**
 * 检查输入数据
 * @param {Et.Worksheet} tempSheet 临时工作表
 * @returns 
 */
function checkInputData(tempSheet) {

    let tempBookName = tempSheet.Parent.Name;

    let tempRowNumMax = tempSheet.UsedRange.Rows.Count;

    let targetColumnNum, targetStartCell, targetEndCell, targetRangeAddr, formulaStr;

    // 检查所有输入数据中，是否存在重复的学号
    targetColumnNum = inputDataTitleList.indexOf("学号") + 1;
    targetStartCell = tempSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = tempSheet.Cells.Item(tempRowNumMax, targetColumnNum);
    targetRangeAddr = tempSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${tempBookName}]temp-sheet'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== tempRowNumMax - 1) {
        closeBookSilently(tempSheet.Parent);
        alert("输入数据中，存在学号相同的重复数据行。");
        return false;
    }

    // 检查所有输入数据中，参与开始时间是否唯一
    targetColumnNum = inputDataTitleList.indexOf("参与开始时间") + 1;
    targetStartCell = tempSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = tempSheet.Cells.Item(tempRowNumMax, targetColumnNum);
    targetRangeAddr = tempSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${tempBookName}]temp-sheet'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        closeBookSilently(tempSheet.Parent);
        alert("输入数据中，参与开始时间不唯一。");
        return false;
    }

    // 检查所有输入数据中，参与结束时间是否唯一
    targetColumnNum = inputDataTitleList.indexOf("参与结束时间") + 1;
    targetStartCell = tempSheet.Cells.Item(2, targetColumnNum);
    targetEndCell = tempSheet.Cells.Item(tempRowNumMax, targetColumnNum);
    targetRangeAddr = tempSheet.Range(targetStartCell, targetEndCell).Address(false, false);
    formulaStr = `COUNTA(UNIQUE('[${tempBookName}]temp-sheet'!${targetRangeAddr}))`;
    if (etApp.Evaluate(formulaStr) !== 1) {
        closeBookSilently(tempSheet.Parent);
        alert("输入数据中，参与结束时间不唯一。");
        return false;
    }
    // 如果勾选了【计算学年平均有效学分】，检查参与结束时间是否落在6-9月
    if (document.getElementById("calculate-annual-average-valid-credits").checked) {
        targetRangeAddr = targetStartCell.Address(false, false);
        formulaStr = `MONTH('[${tempBookName}]temp-sheet'!${targetRangeAddr})`;
        if (etApp.Evaluate(formulaStr) < 6 || etApp.Evaluate(formulaStr) > 9) {
            closeBookSilently(tempSheet.Parent);
            alert("输入数据中，参与结束时间应落在6-9月。");
            return false;
        }
    }

    return true;
}

/**
 * 静默关闭工作簿
 * @param {Et.Workbook} book 要关闭的工作簿
 */
function closeBookSilently(book) {

    etApp.DisplayAlerts = false;
    book.Close();
    etApp.DisplayAlerts = true;
}

/**
 * 获取筛选后的数据
 * @param {Et.Worksheet} tempSheet 临时工作簿
 * @returns 筛选后的数据（二维数组）
 */
function getFilteredInputData(tempSheet) {

    let tempRowNumMax = tempSheet.UsedRange.Rows.Count;

    let targetStartCell = tempSheet.Range("A2");
    let targetEndCell = tempSheet.Cells.Item(tempRowNumMax, inputDataTitleList.length);

    // 将原始数据存储到数组（二维数组）
    let originalInputData = tempSheet.Range(targetStartCell, targetEndCell).Value2;

    // 如果没启用任何数据筛选器，直接返回原始数据
    if (document.querySelectorAll(`input[name="input-data-filter"]:checked`).length === 0) {
        return originalInputData;
    }

    let positionIndex, filterValue;

    let filteredInputData = originalInputData;

    // 根据设置的数据筛选器，筛选【学院】符合条件的数据
    if (document.getElementById("use-college-filter").checked) {
        positionIndex = inputDataTitleList.indexOf("学院");
        filterValue = String(document.getElementById("college-filter-input").value);
        filteredInputData = filteredInputData.filter((inputDataRow) => String(inputDataRow[positionIndex]) === filterValue);
    }

    // 根据设置的数据筛选器，筛选【年级】符合条件的数据
    if (document.getElementById("use-grade-filter").checked) {
        positionIndex = inputDataTitleList.indexOf("年级");
        filterValue = Number(document.getElementById("grade-filter-input").value);
        switch (document.getElementById("grade-filter-operator").value) {
            case "greater":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) > filterValue);
                break;
            case "less":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) < filterValue);
                break;
            case "equal":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) === filterValue);
                break;
            case "greater-or-equal":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) >= filterValue);
                break;
            case "less-or-equal":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) <= filterValue);
                break;
            case "not-equal":
                filteredInputData = filteredInputData.filter((inputDataRow) => Number(inputDataRow[positionIndex]) !== filterValue);
                break;
            default:
                alert("unknown grade filter operator");
                return [];
        }
    }

    // 根据设置的数据筛选器，筛选【班级】符合条件的数据
    if (document.getElementById("use-class-filter").checked) {
        positionIndex = inputDataTitleList.indexOf("班级");
        filterValue = String(document.getElementById("class-filter-input").value);
        switch (document.getElementById("class-filter-operator").value) {
            case "exact-match":
                filteredInputData = filteredInputData.filter((inputDataRow) => String(inputDataRow[positionIndex]) === filterValue);
                break;
            case "partial-match":
                filteredInputData = filteredInputData.filter((inputDataRow) => String(inputDataRow[positionIndex]).includes(filterValue));
                break;
            default:
                alert("unknown class filter operator");
                return [];
        }
    }

    // 如果筛选后的数据为空，数组长度为0，弹窗提示，解锁用户界面，返回两次终止主函数
    if (filteredInputData.length === 0) {
        alert("没有符合筛选条件的输入数据。");
        updateUserInterface("unlock-user-interface");
        return [];
    }

    // 返回筛选后的输入数据
    return filteredInputData;
}

/**
 * 创建输出工作簿
 * @param {Array} filteredInputData 筛选后的数据
 * @returns 输出工作簿
 */
function createOutputBook(filteredInputData) {

    // 新建工作簿作为输出工作簿
    let outputBook = etApp.Workbooks.Add();

    // 获取设置的表格拆分方式，这将决定输出工作表的名称
    let outputSheetType = document.getElementById("output-sheet-type").value;

    let outputSheetName, outputSheetNameList = [], outputSheet, outputRowNum, targetStartCell, targetEndCell;

    for (let i = 0; i < filteredInputData.length; i++) {

        // 获取输出工作表名称
        switch (outputSheetType) {
            case "sheet-type-college":
                outputSheetName = filteredInputData[i][inputDataTitleList.indexOf("学院")];
                break;
            case "sheet-type-grade":
                outputSheetName = filteredInputData[i][inputDataTitleList.indexOf("年级")];
                break;
            case "sheet-type-class":
                outputSheetName = filteredInputData[i][inputDataTitleList.indexOf("班级")];
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
            outputSheet = createOutputSheet(outputBook, outputSheetName);
            outputSheetNameList.push(outputSheetName);
            outputRowNum = 2;
        }

        targetStartCell = outputSheet.Cells.Item(outputRowNum, 1);
        targetEndCell = outputSheet.Cells.Item(outputRowNum, inputDataTitleList.length);

        // 向输出工作表写入此行数据
        outputSheet.Range(targetStartCell, targetEndCell).Value2 = filteredInputData[i];
    }

    // 逐个处理输出工作表
    for (let outputSheetNum = outputBook.Worksheets.Count; outputSheetNum >= 1; outputSheetNum--) {

        outputSheet = outputBook.Worksheets.Item(outputSheetNum);

        if (outputSheetNameList.includes(outputSheet.Name)) {
            // 计算输出工作表，新增列并填入公式
            calculateOutputSheet(outputSheet);
            // 整理输出工作表，删除设置的行列，设置表格标题、表格列宽、单元格格式
            organizeOutputSheet(outputSheet);
        }
        else {
            // 删除名字不在数组中的工作表，通常为新建工作簿时的sheet1。因为其中没有数据，不用关闭警告
            outputSheet.Delete();
        }
    }

    return outputBook;
}

/**
 * 计算输出工作表
 * @param {Et.Worksheet} outputSheet 输出工作表
 */
function calculateOutputSheet(outputSheet) {

    // 获取此工作表的末行数据行的号码（表格边界）
    let outputRowNumMax = outputSheet.UsedRange.Rows.Count;

    // 对应位置新增列
    let targetColumnNumList = [15, 15, 11, 10, 9, 6, 6];
    for (const targetColumnNum of targetColumnNumList) {
        outputSheet.Columns.Item(targetColumnNum).Insert(-4161);
    }

    let targetStartCell, targetEndCell;

    // 根据设置的学分标准版本，计算入学年份与标准版本，自动填充到末行
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

    // 将位置和公式存到map
    let targetColumnNumMap = new Map();
    targetColumnNumMap.set(11, `=MIN(J2,2)`); // 思想成长有效学分
    targetColumnNumMap.set(13, `=MIN(L2,2)`); // 创新创业有效学分
    targetColumnNumMap.set(15, `=MIN(N2,1)`); // 志愿公益有效学分
    targetColumnNumMap.set(20, `=SUM(P2:S2)`); // 总选修学分
    targetColumnNumMap.set(21, `=IFS(G2="2023版",MIN(P2+SUM(Q2:S2),P2+2,3),G2="2019版",MIN(T2,3))`); // 选修学分有效学分
    if (document.getElementById("recalculate-total-credits").checked) {
        targetColumnNumMap.set(22, `=SUM(J2,L2,N2,T2)`); // 总学分
    }
    targetColumnNumMap.set(23, `=SUM(K2,M2,O2,U2)`); // 总有效学分
    if (document.getElementById("calculate-annual-average-valid-credits").checked) {
        targetColumnNumMap.set(24, `=YEAR(I2)-D2`); // 学年数
        targetColumnNumMap.set(25, `=ROUNDDOWN(W2/X2,2)`); // 年均有效学分
        targetColumnNumMap.set(26, `=IF(Y2>=2,"OK","NG")`); // 参评资格
    }

    // 向对应位置填入公式，自动填充到末行
    for (const [targetColumnNum, formulaStr] of targetColumnNumMap) {

        targetStartCell = outputSheet.Cells.Item(2, targetColumnNum);
        targetEndCell = outputSheet.Cells.Item(outputRowNumMax, targetColumnNum);
        targetStartCell.Formula = formulaStr;
        targetStartCell.AutoFill(outputSheet.Range(targetStartCell, targetEndCell), 1);
    }
}

/**
 * 整理输出工作表
 * @param {Et.Worksheet} outputSheet 输出工作表
 */
function organizeOutputSheet(outputSheet) {

    if (document.getElementById("delete-not-good-result-rows").checked) {

        let outputRowNumMax = outputSheet.UsedRange.Rows.Count;

        outputSheet.Range(`A2:Z${outputRowNumMax}`).Sort(outputSheet.Range("Z2"), 2);

        let targetStartCell = outputSheet.Range("Z:Z").Find("NG", undefined, -4163);

        if (targetStartCell !== null) {
            if (targetStartCell.Row === 2) {
                etApp.DisplayAlerts = false;
                outputSheet.Delete();
                etApp.DisplayAlerts = true;
                return;
            }
            else {
                outputSheet.Range(`${targetStartCell.Row}:${outputRowNumMax}`).Delete();
            }
        }
    }

    if (document.getElementById("convert-output-data-into-values").checked) {
        outputSheet.UsedRange.Copy();
        outputSheet.Range("A1").PasteSpecial(-4163, -4142, false, false);
    }

    let annualAverageValidFlag = document.getElementById("calculate-annual-average-valid-credits").checked;

    // 配置第1行标题
    outputSheet.Range("H1:V1").Clear();
    outputSheet.Range("F1:W1").Value2 = ["入学年份", "标准版本", "统计开始时间", "统计结束时间", "必修学分", "", "", "", "", "", "选修学分", "", "", "", "", "", "总学分", "总有效学分"];
    if (annualAverageValidFlag) {
        outputSheet.Range("X1:Z1").Value2 = ["学年数", "年均有效学分", "参评资格"];
    }

    // 配置第2行标题
    outputSheet.Rows.Item(2).Insert(-4121);
    outputSheet.Range("J2:U2").Value2 = ["思想成长", "有效学分", "创新创业", "有效学分", "志愿公益", "有效学分", "实践实习", "文体活动", "工作履历", "技能特长", "总选修学分", "有效学分"];

    // 合并标题栏对应单元格
    let targetRangeAddrList = ["A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:F2", "G1:G2", "H1:H2", "I1:I2", "J1:O1", "P1:U1", "V1:V2", "W1:W2"];
    if (annualAverageValidFlag) {
        targetRangeAddrList.push("X1:X2", "Y1:Y2", "Z1:Z2");
    }
    for (const targetRangeAddr of targetRangeAddrList) {
        outputSheet.Range(targetRangeAddr).Merge();
    }

    // 为对应列设置列宽
    let columnWidthList = [10, 8, 12, 5, 9, 8, 8, 12, 12, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 8, 8, 10];
    if (annualAverageValidFlag) {
        columnWidthList.push(6, 12, 8);
    }
    for (let i = columnWidthList.length; i >= 1; i--) {
        outputSheet.Columns.Item(i).ColumnWidth = columnWidthList[i - 1];
    }
    // 设置行高
    outputSheet.UsedRange.RowHeight = 20;

    // 设置对齐方式，除姓名、学院、班级列的数据部分左对齐，其他单元格水平居中对齐
    outputSheet.UsedRange.HorizontalAlignment = -4108;
    outputSheet.Range("B:C").HorizontalAlignment = -4131;
    outputSheet.Range("E:E").HorizontalAlignment = -4131;
    outputSheet.Range("1:2").HorizontalAlignment = -4108;

    // 设置学分数据的单元格格式，统一2位小数
    outputSheet.Range("J:W").NumberFormatLocal = "0.00";
    if (annualAverageValidFlag) {
        outputSheet.Range("Y:Y").NumberFormatLocal = "0.00";
    }

    // 按照设置，删除对应列数据
    if (document.getElementById("delete-statistical-time-columns").checked) {
        outputSheet.Range("H:I").Delete();
    }
    if (document.getElementById("delete-enrollment-year-column").checked) {
        outputSheet.Range("F:F").Delete();
    }
    if (document.getElementById("delete-class-column").checked) {
        outputSheet.Range("E:E").Delete();
    }
    if (document.getElementById("delete-grade-column").checked) {
        outputSheet.Range("D:D").Delete();
    }
    if (document.getElementById("delete-college-column").checked) {
        outputSheet.Range("C:C").Delete();
    }

}
/**
 * @author LiBaode <libaode2013@163.com>
 * @version 20260116A
 */

var etApp = window.Application;
const collegeNameList = ["机械工程与自动化学院", "汽车与交通工程学院", "材料科学与工程学院", "化学与环境工程学院", "电气工程学院", "电子与信息工程学院", "经济管理学院", "文化传媒与艺术设计学院", "土木建筑工程学院", "外国语学院", "理学院", "交叉科学学院"];
const dataTitleList = ["学号", "姓名", "学院", "年级", "班级", "参与开始时间", "参与结束时间", "思想成长类-学分", "创新创业-学分", "志愿公益类-学分", "实践实习类-学分", "文体活动-学分", "工作履历-学分", "技能特长-学分", "学分总和"];

// 初始化用户界面：选择输入工作簿
try {
    const inputBookListNode = document.getElementById("select-input-books");
    let inputBookName, inputBookNode, checkboxNode, labelNode;
    let inputBookNumMax = etApp.Workbooks.Count;
    for (let inputBookNum = 1; inputBookNum <= inputBookNumMax; inputBookNum++) {
        inputBookName = etApp.Workbooks.Item(inputBookNum).Name;
        // create HTML elements
        inputBookNode = document.createElement("div");
        checkboxNode = document.createElement("input");
        labelNode = document.createElement("label");
        // set attributes to elements
        checkboxNode.setAttribute("type", "checkbox");
        checkboxNode.setAttribute("id", `input-book-${inputBookNum}`);
        checkboxNode.setAttribute("name", "input-book");
        checkboxNode.setAttribute("value", inputBookName);
        labelNode.setAttribute("for", `input-book-${inputBookNum}`);
        labelNode.textContent = inputBookName;
        // 如果检查不通过，则禁用复选框，标签显示为灰色
        if (checkInputBook(inputBookName) === false) {
            checkboxNode.setAttribute("disabled", "");
            labelNode.setAttribute("style", "color: gray;");
        }
        // write elements to document
        inputBookNode.appendChild(checkboxNode);
        inputBookNode.appendChild(labelNode);
        inputBookListNode.appendChild(inputBookNode);
    }
} catch (error) {
    alert(`初始化用户界面（选择输入工作簿）时出错，错误信息：\n\n${error.stack}\n\n可联系开发者获取更多支持。libaode2013@163.com`);
    window.close();
}

/**
 * 检查输入工作簿
 * @param {string} inputBookName 输入工作簿名称
 * @returns 
 */
function checkInputBook(inputBookName) {
    let inputBook = etApp.Workbooks.Item(inputBookName);
    // 检查输入工作簿中的工作表数量，应只有1个
    if (inputBook.Worksheets.Count !== 1) {
        return false;
    }
    let inputSheet = inputBook.Worksheets.Item(1);
    // 检查输入工作表的名称，应为学分汇总
    if (inputSheet.Name !== "学分汇总") {
        return false;
    }
    // 检查输入工作表的表格标题，应有所有数据标题数组中的元素
    for (const dataTitle of dataTitleList) {
        if (inputSheet.Rows.Item(1).Find(dataTitle) === null) {
            return false;
        }
    }
    // 检查输入工作表的数据行数，除标题行应至少有1行数据
    if (inputSheet.UsedRange.Rows.Count < 2) {
        return false;
    }
    // 全部检查通过，返回true
    return true;
}

// 初始化用户界面：筛选输入数据
try {
    // HTML elements
    const useCollegeFilter = document.getElementById("use-college-filter");
    const collegeFilter = document.getElementById("college-filter");
    const useGradeFilter = document.getElementById("use-grade-filter");
    const gradeFilter = document.getElementById("grade-filter");
    const useClassFilter = document.getElementById("use-class-filter");
    const classFilter = document.getElementById("class-filter");

    // 初始化学院筛选器，添加学院名
    let collegeNameNode;
    for (const collegeName of collegeNameList) {
        collegeNameNode = document.createElement("option");
        collegeNameNode.setAttribute("value", collegeName);
        collegeNameNode.textContent = collegeName;
        collegeFilter.appendChild(collegeNameNode);
    }

    // 数据筛选器，选中则解锁选项，取消选中则锁定选项
    useCollegeFilter.addEventListener("change", () => {
        if (useCollegeFilter.checked) {
            collegeFilter.removeAttribute("disabled");
        }
        else {
            collegeFilter.setAttribute("disabled", "");
        }
    });
    useGradeFilter.addEventListener("change", () => {
        if (useGradeFilter.checked) {
            gradeFilter.removeAttribute("disabled");
        }
        else {
            gradeFilter.setAttribute("disabled", "");
        }
    });
    useClassFilter.addEventListener("change", () => {
        if (useClassFilter.checked) {
            classFilter.removeAttribute("disabled");
        }
        else {
            classFilter.setAttribute("disabled", "");
        }
    });
} catch (error) {
    alert(`初始化用户界面（筛选输入数据）时出错，错误信息：\n\n${error.stack}\n\n可联系开发者获取更多支持。libaode2013@163.com`);
    window.close();
}

// 初始化用户界面：其他 HTML 元素
try {

    // 设置输出工作表，显示高级选项按钮，点击后显示隐藏的选项
    document.getElementById("show-advanced-options-for-output-sheets").addEventListener("click", () => {
        document.getElementById("show-advanced-options-for-output-sheets").setAttribute("hidden", "");
        document.getElementById("advanced-options-for-output-sheets").removeAttribute("hidden");
    });

    // 解锁开始按钮
    document.getElementById("start-button").removeAttribute("disabled");
} catch (error) {
    alert(`初始化用户界面（其他 HTML 元素）时出错，错误信息：\n\n${error.stack}\n\n可联系开发者获取更多支持。libaode2013@163.com`);
    window.close();
}

function main() {
    try {
        // 将选择的输入工作簿名称存入数组
        let inputBookNameList = [];
        for (const checkbox of document.getElementsByName("input-book")) {
            if (checkbox.checked) {
                inputBookNameList.push(checkbox.value);
            }
        }
        // 检查用户配置，应至少选择1个输入工作簿，数组长度不能为0
        if (inputBookNameList.length === 0) {
            alert("未选择输入工作簿。");
            return;
        }
        // 检查每个输入工作簿中，参与开始时间、参与结束时间是否唯一，并存储到数组
        let inputSheet, inputRowNumMax, timeColumnStr, timeRangeStr, startTimeList = [], endTimeList = [];
        for (const inputBookName of inputBookNameList) {
            // 重定向输入工作表
            inputSheet = etApp.Workbooks.Item(inputBookName).Worksheets.Item(1);
            // 获取表格边界
            inputRowNumMax = inputSheet.UsedRange.Rows.Count;

            // 检查在此输入工作表中，参与开始时间是否唯一
            timeColumnStr = inputSheet.Rows.Item(1).Find("参与开始时间").Address(false, false).slice(0, -1);
            timeRangeStr = `'[${inputBookName}]${inputSheet.Name}'!${timeColumnStr}2:${timeColumnStr}${inputRowNumMax}`;
            if (etApp.Evaluate(`COUNTA(UNIQUE(${timeRangeStr}))`) !== 1) {
                alert(`在【${inputBookName}】中的参与开始时间不唯一`);
                return;
            }
            // 将此工作表中的参与开始时间存储到数组
            startTimeList.push(inputSheet.Range(`${timeColumnStr}2`).Value2);

            // 检查在此输入工作表中，参与结束时间是否唯一
            timeColumnStr = inputSheet.Rows.Item(1).Find("参与结束时间").Address(false, false).slice(0, -1);
            timeRangeStr = `'[${inputBookName}]${inputSheet.Name}'!${timeColumnStr}2:${timeColumnStr}${inputRowNumMax}`;
            if (etApp.Evaluate(`COUNTA(UNIQUE(${timeRangeStr}))`) !== 1) {
                alert(`在【${inputBookName}】中的参与结束时间不唯一`);
                return;
            }
            // 将此工作表中的参与结束时间存储到数组
            endTimeList.push(inputSheet.Range(`${timeColumnStr}2`).Value2);
        }

        // 去除2个时间数组中的重复元素
        startTimeList = Array.from(new Set(startTimeList));
        endTimeList = Array.from(new Set(endTimeList));
        // 检查参与开始时间、参与结束时间的数量
        if (startTimeList.length !== 1) {
            alert("已选中的输入工作簿中，参与开始时间不唯一");
            return;
        }
        if (endTimeList.length !== 1) {
            alert("已选中的输入工作簿中，参与结束时间不唯一");
            return;
        }
        // 如果勾选了计算学年平均有效学分，检查结束时间，应落在6-9月
        let endTime = new Date(endTimeList[0]);
        if (document.getElementById("calculate-annual-average-valid-credits").checked && endTime.getMonth() < 5 || endTime.getMonth() > 8) {
            alert(`参与结束时间（${endTimeList[0]}）异常，应落在6-9月`);
            return;
        }

        // 复制所有输入工作表到临时工作簿，存储新工作表名到数组
        let tempBook = etApp.Workbooks.Add();
        let inputSheetNameList = [];

        for (const inputBookName of inputBookNameList) {
            // 重定向输入工作表
            inputSheet = etApp.Workbooks.Item(inputBookName).Worksheets.Item(1);
            // 复制，存储新表名
            inputSheet.Copy(tempBook.Worksheets.Item(1), null);
            inputSheetNameList.push(etApp.ActiveSheet.Name);
        }

        // 整理表格，复制所有输入数据到临时工作表，检查学号是否有重复
        let tempSheet = tempBook.Worksheets.Add();
        tempSheet.Name = "tempSheet";
        tempSheet.Range("A1:O1").Value2 = dataTitleList;

        let inputColumnNumMax, tempRowNum;
        for (const inputSheetName of inputSheetNameList) {
            // 重定向输入工作表
            inputSheet = tempBook.Worksheets.Item(inputSheetName);

            // 获取表格边界
            inputColumnNumMax = inputSheet.UsedRange.Columns.Count;

            // 删除数据标题不在列表中的数据列
            for (let inputColumnNum = inputColumnNumMax; inputColumnNum >= 1; inputColumnNum--) {
                if (dataTitleList.includes(inputSheet.Cells.Item(1, inputColumnNum).Value2) === false) {
                    inputSheet.Columns.Item(inputColumnNum).Delete();
                }
            }

            // 调整数据列位置
            for (let i = dataTitleList.length - 1; i >= 0; i--) {
                inputSheet.Columns.Item(1).Insert(-4161);
                inputSheet.Columns.Item(inputSheet.Rows.Item(1).Find(dataTitleList[i]).Column).Cut(inputSheet.Columns.Item(1));
            }

            // 复制数据到临时工作表
            inputRowNumMax = inputSheet.UsedRange.Rows.Count;
            tempRowNum = tempSheet.UsedRange.Rows.Count + 1;
            inputSheet.Range(`A2:O${inputRowNumMax}`).Copy(tempSheet.Range(`A${tempRowNum}`));
        }

        // 检查是否存在重复的学号
        let tempRowNumMax = tempSheet.UsedRange.Rows.Count;
        let stunumRangeStr = `'[${tempBook.Name}]tempSheet'!A2:A${tempRowNumMax}`;
        if (etApp.Evaluate(`COUNTA(UNIQUE(${stunumRangeStr}))`) !== tempRowNumMax - 1) {
            alert("输入数据异常：存在学号相同的数据行");
            etApp.DisplayAlerts = false;
            tempBook.Close();
            etApp.DisplayAlerts = true;
            return;
        }

        // 检查是否存在满足所有筛选条件的数据行
        let rangeCriteriaStr = [];

        if (document.querySelectorAll(`input[name="data-filter"]:checked`).length !== 0) {

            if (document.getElementById("use-college-filter").checked === true) {
                rangeCriteriaStr.push(`'[${tempBook.Name}]tempSheet'!C2:C${tempRowNumMax}`);
                rangeCriteriaStr.push(`"${document.getElementById("college-filter").value}"`);
            }
            if (document.getElementById("use-grade-filter").checked === true) {
                rangeCriteriaStr.push(`'[${tempBook.Name}]tempSheet'!D2:D${tempRowNumMax}`);
                rangeCriteriaStr.push(`"${document.getElementById("grade-filter").value}"`);
            }
            if (document.getElementById("use-class-filter").checked === true) {
                rangeCriteriaStr.push(`'[${tempBook.Name}]tempSheet'!E2:E${tempRowNumMax}`);
                rangeCriteriaStr.push(`"${document.getElementById("class-filter").value}"`);
            }

            rangeCriteriaStr = String(rangeCriteriaStr);
            if (etApp.Evaluate(`COUNTIFS(${rangeCriteriaStr})`) === 0) {
                alert("输入数据异常：未找到符合筛选条件的数据");
                etApp.DisplayAlerts = false;
                tempBook.Close();
                etApp.DisplayAlerts = true;
                return;
            }
        }

        // 检查结束，锁定用户界面，开始处理
        for (const fieldset of document.getElementsByName("option-group")) {
            fieldset.setAttribute("disabled", "");
        }
        document.getElementById("start-button").setAttribute("disabled", "");


        // 配置输出工作簿，重定向输入工作表
        let outputBook = etApp.Workbooks.Add();
        inputSheet = tempSheet;

        // 读取设置的数据筛选器
        let collegeFilter = "", gradeFilter = "", classFilter = "";
        if (document.getElementById("use-college-filter").checked) {
            collegeFilter = document.getElementById("college-filter").value;
        }
        if (document.getElementById("use-grade-filter").checked) {
            gradeFilter = document.getElementById("grade-filter").value;
        }
        if (document.getElementById("use-class-filter").checked) {
            classFilter = document.getElementById("class-filter").value;
        }

        // 读取设置的分表方式
        let outputSheetType = document.getElementById("output-sheet-type").value;

        // 开始逐行复制数据到输出工作表
        let outputSheet, outputSheetName, outputSheetNameList = [], outputRowNum;

        // 将整个输入工作表的数据，都存储到数组
        // 减少与JSAPI的交互次数，提高处理性能
        inputRowNumMax = inputSheet.UsedRange.Rows.Count;
        let inputData = inputSheet.Range(`A2:O${inputRowNumMax}`).Value2;
        for (let i = 0; i < inputRowNumMax - 1; i++) {

            // 检查此行数据是否符合筛选条件，不符合则跳过此行
            if (collegeFilter !== "" && inputData[i][2] !== collegeFilter) {
                continue;
            }
            if (gradeFilter !== "" && inputData[i][3] !== gradeFilter) {
                continue;
            }
            if (classFilter !== "" && inputData[i][4] !== classFilter) {
                continue;
            }

            // 获取输出工作表名
            switch (outputSheetType) {
                case "sheet-type-college":
                    outputSheetName = inputData[i][2];
                    break;
                case "sheet-type-grade":
                    outputSheetName = inputData[i][3];
                    break;
                case "sheet-type-class":
                    outputSheetName = inputData[i][4];
                    break;
                case "sheet-type-all-in-one":
                    outputSheetName = "output-sheet";
                    break;
                default:
                    alert("unknown output sheet type");
            }

            // 重定向输出工作表，若不存在则初始化输出工作表。获取输出位置
            if (outputSheetNameList.includes(outputSheetName)) {
                outputSheet = outputBook.Worksheets.Item(outputSheetName);
                outputRowNum = outputSheet.UsedRange.Rows.Count + 1;
            }
            else {
                outputSheet = outputBook.Worksheets.Add();
                outputSheet.Name = outputSheetName;
                outputSheet.Range("A1:O1").Value2 = dataTitleList;
                outputSheetNameList.push(outputSheetName);
                outputRowNum = 2;
            }

            // 强制转换学分数据为文本格式，否则公式计算会出错
            // inputData[i][0] = `'${inputData[i][0]}`;


            // 向输出工作表写入这行数据
            outputSheet.Range(`A${outputRowNum}:O${outputRowNum}`).Value2 = inputData[i];
        }

        // 处理输出工作表
        let outputRowNumMax;
        for (const outputSheetName of outputSheetNameList) {

            // 重定向输出工作表
            outputSheet = outputBook.Worksheets.Item(outputSheetName);
            // 获取表格边界
            outputRowNumMax = outputSheet.UsedRange.Rows.Count;

            // 计算选修学分总和
            outputSheet.Range("O:O").Insert(-4161);
            outputSheet.Range("O2").Formula = "=SUM(K2,L2,M2,N2)";
            outputSheet.Range("O2").AutoFill(outputSheet.Range(`O2:O${outputRowNumMax}`), 1);

            // 如果勾选了选项，重新计算学分总和
            if (document.getElementById("recalculate-total-credits").checked) {
                outputSheet.Range("P2").Formula = "=SUM(H2,I2,J2,O2)";
                outputSheet.Range("P2").AutoFill(outputSheet.Range(`P2:P${outputRowNumMax}`), 1);
            }

            // 插入列，用于存放有效学分数据
            outputSheet.Range("P:P").Insert(-4161); // 选修学分-总选修学分-有效学分
            outputSheet.Range("K:K").Insert(-4161); // 必修学分-志愿公益-有效学分
            outputSheet.Range("J:J").Insert(-4161); // 必修学分-创新创业-有效学分
            outputSheet.Range("I:I").Insert(-4161); // 必修学分-思想成长-有效学分

            // 向新插入的列写入公式，计算必修学分各类的有效学分
            outputSheet.Range("I2").Formula = "=MIN(H2,2)";
            outputSheet.Range("I2").AutoFill(outputSheet.Range(`I2:I${outputRowNumMax}`), 1);
            outputSheet.Range("K2").Formula = "=MIN(J2,2)";
            outputSheet.Range("K2").AutoFill(outputSheet.Range(`K2:K${outputRowNumMax}`), 1);
            outputSheet.Range("M2").Formula = "=MIN(L2,1)";
            outputSheet.Range("M2").AutoFill(outputSheet.Range(`M2:M${outputRowNumMax}`), 1);

            // 根据设定的学分有效性标准，写入不同公式，计算选修学分的有效学分
            switch (document.getElementById("credit-standred-version").value) {
                case "standred-version-2019":
                    // 写入计算公式
                    outputSheet.Range("S2").Formula = "=MIN(R2,3)";
                    outputSheet.Range("S2").AutoFill(outputSheet.Range(`S2:S${outputRowNumMax}`), 1);
                    // 写入有效性标准
                    outputSheet.Range("V2").Value2 = "2019版";
                    outputSheet.Range("V2").AutoFill(outputSheet.Range(`V2:V${outputRowNumMax}`), 1);
                    break;
                case "standred-version-2023":
                    // 写入计算公式
                    outputSheet.Range("S2").Formula = "=IF(N2>=1,N2+MIN(SUM(O2,P2,Q2),3-N2),N2+MIN(SUM(O2,P2,Q2),2))";
                    outputSheet.Range("S2").AutoFill(outputSheet.Range(`S2:S${outputRowNumMax}`), 1);
                    // 写入有效性标准
                    outputSheet.Range("V2").Value2 = "2023版";
                    outputSheet.Range("V2").AutoFill(outputSheet.Range(`V2:V${outputRowNumMax}`), 1);
                    break;
                case "standred-version-auto":
                    for (let outputRowNum = 2; outputRowNum <= outputRowNumMax; outputRowNum++) {
                        if (`20${String(outputSheet.Range(`A${outputRowNum}`).Value2).substring(0, 2)}` >= 2023) {
                            outputSheet.Range(`S${outputRowNum}`).Formula = `=IF(N${outputRowNum}>=1,N${outputRowNum}+MIN(SUM(O${outputRowNum},P${outputRowNum},Q${outputRowNum}),3-N${outputRowNum}),N${outputRowNum}+MIN(SUM(O${outputRowNum},P${outputRowNum},Q${outputRowNum}),2))`;
                            outputSheet.Range(`V${outputRowNum}`).Value2 = "2023版";
                        }
                        else {
                            outputSheet.Range(`S${outputRowNum}`).Formula = `=MIN(R${outputRowNum},3)`;
                            outputSheet.Range(`V${outputRowNum}`).Value2 = "2019版";
                        }
                    }
                    break;
                default:
                    alert("unknown credit standerd version");
            }
            // 写入公式，计算总有效学分
            outputSheet.Range("U2").Formula = "=SUM(I2,K2,M2,S2)";
            outputSheet.Range("U2").AutoFill(outputSheet.Range(`U2:U${outputRowNumMax}`), 1);

            // 计算学年平均有效学分
            if (document.getElementById("calculate-annual-average-valid-credits").checked) {
                // 写入公式，计算学年数
                outputSheet.Range("W2").Formula = "=YEAR(G2)-D2";
                outputSheet.Range("W2").AutoFill(outputSheet.Range(`W2:W${outputRowNumMax}`), 1);
                // 计算学年有效学分
                outputSheet.Range("X2").Formula = "=ROUNDDOWN(U2/W2,2)";
                outputSheet.Range("X2").AutoFill(outputSheet.Range(`X2:X${outputRowNumMax}`), 1);
                // 判断参评资格
                outputSheet.Range("Y2").Formula = "=IF(X2>=2,\"OK\",\"NG\")";
                outputSheet.Range("Y2").AutoFill(outputSheet.Range(`Y2:Y${outputRowNumMax}`), 1);
            }

            // 配置输出工作表，输出数据存储为值
            if (document.getElementById("store-output-data-as-values").checked) {
                outputSheet.UsedRange.Copy();
                outputSheet.Range("A1").PasteSpecial(-4163, -4142, false, false);
            }


            // 配置表格标题，第一行
            outputSheet.Range("F1").Value2 = "统计开始时间";
            outputSheet.Range("G1").Value2 = "统计结束时间";
            outputSheet.Range("H1:Q1").Clear();
            outputSheet.Range("H1").Value2 = "必修学分";
            outputSheet.Range("N1").Value2 = "选修学分";
            outputSheet.Range("T1").Value2 = "总学分";
            outputSheet.Range("U1").Value2 = "总有效学分";
            outputSheet.Range("V1").Value2 = "学分标准版本";
            if (document.getElementById("calculate-annual-average-valid-credits").checked) {
                outputSheet.Range("W1").Value2 = "学年数";
                outputSheet.Range("X1").Value2 = "年均有效学分";
                outputSheet.Range("Y1").Value2 = "参评资格";
            }
            // 配置表格标题，第二行
            outputSheet.Rows.Item(2).Insert(-4121);
            outputSheet.Range("H2:S2").Value2 = ["思想成长", "有效学分", "创新创业", "有效学分", "志愿公益", "有效学分", "实践实习", "文体活动", "工作履历", "技能特长", "选修学分", "有效学分"];
            // 合并单元格
            outputSheet.Range("A1:A2").Merge();
            outputSheet.Range("B1:B2").Merge();
            outputSheet.Range("C1:C2").Merge();
            outputSheet.Range("D1:D2").Merge();
            outputSheet.Range("E1:E2").Merge();
            outputSheet.Range("F1:F2").Merge();
            outputSheet.Range("G1:G2").Merge();
            outputSheet.Range("H1:M1").Merge();
            outputSheet.Range("N1:S1").Merge();
            outputSheet.Range("T1:T2").Merge();
            outputSheet.Range("U1:U2").Merge();
            outputSheet.Range("V1:V2").Merge();
            if (document.getElementById("calculate-annual-average-valid-credits").checked) {
                outputSheet.Range("W1:W2").Merge();
                outputSheet.Range("X1:X2").Merge();
                outputSheet.Range("Y1:Y2").Merge();
            }

            // 配置列宽
            outputSheet.Range("A:A").ColumnWidth = 10;
            outputSheet.Range("B:B").ColumnWidth = 8;
            outputSheet.Range("C:C").ColumnWidth = 10;
            outputSheet.Range("D:D").ColumnWidth = 6;
            outputSheet.Range("E:E").ColumnWidth = 10;
            outputSheet.Range("F:G").ColumnWidth = 12;
            outputSheet.Range("H:U").ColumnWidth = 8;
            outputSheet.Range("U:U").ColumnWidth = 10;
            outputSheet.Range("V:V").ColumnWidth = 12;
            if (document.getElementById("calculate-annual-average-valid-credits").checked) {
                outputSheet.Range("W:W").ColumnWidth = 6;
                outputSheet.Range("X:X").ColumnWidth = 12;
                outputSheet.Range("Y:Y").ColumnWidth = 8;
            }
            // 配置行高
            outputSheet.UsedRange.RowHeight = 20;

            // 配置对齐方式
            outputSheet.UsedRange.HorizontalAlignment = -4108; // 水平居中对齐
            // 姓名、学院、班级列数据长度不固定，配置为水平左对齐
            outputSheet.Range("B:C").HorizontalAlignment = -4131; // 水平左对齐
            outputSheet.Range("E:E").HorizontalAlignment = -4131;
            outputSheet.Range("1:2").HorizontalAlignment = -4108; // 标题恢复水平居中对齐


            // 配置数据格式
            outputSheet.Range("H:U").NumberFormatLocal = "0.00";
            outputSheet.Range("X:X").NumberFormatLocal = "0.00";

            // 表格添加边框
            // outputSheet.UsedRange.Borders.LineStyle = 1;


            // 配置输出工作表，删除设置的删除列
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

        // 删除多余的工作簿与工作表
        etApp.DisplayAlerts = false;
        // 关闭临时工作簿
        tempBook.Close();
        // 删除输出工作簿中，多余的工作表
        for (let outputSheetNum = outputBook.Worksheets.Count; outputSheetNum >= 1; outputSheetNum--) {
            outputSheetName = outputBook.Worksheets.Item(outputSheetNum).Name;
            if (outputSheetNameList.includes(outputSheetName) === false) {
                outputBook.Worksheets.Item(outputSheetNum).Delete();
            }
        }
        etApp.DisplayAlerts = true;

        alert("处理完成");
        window.close();

    } catch (error) {
        alert(`执行主过程函数时出错，错误信息：\n\n${error.stack}\n\n可联系开发者获取更多支持。libaode2013@163.com`);
        window.close();
    }
}

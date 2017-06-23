/**
 * Created by xianbr on 2017/5/24.
 */

// 导入通用列表
$(function () {
    // 选择导入文件
    $("#importNormalList").click(function () {
        $("#importNormalList2").click();
    });

    $("#importNormalList2").change(function (e) {
        var importValue = $("#importNormalList2").val();
        var fileName = importValue.split("\\")[2];

        // 列表的相对路径
        var url = "./Excels/" + fileName;       
        $("#pathArea").val(fileName);

        var oReq = new XMLHttpRequest();
        oReq.open("GET", url, true);
        oReq.responseType = "arraybuffer";
        oReq.onload = function (e) {
            var arraybuffer = oReq.response;
            var data = new Uint8Array(arraybuffer);
            var arr = [];
            for (var i = 0; i != data.length; ++i)
                arr[i] = String.fromCharCode(data[i]);
            var bstr = arr.join("");
            // 获取到的Excel对象
            var workbook = XLSX.read(bstr, {type: "binary"});

            // 每次先清空显示
            if ($("#selectSheets").children().length > 0) {
                $("#selectSheets").empty();
            }
            // 获得所有的sheet,并添加至下拉框
            $.each(workbook.SheetNames, function (index, value) {
                var $optionSheet = $("<option>" + value + "</option>");
                $("#selectSheets").append($optionSheet);
            });

            alert("表格导入成功！");


/************************************************************************************************************************************
            通用选项的默认选项的事件
************************************************************************************************************************************/

            // 解除每个按钮事件的引用
            $("#checkNewCnTxt").unbind("click");
            $("#checkNewEnTxt").unbind("click");
            $("#convertNewTxt").unbind("click");
            $("#exportNewList").unbind("click");
            $("#exportGroupList").unbind("click");

            // 获得要处理的sheet名字
            var selectSheetNameGobal = $("#selectSheets").val();
            // 标题所在的行， value值实际显示值“-1”
            var selectTitleRowGobal = $("#titleRow").val();
            // 有效数据所在的行， value值实际显示值“-1”
            var selectDataRowGobal = $("#dataRow").val();
            // 要处理的sheet对象
            var sheet = workbook.Sheets[selectSheetNameGobal];
            // 将要处理的sheet转换为数组json对象：[{ }, { }, { }]
            var sheetArrayJson = XLSX.utils.sheet_to_json(sheet, {header: "A"});

            // 计算sheet的总列数： lengthCol
            var charCol = "";
            var lengthCol = 0;
            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                charCol = index;
            });
            if (charCol.length != 1) {
                // 列数大于"Z"，"AA-ZZ"
                lengthCol = (charCol[0].charCodeAt() - 64) * 26 + (charCol[1].charCodeAt() - 64);
            } else {
                // 列数少于等于"Z"
                lengthCol = charCol.charCodeAt() - 64;
            }

            // 计算sheet的总行数：lengthRow
            var lengthRow = sheetArrayJson.length;

            // 将表格的标题添加到基准列、参数列的下拉框
            $(".standerCol").empty();
            // 非基准列下拉框，先将原有的全部非“空”选项删除
            $(".appendTitles option:first-child").nextAll().remove();

            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                // 参数列下拉框
                var $appendTitles = $("<option></option>");
                // 基准列下拉框
                var $appendTitles2 = $("<option></option>");
                $appendTitles.html(value).val(index);
                $appendTitles2.html(value).val(index);
                // DOM追加子元素
                $(".appendTitles").append($appendTitles);
                $(".standerCol").append($appendTitles2);
            });


            // 通用列表的中、英文全名与缩写的列
            var enFullNameCol = "";
            var enShortNameCol = "";
            var cnFullNameCol = "";
            var cnShortNameCol = "";

            //  英文全名列
            $("#checkEnFullName").children().remove();
            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                // 只匹配关键字："英文全名"和"Full Name"
                if (value == "英文全名" || value == "Full Name") {
                    var enFullName = sheetArrayJson[selectTitleRowGobal][index];
                    var $enFullNameOption = $("<option>" + enFullName + "</option>");
                    $("#checkEnFullName").append($enFullNameOption);
                    enFullNameCol = index;
                    // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    return false;
                }else {

                }
            });

            //  英文缩写列
            // 添加前先清空
            $("#checkEnShortName").children().remove();
            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                if (value == "英文缩写" || value == "Short Name") {
                    var enShortName = sheetArrayJson[selectTitleRowGobal][index];       //
                    var $enShortNameOption = $("<option>" + enShortName + "</option>");
                    $("#checkEnShortName").append($enShortNameOption);
                    enShortNameCol = index;
                    return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                }
            });

            //  中文全名列
            // 添加前先清空
            $("#checkCnFullName").children().remove();
            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                if (value == "中文全名" || value == "Full Name Chinese") {
                    var cnFullName = sheetArrayJson[selectTitleRowGobal][index];       //
                    var $cnFullNameOption = $("<option>" + cnFullName + "</option>");
                    $("#checkCnFullName").append($cnFullNameOption);
                    cnFullNameCol = index;
                    return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                }
            });

            //  中文缩写列
            // 添加前先清空
            $("#checkCnShortName").children().remove();
            $.each(sheetArrayJson[selectTitleRowGobal], function (index, value) {
                if (value == "中文缩写" || value == "Short Name Chinese") {
                    var cnShortName = sheetArrayJson[selectTitleRowGobal][index];       //
                    var $cnShortNameOption = $("<option>" + cnShortName + "</option>");
                    $("#checkCnShortName").append($cnShortNameOption);
                    cnShortNameCol = index;
                    return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                }
            });


            // 基础检查
//                    $("#basicCheck").click(function () {
//                        // 清空显示区
//                        $("#sheetjs").children().remove();
//
//                        var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
//                        var errAddressArray = ["异常表格位置：\r\n"];    // 保存名称的数组，用于导出至.
//                        var enMinLength = $("#enMinLength").val();
//                        var enMaxLength = $("#enMaxLength").val();
//                        var cnMinLength = $("#cnMinLength").val();
//                        var cnMaxLength = $("#cnMaxLength").val();
//                        for (var i = 2; i < lengthRow; i++) {
//                            if (sheetArrayJson[i][valueCol] == undefined) {
//                                continue;
//                            } else {
//                                // 英文
//                                var pattern = /[^\x00-\xff]/;        // 全角字符范围
//                                if (pattern.test(sheetArrayJson[i]["D"])) {      // 英文缩写是否有全角字符
//                                    errAddressArray.push("D" + (i + 1) + "\t");
//                                } else if (sheetArrayJson[i]["D"] != undefined) {    // 空 缩写
//                                    if (sheetArrayJson[i]["D"].length < enMinLength || sheetArrayJson[i]["D"].length > enMaxLength) {    // 英文缩写长度是否在范围内
//                                        errAddressArray.push("D" + (i + 1) + "\t");
//                                    }
//                                } else if (sheetArrayJson[i]["D"] == undefined) {
//                                    errAddressArray.push("D" + (i + 1) + "\t");
//                                } else if (pattern.test(sheetArrayJson[i]["C"])) {      // 英文全名是否有全角字符
//                                    errAddressArray.push("C" + (i + 1) + "\t");
//                                } else if (sheetArrayJson[i]["C"] == undefined) {       // 空 全名
//                                    errAddressArray.push("C" + (i + 1) + "\t");
//                                }
//
//                                // 中文
//                                else if (sheetArrayJson[i]["F"] != undefined) {
//                                    if (sheetArrayJson[i]["F"].length < cnMinLength || sheetArrayJson[i]["F"].length > cnMaxLength) {    // 英文缩写长度是否在范围内
//                                        errAddressArray.push("F" + (i + 1) + "\t");
//                                    }
//                                } else if (sheetArrayJson[i]["F"] == undefined) {     // 空 缩写
//                                    errAddressArray.push("F" + (i + 1) + "\t");
//                                } else if (sheetArrayJson[i]["E"] == undefined) {     // 空 全名
//                                    errAddressArray.push("E" + (i + 1) + "\t");
//                                }
//                            }
//                        }
//
//                        if (errAddressArray.length <= 1) {
//                            alert("There is no mistake!");
//                        } else {
//                            saveAs(new Blob(errAddressArray, {type: "text/plain;charset=utf-8"}), "错误列表.txt");
//                        }
//                    });


            // 生成英文列表
            var txtArray = [];    // 保存名称的数组，用于导出至.txt
            $("#checkNewEnTxt").click(function () {
                // 清空显示区
                $("#sheetjs").children().remove();

                var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                for (var i = selectDataRowGobal; i < lengthRow; i++) {
                    if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                        continue;
                    } else {
                        var enFullName = sheetArrayJson[i][enFullNameCol] + "*";
                        var enShortName = sheetArrayJson[i][enShortNameCol] + "\r\n";
                        txtArray.push(enFullName);
                        txtArray.push(enShortName);
                        var $p1 = $("<tr>" + "<td>" + enFullName + enShortName + "</td>" + "</tr>");
                        $("#sheetjs").append($p1);
                    }
                }
//                        console.log(txtArray);
                // 导出名字文本
                saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "英文名称列表.txt");
            });


            // 生成中文列表
            $("#checkNewCnTxt").click(function () {
                // 清空显示区
                $("#sheetjs").children().remove();

                var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                var txtArray1 = [];    // 保存名称的数组，用于导出至.txt
                for (var i = selectDataRowGobal; i < lengthRow; i++) {
                    if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                        continue;
                    } else {
                        var cnFullName = sheetArrayJson[i][cnFullNameCol] + "*";
                        var cnShortName = sheetArrayJson[i][cnShortNameCol] + "\r\n";
                        txtArray1.push(cnFullName);
                        txtArray1.push(cnShortName);
                        var $p1 = $("<tr>" + "<td>" + cnFullName + cnShortName + "</td>" + "</tr>");
                        $("#sheetjs").append($p1);
                    }
                }
//                        console.log(txtArray);
                // 导出名字文本
                saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "中文名称列表.txt");
            });


            // 生成转换列表
            $("#convertNewTxt").click(function () {
                // 清空显示区
//                        $("#sheetjs").children().remove();
                var convertParams = $("#form3 .appendTitles option:selected");   // 参数下拉框select选中的option
                var paramPreConfigs = $("#form3 .paramPreConfig option:selected");   // 数据预处理下拉框select选中的option
//                        console.log(convertParams[selectTitleRowGobal]["value"]);    // option的value,即col
//                        console.log(paramPreConfigs[selectTitleRowGobal]["value"]);    // option的value,即col
                var lenCol = convertParams.length;   // 19
                var valueCol = $("#convertStander option:selected").val();   // 基准列选中的值（col）

                var textArray2 = [];
                var titleArray = [];
                for (var i = selectDataRowGobal; i < lengthRow; i++) {
                    var textArray1 = [];     // 参数下拉框选项值
                    var titleArray1 = [];
                    if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                        continue;
                    } else {
                        for (var j = 0; j < lenCol; j++) {

                            if (convertParams[j].value == "空") {
                                continue;
                            } else {
                                if (paramPreConfigs[j].text == "不处理") {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === undefined) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        var preValue1 = sheetArrayJson[i][titleValue1];
                                        textArray1.push(preValue1);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                } else if (paramPreConfigs[j].text == "参数列数据加1") {
                                    var titleValue2 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue2] === undefined || sheetArrayJson[i][titleValue2] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        var preValue2 = Number(sheetArrayJson[i][titleValue2]) + 1;
                                        textArray1.push(preValue2);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                } else if (paramPreConfigs[j].text == "参数列数据减1") {
                                    var titleValue3 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue3] === undefined || sheetArrayJson[i][titleValue3] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        var preValue3 = sheetArrayJson[i][titleValue3] - 1;
                                        textArray1.push(preValue3);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                        }
                        textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                        titleArray.push(titleArray1);
                    }
                }

                var textArray = [];
                var language = "";
                for (var m = 0; m < textArray2[0].length; m++) {
                    var temp = [];
                    var cellValue = "";
                    var titleShow = "\r\n" + "\r\n" + ";" + titleArray[0][m] + "\r\n";
                    textArray.push(titleShow);
                    var $p1 = $("<tr>" + "<td>" + titleShow + "</td>" + "</tr>");
                    $("#sheetjs").append($p1);
                    for (var n = 0; n < textArray2.length; n++) {
                        // 数据预处理配置
                        temp.push(textArray2[n][m]);
                    }

                    var temp1 = chunk(temp, 20);    // 将temp数组分割成以20个数据为一个数组的二维数组temp1

                    for (var k = 0; k < temp1.length; k++) {
                        if ($('input[type="radio"][name="language"]:checked')) {
                            language = $('input[type="radio"][name="language"]:checked').val();   // "db" or "dw"
//                                    console.log("language: " + language);
                            cellValue = "\t" + language + "\t" + temp1[k].join(",") + "\t" + "\t" + ";" + (k * 20 + temp1[k].length) + "\r\n";
                            textArray.push(cellValue);    // 用于导出至txt文本
                            var $p2 = $("<tr>" + "<td>" + cellValue + "</td>" + "</tr>");
                            $("#sheetjs").append($p2);
                        }
                    }
                }
                console.log(textArray);
                // 导出文本
                saveAs(new Blob(textArray, {type: "text/plain;charset=utf-8"}), "汇编" + language + ".asm");
            });


            // 导出新Excel列表
            $("#exportNewList").click(function () {
                // 清空显示区
                $("#sheetjs").children().remove();
                var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                        console.log(convertParams[selectTitleRowGobal]["value"]);    // option的value,即col
                var lenCol = convertParams.length;   //
                var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                var textArray2 = [];
                var titleArray = [];
                for (var i = selectDataRowGobal; i < lengthRow; i++) {
                    var textArray1 = [];     // 参数下拉框选项值
                    var titleArray1 = [];
                    if (sheetArrayJson[i][valueCol] ===undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                        continue;
                    } else {
                        for (var j = 0; j < lenCol; j++) {
                            if (convertParams[j].value == "空") {
                                continue;
                            } else {
                                var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                titleArray1.push(convertParams[j].text);        // 标题名称数组
                                if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                    textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                } else {
                                    textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                }
                            }
                        }
                        textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                        titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                    }
                }
//                        console.log("textArray2: " + textArray2);
//                        console.log("length1: " + textArray1.length);
//                        console.log("length2: " + textArray2.length);
//                        console.log("titlesArray: " + titleArray);


                // 生成table表格
                $("#sheetjs").empty();
                var $trTitle = $("<tr>" + "</tr>");
                for (var m = 0; m < titleArray[0].length; m++) {
                    var $thTitle = $("<th>" + titleArray[0][m] + "</th>");
                    $trTitle.append($thTitle);
                }
                $("#sheetjs").append($trTitle);

                for (var p = 0; p < textArray2.length; p++) {
                    var $trContent = $("<tr>" + "</tr>");
                    for (var q = 0; q < textArray2[0].length; q++) {
                        var tdContent = $("<td>" + textArray2[p][q] + "</td>");
                        $trContent.append(tdContent);
                    }
                    $("#sheetjs").append($trContent);
                }


                // table_to_book
                var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                var wb = XLSX.utils.table_to_book(tbl);

                /* bookType can be any supported output type */
                var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                var wbout = XLSX.write(wb, wopts);

                function s2ab(s) {
                    var buf = new ArrayBuffer(s.length);
                    var view = new Uint8Array(buf);
                    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                    return buf;
                }

                /* the saveAs call downloads a file on the local machine */
                saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), selectSheetNameGobal + ".xlsx");
            });

//                    $("#selectSheets option:eq(1)").attr("selected", true);
//                    $("#selectSheets option:eq(0)").attr("selected", true);


            // sheetName的change事件
            $("#selectSheets").change(function () {

                $("#checkNewCnTxt").unbind("click");
                $("#checkNewEnTxt").unbind("click");
                $("#convertNewTxt").unbind("click");
                $("#exportNewList").unbind("click");
                $("#exportGroupList").unbind("click");

                var selectSheetName = $("#selectSheets").val();    // 要处理的sheet名字
                var selectTitleRow = $("#titleRow").val();    // 标题所在的行， value值实际显示值“-1”
                var selectDataRow = $("#dataRow").val();     // 有效数据所在的行， value值实际显示值“-1”


                // 要处理的sheet
                var sheet = workbook.Sheets[selectSheetName];

                // 将要处理的sheet转换为数组json对象
                var sheetArrayJson = XLSX.utils.sheet_to_json(sheet, {header: "A"});    // [{ }, { }, { }]



                // console.log(sheetArrayJson[selectTitleRow]);


                // 总列数
                var charCol = "";
                var lengthCol = 0;
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    charCol = index;
//                        console.log(XLSX.utils.decode_cell("" + index));
                });
                if (charCol.length != 1) {               // "AA-ZZ"的情况
                    lengthCol = (charCol[0].charCodeAt() - 64) * 26 + (charCol[1].charCodeAt() - 64);
                } else {
                    lengthCol = charCol.charCodeAt() - 64;
                }

                console.log("总列数：" + lengthCol);    // 总列数
                // console.log("selectSheetName: " + selectSheetName);



                // 总行数
                var lengthRow = sheetArrayJson.length;
                console.log("lengthRow: " + lengthRow);      // 总行数


                // 将表格的标题添加到基准Select、参数Select
                // 添加前先清空
                $(".standerCol").empty();    // 基准列
                $(".appendTitles option:first-child").nextAll().remove();

                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    var $appendTitles = $("<option></option>");
                    var $appendTitles2 = $("<option></option>");
                    $appendTitles.html(value).val(index);    // 给参数下拉的option添加value值"A"-"ZZ"
                    $appendTitles2.html(value).val(index);    // 给基准下拉的option添加value值"A"-"ZZ"
                    $(".appendTitles").append($appendTitles);
                    $(".standerCol").append($appendTitles2);
                });


                // 通用列表的中、英文全名与缩写的列
                var enFullNameCol = "";
                var enShortNameCol = "";
                var cnFullNameCol = "";
                var cnShortNameCol = "";

                //  英文全名列
                // 添加前先清空
                $("#checkEnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文全名" || value == "Full Name") {
                        var enFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $enFullNameOption = $("<option>" + enFullName + "</option>");
                        $("#checkEnFullName").append($enFullNameOption);
                        enFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  英文缩写列
                // 添加前先清空
                $("#checkEnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文缩写" || value == "Short Name") {
                        var enShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $enShortNameOption = $("<option>" + enShortName + "</option>");
                        $("#checkEnShortName").append($enShortNameOption);
                        enShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文全名列
                // 添加前先清空
                $("#checkCnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文全名" || value == "Full Name Chinese") {
                        var cnFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnFullNameOption = $("<option>" + cnFullName + "</option>");
                        $("#checkCnFullName").append($cnFullNameOption);
                        cnFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文缩写列
                // 添加前先清空
                $("#checkCnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文缩写" || value == "Short Name Chinese") {
                        var cnShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnShortNameOption = $("<option>" + cnShortName + "</option>");
                        $("#checkCnShortName").append($cnShortNameOption);
                        cnShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });


                // 生成英文列表
                $("#checkNewEnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var enFullName = sheetArrayJson[i][enFullNameCol] + "*";
                            var enShortName = sheetArrayJson[i][enShortNameCol] + "\r\n";
                            txtArray.push(enFullName);
                            txtArray.push(enShortName);
                            var $p1 = $("<tr>" + "<td>" + enFullName + enShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
//                        console.log(txtArray);
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "英文名称列表.txt");
                });


                // 生成中文列表
                $("#checkNewCnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var cnFullName = sheetArrayJson[i][cnFullNameCol] + "*";
                            var cnShortName = sheetArrayJson[i][cnShortNameCol] + "\r\n";
                            txtArray.push(cnFullName);
                            txtArray.push(cnShortName);
                            var $p1 = $("<tr>" + "<td>" + cnFullName + cnShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
//                        console.log(txtArray);
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "中文名称列表.txt");
                });


                // 生成转换列表
                $("#convertNewTxt").click(function () {
                    // 清空显示区
//                        $("#sheetjs").children().remove();
                    var convertParams = $("#form3 .appendTitles option:selected");   // 参数下拉框select选中的option
                    var paramPreConfigs = $("#form3 .paramPreConfig option:selected");   // 数据预处理下拉框select选中的option
//                        console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
//                        console.log(paramPreConfigs[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   // 19
                    var valueCol = $("#convertStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {

                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    if (paramPreConfigs[j].text == "不处理") {
                                        var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue1 = sheetArrayJson[i][titleValue1];
                                            textArray1.push(preValue1);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据加1") {
                                        var titleValue2 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue2] === undefined || sheetArrayJson[i][titleValue2] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue2 = Number(sheetArrayJson[i][titleValue2]) + 1;
                                            textArray1.push(preValue2);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据减1") {
                                        var titleValue3 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue3] === undefined || sheetArrayJson[i][titleValue3] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue3 = sheetArrayJson[i][titleValue3] - 1;
                                            textArray1.push(preValue3);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);
                        }
                    }

                    var textArray = [];
                    var langusge = "";
                    for (var m = 0; m < textArray[2][0].length; m++) {
                        var temp = [];
                        var cellValue = "";
                        var titleShow = "\r\n" + "\r\n" + ";" + titleArray[0][m] + "\r\n";
                        textArray.push(titleShow);
                        var $p1 = $("<tr>" + "<td>" + titleShow + "</td>" + "</tr>");
                        $("#sheetjs").append($p1);
                        for (var n = 0; n < textArray2.length; n++) {
                            // 数据预处理配置
                            temp.push(textArray2[n][m]);
                        }

                        var temp1 = chunk(temp, 20);    // 将temp数组分割成以20个数据为一个数组的二维数组temp1

                        for (var k = 0; k < temp1.length; k++) {
                            if ($('input[type="radio"][name="language"]:checked')) {
                                language = $('input[type="radio"][name="language"]:checked').val();   // "db" or "dw"
//                                    console.log("language: " + language);
                                cellValue = "\t" + language + "\t" + temp1[k].join(",") + "\t" + "\t" + ";" + (k * 20 + temp1[k].length) + "\r\n";
                                textArray.push(cellValue);    // 用于导出至txt文本
                                var $p2 = $("<tr>" + "<td>" + cellValue + "</td>" + "</tr>");
                                $("#sheetjs").append($p2);
                            }
                        }
                    }
                    console.log(textArray);
                    // 导出文本
                    saveAs(new Blob(textArray, {type: "text/plain;charset=utf-8"}), "汇编" + language + ".txt");
                });


                // 导出新Excel列表
                $("#exportNewList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                            console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);
                        }
                    }



                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    for (var m = 0; m < titleArray.length; m++) {
                        var $thTitle = $("<th>" + titleArray[0][m] + "</th>");
                        $trTitle.append($thTitle);
                    }
                    $("#sheetjs").append($trTitle);

                    for (var p = 0; p < textArray2.length; p++) {
                        var $trContent = $("<tr>" + "</tr>");
                        for (var q = 0; q < textArray2[0].length; q++) {
                            var tdContent = $("<td>" + textArray2[p][q] + "</td>");
                            $trContent.append(tdContent);
                        }
                        $("#sheetjs").append($trContent);
                    }


                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), selectSheetName + ".xlsx");
                });

                // 导出分组
                $("#exportGroupList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                    console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < 2; j++) {       // 只获取参数1列与参数2列的值
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefine || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                        }
                    }



                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    $trTitle.append("<th>" + "组序号" + "</th>");
                    $trTitle.append("<th>" + "组的数值范围" + "</th>");
                    $trTitle.append("<th>" + titleArray[0][0] + "</th>");
                    $trTitle.append("<th>" + titleArray[0][1] + "</th>");

                    $("#sheetjs").append($trTitle);


                    for (var m = 0, n = 1, l = 0; m < textArray2.length; m++){
                        var $trContent = $("<tr>" + "</tr>");
                        if ((m+1) < textArray2.length && textArray2[m][0] == textArray2[m+1][0]){
                            l++;  // 相邻两项相同的次数
                            continue;
                        }else{
                            var $tdContent1 = $("<td>" + n + "</td>");
                            n++;
                            var $tdContent2 = $("<td>" + (m+1-l) + "-" + (m+1) + "</td>");
                            l = 0;
                            var $tdContent3 = $("<td>" + textArray2[m][0] + "</td>");
                            var $tdContent4 = $("<td>" + textArray2[m][1] + "</td>");
                            $trContent.append($tdContent1);
                            $trContent.append($tdContent2);
                            $trContent.append($tdContent3);
                            $trContent.append($tdContent4);

                            $("#sheetjs").append($trContent);
                        }
                    }



                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}),  "分组统计.xlsx");
                });

            });
            // titleRow的change事件
            $("#titleRow").change(function () {

                $("#checkNewCnTxt").unbind("click");
                $("#checkNewEnTxt").unbind("click");
                $("#convertNewTxt").unbind("click");
                $("#exportNewList").unbind("click");
                $("#exportGroupList").unbind("click");


                var selectSheetName = $("#selectSheets").val();    // 要处理的sheet名字
                var selectTitleRow = $("#titleRow").val();    // 标题所在的行， value值实际显示值“-1”
                var selectDataRow = $("#dataRow").val();     // 有效数据所在的行， value值实际显示值“-1”


//                        console.log(selectSheetName);
//                        console.log(selectTitleRow);     // 1
//                        console.log(selectDataRow);     // 2

                // 要处理的sheet
                var sheet = workbook.Sheets[selectSheetName];

                // 将要处理的sheet转换为数组json对象
                var sheetArrayJson = XLSX.utils.sheet_to_json(sheet, {header: "A"});    // [{ }, { }, { }]


                // 总列数
                var charCol = "";
                var lengthCol = 0;
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    charCol = index;
//                        console.log(XLSX.utils.decode_cell("" + index));
                });
                if (charCol.length != 1) {               // "AA-ZZ"的情况
                    lengthCol = (charCol[0].charCodeAt() - 64) * 26 + (charCol[1].charCodeAt() - 64);
                } else {
                    lengthCol = charCol.charCodeAt() - 64;
                }

                console.log("总列数：" + lengthCol);    // 总列数


                // 总行数
                var lengthRow = sheetArrayJson.length;
                console.log("lengthRow: " + lengthRow);      // 总行数


                // 将表格的标题添加到基准Select、参数Select
                // 添加前先清空
                $(".standerCol").empty();    // 基准列
                $(".appendTitles option:first-child").nextAll().remove();

                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    var $appendTitles = $("<option></option>");
                    var $appendTitles2 = $("<option></option>");
                    $appendTitles.html(value).val(index);    // 给参数下拉的option添加value值"A"-"ZZ"
                    $appendTitles2.html(value).val(index);    // 给基准下拉的option添加value值"A"-"ZZ"
                    $(".appendTitles").append($appendTitles);
                    $(".standerCol").append($appendTitles2);
                });


                // 通用列表的中、英文全名与缩写的列
                var enFullNameCol = "";
                var enShortNameCol = "";
                var cnFullNameCol = "";
                var cnShortNameCol = "";

                //  英文全名列
                // 添加前先清空
                $("#checkEnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文全名" || value == "Full Name") {
                        var enFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $enFullNameOption = $("<option>" + enFullName + "</option>");
                        $("#checkEnFullName").append($enFullNameOption);
                        enFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  英文缩写列
                // 添加前先清空
                $("#checkEnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文缩写" || value == "Short Name") {
                        var enShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $enShortNameOption = $("<option>" + enShortName + "</option>");
                        $("#checkEnShortName").append($enShortNameOption);
                        enShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文全名列
                // 添加前先清空
                $("#checkCnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文全名" || value == "Full Name Chinese") {
                        var cnFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnFullNameOption = $("<option>" + cnFullName + "</option>");
                        $("#checkCnFullName").append($cnFullNameOption);
                        cnFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文缩写列
                // 添加前先清空
                $("#checkCnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文缩写" || value == "Short Name Chinese") {
                        var cnShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnShortNameOption = $("<option>" + cnShortName + "</option>");
                        $("#checkCnShortName").append($cnShortNameOption);
                        cnShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });


                // 生成英文列表
                $("#checkNewEnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    console.log("valueCol: " + valueCol);
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var enFullName = sheetArrayJson[i][enFullNameCol] + "*";
                            var enShortName = sheetArrayJson[i][enShortNameCol] + "\r\n";
                            txtArray.push(enFullName);
                            txtArray.push(enShortName);
                            var $p1 = $("<tr>" + "<td>" + enFullName + enShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
//                        console.log(txtArray);
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "英文名称列表.txt");
                });


                // 生成中文列表
                $("#checkNewCnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var cnFullName = sheetArrayJson[i][cnFullNameCol] + "*";
                            var cnShortName = sheetArrayJson[i][cnShortNameCol] + "\r\n";
                            txtArray.push(cnFullName);
                            txtArray.push(cnShortName);
                            var $p1 = $("<tr>" + "<td>" + cnFullName + cnShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "中文名称列表.txt");
                });


                // 生成转换列表
                $("#convertNewTxt").click(function () {
                    // 清空显示区
//                        $("#sheetjs").children().remove();
                    var convertParams = $("#form3 .appendTitles option:selected");   // 参数下拉框select选中的option
                    var paramPreConfigs = $("#form3 .paramPreConfig option:selected");   // 数据预处理下拉框select选中的option
//                            console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
//                            console.log(paramPreConfigs[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   // 19
                    var valueCol = $("#convertStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    var textArray1 = [];     // 参数下拉框选项值
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {

                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    if (paramPreConfigs[j].text == "不处理") {
                                        var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue1 = sheetArrayJson[i][titleValue1];
                                            textArray1.push(preValue1);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据加1") {
                                        var titleValue2 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue2] === undefined || sheetArrayJson[i][titleValue2] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue2 = Number(sheetArrayJson[i][titleValue2]) + 1;
                                            textArray1.push(preValue2);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据减1") {
                                        var titleValue3 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue3] === undefined || sheetArrayJson[i][titleValue3] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue3 = sheetArrayJson[i][titleValue3] - 1;
                                            textArray1.push(preValue3);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);
                        }
                    }

                    var textArray = [];
                    var language = "";
                    for (var m = 0; m < textArray1.length; m++) {
                        var temp = [];
                        var cellValue = "";
                        var titleShow = "\r\n" + "\r\n" + ";" + titleArray[0][m] + "\r\n";
                        textArray.push(titleShow);
                        var $p1 = $("<tr>" + "<td>" + titleShow + "</td>" + "</tr>");
                        $("#sheetjs").append($p1);
                        for (var n = 0; n < textArray2.length; n++) {
                            // 数据预处理配置
                            temp.push(textArray2[n][m]);
                        }

                        var temp1 = chunk(temp, 20);    // 将temp数组分割成以20个数据为一个数组的二维数组temp1

                        for (var k = 0; k < temp1.length; k++) {
                            if ($('input[type="radio"][name="language"]:checked')) {
                                language = $('input[type="radio"][name="language"]:checked').val();   // "db" or "dw"
//                                    console.log("language: " + language);
                                cellValue = "\t" + language + "\t" + temp1[k].join(",") + "\t" + "\t" + ";" + (k * 20 + temp1[k].length) + "\r\n";
                                textArray.push(cellValue);    // 用于导出至txt文本
                                var $p2 = $("<tr>" + "<td>" + cellValue + "</td>" + "</tr>");
                                $("#sheetjs").append($p2);
                            }
                        }
                    }
                    console.log(textArray);
                    // 导出文本
                    saveAs(new Blob(textArray, {type: "text/plain;charset=utf-8"}), "汇编" + language + ".txt");
                });


                // 导出新Excel列表
                $("#exportNewList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                            console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);
                        }
                    }


                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    for (var m = 0; m < titleArray[0].length; m++) {
                        var $thTitle = $("<th>" + titleArray[0][m] + "</th>");
                        $trTitle.append($thTitle);
                    }
                    $("#sheetjs").append($trTitle);

                    for (var p = 0; p < textArray2.length; p++) {
                        var $trContent = $("<tr>" + "</tr>");
                        for (var q = 0; q < textArray2[0].length; q++) {
                            var tdContent = $("<td>" + textArray2[p][q] + "</td>");
                            $trContent.append(tdContent);
                        }
                        $("#sheetjs").append($trContent);
                    }


                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), selectSheetName + ".xlsx");
                });

                // 导出分组
                $("#exportGroupList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                    console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < 2; j++) {       // 只获取参数1列与参数2列的值
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                        }
                    }



                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    $trTitle.append("<th>" + "组序号" + "</th>");
                    $trTitle.append("<th>" + "组的数值范围" + "</th>");
                    $trTitle.append("<th>" + titleArray[0][0] + "</th>");
                    $trTitle.append("<th>" + titleArray[0][1] + "</th>");

                    $("#sheetjs").append($trTitle);


                    for (var m = 0, n = 1, l = 0; m < textArray2.length; m++){
                        var $trContent = $("<tr>" + "</tr>");
                        if ((m+1) < textArray2.length && textArray2[m][0] == textArray2[m+1][0]){
                            l++;  // 相邻两项相同的次数
                            continue;
                        }else{
                            var $tdContent1 = $("<td>" + n + "</td>");
                            n++;
                            var $tdContent2 = $("<td>" + (m+1-l) + "-" + (m+1) + "</td>");
                            l = 0;
                            var $tdContent3 = $("<td>" + textArray2[m][0] + "</td>");
                            var $tdContent4 = $("<td>" + textArray2[m][1] + "</td>");
                            $trContent.append($tdContent1);
                            $trContent.append($tdContent2);
                            $trContent.append($tdContent3);
                            $trContent.append($tdContent4);

                            $("#sheetjs").append($trContent);
                        }
                    }



                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}),  "分组统计.xlsx");
                });

            });
            // dataRow的change事件
            $("#dataRow").change(function () {

                $("#checkNewCnTxt").unbind("click");
                $("#checkNewEnTxt").unbind("click");
                $("#convertNewTxt").unbind("click");
                $("#exportNewList").unbind("click");
                $("#exportGroupList").unbind("click");

                var selectSheetName = $("#selectSheets").val();    // 要处理的sheet名字
                var selectTitleRow = $("#titleRow").val();    // 标题所在的行， value值实际显示值“-1”
                var selectDataRow = $("#dataRow").val();     // 有效数据所在的行， value值实际显示值“-1”


                // 要处理的sheet
                var sheet = workbook.Sheets[selectSheetName];
                // 将要处理的sheet转换为数组json对象
                var sheetArrayJson = XLSX.utils.sheet_to_json(sheet, {header: "A"});    // [{ }, { }, { }]


                // 总列数
                var charCol = "";
                var lengthCol = 0;
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    charCol = index;
                });
                if (charCol.length != 1) {               // "AA-ZZ"的情况
                    lengthCol = (charCol[0].charCodeAt() - 64) * 26 + (charCol[1].charCodeAt() - 64);
                } else {
                    lengthCol = charCol.charCodeAt() - 64;
                }
                console.log("总列数：" + lengthCol);    // 总列数


                // 总行数
                var lengthRow = sheetArrayJson.length;
                console.log("lengthRow: " + lengthRow);      // 总行数


                // 将表格的标题添加到基准Select、参数Select
                // 添加前先清空
                $(".standerCol").empty();    // 基准列
                $(".appendTitles option:first-child").nextAll().remove();

                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    var $appendTitles = $("<option></option>");
                    var $appendTitles2 = $("<option></option>");
                    $appendTitles.html(value).val(index);    // 给参数下拉的option添加value值"A"-"ZZ"
                    $appendTitles2.html(value).val(index);    // 给基准下拉的option添加value值"A"-"ZZ"
                    $(".appendTitles").append($appendTitles);
                    $(".standerCol").append($appendTitles2);
                });


                // 通用列表的中、英文全名与缩写的列
                var enFullNameCol = "";
                var enShortNameCol = "";
                var cnFullNameCol = "";
                var cnShortNameCol = "";

                //  英文全名列
                // 添加前先清空
                $("#checkEnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文全名" || value == "Full Name") {
                        var enFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $enFullNameOption = $("<option>" + enFullName + "</option>");
                        $("#checkEnFullName").append($enFullNameOption);
                        enFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  英文缩写列
                // 添加前先清空
                $("#checkEnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "英文缩写" || value == "Short Name") {
                        var enShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $enShortNameOption = $("<option>" + enShortName + "</option>");
                        $("#checkEnShortName").append($enShortNameOption);
                        enShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文全名列
                // 添加前先清空
                $("#checkCnFullName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文全名" || value == "Full Name Chinese") {
                        var cnFullName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnFullNameOption = $("<option>" + cnFullName + "</option>");
                        $("#checkCnFullName").append($cnFullNameOption);
                        cnFullNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });

                //  中文缩写列
                // 添加前先清空
                $("#checkCnShortName").children().remove();
                $.each(sheetArrayJson[selectTitleRow], function (index, value) {
                    if (value == "中文缩写" || value == "Short Name Chinese") {
                        var cnShortName = sheetArrayJson[selectTitleRow][index];       //
                        var $cnShortNameOption = $("<option>" + cnShortName + "</option>");
                        $("#checkCnShortName").append($cnShortNameOption);
                        cnShortNameCol = index;
                        return false;     // 跳出所有循环,相当于javascript的"break";   return true相当于"continue"
                    }
                });


                // 生成英文列表
                $("#checkNewEnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var enFullName = sheetArrayJson[i][enFullNameCol] + "*";
                            var enShortName = sheetArrayJson[i][enShortNameCol] + "\r\n";
                            txtArray.push(enFullName);
                            txtArray.push(enShortName);
                            var $p1 = $("<tr>" + "<td>" + enFullName + enShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
//                        console.log(txtArray);
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "英文名称列表.txt");
                });


                // 生成中文列表
                $("#checkNewCnTxt").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();

                    var valueCol = $("#checkStander option:selected").val();   // 基准列选中的值（col）
                    var txtArray = [];    // 保存名称的数组，用于导出至.txt
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            var cnFullName = sheetArrayJson[i][cnFullNameCol] + "*";
                            var cnShortName = sheetArrayJson[i][cnShortNameCol] + "\r\n";
                            txtArray.push(cnFullName);
                            txtArray.push(cnShortName);
                            var $p1 = $("<tr>" + "<td>" + cnFullName + cnShortName + "</td>" + "</tr>");
                            $("#sheetjs").append($p1);
                        }
                    }
//                        console.log(txtArray);
                    // 导出名字文本
                    saveAs(new Blob(txtArray, {type: "text/plain;charset=utf-8"}), "中文名称列表.txt");
                });


                // 生成转换列表
                $("#convertNewTxt").click(function () {
                    // 清空显示区
//                        $("#sheetjs").children().remove();
                    var convertParams = $("#form3 .appendTitles option:selected");   // 参数下拉框select选中的option
                    var paramPreConfigs = $("#form3 .paramPreConfig option:selected");   // 数据预处理下拉框select选中的option
//                    console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
//                    console.log(paramPreConfigs[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   // 19
                    console.log("lenCol: " + lenCol);
                    var valueCol = $("#convertStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {

                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    if (paramPreConfigs[j].text == "不处理") {
                                        var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue1 = sheetArrayJson[i][titleValue1];
                                            textArray1.push(preValue1);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据加1") {
                                        var titleValue2 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue2] === undefined || sheetArrayJson[i][titleValue2] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue2 = Number(sheetArrayJson[i][titleValue2]) + 1;
                                            textArray1.push(preValue2);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    } else if (paramPreConfigs[j].text == "参数列数据减1") {
                                        var titleValue3 = convertParams[j].value;       // 已选择的参数标题
                                        titleArray1.push(convertParams[j].text);        // 标题名称数组
                                        if (sheetArrayJson[i][titleValue3] === undefined || sheetArrayJson[i][titleValue3] === null) {
                                            textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                        } else {
                                            var preValue3 = sheetArrayJson[i][titleValue3] - 1;
                                            textArray1.push(preValue3);      // 得到一行数据，以基准列为参考,保存至数组中
                                        }
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                        }
                    }


                    var textArray = [];
                    for (var m = 0; m < textArray2[0].length; m++) {
                        var temp = [];
                        var cellValue = "";
                        var titleShow = "\r\n" + "\r\n" + ";" + titleArray[0][m] + "\r\n";
                        textArray.push(titleShow);
                        var $p1 = $("<tr>" + "<td>" + titleShow + "</td>" + "</tr>");
                        $("#sheetjs").append($p1);
                        for (var n = 0; n < textArray2.length; n++) {
                            // 数据预处理配置
                            temp.push(textArray2[n][m]);
                        }

                        var temp1 = chunk(temp, 20);    // 将temp数组分割成以20个数据为一个数组的二维数组temp1
//                                console.log("temp1: " + temp1);

                        for (var k = 0; k < temp1.length; k++) {
                            if ($('input[type="radio"][name="language"]:checked')) {
                                if ($('input[type="radio"][name="language"]:checked').val() != "c"){
                                    var language = $('input[type="radio"][name="language"]:checked').val();   // "db" or "dw"
//                                    console.log("language: " + language);
                                    cellValue = "\t" + language + "\t" + temp1[k].join(",") + "\t" + "\t" + ";" + (k * 20 + temp1[k].length) + "\r\n";
                                    textArray.push(cellValue);    // 用于导出至txt文本
                                    var $p2 = $("<tr>" + "<td>" + cellValue + "</td>" + "</tr>");
                                    $("#sheetjs").append($p2);
                                }else{

                                }
                            }
                        }
                    }
                    // 导出文本
                    saveAs(new Blob(textArray, {type: "text/plain;charset=utf-8"}), "生成数据表格.asm");
                });


                // 导出新Excel列表
                $("#exportNewList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                    console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < lenCol; j++) {
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                        }
                    }

                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    for (var m = 0; m < titleArray[0].length; m++) {
                        var $thTitle = $("<th>" + titleArray[0][m] + "</th>");
                        $trTitle.append($thTitle);
                    }
                    $("#sheetjs").append($trTitle);

                    for (var p = 0; p < textArray2.length; p++) {
                        var $trContent = $("<tr>" + "</tr>");
                        for (var q = 0; q < textArray2[0].length; q++) {
                            var tdContent = $("<td>" + textArray2[p][q] + "</td>");
                            $trContent.append(tdContent);
                        }
                        $("#sheetjs").append($trContent);
                    }


                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), selectSheetName + ".xlsx");
                });


                // 导出分组
                $("#exportGroupList").click(function () {
                    // 清空显示区
                    $("#sheetjs").children().remove();
                    var convertParams = $("#exportArea .appendTitles option:selected");   // 参数下拉框select选中的option
//                    console.log(convertParams[selectTitleRow]["value"]);    // option的value,即col
                    var lenCol = convertParams.length;   //
                    var valueCol = $("#exportStander option:selected").val();   // 基准列选中的值（col）

                    var textArray2 = [];
                    var titleArray = [];
                    for (var i = selectDataRow; i < lengthRow; i++) {
                        var textArray1 = [];     // 参数下拉框选项值
                        var titleArray1 = [];
                        if (sheetArrayJson[i][valueCol] === undefined || sheetArrayJson[i][valueCol] === null || sheetArrayJson[i][valueCol] == "DEL") {
                            continue;
                        } else {
                            for (var j = 0; j < 2; j++) {       // 只获取参数1列与参数2列的值
                                if (convertParams[j].value == "空") {
                                    continue;
                                } else {
                                    var titleValue1 = convertParams[j].value;       // 已选择的参数标题
                                    titleArray1.push(convertParams[j].text);        // 标题名称数组
                                    if (sheetArrayJson[i][titleValue1] === undefined || sheetArrayJson[i][titleValue1] === null) {
                                        textArray1.push(" ");        // 对应的cell为空时，保存一个空格
                                    } else {
                                        textArray1.push(sheetArrayJson[i][titleValue1]);      // 得到一行数据，以基准列为参考,保存至数组中
                                    }
                                }
                            }
                            textArray2.push(textArray1);    // 二维数组，里面的每一个数组是一列参数
                            titleArray.push(titleArray1);    // 二维数组，里面的每一个数组是一列参数
                        }
                    }



                    // 生成table表格
                    $("#sheetjs").empty();
                    var $trTitle = $("<tr>" + "</tr>");
                    $trTitle.append("<th>" + "组序号" + "</th>");
                    $trTitle.append("<th>" + "组的数值范围" + "</th>");
                    $trTitle.append("<th>" + titleArray[0][0] + "</th>");
                    $trTitle.append("<th>" + titleArray[0][1] + "</th>");

                    $("#sheetjs").append($trTitle);


                    for (var m = 0, n = 1, l = 0; m < textArray2.length; m++){
                        var $trContent = $("<tr>" + "</tr>");
                        if ((m+1) < textArray2.length && textArray2[m][0] == textArray2[m+1][0]){
                            l++;  // 相邻两项相同的次数
                            continue;
                        }else{
                            var $tdContent1 = $("<td>" + n + "</td>");
                            n++;
                            var $tdContent2 = $("<td>" + (m+1-l) + "-" + (m+1) + "</td>");
                            l = 0;
                            var $tdContent3 = $("<td>" + textArray2[m][0] + "</td>");
                            var $tdContent4 = $("<td>" + textArray2[m][1] + "</td>");
                            $trContent.append($tdContent1);
                            $trContent.append($tdContent2);
                            $trContent.append($tdContent3);
                            $trContent.append($tdContent4);

                            $("#sheetjs").append($trContent);
                        }
                    }



                    // table_to_book
                    var tbl = document.getElementById('sheetjs');    // 这里要用元素js读取节点，否则xlsx.js不能识别到
                    var wb = XLSX.utils.table_to_book(tbl);

                    /* bookType can be any supported output type */
                    var wopts = {bookType: 'xlsx', bookSST: false, type: 'binary'};

                    var wbout = XLSX.write(wb, wopts);

                    function s2ab(s) {
                        var buf = new ArrayBuffer(s.length);
                        var view = new Uint8Array(buf);
                        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
                        return buf;
                    }

                    /* the saveAs call downloads a file on the local machine */
                    saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}),  "分组统计.xlsx");
                });
            });
        };
        oReq.send();

    });
});


/* 拓展设置 */
$(function () {
    $("#convertExpand").click(function () {
//                $("#dialog").hide(200);
        $("#dialog2").css("z-index", "1").toggle(100);
    });
});


// 分割数组函数
// chunk([1,2,3],2)   >>>   [ [1,2], [3] ]
var chunk = function (array, size) {
    var result = [];
    for (var x = 0; x < Math.ceil(array.length / size); x++) {
        var start = x * size;
        var end = start + size;
        result.push(array.slice(start, end));
    }
    return result;
};
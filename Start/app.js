/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

	Office.onReady()
		    .then(function() {
		        $(document).ready(function () {

            // TODO1: Determine if the user's version of Office supports all the 确定用户版本的Office是否支持所有
            //        Office.js APIs that are used in the tutorial.本教程中使用的Office.js api。
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
                console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
            }
            // TODO2: Assign event handlers and other initialization logic.分配事件处理程序和其他初始化逻辑。
            $('#create-table').click(createTable);
            $('#freeze-header').click(freezeHeader);/*冻结表头*/
            $('#open-dialog').click(openDialog);/**打开一个对话框 */
        });
    });

    // TODO3: Add handlers and business logic functions here.在这里添加处理程序和业务逻辑函数
    function createTable() {
        Excel.run(function (context) {
    
            // TODO4: Queue table creation logic here.这里是队列表创建逻辑。
            var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            var expensesTable = currentWorksheet.tables.add("A1:I1", true /*hasHeaders标题*/);
            expensesTable.name = "ExpensesTable";
            // TODO5: Queue commands to populate the table with data.用数据填充表的队列命令。
            expensesTable.getHeaderRowRange().values =
            [["序号", "年", "月", "日","摘要","一级科目","次级科目","借方金额","贷方金额"]];

            expensesTable.rows.add(null /*add at the end*/, [
                ["1/1/2017", "The Phone Company", "Communications", "120"],
                ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
                ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
                ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
                ["1/11/2017", "Bellows College", "Education", "350.1"],
                ["1/15/2017", "Trey Research", "Other", "135"],
                ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
            ]);
            // TODO6: Queue commands to format the table.格式化表的队列命令。
            expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
            expensesTable.getRange().format.autofitColumns();
            expensesTable.getRange().format.autofitRows();
            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    /***冻结表头函数 */
    function freezeHeader() {
        Excel.run(function (context) {
    
            // TODO1: Queue commands to keep the header visible when the user scrolls.当用户滚动时，保持头部可见的队列命令。
            var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            currentWorksheet.freezePanes.freezeRows(1);
            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    var dialog = null;  /**此变量用于在父页面的执行上下文中保存对象，该对象充当对话框页面执行上下文的中介 */
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
        Office.context.ui.displayDialogAsync(
            'https://localhost:3000/popup.html',
            {height: 45, width: 55},
        
            // TODO2: Add callback parameter.
            function (result) {
                dialog = result.value;
                dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
            }
        );
    }

    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
        

})();
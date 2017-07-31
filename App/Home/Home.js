/// <reference path="../App.js" />
var language;
var listStockViewOnExcel;
var listStockFollow;

var ssStockServer = "https://stock.spreadsheetspace.net/stockondemand/stockOnDemand";
var sssServer = "https://jarvis.spreadsheetspace.net/";

var strToday;

var stockMap;
var platform;

var selectedTicker;
var nameSelectedTicker;

(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            selectedTicker = "";
            nameSelectedTicker = "";

            platform = navigator.platform;
            language = Office.context.displayLanguage;
            setHtmlLanguage();

            listStockViewOnExcel = [];
            listStockFollow = [];
            stockMap = {};

            var today = new Date();
            strToday = getFormatDate(today);

            $("#from").val(strToday);
            $("#to").val(strToday);

            document.getElementById("infoTab").style.borderRight = "none";
            document.getElementById("homeTabSSS").style.display = "block";

            getListStockViews();

            if (Office.context.document.settings.get("stockViews") != null) {
                listStockFollow = JSON.parse(Office.context.document.settings.get("stockViews"));
                tableStockView();
                updateFollowStock();
            }

            if (Office.context.document.settings.get("worksheetsID") != null) {
                compareWorksheetID();
            }

            checkWorksheetChange();

            setInterval(function () { updateFollowStock(); }, 15 * 60 * 1000); //15 * 60 * 1000

            $('#ok').click(getStockData);

            $('#homeTab').click(clickTab);
            $('#followedTab').click(clickTab);
            $('#infoTab').click(clickTab);

            $('#refresh-tickers').click(getListStockViews);
        });
    };

    function clickTab(evt) {
        // Declare all variables
        var tabName = this.id;
        var i, tabcontent, tablinks;

        // Get all elements with class="tabcontent" and hide them
        tabcontent = document.getElementsByClassName("tabcontent");
        for (i = 0; i < tabcontent.length; i++) {
            tabcontent[i].style.display = "none";
        }

        // Get all elements with class="tablinks" and remove the class "active"
        tablinks = document.getElementsByClassName("tablinks");
        for (i = 0; i < tablinks.length; i++) {
            tablinks[i].className = tablinks[i].className.replace(" active", "");
        }

        // Show the current tab, and add an "active" class to the button that opened the tab
        document.getElementById(tabName+"SSS").style.display = "block";
        evt.currentTarget.className += " active";
    }

    function setHtmlLanguage() {
        if (language == "it-IT") {
            document.getElementById("homeTab").innerHTML = "Home";
            document.getElementById("followedTab").innerHTML = "Seguite";
            document.getElementById("infoTab").innerHTML = "Info";
            document.getElementById("stock_title").innerHTML = "Il servizio di borsa ti permette di monitorare l'andamento dei titoli della Borsa di Milano (i dati riportati sono ritardati di 15 minuti). Seleziona il titolo che ti interessa monitorare, la finestra temporale di interesse e premi \"OK\" per ricevere i dati nella cella che hai selezionato sul tuo foglio di lavoro.";
            document.getElementById("refresh-tickers").innerHTML = "Aggiorna lista titoli";
            document.getElementById("label_warning_1").innerHTML = "Ogni operazione come:";
            document.getElementById("label_warning_li1").innerHTML = "inserire/rimuovere righe/colonne;";
            document.getElementById("label_warning_li2").innerHTML = "taglia/incolla;";
            document.getElementById("label_warning_li3").innerHTML = "modifiche;";
            document.getElementById("label_warning_2").innerHTML = "che coinvolgono uno o più elementi importati dalla borsa, potrebbero compromettere l'aggiornamento dei dati.";
            document.getElementById("p_from").innerHTML = "da:";
            document.getElementById("p_to").innerHTML = "a:";
            document.getElementById("to_follow_label").innerHTML = "segui, aggiornamenti ogni 15 minuti.";
            document.getElementById("follow_title").innerHTML = "Viste di borsa seguite:";
        } else {
            document.getElementById("homeTab").innerHTML = "Home";
            document.getElementById("followedTab").innerHTML = "Followed";
            document.getElementById("infoTab").innerHTML = "Info";
            document.getElementById("stock_title").innerHTML = "The Stock Exchange service allows you to track the progress of Milan Stock Exchange tickers (the data is delayed by 15 minutes). Select the title that interests you, the time window of interest and click \"OK\" to receive the data in the selected cell on your worksheet.";
            document.getElementById("refresh-tickers").innerHTML = "Refresh titles list";
            document.getElementById("label_warning_1").innerHTML = "Any operations like:";
            document.getElementById("label_warning_li1").innerHTML = "insert/remove rows/columns;";
            document.getElementById("label_warning_li2").innerHTML = "cut/paste;";
            document.getElementById("label_warning_li3").innerHTML = "edting;";
            document.getElementById("label_warning_2").innerHTML = "that involved one or more stock element imported, may compromise the data updates.";
            document.getElementById("p_from").innerHTML = "from:";
            document.getElementById("p_to").innerHTML = "to:";
            document.getElementById("to_follow_label").innerHTML = "follow, values updated every 15 minutes.";
            document.getElementById("follow_title").innerHTML = "Followed stock views";
        }
    }

    function getExcelFormatDate(time) {
        var dd = time.getDate();
        var mm = time.getMonth() + 1;
        var yyyy = time.getFullYear();
        if (dd < 10) {
            dd = '0' + dd
        }
        if (mm < 10) {
            mm = '0' + mm
        }
        var newDate = mm + '/' + dd + '/' + yyyy;
        return newDate;
    }

    function getFormatDate(time) {
        var dd = time.getDate();
        var mm = time.getMonth() + 1;
        var yyyy = time.getFullYear();
        if (dd < 10) {
            dd = '0' + dd
        }
        if (mm < 10) {
            mm = '0' + mm
        }
        var newDate = "";
        if (language == "it-IT") {
            newDate = dd + '/' + mm + '/' + yyyy;
        } else {
            newDate = mm + '/' + dd + '/' + yyyy;
        }
        return newDate;
    }

    function getHourMinuteDate(time) {
        var h = time.getHours();
        var m = time.getMinutes();
        if (h < 10) {
            h = '0' + h
        }
        if (m < 10) {
            m = '0' + m
        }

        return h + ':' + m;
    }

    function checkWorksheetIdToRemove() {
        var listWorksheetID = [];
        var listStockToDelete = [];

        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    var ws = worksheets.items[i];
                    ws.load("id");
                }
                ctx.sync().then(function () {
                    for (var i = 0; i < worksheets.items.length; i++) {
                        var id = worksheets.items[i].id;
                        listWorksheetID.push(id);
                    }

                    Office.context.document.settings.set('worksheetsID', JSON.stringify(listWorksheetID));
                    Office.context.document.settings.saveAsync(function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            console.log("Error: " + asyncResult.error.message);
                        } else {
                            console.log("Settings saved");
                        }
                    });
                    
                    //var count_remove = 0;
                    var listStockFollow_size = listStockFollow.length;
                    for (var i = 0; i < listStockFollow_size; i++) {
                        var stock_id = listStockFollow[i].worksheet_id;

                        if (listWorksheetID.indexOf(stock_id) <= -1) {
                            listStockToDelete.push(i);
                        }
                    }

                    if (listStockToDelete.length > 0) {
                        var listStock_size = listStockToDelete.length;

                        for (var i=0; i < listStock_size; i++) {
                            var id = listStockToDelete[i];

                            listStockFollow.splice(id-i, 1);
                        }

                        tableStockView();
                    }
                })
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function checkWorksheetChange() {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
            checkWorksheetIdToRemove();
        });
    }

    function getListStockViews() {
        var token = "22c485fa-f852-4010-a609-61863dd8a4be";
        listStockViewOnExcel = [];

        var url = sssServer + 'stock-ws/last/getNamesAndTickers';

        $.ajax({
            url: url,
            type: 'POST',
            data: null,
            headers: { 'X-Token': token },
            success: function (data, textStatus, jqXHR) {
                createListHtml(data);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                app.showNotification('error');
                document.getElementById("div-list-stock").innerHTML = "";

            }
        });
    }

    function createListHtml(data) {
        stockMap = {};
        
        var stocks = JSON.parse(data);
        for (var i = 0; i < stocks.length; i++) {
            var stock = stocks[i];
            stockMap[stock.name] = stock.ticker;
        }

        var keys = Object.keys(stockMap).sort();

        var div = "";
        if (platform.toLocaleLowerCase().indexOf("mac") > -1) {
            div = "<select id=\"select-list-stock\">\n";
            if (language == "it-IT") {
                div += "<option>" + "Seleziona un titolo..." + "</option>\n";
            } else {
                div += "<option>" + "Select a title..." + "</option>\n";
            }

            for (var key in keys) {
                div += "<option>" + keys[key] + "</option>\n";
            }
            div += "</select>\n"
        } else {
            if (language == "it-IT") {
                div = "<input id=\"input-list-stock\" list=\"tickersName\" placeholder=\"Seleziona un titolo...\"></input>\n";
            } else {
                div = "<input id=\"input-list-stock\" list=\"tickersName\" placeholder=\"Select a title...\"></input>\n";
            }
            div += "<datalist id=tickersName>\n";
            for (var key in keys) {
                div += "<option value=\"" + keys[key] + "\"\>\n";
            }
            div += "</datalist>\n";
        }
     
        document.getElementById("div-list-stock").innerHTML = div;
        if (platform.toLocaleLowerCase().indexOf("mac") > -1) {
            document.getElementById("select-list-stock").onchange = function () {
                nameSelectedTicker = keys[this.selectedIndex - 1];
                selectedTicker = stockMap[nameSelectedTicker];
            }
        }
    }

    function getStockData() {
        var isUpdated = false;
        var from = $("#from").val();
        var to = $("#to").val();

        var epoch_today = "";
        var epoch_from = "";
        var epoch_to = "";

        var dateTodayParts = strToday.split("/");
        var dateFromParts = from.split("/");
        var dateToParts = to.split("/");
        if (language == "it-IT") {
            epoch_today = (new Date(Date.UTC(dateTodayParts[2], dateTodayParts[1] - 1, dateTodayParts[0], 0, 0, 0))).getTime();
            epoch_from = (new Date(Date.UTC(dateFromParts[2], dateFromParts[1] - 1, dateFromParts[0], 0, 0, 0))).getTime();
            epoch_to = (new Date(Date.UTC(dateToParts[2], dateToParts[1] - 1, dateToParts[0], 19, 0, 0))).getTime();
        } else {
            epoch_today = (new Date(Date.UTC(dateTodayParts[2], dateTodayParts[0] - 1, dateTodayParts[1], 0, 0, 0))).getTime();
            epoch_from = (new Date(Date.UTC(dateFromParts[2], dateFromParts[0] - 1, dateFromParts[1], 0, 0, 0))).getTime();
            epoch_to = (new Date(Date.UTC(dateToParts[2], dateToParts[0] - 1, dateToParts[1], 19, 0, 0))).getTime();
        }

        if (!(platform.toLocaleLowerCase().indexOf("mac") > -1)) {
            nameSelectedTicker = $("#input-list-stock").val();
            selectedTicker = stockMap[nameSelectedTicker.toUpperCase()];
        }

        if (nameSelectedTicker == "") {
            if (language == "it-IT") {
                app.showNotification('Nessun titolo selezionato');
            } else {
                app.showNotification('No title selected');
            }
        } else if (selectedTicker == undefined) {
            if (language == "it-IT") {
                app.showNotification('Ticker non presente, selezionare un ticker corretto dal menù');
            } else {
                app.showNotification('Ticker not found, select a correct ticker from the menu');
            }
        } else if (epoch_from > epoch_today) {
            if (language == "it-IT") {
                app.showNotification('Non sono presenti valori nel periodo selezionato per il titolo ' + nameSelectedTicker);
            } else {
                app.showNotification('No values for the selected period for title ' + nameSelectedTicker);
            }
        } else {

            var to_type = document.querySelector('input[name = "date_to"]:checked').id;

            var toFollow;
            if (to_type == "to_follow") {
                toFollow = true;
            } else {
                toFollow = false;
            }

            retrieveStockData(nameSelectedTicker, selectedTicker, epoch_from, epoch_to, isUpdated, toFollow, "", "", "", "");
        }
    }

    function retrieveStockData(nameTicker, ticker, epoch_from, epoch_to, isUpdated, toFollow, savedAddress, savedRowIndex, savedColumnIndex, savedWorksheetID) {
        var objStockData = new ObjStockData(ticker, epoch_from, epoch_to);

        var jsonStock = JSON.stringify(objStockData);

        $.ajax({
            url: ssStockServer,
            type: 'POST',
            data: jsonStock,
            headers: null,
            success: function (data, textStatus, jqXHR) {
                if (data == undefined) {
                    if (language == "it-IT") { 
                        app.showNotification('Si è verificato un errore. Riprova più tardi.');
                    } else {
                        app.showNotification('Error occurred. Try again later.');
                    }
                } else {
                    createTable(data, isUpdated, objStockData, nameTicker, toFollow, savedAddress, savedRowIndex, savedColumnIndex, savedWorksheetID);
                }
            },
            error: function (jqXHR, textStatus, errorThrown) {
                if (language == "it-IT") {
                    app.showNotification('Si è verificato un errore. Riprova più tardi.');
                } else {
                    app.showNotification('Error occurred. Try again later.');
                }
            }
        });
    }

    function createTable(data, isUpdated, objStockData, nameTicker, toFollow, savedAddress, savedRowIndex, savedColumnIndex, savedWorksheetID) {
        var value = [];
        var v = [];
        v[0] = objStockData.ticker;
        if (language == "it-IT") {
            v[1] = 'Apertura';
            v[2] = 'Chiusura';
            v[3] = 'Minimo';
            v[4] = 'Massimo';
            v[5] = 'Volume';
            v[6] = 'Stato';
        } else {
            v[1] = 'Open';
            v[2] = 'Close';
            v[3] = 'Min';
            v[4] = 'Max';
            v[5] = 'TotQuantity';
            v[6] = 'State';
        }

        value.push(v);

        var rowNum = data.length + 1;
        for (var i = 0; i < data.length; i++) {
            var d = data[i];

            var epochDate = new Date(d.timestamp);
            var day = getExcelFormatDate(epochDate);

            v = [];
            v[0] = day;
            v[1] = d.open;
            v[2] = d.close;
            v[3] = d.min;
            v[4] = d.max;
            v[5] = d.totQuantity;
            if (i == data.length - 1) {
                var now = new Date();
                var day = now.getUTCDay();
                var hour = now.getUTCHours();

                if (day != 0 && day != 6) {
                    if (hour >= 7 && hour <= 17) {
                        var hourMinuteNow = getHourMinuteDate(now);
                        if (toFollow) {
                            if (language == "it-IT") {
                                v[6] = 'Valori aggiornati alle ' + hourMinuteNow;
                            } else {
                                v[6] = 'Values updated at ' + hourMinuteNow;
                            }
                        } else {
                            if (language == "it-IT") {
                                v[6] = 'Valori rilevati alle ' + hourMinuteNow;
                            } else {
                                v[6] = 'Values observed at ' + hourMinuteNow;
                            }
                        }
                    } else {
                        v[6] = '';
                    }
                } else {
                    v[6] = '';
                }

            } else {
                v[6] = '';
            }
            
            value.push(v);
            
        }

        if (toFollow && isUpdated) {
            Excel.run(function (ctx) {
                var selectedRange = ctx.workbook.getSelectedRange();
                var worksheet = ctx.workbook.worksheets.getItem(savedWorksheetID)
                worksheet.load('name');

                return ctx.sync().then(function () {
                    insertData(value, worksheet.name, savedWorksheetID, savedAddress, savedRowIndex, savedColumnIndex, rowNum, isUpdated, toFollow, objStockData, nameTicker);

                });
            }).catch(function (error) {
                console.log("Error: " + error);
            });

        } else {
            var address, rowIndex, columnIndex;
            Excel.run(function (ctx) {
                var selectedRange = ctx.workbook.getSelectedRange();
                ctx.workbook.worksheets.getItem
                var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
                worksheet.load('id');
                worksheet.load('name');
                selectedRange.load('address');
                selectedRange.load('rowIndex');
                selectedRange.load('columnIndex');

                return ctx.sync().then(function () {
                    address = selectedRange.address;
                    rowIndex = selectedRange.rowIndex;
                    columnIndex = selectedRange.columnIndex;

                    insertData(value, worksheet.name, worksheet.id, address, rowIndex, columnIndex, rowNum, isUpdated, toFollow, objStockData, nameTicker);

                });
            }).catch(function (error) {
                console.log("Error: " + error);
            });
        }
    }

    function insertData(value, worksheetName, worksheetID, address, rowIndex, columnIndex, rowNum, isUpdated, toFollow, objStockData, nameTicker) {
        Excel.run(function (ctx) {
            //a partire dall'address salvato in precedenza, ricavo la cella di partenza in cui copiare il risultato
            var index = address.indexOf("!") + 1;
            //var wb = address.substring(0, index - 1);
            var wb = worksheetName;
            var cell = address.substring(index);

            //utilizzando rowIndex, columnIndex e le dimensioni del dato ottenuto ricavo l'address dell'ultima cella in cui andro' ad incollare i dati
            var sheet = ctx.workbook.worksheets.getItem(wb);

            var firstCellRange = sheet.getRange(cell + ":" + cell);
            var firstCell = sheet.getCell(rowIndex, columnIndex);
            var lastCell = sheet.getCell(rowIndex + rowNum - 1, columnIndex + 7 - 1);
            var secondRowFirstColumn = sheet.getCell(rowIndex + 1, columnIndex);
            var secondRowSecondColumn = sheet.getCell(rowIndex + 1, columnIndex + 1);
            var lastRowFirstColumn = sheet.getCell(rowIndex + rowNum - 1, columnIndex);
            var lastRowFifthColumn = sheet.getCell(rowIndex + rowNum - 1, columnIndex + 5 - 1);

            lastCell.load('address');
            secondRowFirstColumn.load('address');
            secondRowSecondColumn.load('address');
            lastRowFirstColumn.load('address');
            lastRowFifthColumn.load('address');

            return ctx.sync().then(function () {
                if (!isUpdated) {
                    var table = ctx.workbook.tables.add(cell + ":" + lastCell.address, true);
                    /*if (language == "it-IT") {
                        table.columns.getItemAt(0).getRange().numberFormat = "dd/MM/yyyy";
                    } else {
                        table.columns.getItemAt(0).getRange().numberFormat = "MM/dd/yyyy";
                    }
                    table.columns.getItemAt(1).getRange().numberFormat = "€ #,##0.00";
                    table.columns.getItemAt(2).getRange().numberFormat = "€ #,##0.00";
                    table.columns.getItemAt(3).getRange().numberFormat = "€ #,##0.00";
                    table.columns.getItemAt(4).getRange().numberFormat = "€ #,##0.00";*/
                }

                //creo il range con i dati calcolati prima ed incollo il risultato ottenuto dal server
                var range = sheet.getRange(cell + ":" + lastCell.address);
                range.values = value;

                //modifico i formati del range (data la prima colonna e le 4 colonne successive numero con due cifre dopo la virgola)
                if (language == "it-IT") {
                    sheet.getRange(secondRowFirstColumn.address + ":" + lastRowFirstColumn.address).numberFormat = "dd/MM/yyyy";
                } else {
                    sheet.getRange(secondRowFirstColumn.address + ":" + lastRowFirstColumn.address).numberFormat = "MM/dd/yyyy";
                }
                sheet.getRange(secondRowSecondColumn.address + ":" + lastRowFifthColumn.address).numberFormat = "€ #,##0.00";

                //metto bold solo l'ultima riga nel caso volessi seguire
                if (toFollow) {
                    sheet.getRange(cell + ":" + lastCell.address).format.font.bold = false;
                    sheet.getRange(lastRowFirstColumn.address + ":" + lastCell.address).format.font.bold = true;
                }

                //salvo i metadati
                if (toFollow && !isUpdated) {
                    listStockFollow.push({
                        address: address,
                        rowIndex: rowIndex,
                        columnIndex: columnIndex,
                        epoch_from: objStockData.begin_date,
                        nameTicker: nameTicker,
                        ticker: objStockData.ticker,
                        worksheet_id: worksheetID
                    });

                    tableStockView();
                }
            });
        }).catch(function (error) {
            if (language == "it-IT") {
                app.showNotification("L'intervallo selezionato contiene filtri o tabelle");
            } else {
                app.showNotification("The selected range contains filters or tables");
            }
            console.log("Error: " + error);
        });
    }

    function updateFollowStock() {
        for (var i = 0; i < listStockFollow.length; i++) {
            var stock = listStockFollow[i];

            var epoch_from = stock.epoch_from;
            var epoch_to = (new Date()).getTime();
            var nameTicker = stock.nameTicker;
            var ticker = stock.ticker;
            var savedAddress = stock.address;
            var savedRowIndex = stock.rowIndex;
            var savedColumnIndex = stock.columnIndex;
            var savedWorksheetID = stock.worksheet_id;

            retrieveStockData(nameTicker, ticker, epoch_from, epoch_to, true, true, savedAddress, savedRowIndex, savedColumnIndex, savedWorksheetID);
        }
    }

    function compareWorksheetID() {
        var loadWorksheetsID = JSON.parse(Office.context.document.settings.get("worksheetsID"));

        Excel.run(function (ctx) {
            var worksheets = ctx.workbook.worksheets;
            worksheets.load('items');
            return ctx.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    var ws = worksheets.items[i];
                    ws.load("id");
                }
                ctx.sync().then(function () {
                    var toUpdated = false;
                    for (var i = 0; i < worksheets.items.length; i++) {
                        var current_id = worksheets.items[i].id;
                        var saved_id = loadWorksheetsID[i];

                        if (current_id != saved_id) {
                            for (var j = 0; j < listStockFollow.length; j++) {
                                var worksheet_id = listStockFollow[j].worksheet_id;

                                if (worksheet_id == saved_id) {
                                    listStockFollow[j].worksheet_id = current_id;
                                    toUpdated = true;
                                }
                            }
                        }
                    }

                    if (toUpdated) {
                        updateFollowStock();
                    }

                    Office.context.document.settings.set('worksheetsID', JSON.stringify(listWorksheetID));
                    Office.context.document.settings.saveAsync(function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            console.log("Error: " + asyncResult.error.message);
                        } else {
                            console.log("Settings saved");
                        }
                    });
                })
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function tableStockView() {
        if (listStockFollow.length > 0) {
            //var table = '<table id="tableView">\n<thead>\n<tr>\n<th>Name</th>\n<th></th>\n</tr>\n</thead>\n<tbody>\n';
            var table = '<table id="tableView">\n<tbody>\n';
            for (var i = 0; i < listStockFollow.length; i++) {
                var listStockFollow_i = listStockFollow[i];
                
                if (language == "it-IT") {
                    table += '<tr id="tr">\n<td>' + listStockFollow_i.nameTicker + '</br>(da ' + getFormatDate(new Date(listStockFollow_i.epoch_from)) + ')\t' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '<button id="remove_' + i + '" title="Blocca aggiornamenti"><img src="../../Images/stop.png" style="width:16px;height:16px"></button>' + '</td>\n';
                    table += '<td>' + '<button id="update_' + i + '" title="Aggiorna"><img src="../../Images/update.png" style="width:16px;height:16px"></button>' + '</td>\n';
                    table += '<td>' + '<button id="select_' + i + '" title="Seleziona"><img src="../../Images/select.png" style="width:16px;height:16px"></button>' + '</td>\n</tr>\n';
                } else {
                    table += '<tr id="tr">\n<td>' + listStockFollow_i.nameTicker + '</br>(from ' + getFormatDate(new Date(listStockFollow_i.epoch_from)) + ')\t' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '</td>\n';
                    table += '<td>' + '<button id="remove_' + i + '" title="Stop updating"><img src="../../Images/stop.png" style="width:16px;height:16px"></button>' + '</td>\n';
                    table += '<td>' + '<button id="update_' + i + '" title="Update"><img src="../../Images/update.png" style="width:16px;height:16px"></button>' + '</td>\n';
                    table += '<td>' + '<button id="select_' + i + '" title="Select"><img src="../../Images/select.png" style="width:16px;height:16px"></button>' + '</td>\n</tr>\n';

                }
            }
            table += '</tbody>\n</table>\n';
            table += '<p></p>\n';
            if (language == "it-IT") {
                table += '<button id="updateAll">Aggiorna tutto</button>' + '\n';
            } else {
                table += '<button id="updateAll">Update All</button>' + '\n';
            }

            document.getElementById("div-follow-stock").innerHTML = table;

            $('#updateAll').on('click', updateFollowStock);
            for (var i = 0; i < listStockFollow.length; i++) {
                var listStockFollow_i = listStockFollow[i];
                $('#remove_' + i).on('click', { index: i }, removeStockView);
                $('#update_' + i).on('click', { index: i }, updateStockView);
                $('#select_' + i).on('click', { index: i }, selectStockView);
            }
        } else {
            var table = "";
            document.getElementById("div-follow-stock").innerHTML = table;
        }

        Office.context.document.settings.set('stockViews', JSON.stringify(listStockFollow));
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log("Error: " + asyncResult.error.message);
            } else {
                console.log("Settings saved");
            }
        });
    }

    function removeStockView(event) {
        var index = event.data.index;

        listStockFollow.splice(index, 1);

        tableStockView();
    }

    function updateStockView(event) {
        var index = event.data.index;
        var stock = listStockFollow[index];

        var epoch_from = stock.epoch_from;
        var epoch_to = (new Date()).getTime();
        var nameTicker = stock.nameTicker;
        var ticker = stock.ticker;
        var savedAddress = stock.address;
        var savedRowIndex = stock.rowIndex;
        var savedColumnIndex = stock.columnIndex;
        var savedWorksheetID = stock.worksheet_id;

        retrieveStockData(nameTicker, ticker, epoch_from, epoch_to, true, true, savedAddress, savedRowIndex, savedColumnIndex, savedWorksheetID);
    }

    function selectStockView(event) {
        var index = event.data.index;
        var stock = listStockFollow[index];

        var savedAddress = stock.address;
        var savedWorksheetID = stock.worksheet_id;

        Excel.run(function (ctx) {
            var selectedRange = ctx.workbook.getSelectedRange();
            var worksheet = ctx.workbook.worksheets.getItem(savedWorksheetID)
            worksheet.load('name');

            return ctx.sync().then(function () {
                Excel.run(function (ctx) {
                    //a partire dall'address salvato in precedenza, ricavo la cella di partenza in cui copiare il risultato
                    var index = savedAddress.indexOf("!") + 1;
                    //var wb = address.substring(0, index - 1);
                    var wb = worksheet.name;
                    var cell = savedAddress.substring(index);

                    var movedAddress = wb + "!" + cell;

                    Office.context.document.goToByIdAsync(movedAddress, Office.GoToType.NamedItem, function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            console.log("Error: " + asyncResult.error.message);
                        } else {
                            console.log("Selected new stock");
                        }
                    });

                    return ctx.sync();
                }).catch(function (error) {
                    if (language == "it-IT") {
                        app.showNotification("L'intervallo selezionato contiene filtri o tabelle");
                    } else {
                        app.showNotification("The selected range contains filters or tables");
                    }
                    console.log("Error: " + error);
                });

            });
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    }

    function ObjStockData(ticker, begin_date, end_date) {
        this.ticker = ticker;
        this.begin_date = begin_date;
        this.end_date = end_date;
    }
})();
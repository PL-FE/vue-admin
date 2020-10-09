var spreadNS = GC.Spread.Sheets;

function readData(sheet, url, callback) {
    $.getJSON(url, function (event) {
        if (event && event.data) {
            sheet.suspendPaint();
            var columnInfos = [
                    { name: "FirstName", size: 100 },
                    { name: "MiddleName", size: 100 },
                    { name: "LastName", size: 140 },
                    { name: "Gender", size: 100 },
                    { name: "Title", size: 100 },
                    { name: "PersonType", size: 100 },
                    { name: "Age", size: 100 },
                    { name: "Major", size: 100 },
                    { name: "Phone", size: 100 },
                    { name: "Email", size: 200 },
                    { name: "Location", size: 100 }
            ];
            sheet.autoGenerateColumns = false;
            sheet.setDataSource(event.data);
            sheet.bindColumns(columnInfos);
            var defaultStyle = new spreadNS.Style();
            defaultStyle.hAlign = spreadNS.HorizontalAlign.center;
            defaultStyle.vAlign = spreadNS.VerticalAlign.center;
            defaultStyle.font = "10pt Comic Sans MS";
            sheet.setDefaultStyle(defaultStyle);
            sheet.getCell(-1, 0).hAlign(spreadNS.HorizontalAlign.right);
            sheet.getCell(-1, 2).hAlign(spreadNS.HorizontalAlign.left);
            sheet.defaults.rowHeight = 25;
            sheet.resumePaint();

            if (callback) {
                callback();
            }
        }
    });
}

function createData(sheet, url, callback) {
    var rowCount = sheet.getRowCount();
    sheet.suspendPaint();
    sheet.addRows(rowCount, 1);
    sheet.setValue(rowCount, 0, "Frank");
    sheet.setValue(rowCount, 1, "M");
    sheet.setValue(rowCount, 2, "Lewis");
    sheet.setValue(rowCount, 3, "male");
    sheet.setValue(rowCount, 4, "Mr.");
    sheet.setValue(rowCount, 5, "VC");
    sheet.setValue(rowCount, 6, 26);
    sheet.setValue(rowCount, 7, "CS");
    sheet.setValue(rowCount, 8, "88331988");
    sheet.setValue(rowCount, 9, "someone@grapecity.com");
    sheet.setValue(rowCount, 10, "XiAn");
    sheet.resumePaint();

    var insertRows = sheet.getInsertRows();
    var insertItems = [];
    for (var i = 0; i < insertRows.length; i++) {
        insertItems[i] = insertRows[i].item;
    }
    if (insertItems.length > 0) {
        sheet.clearPendingChanges();
        $.ajax({
            url: url,
            type: "POST",
            data: JSON.stringify(insertItems),
            contentType: "application/json,charset=UTF-8",
            success: callback
        });
    }
}

function updateData(sheet, url, callback) {
    var dirtyRows = sheet.getDirtyRows();
    var dirtyItems = [];
    for (var i = 0; i < dirtyRows.length; i++) {
        dirtyItems[i] = dirtyRows[i].item;
    }
    if (dirtyItems.length > 0) {
        sheet.clearPendingChanges();
        $.ajax({
            url: url,
            type: "POST",
            data: JSON.stringify(dirtyItems),
            dataType: "json",
            contentType: "application/json,charset=UTF-8",
            success: callback
        });
    } else {
        alert("There is no change.")
    }
}

function deleteData(sheet, url, callback) {
    var deletingRowIndex = sheet.getActiveRowIndex();
    var deletingItem;
    if (0 <= deletingRowIndex && deletingRowIndex < sheet.getRowCount()) {
        deletingItem = sheet.getDataItem(deletingRowIndex);
        sheet.deleteRows(deletingRowIndex, 1);
    }
    if (deletingItem) {
        sheet.clearPendingChanges();
        $.ajax({
            url: url,
            type: "POST",
            data: JSON.stringify(deletingItem),
            contentType: "application/json,charset=UTF-8",
            success: callback
        });
    }
}

function getPosition(el) {
    var left = 0, top = 0;
    while (el) {
        left += el.offsetLeft;
        top += el.offsetTop;
        el = el.offsetParent;
    }
    return { left: left, top: top };
}

function showDelayDiv() {
    var $spread = $("#ss");
    var position = getPosition($spread[0]),
        width = $spread.width(),
        height = $spread.height();
    $("<span id='delaySpan'><span id='icon' class='ui-icon ui-icon-clock' style='display:inline-block'></span>Loading...</span>")
        .addClass("busyIndicator")
        .css("left", position.left + width / 2 - 70)
        .css("top", position.top + height / 2 - 30)
        .insertAfter("#ss");
    $("<div id='delayDiv'></div>")
        .css("background", "#2D5972")
        .css("opacity", 0.3)
        .css("position", "absolute")
        .css("top", position.top)
        .css("left", position.left)
        .css("width", width)
        .css("height", height)
        .insertAfter("#ss");
}

function hideDelayDiv() {
    $("#delayDiv").remove();
    $("#delaySpan").remove();
}
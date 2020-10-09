/*jshint -W030 */   // Expected an assignment or function call and instead saw an expression (a && a.fun1())
/*jshint -W004 */   // {a} is already defined (can use let instead of var in es6)
var spreadNS = GC.Spread.Sheets;
var DataValidation = spreadNS.DataValidation;
var ConditionalFormatting = spreadNS.ConditionalFormatting;
var ComparisonOperators = ConditionalFormatting.ComparisonOperators;
var Calc = GC.Spread.CalcEngine;
var ExpressionType = Calc.ExpressionType;
var SheetsCalc = spreadNS.CalcEngine;
var Sparklines = spreadNS.Sparklines;
var Barcode = spreadNS.Barcode;
var isSafari = (function () {
    var tem, M = navigator.userAgent.match(/(opera|chrome|safari|firefox|msie|trident(?=\/))\/?\s*(\d+)/i) || [];
    if (!/trident/i.test(M[1]) && M[1] !== 'Chrome') {
        M = M[2] ? [M[1], M[2]] : [navigator.appName, navigator.appVersion, '-?'];
        if ((tem = navigator.userAgent.match(/version\/(\d+)/i)) != null) M.splice(1, 1, tem[1]);
        return M[0].toLowerCase() === "safari";
    }
    return false;
})();
var isIE = navigator.userAgent.toLowerCase().indexOf('compatible') < 0 && /(trident)(?:.*? rv ([\w.]+)|)/.exec(navigator.userAgent.toLowerCase()) !== null;
var DOWNLOAD_DIALOG_WIDTH = 300;

var spread, excelIO;
var tableIndex = 1, pictureIndex = 1;
var fbx, isShiftKey = false;
var resourceMap = {},
    conditionalFormatTexts = {};
var mergable = false, unmergable = false;
var isFirstChart = true;
var showValue = false;
var showSeriesName = false;
var showCategoryName = false;
var defaultParagraphSeparator = 'p';

function getRichText() {
    var iterator = document.createNodeIterator(document.getElementsByClassName('rich-editor-content')[0], NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, null, false);
    var root = iterator.nextNode();// root
    var richText = [];
    var style = {};
    var text = '';
    var node = iterator.nextNode();
    var underlineNode = null, lineThroughNode = null, pNode = null;
    while (node !== null) {
        if (node.nodeType === 3/*TextNode*/) {
            text = node.nodeValue;
            style = document.defaultView.getComputedStyle(node.parentElement, null);
            if (underlineNode && underlineNode.contains(node) === false) {
                underlineNode = null;
            }
            if (lineThroughNode && lineThroughNode.contains(node) === false) {
                lineThroughNode = null;
            }
            if (pNode && getLastTextNode(pNode) === node && getLastTextNode(root) !== node) {
                text = text + '\r\n';
                pNode = null;
            }
            var richTextStyle = getRichStyle(style, underlineNode, lineThroughNode);
            handleSuperAndSubScript(root,node,richTextStyle);
            richText.push({
                style: richTextStyle,
                text: text
            });
        } else if (node.nodeName.toLowerCase() === defaultParagraphSeparator) {
            pNode = node;
        } else if (node.nodeName.toLowerCase() === 'u') {
            underlineNode = node;
        } else if (node.nodeName.toLowerCase() === 'strike') {
            lineThroughNode = node;
        }

        node = iterator.nextNode();
    }
    return richText;
}

function handleSuperAndSubScript(root,node,style){
    if (root === node){
        return;
    }
    while(node.parentNode !== root){
        if(node.nodeName.toLowerCase() === 'sub'){
            style.vertAlign = 2;
            break;
        }
        if(node.nodeName.toLowerCase() === 'sup'){
            style.vertAlign = 1;
            break;
        }
        node = node.parentNode;
    }
}

function getRichStyle(style, isUnderlineNode, isLineThroughNode) {// getComputedStyle can't get inherit textDecoration
    return {
        font: (style.fontWeight === '700' ? 'bold ' : '') + (style.fontStyle === 'italic' ? 'italic ' : '') + style.fontSize + ' ' + style.fontFamily,
        foreColor: style.color,
        textDecoration: (isUnderlineNode ? 1 : 0) | (isLineThroughNode ? 2 : 0)
    };
}

function getLastTextNode(root) {
    if (root && root.nodeType === 1) {
        var child = root.lastChild;
        return getLastTextNode(child);
    } else {
        return root;
    }
}

function toggleState() {
    var $element = $(this),
        $parent = $element.parent(),
        $content = $parent.siblings(".insp-group-content"),
        $target = $parent.find("span.group-state"),
        collapsed = $target.hasClass("fa-caret-right");

    if (collapsed) {
        $target.removeClass("fa-caret-right").addClass("fa-caret-down");
        $content.slideToggle("fast");
    } else {
        $target.addClass("fa-caret-right").removeClass("fa-caret-down");
        $content.slideToggle("fast");
    }
}

function updateMergeButtonsState() {
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();
    mergable = false, unmergable = false;
    sels.forEach(function (range) {
        var ranges = sheet.getSpans(range),
            spanCount = ranges.length;

        if (!mergable) {
            if (spanCount > 1 || (spanCount === 0 && (range.rowCount > 1 || range.colCount > 1))) {
                mergable = true;
            } else if (spanCount === 1) {
                var range2 = ranges[0];
                if (range2.row !== range.row || range2.col !== range.col ||
                    range2.rowCount !== range2.rowCount || range2.colCount !== range.colCount) {
                    mergable = true;
                }
            }
        }
        if (!unmergable) {
            unmergable = spanCount > 0;
        }
    });

    $("#mergeCells").attr("disabled", mergable ? null : "disabled");
    $("#unmergeCells").attr("disabled", unmergable ? null : "disabled");
}

function updateCellStyleState(sheet, row, column) {
    var style = sheet.getActualStyle(row, column);

    if (style) {
        var sfont = style.font;

        // Font
        var font
        if (sfont) {
            font = parseFont(sfont);

            setFontStyleButtonActive("bold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
            setFontStyleButtonActive("italic", font.fontStyle !== 'normal');
            setDropDownText($("#cellTab div.insp-dropdown-list[data-name='fontFamily']"), font.fontFamily);
            setDropDownText($("#cellTab div.insp-dropdown-list[data-name='fontSize']"), parseFloat(font.fontSize));
        }

        var underline = spreadNS.TextDecorationType.underline,
            linethrough = spreadNS.TextDecorationType.lineThrough,
            overline = spreadNS.TextDecorationType.overline,
            doubleUnderline = spreadNS.TextDecorationType.doubleUnderline,
            textDecoration = style.textDecoration;
        setFontStyleButtonActive("underline", textDecoration && ((textDecoration & underline) === underline));
        setFontStyleButtonActive("strikethrough", textDecoration && ((textDecoration & linethrough) === linethrough));
        setFontStyleButtonActive("overline", textDecoration && ((textDecoration & overline) === overline));
        setFontStyleButtonActive("double-underline", textDecoration && ((textDecoration & doubleUnderline) === doubleUnderline));

        setColorValue("foreColor", style.foreColor || "#000");
        setColorValue("backColor", style.backColor || "#fff");

        // Alignment
        setRadioButtonActive("hAlign", style.hAlign);   // general (3, auto detect) without setting button just like Excel
        setRadioButtonActive("vAlign", style.vAlign);
        setCheckValue("wrapText", style.wordWrap);

        //cell padding
        var cellPadding = style.cellPadding;
        if (cellPadding) {
            setTextValue("cellPadding", cellPadding);
        } else {
            setTextValue("cellPadding", "");
        }
        //watermark
        var watermark = style.watermark;
        if (watermark) {
            setTextValue("watermark", watermark);
        } else {
            setTextValue("watermark", "");
        }
        //label options
        var labelOptions = style.labelOptions;
        if (labelOptions) {
            var lFont = labelOptions.font;
            if (lFont) {
                font = parseFont(lFont);
                setFontStyleButtonActive("labelBold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
                setFontStyleButtonActive("labelItalic", font.fontStyle !== 'normal');
                setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontFamily']"), font.fontFamily.replace(/'/g, ""));
                setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontSize']"), parseFloat(font.fontSize));
            } else {
                setFontStyleButtonActive("labelBold", ["bold", "bolder", "700", "800", "900"].indexOf(font.fontWeight) !== -1);
                setFontStyleButtonActive("labelItalic", font.fontStyle !== 'normal');
                setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontFamily']"), font.fontFamily.replace(/'/g, ""));
                setDropDownText($("#cellTab div.insp-dropdown-list[data-name='labelFontSize']"), parseFloat(font.fontSize));
            }
            setColorValue("labelForeColor", labelOptions.foreColor || "#000");
            setTextValue("labelMargin", labelOptions.margin || "");
            setDropDownValueByIndex($("#cellLabelVisibility"), labelOptions.visibility === undefined ? 2 : labelOptions.visibility);
            setDropDownValueByIndex($("#cellLabelAlignment"), labelOptions.alignment || 0);
        }
    }
}

function setFontStyleButtonActive(name, active) {
    var $target = $("div.group-container>span[data-name='" + name + "']");

    if (active) {
        $target.addClass("active");
    } else {
        $target.removeClass("active");
    }
}

function setRadioButtonActive(name, index) {
    var $items = $("div.insp-radio-button-group[data-name='" + name + "'] div>span");

    $items.removeClass("active");
    $($items[index]).addClass("active");
}

function parseFont(font) {
    var fontFamily = null,
        fontSize = null,
        fontStyle = "normal",
        fontWeight = "normal",
        fontVariant = "normal",
        lineHeight = "normal";

    var elements = font.split(/\s+/);
    var element;
    while ((element = elements.shift())) {
        switch (element) {
            case "normal":
                break;

            case "italic":
            case "oblique":
                fontStyle = element;
                break;

            case "small-caps":
                fontVariant = element;
                break;

            case "bold":
            case "bolder":
            case "lighter":
            case "100":
            case "200":
            case "300":
            case "400":
            case "500":
            case "600":
            case "700":
            case "800":
            case "900":
                fontWeight = element;
                break;

            default:
                if (!fontSize) {
                    var parts = element.split("/");
                    fontSize = parts[0];
                    if (fontSize.indexOf("px") !== -1) {
                        fontSize = px2pt(parseFloat(fontSize)) + 'pt';
                    }
                    if (parts.length > 1) {
                        lineHeight = parts[1];
                        if (lineHeight.indexOf("px") !== -1) {
                            lineHeight = px2pt(parseFloat(lineHeight)) + 'pt';
                        }
                    }
                    break;
                }

                fontFamily = element;
                if (elements.length)
                    fontFamily += " " + elements.join(" ");

                return {
                    "fontStyle": fontStyle,
                    "fontVariant": fontVariant,
                    "fontWeight": fontWeight,
                    "fontSize": fontSize,
                    "lineHeight": lineHeight,
                    "fontFamily": fontFamily
                };
        }
    }

    return {
        "fontStyle": fontStyle,
        "fontVariant": fontVariant,
        "fontWeight": fontWeight,
        "fontSize": fontSize,
        "lineHeight": lineHeight,
        "fontFamily": fontFamily
    };
}

var tempSpan = $("<span></span>");
function px2pt(pxValue) {
    tempSpan.css({
        "font-size": "96pt",
        "display": "none"
    });
    tempSpan.appendTo($(document.body));
    var tempPx = tempSpan.css("font-size");
    if (tempPx.indexOf("px") !== -1) {
        var tempPxValue = parseFloat(tempPx);
        return Math.round(pxValue * 96 / tempPxValue);
    }
    else {  // when browser have not convert pt to px, use 96 DPI.
        return Math.round(pxValue * 72 / 96);
    }
}

function processRadioButtonClicked(key, $item, $group) {
    var name = $item.data("name");

    // only need process when click on radio button or relate label like text
    if ($item.hasClass("radiobutton") || $item.hasClass("text")) {
        $group.find("div.radiobutton").removeClass("checked");
        $group.find("div.radiobutton[data-name='" + name + "']").addClass("checked");

        switch (key) {
            case "referenceStyle":
                setReferenceStyle(name);
                break;
            case "slicerMoveAndSize":
                setSlicerSetting("moveSize", name);
                break;
            case "pictureMoveAndSize":
                var picture = _activePicture;
                if (name === "picture-move-size") {
                    picture.dynamicMove(true);
                    picture.dynamicSize(true);
                }
                if (name === "picture-move-nosize") {
                    picture.dynamicMove(true);
                    picture.dynamicSize(false);
                }
                if (name === "picture-nomove-size") {
                    picture.dynamicMove(false);
                    picture.dynamicSize(false);
                }
                break;
        }
    }
}

function setReferenceStyle(name) {
    var referenceStyle, columnHeaderAutoText;

    if (name === "a1style") {
        referenceStyle = spreadNS.ReferenceStyle.a1;
        columnHeaderAutoText = spreadNS.HeaderAutoText.letters;
    } else {
        referenceStyle = spreadNS.ReferenceStyle.r1c1;
        columnHeaderAutoText = spreadNS.HeaderAutoText.numbers;
    }

    spread.options.referenceStyle = referenceStyle;
    spread.sheets.forEach(function (sheet) {
        sheet.options.colHeaderAutoText = columnHeaderAutoText;
    });
    updatePositionBox(spread.getActiveSheet());
}

function checkedChanged() {
    var $element = $(this),
        name = $element.data("name");

    if ($element.hasClass("disabled")) {
        return;
    }

    // radio buttons need special process
    switch (name) {
        case "referenceStyle":
        case "slicerMoveAndSize":
        case "pictureMoveAndSize":
            processRadioButtonClicked(name, $(event.target), $element);
            return;
    }


    var $target = $("div.button", $element),
        value = !$target.hasClass("checked");

    var sheet = spread.getActiveSheet();

    $target.toggleClass("checked");

    spread.suspendPaint();

    var options = spread.options;

    switch (name) {

        case  "allowCopyPasteExcelStyle":
            options.allowCopyPasteExcelStyle = value;
            break;

        case "allowExtendPasteRange":
            options.allowExtendPasteRange = value;
            break;

        case "referenceStyle":
            options.referenceStyle = (value ? spreadNS.ReferenceStyle.r1c1 : spreadNS.ReferenceStyle.a1);
            break;

        case "cutCopyIndicatorVisible":
            options.cutCopyIndicatorVisible = value;
            break;

        case "showVerticalScrollbar":
            options.showVerticalScrollbar = value;
            break;

        case "showHorizontalScrollbar":
            options.showHorizontalScrollbar = value;
            break;

        case "scrollIgnoreHidden":
            options.scrollIgnoreHidden = value;
            break;

        case "scrollbarMaxAlign":
            options.scrollbarMaxAlign = value;
            break;

        case "scrollbarShowMax":
            options.scrollbarShowMax = value;
            break;

        case "tabStripVisible":
            options.tabStripVisible = value;
            break;

        case "newTabVisible":
            options.newTabVisible = value;
            break;

        case "tabEditable":
            options.tabEditable = value;
            break;

        case "showTabNavigation":
            options.tabNavigationVisible = value;
            break;

        case "showDragDropTip":
            options.showDragDropTip = value;
            break;

        case "showDragFillTip":
            options.showDragFillTip = value;
            break;

        case "sheetVisible":
            var sheetIndex = $target.data("sheetIndex"),
                sheetName = $target.data("sheetName"),
                selectedSheet = spread.sheets[sheetIndex];

            // be sure related sheet not changed (such add / remove sheet, rename sheet)
            if (selectedSheet && selectedSheet.name() === sheetName) {
                selectedSheet.visible(value);
            } else {
                //console.log("selected sheet' info was changed, please select the sheet and set visible again.");
            }
            break;

        case "allowUserDragDrop":
            spread.options.allowUserDragDrop = value;
            break;

        case "allowUserDragFill":
            spread.options.allowUserDragFill = value;
            break;

        case "allowZoom":
            spread.options.allowUserZoom = value;
            break;

        case "allowOverflow":
            spread.sheets.forEach(function (sheet) {
                sheet.options.allowCellOverflow = value;
            });
            break;

        case "showDragFillSmartTag":
            spread.options.showDragFillSmartTag = value;
            break;

        case "allowDragMerge":
            spread.options.allowUserDragMerge = value;
            break;

        case "allowContextMenu":
            spread.options.allowContextMenu = value;
            break;

        case "allowUserDeselect":
            spread.options.allowUserDeselect = value;
            break;

        case "showVerticalGridline":
            sheet.options.gridline.showVerticalGridline = value;
            break;

        case "showHorizontalGridline":
            sheet.options.gridline.showHorizontalGridline = value;
            break;

        case "showRowHeader":
            sheet.options.rowHeaderVisible = value;
            break;

        case "showColumnHeader":
            sheet.options.colHeaderVisible = value;
            break;

        case "wrapText":
            setWordWrap(sheet);
            break;
        case "hideSelection":
            spread.options.hideSelection = value;
            break;

        case "showRowOutline":
            sheet.showRowOutline(value);
            break;

        case "showColumnOutline":
            sheet.showColumnOutline(value);
            break;

        case "highlightInvalidData":
            spread.options.highlightInvalidData = value;
            break;

        /* table realted items */
        case "tableFilterButton":
            _activeTable && _activeTable.filterButtonVisible(value);
            break;

        case "tableHeaderRow":
            _activeTable && _activeTable.showHeader(value);
            break;

        case "tableTotalRow":
            _activeTable && _activeTable.showFooter(value);
            break;

        case "tableBandedRows":
            _activeTable && _activeTable.bandRows(value);
            break;

        case "tableBandedColumns":
            _activeTable && _activeTable.bandColumns(value);
            break;

        case "tableFirstColumn":
            _activeTable && _activeTable.highlightFirstColumn(value);
            break;

        case "tableLastColumn":
            _activeTable && _activeTable.highlightLastColumn(value);
            break;
        /* table realted items (end) */

        /* comment related items */
        case "commentDynamicSize":
            _activeComment && _activeComment.dynamicSize(value);
            break;

        case "commentDynamicMove":
            _activeComment && _activeComment.dynamicMove(value);
            break;

        case "commentLockText":
            _activeComment && _activeComment.lockText(value);
            break;

        case "commentShowShadow":
            _activeComment && _activeComment.showShadow(value);
            break;
        /* comment related items (end) */

        /* picture related items */
        case "pictureDynamicSize":
            _activePicture && _activePicture.dynamicSize(value);
            break;

        case "pictureDynamicMove":
            _activePicture && _activePicture.dynamicMove(value);
            break;

        case "pictureFixedPosition":
            _activePicture && _activePicture.fixedPosition(value);
            break;
        /* picture related items (end) */

        /* protect sheet realted items */
        case "checkboxProtectSheet":
            syncProtectSheetRelatedItems(sheet, value);
            break;

        case "checkboxSelectLockedCells":
            setProtectionOption(sheet, "allowSelectLockedCells", value);
            break;

        case "checkboxSelectUnlockedCells":
            setProtectionOption(sheet, "allowSelectUnlockedCells", value);
            break;

        case "checkboxSort":
            setProtectionOption(sheet, "allowSort", value);
            break;

        case "checkboxUseAutoFilter":
            setProtectionOption(sheet, "allowFilter", value);
            break;

        case "checkboxResizeRows":
            setProtectionOption(sheet, "allowResizeRows", value);
            break;

        case "checkboxResizeColumns":
            setProtectionOption(sheet, "allowResizeColumns", value);
            break;

        case "checkboxEditObjects":
            setProtectionOption(sheet, "allowEditObjects", value);
            break;

        case "checkboxDragInsertRows":
            setProtectionOption(sheet, "allDragInsertRows", value);
            break;

        case "checkboxDragInsertColumns":
            setProtectionOption(sheet, "allowDragInsertColumns", value);
            break;

        case "checkboxInsertRows":
            setProtectionOption(sheet, "allowInsertRows", value);
            break;

        case "checkboxInsertColumns":
            setProtectionOption(sheet, "allowInsertColumns", value);
            break;

        case "checkboxDeleteRows":
            setProtectionOption(sheet, "allowDeleteRows", value);
            break;

        case "checkboxDeleteColumns":
            setProtectionOption(sheet, "allowDeleteColumns", value);
            break;
        /* protect sheet realted items (end) */

        /* slicer related items */
        case "displaySlicerHeader":
            setSlicerSetting("showHeader", value);
            break;

        case "lockSlicer":
            setSlicerSetting("lock", value);
            break;
        /* slicer related items (end) */

        case "showDataLabelsValue":
            var isShow = judjeDataLabelsIsShow({item:"showDataLabelsValue",isShow:value});
            updateDataLabelsPositionDropDown(isShow);
            break;
        case "showDataLabelsSeriesName":
            var isShow = judjeDataLabelsIsShow({item:"showDataLabelsSeriesName",isShow:value});
            updateDataLabelsPositionDropDown(isShow);
            break;
        case "showDataLabelsCategoryName":
            var isShow = judjeDataLabelsIsShow({item:"showDataLabelsCategoryName",isShow:value});
            updateDataLabelsPositionDropDown(isShow);
            break;
        case "useChartAnimation":
            applyChartAnimationSetting(value);
            break;

        default:
            //console.log("not added code for", name);
            break;

    }
    spread.resumePaint();
}

function updateNumberProperty() {
    var $element = $(this),
        $parent = $element.parent(),
        name = $parent.data("name"),
        value = parseInt($element.val(), 10);

    if (isNaN(value)) {
        return;
    }

    var sheet = spread.getActiveSheet();

    spread.suspendPaint();
    switch (name) {
        case "rowCount":
            sheet.setRowCount(value);
            break;

        case "columnCount":
            sheet.setColumnCount(value);
            break;

        case "frozenRowCount":
            sheet.frozenRowCount(value);
            break;

        case "frozenColumnCount":
            sheet.frozenColumnCount(value);
            break;

        case "trailingFrozenRowCount":
            sheet.frozenTrailingRowCount(value);
            break;

        case "trailingFrozenColumnCount":
            sheet.frozenTrailingColumnCount(value);
            break;

        case "commentBorderWidth":
            _activeComment && _activeComment.borderWidth(value);
            break;

        case "commentOpacity":
            _activeComment && _activeComment.opacity(value / 100);
            break;

        case "pictureBorderWidth":
            _activePicture && _activePicture.borderWidth(value);
            break;

        case "pictureBorderRadius":
            _activePicture && _activePicture.borderRadius(value);
            break;

        case "slicerColumnNumber":
            setSlicerSetting("columnCount", value);
            break;

        case "slicerButtonHeight":
            setSlicerSetting("itemHeight", value);
            break;

        case "slicerButtonWidth":
            setSlicerSetting("itemWidth", value);
            break;

        default:
            //console.log("updateNumberProperty need add for", name);
            break;
    }
    spread.resumePaint();
}

function updateStringProperty() {
    var $element = $(this),
        $parent = $element.parent(),
        name = $parent.data("name"),
        value = $element.val();

    var sheet = spread.getActiveSheet();

    switch (name) {
        case "sheetName":
            if (value && value !== sheet.name()) {
                try {
                    sheet.name(value);
                } catch (ex) {
                    alert(getResource("messages.duplicatedSheetName"));
                    $element.val(sheet.name());
                }
            }
            break;

        case "tableName":
            if (value && _activeTable && value !== _activeTable.name()) {
                if (!sheet.tables.findByName(value)) {
                    _activeTable.name(value);
                } else {
                    alert(getResource("messages.duplicatedTableName"));
                    $element.val(_activeTable.name());
                }
            }
            break;

        case "commentPadding":
            setCommentPadding(value);
            break;

        case "customFormat":
            setFormatter(value);
            break;

        case "slicerName":
            setSlicerSetting("name", value);
            break;

        case "slicerCaptionName":
            setSlicerSetting("captionName", value);
            break;

        case "watermark":
            setWatermark(sheet, value);
            break;

        case "cellPadding":
            setCellPadding(sheet, value);
            break;

        case "labelmargin":
            setLabelOptions(sheet, value, "margin");
            break;
        case "shapeText":
            setTextValue("shapeText",value);
            break;

        default:
            //console.log("updateStringProperty w/o process of ", name);
            break;
    }
}

function setCommentPadding(padding) {
    if (_activeComment && padding) {
        var para = padding.split(",");
        if (para.length === 1) {
            _activeComment.padding(new spreadNS.Comments.Padding(parseInt(para[0], 10)));
        } else if (para.length === 4) {
            _activeComment.padding(new spreadNS.Comments.Padding(parseInt(para[0], 10), parseInt(para[1], 10), parseInt(para[2], 10), parseInt(para[3], 10)));
        }
    }
}

function fillSheetNameList($container) {
    var html = "";

    // unbind event if present
    $container.find(".menu-item").off('click');

    spread.sheets.forEach(function (sheet, index) {
        html += '<div class="menu-item"><div class="image"></div><div class="text" data-value="' + index + '">' + sheet.name() + '</div></div>';
    });
    $container.html(html);

    // bind event for new added elements
    $container.find(".menu-item").on('click', itemSelected);
}

function syncSpreadPropertyValues() {
    var options = spread.options;
    // General
    setCheckValue("allowUserDragDrop", options.allowUserDragDrop);
    setCheckValue("allowUserDragFill", options.allowUserDragFill);
    setCheckValue("allowZoom", options.allowUserZoom);
    setCheckValue("allowOverflow", spread.getActiveSheet().options.allowCellOverflow);
    setCheckValue("showDragFillSmartTag", options.showDragFillSmartTag);
    setCheckValue("allowDragMerge", options.allowUserDragMerge);
    setDropDownValue("resizeZeroIndicator", options.resizeZeroIndicator);

    // Calculation
    setRadioItemChecked("referenceStyle", options.referenceStyle === spreadNS.ReferenceStyle.r1c1 ? "r1c1style" : "a1style");

    // Scroll Bar
    setCheckValue("showVerticalScrollbar", options.showVerticalScrollbar);
    setCheckValue("showHorizontalScrollbar", options.showHorizontalScrollbar);
    setCheckValue("scrollbarMaxAlign", options.scrollbarMaxAlign);
    setCheckValue("scrollbarShowMax", options.scrollbarShowMax);
    setCheckValue("scrollIgnoreHidden", options.scrollIgnoreHidden);

    // TabStrip
    setCheckValue("tabStripVisible", options.tabStripVisible);
    setCheckValue("newTabVisible", options.newTabVisible);
    setCheckValue("tabEditable", options.tabEditable);
    setCheckValue("allowSheetReorder", options.allowSheetReorder);
    setCheckValue("showTabNavigation", options.tabNavigationVisible);

    // Color
    setColorValue("spreadBackcolor", options.backColor);
    setColorValue("grayAreaBackcolor", options.grayAreaBackColor);

    // Tip
    setDropDownValue($("div.insp-dropdown-list[data-name='scrollTip']"), options.showScrollTip);
    setDropDownValue($("div.insp-dropdown-list[data-name='resizeTip']"), options.showResizeTip);
    setCheckValue("showDragDropTip", options.showDragDropTip);
    setCheckValue("showDragFillTip", options.showDragFillTip);

    // Cut / Copy Indicator
    setCheckValue("cutCopyIndicatorVisible", options.cutCopyIndicatorVisible);
    setColorValue("cutCopyIndicatorBorderColor", options.cutCopyIndicatorBorderColor);

    // Data validation
    setCheckValue("highlightInvalidData", options.highlightInvalidData);
}

function syncForzenProperties(sheet) {
    setNumberValue("frozenRowCount", sheet.frozenRowCount());
    setNumberValue("frozenColumnCount", sheet.frozenColumnCount());
    setNumberValue("trailingFrozenRowCount", sheet.frozenTrailingRowCount());
    setNumberValue("trailingFrozenColumnCount", sheet.frozenTrailingColumnCount());
}

function syncSheetPropertyValues() {
    var sheet = spread.getActiveSheet(),
        options = sheet.options;

    // General
    setNumberValue("rowCount", sheet.getRowCount());
    setNumberValue("columnCount", sheet.getColumnCount());
    setTextValue("sheetName", sheet.name());
    setColorValue("sheetTabColor", options.sheetTabColor);

    // Grid Line
    setCheckValue("showVerticalGridline", options.gridline.showVerticalGridline);
    setCheckValue("showHorizontalGridline", options.gridline.showHorizontalGridline);
    setColorValue("gridlineColor", options.gridline.color);

    // Header
    setCheckValue("showRowHeader", options.rowHeaderVisible);
    setCheckValue("showColumnHeader", options.colHeaderVisible);

    // Freeze
    setColorValue("frozenLineColor", options.frozenlineColor);

    syncForzenProperties(sheet);

    // Selection
    setDropDownValue($("#sheetTab div.insp-dropdown-list[data-name='selectionPolicy']"), sheet.selectionPolicy());
    setDropDownValue($("#sheetTab div.insp-dropdown-list[data-name='selectionUnit']"), sheet.selectionUnit());
    setColorValue("selectionBorderColor", options.selectionBorderColor);
    setColorValue("selectionBackColor", options.selectionBackColor);
    setCheckValue("hideSelection", spread.options.hideSelection);

    // Protection
    var isProtected = options.isProtected;
    setCheckValue("checkboxProtectSheet", isProtected);
    syncProtectSheetRelatedItems(sheet, isProtected);
    getCurrentSheetProtectionOption(sheet);

    updateCellStyleState(sheet, sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());

    // Group
    setCheckValue("showRowOutline", sheet.showRowOutline());
    setCheckValue("showColumnOutline", sheet.showColumnOutline());

    if (!$(sheet).data("bind")) {
        $(sheet).data("bind", true);
        sheet.bind(spreadNS.Events.RangeChanged, function (event, args) {
            if (args.action === spreadNS.RangeChangedAction.clear) {
                // check special type items and switch to cell tab (laze process)
                if (isSpecialTabSelected()) {
                    onCellSelected();
                }
            }
        });
        sheet.bind(spreadNS.Events.FloatingObjectRemoved, function (event, args) {
            // check special type items and switch to cell tab (laze process)
            if (isSpecialTabSelected()) {
                onCellSelected();
            }
        });
        sheet.bind(spreadNS.Events.CommentRemoved, function (event, args) {
            // check special type items and switch to cell tab (laze process)
            if (isSpecialTabSelected()) {
                onCellSelected();
            }
        });
    }
}

function setNumberValue(name, value) {
    $("div.insp-number[data-name='" + name + "'] input.editor").val(value);
}

function getNumberValue(name) {
    return +$("div[data-name='" + name + "'] input.editor").val();
}

function setTextValue(name, value) {
    $("div.insp-text[data-name='" + name + "'] input.editor").val(value);
}

function getTextValue(name) {
    return $("div.insp-text[data-name='" + name + "'] input.editor").val();
}

function setCheckValue(name, value, options) {
    var $target = $("div.insp-checkbox[data-name='" + name + "'] div.button");
    if (value) {
        $target.addClass("checked");
    } else {
        $target.removeClass("checked");
    }
    if (options) {
        $target.data(options);
    }
}

function getCheckValue(name) {
    var $target = $("div.insp-checkbox[data-name='" + name + "'] div.button");

    return $target.hasClass("checked");
}

function setColorValue(name, value) {
    $("div.insp-color-picker[data-name='" + name + "'] div.color-view").css("background-color", value || "");
}

var _dropdownitem;
var _colorpicker;
var _needShow = true;

var _handlePopupCloseEvents = 'mousedown touchstart MSPointerDown pointerdown'.split(' ');

function processEventListenerHandleClosePopup(add) {
    if (add) {
        _handlePopupCloseEvents.forEach(function (value) {
            document.addEventListener(value, documentMousedownHandler, true);
        });
    } else {
        _handlePopupCloseEvents.forEach(function (value) {
            document.removeEventListener(value, documentMousedownHandler, true);
        });
    }
}

function showDropdown() {
    if (!_needShow) {
        _needShow = true;
        return;
    }

    var DROPDOWN_OFFSET = 10;
    var $element = $(this),
        $container = $element.parent(),
        name = $container.data("name"),
        targetId = $container.data("list-ref"),
        $target = $("#" + targetId);

    if ($target && !$target.hasClass("show")) {
        if (name === "sheetName") {
            fillSheetNameList($target);
        }

        $target.data("dropdown", this);
        _dropdownitem = $target[0];

        var $dropdown = $element,
            offset = $dropdown.offset();

        var height = $element.outerHeight(),
            targetHeight = $target.outerHeight(),
            width = $element.outerWidth(),
            targetWidth = $target.outerWidth(),
            top = offset.top + height;

        // adjust drop down' width to same
        if (targetWidth < width) {
            $target.width(width);
        }

        var $inspContainer = $(".insp-container"),
            maxTop = $inspContainer.height() + $inspContainer.offset().top;

        // adjust top when out of bottom range
        if (top + targetHeight + DROPDOWN_OFFSET > maxTop) {
            top = offset.top - targetHeight;
        }

        $target.css({
            top: top,
            left: offset.left - $target.width() + $dropdown.width() + 16
        });

        // select corresponding item
        if (name === "borderLine" || name === "shapeBorder" || name === "beginArrowStyle" || name === "endArrowStyle") {
            var text;
            switch(name){
                case "shapeBorder":
                    text = $("#shape-border-line-type").attr("class");
                    break;
                case "borderLine":
                    text = $("#border-line-type").attr("class");
                    break;
                case "beginArrowStyle":
                    text = $("#begin-arrow-style-type").attr("class");
                    break;
                case "endArrowStyle":
                    text = $("#end-arrow-style-type").attr("class");
                    break;
            }
            $("div.image", $target).removeClass("fa-check");
            $("div.text", $target)
                .filter(function () {
                    return $(this).find("div").attr("class") === text;
                })
                .siblings("div.image")
                .addClass("fa fa-check");
            $("div.image.nocheck", $target).removeClass("fa-check");
        }
        else {
            var text = $("span.display", $dropdown).text();
            $("div.image", $target).removeClass("fa-check");
            $("div.text", $target)
                .filter(function () {
                    return $(this).text() === text;
                })
                .siblings("div.image")
                .addClass("fa fa-check");
            // remove check for special items mark with nocheck class
            $("div.image.nocheck", $target).removeClass("fa-check");
        }

        $target.addClass("show");

        processEventListenerHandleClosePopup(true);
    }
}

function documentMousedownHandler(event) {
    var target = event.target,
        container = _dropdownitem || _colorpicker || $("#clearActionList:visible")[0] || $("#exportActionList:visible")[0];

    if (container) {
        if (container === target || $.contains(container, target)) {
            return;
        }

        // click on related item popup the dropdown, close it
        var dropdown = $(container).data("dropdown");
        if (dropdown && $.contains(dropdown, target)) {
            hidePopups();
            _needShow = false;
            return false;
        }
    }

    hidePopups();
    $("#passwordError").hide();
}

function hidePopups() {
    hideDropdown();
    hideColorPicker();
    hideClearActionDropDown();
    hideExportActionDropDown();
}

function hideClearActionDropDown() {
    if ($("#clearActionList:visible").length > 0) {
        $("#clearActionList").hide();
        processEventListenerHandleClosePopup(false);
    }
}

function hideExportActionDropDown() {
    if ($("#exportActionList:visible").length > 0) {
        $("#exportActionList").hide();
        processEventListenerHandleClosePopup(false);
    }
}

function hideDropdown() {
    if (_dropdownitem) {
        $(_dropdownitem).removeClass("show");
        _dropdownitem = null;
    }

    processEventListenerHandleClosePopup(false);
}

function showColorPicker() {
    if (!_needShow) {
        _needShow = true;
        return;
    }

    var MIN_TOP = 30, MIN_BOTTOM = 4;
    var $element = $(this),
        $container = $element.parent(),
        name = $container.data("name"),
        $target = $("#colorpicker");

    if ($target && !$target.hasClass("colorpicker-visible")) {
        $target.data("dropdown", this);
        // save related name for later use
        $target.data("name", name);

        var $nofill = $target.find("div.nofill-color");
        if ($container.hasClass("show-nofill-color")) {
            $nofill.show();
        } else {
            $nofill.hide();
        }

        var $opacity = $target.find("#colorpickerTransparencyContainer");
        if ($container.hasClass("show-transparency-input")) {
            getTransparency(name);
            $opacity.show();
        } else {
            $opacity.hide();
        }

        _colorpicker = $target[0];

        var $dropdown = $element,
            offset = $dropdown.offset();

        var height = $target.height(),
            top = offset.top - (height - $element.height()) / 2 + 3,   // 3 = padding (4) - border-width(1)
            yOffset = 0;

        if (top < MIN_TOP) {
            yOffset = MIN_TOP - top;
            top = MIN_TOP;
        } else {
            var $inspContainer = $(".insp-container"),
                maxTop = $inspContainer.height() + $inspContainer.offset().top;

            // adjust top when out of bottom range
            if (top + height > maxTop - MIN_BOTTOM) {
                var newTop = maxTop - MIN_BOTTOM - height;
                yOffset = newTop - top;
                top = newTop;
            }
        }

        $target.css({
            top: top,
            left: offset.left - $target.width() - 20
        });

        // v-center the pointer
        var $pointer = $target.find(".cp-pointer");
        $pointer.css({top: (height - 24) / 2 - yOffset});   // 24 = pointer height

        $target.addClass("colorpicker-visible");

        processEventListenerHandleClosePopup(true);
    }
}

function hideColorPicker() {
    if (_colorpicker) {
        $(_colorpicker).removeClass("colorpicker-visible");
        _colorpicker = null;
    }
    processEventListenerHandleClosePopup(false);
}

function itemSelected() {
    // get related dropdown item
    var dropdown = $(_dropdownitem).data("dropdown");

    hideDropdown();

    if (this.parentElement.id === "clearActionList") {
        processClearAction($(this.parentElement), $("div.text", this).data("value"));
        return;
    }

    if (this.parentElement.id === "exportActionList") {
        processExportAction($(this.parentElement), $("div.text", this).data("value"));
        return;
    }

    var sheet = spread.getActiveSheet();

    var name = $(dropdown.parentElement).data("name"),
        $text = $("div.text", this),
        dataValue = $text.data("value"),    // data-value includes both number value and string value, should pay attention when use it
        numberValue = +dataValue,
        text = $text.text(),
        value = text,
        nameValue = dataValue || text;

    var options = spread.options;

    switch (name) {
        case "scrollTip":
            options.showScrollTip = numberValue;
            break;

        case "resizeTip":
            options.showResizeTip = numberValue;
            break;

        case "fontFamily":
            setStyleFont(sheet, "font-family", false, [value], value);
            break;

        case "labelFontFamily":
            setStyleFont(sheet, "font-family", true, [value], value);
            break;

        case "fontSize":
            value += "pt";
            setStyleFont(sheet, "font-size", false, [value], value);
            break;

        case "labelFontSize":
            value += "pt";
            setStyleFont(sheet, "font-size", true, [value], value);
            break;

        case "cellLabelVisibility":
            setLabelOptions(sheet, nameValue, "visibility");
            break;

        case "cellLabelAlignment":
            setLabelOptions(sheet, nameValue, "alignment");
            break;

        case "selectionPolicy":
            sheet.selectionPolicy(numberValue);
            break;

        case "selectionUnit":
            sheet.selectionUnit(numberValue);
            break;

        case "sheetName":
            var selectedSheet = spread.sheets[numberValue];
            setCheckValue("sheetVisible", selectedSheet.visible(), {
                sheetIndex: numberValue,
                sheetName: selectedSheet.name()
            });
            break;

        case "commentFontFamily":
            _activeComment && _activeComment.fontFamily(value);
            break;

        case "commentFontSize":
            value += "pt";
            _activeComment && _activeComment.fontSize(value);
            break;

        case "commentDisplayMode":
            _activeComment && _activeComment.displayMode(numberValue);
            break;

        case "commentFontStyle":
            _activeComment && _activeComment.fontStyle(nameValue);
            break;

        case "commentFontWeight":
            _activeComment && _activeComment.fontWeight(nameValue);
            break;

        case "commentBorderStyle":
            _activeComment && _activeComment.borderStyle(nameValue);
            break;

        case "commentHorizontalAlign":
            _activeComment && _activeComment.horizontalAlign(numberValue);
            break;

        case "pictureBorderStyle":
            _activePicture && _activePicture.borderStyle(nameValue);
            break;

        case "pictureStretch":
            _activePicture && _activePicture.pictureStretch(numberValue);
            break;

        case "conditionalFormat":
            processConditionalFormatDetailSetting(nameValue);
            break;

        case "ruleType":
            updateEnumTypeOfCF(numberValue);
            break;

        case "comparisonOperator":
            processComparisonOperator(numberValue);
            break;

        case "iconSetType":
            updateIconCriteriaItems(numberValue);
            break;

        case "minType":
            processMinItems(numberValue, "minValue");
            break;

        case "midType":
            processMidItems(numberValue, "midValue");
            break;

        case "maxType":
            processMaxItems(numberValue, "maxValue");
            break;

        case "cellTypes":
            processCellTypeSetting(nameValue);
            break;

        case "validatorType":
            processDataValidationSetting(nameValue, value);
            break;

        case "numberValidatorComparisonOperator":
            processNumberValidatorComparisonOperatorSetting(numberValue);
            break;

        case "dateValidatorComparisonOperator":
            processDateValidatorComparisonOperatorSetting(numberValue);
            break;

        case "textLengthValidatorComparisonOperator":
            processTextLengthValidatorComparisonOperatorSetting(numberValue);
            break;

        case "customHighlightStyleType":
            processCustomHighlightStyleTypeSetting(numberValue);
            break;
        
        case "sparklineExType":
            $("#richTextContainer").show();
            break;

        case "richText":
            processRichTextSetting(nameValue, value);
            break;

        case "zoomSpread":
            processZoomSetting(nameValue, value);
            break;

        case "commomFormat":
            processFormatSetting(nameValue, value);
            break;

        case "borderLine":
            processBorderLineSetting(nameValue);
            break;

        case "beginArrowStyle":
        case "endArrowStyle":
           processArrowStyleSetting(name,nameValue);
           break;

        case "shapeBorder":
            processShapeBorderLineSetting(nameValue);
            break;

        case "minAxisType":
            updateManual(nameValue, "manualMin");
            break;

        case "maxAxisType":
            updateManual(nameValue, "manualMax");
            break;

        case "slicerItemSorting":
            processSlicerItemSorting(numberValue);
            break;

        case "spreadTheme":
            processChangeSpreadTheme(nameValue);
            break;

        case "resizeZeroIndicator":
            spread.options.resizeZeroIndicator = numberValue;
            break;

        case "copyPasteHeaderOptions":
            spread.options.copyPasteHeaderOptions = GC.Spread.Sheets.CopyPasteHeaderOptions[nameValue]
            break;
        case "chartSeriesIndexValue":
            changeSeriesIndex(dataValue);
            break;
        case "chartAxieType":
            changeAxieTypeIndex(nameValue);
            break;
        case "chartDataPointsValue":
            changeDataPointIndex(dataValue);
            break;
        case "qrCodeSparklineModel":
            changeModelIndex(dataValue);
            break;

        case "shapeCapType":
            changeCapTypeIndex(dataValue);
            break;

        case "shapeJoinType":
            changeJoinTypeIndex(dataValue);
            break;

        case "shapeFontSize":
            changeShapeFontSize(nameValue);
            break;

        case "shapeFontFamily":
            changeShapeFontFamily(nameValue);
            break;

        default:
            //console.log("TODO add itemSelected for ", name, value);
            break;
    }

    setDropDownText(dropdown, text);
}

function setDropDownText(container, value) {
    var refList = "#" + $(container).data("list-ref"),
        $items = $(".menu-item div.text", refList),
        $item = $items.filter(function () {
            return $(this).data("value") === value;
        });

    var text = $item.text() || value;
    $("span.display", container).text(text);
}

function setDropDownValue(container, value, host) {
    if (typeof container === "string") {
        host = host || document;

        container = $(host).find("div.insp-dropdown-list[data-name='" + container + "']");
    }

    var refList = "#" + $(container).data("list-ref");
    $("span.display", container).text($(".menu-item>div.text[data-value='" + value + "']", refList).text());
}

function setDropDownValueByIndex(container, index) {
    var refList = "#" + $(container).data("list-ref"),
        $item = $(".menu-item:eq(" + index + ") div.text", refList);

    $("span.display", container).text($item.text());

    return {text: $item.text(), value: $item.data("value")};
}

function getDropDownValue(name, host) {
    host = host || document;

    var container = $(host).find("div.insp-dropdown-list[data-name='" + name + "']"),
        refList = "#" + $(container).data("list-ref"),
        text = $("span.display", container).text();

    var value = $("div.text", refList)
        .filter(function () {
            return $(this).text() === text;
        })
        .data("value");

    return value;
}

function getDropDownText(name, host) {
    host = host || document;

    var container = $(host).find("div.insp-dropdown-list[data-name='" + name + "']"),
        refList = "#" + $(container).data("list-ref"),
        text = $("span.display", container).text();

    var value = $("div.text", refList).filter(function () {
        return $(this).text() === text;
    }).text();

    return value;
}

function colorSelected() {
    var themeColor = $(this).data("name");
    var value = $(this).css("background-color");

    var name = $(_colorpicker).data("name");
    var sheet = spread.getActiveSheet();

    $("div.color-view", $(_colorpicker).data("dropdown")).css("background-color", value);

    // No Fills need special process
    if ($(this).hasClass("auto-color-cell")) {
        if (name === "backColor") {
            value = undefined;
        }
    }

    var options = spread.options;

    spread.suspendPaint();
    switch (name) {
        case "spreadBackcolor":
            options.backColor = value;
            break;

        case "grayAreaBackcolor":
            options.grayAreaBackColor = value;
            break;

        case "cutCopyIndicatorBorderColor":
            options.cutCopyIndicatorBorderColor = value;
            break;

        case "sheetTabColor":
            sheet.options.sheetTabColor = value;
            break;

        case "frozenLineColor":
            sheet.options.frozenlineColor = value;
            break;

        case "gridlineColor":
            sheet.options.gridline.color = value;
            break;

        case "foreColor":
        case "backColor":
            setColor(sheet, name, themeColor || value);
            break;

        case "labelForeColor":
            setLabelOptions(sheet, value, "foreColor");
            break;

        case "selectionBorderColor":
            sheet.options.selectionBorderColor = value;
            break;

        case "selectionBackColor":
            // change to rgba (alpha: 0.2) to make cell content visible
            value = getRGBAColor(value, 0.2);
            sheet.options.selectionBackColor = value;
            $("div.color-view", $(_colorpicker).data("dropdown")).css("background-color", value);
            break;

        case "commentBorderColor":
            _activeComment && _activeComment.borderColor(value);
            break;

        case "commentForeColor":
            _activeComment && _activeComment.foreColor(value);
            break;

        case "commentBackColor":
            _activeComment && _activeComment.backColor(value);
            break;

        case "pictureBorderColor":
            _activePicture && _activePicture.borderColor(value);
            break;

        case "pictureBackColor":
            _activePicture && _activePicture.backColor(value);
            break;

        default:
            //console.log("TODO colorSelected", name);
            break;
    }
    spread.resumePaint();
}

function updateColorOpacity(e) {
    var transparency = e.target.value;
    var color = $("div.color-view", $(_colorpicker).data("dropdown")).css("backgroundColor");
    var rgbaColor = getRGBAColor(color, 1 - transparency);
    $("div.color-view", $(_colorpicker).data("dropdown")).css("background-color", rgbaColor);
}

function getRGBAColor(color, alpha) {
    if (color === undefined || color === null) {
        return '';
    }
    var result = color,
        prefix = "rgb(",
        rgbaPrefix = "rgba(";

    if (color.indexOf(rgbaPrefix) === 0) {
        color = color.replace(rgbaPrefix, prefix);
        color = color.substr(0, color.lastIndexOf(",")) + ")";
    }

    // get rgb color use jquery
    if (color.substr(0, 4) !== prefix) {
        var $temp = $("#setfontstyle");
        $temp.css("background-color", color);
        color = $temp.css("background-color");
    }

    // adding alpha to make rgba
    if (color.substr(0, 4) === prefix) {
        var length = color.length;
        result = "rgba(" + color.substring(4, length - 1) + ", " + alpha + ")";
    }

    return result;
}

function setColor(sheet, method, value) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount)[method](value);
    }
    sheet.resumePaint();
}

function updateShapeAlign(alignType, alignValue) {
    var shapes = getActiveShapes();
    var _applyShapeAlign = function(_shapes) {
        _shapes.forEach(function(item) {
            var shapeType = getShapeType(item);
            if(shapeType === 'shapeGroup') {
                _applyShapeAlign(item.all());
            }
            if(shapeType === 'shape') {
                var itemStyle = item.style();
                itemStyle.textFrame[alignType] = alignValue;
                item.style(itemStyle);
            }
        });
    }

    _applyShapeAlign(shapes);
}

function buttonClicked() {
    var $element = $(this),
        name = $element.data("name"),
        container;

    var sheet = spread.getActiveSheet();

    // get group
    if ((container = $element.parents(".insp-radio-button-group")).length > 0) {
        name = container.data("name");
        $element.siblings().removeClass("active");
        $element.addClass("active");
        switch (name) {
            case "vAlign":
            case "hAlign":
                setAlignment(sheet, name, $element.data("name"));
                break;
            case "shapeVAlign":
                updateShapeAlign('vAlign', GC.Spread.Sheets.VerticalAlign[$element.data("name")]);
                break;
            case "shapeHAlign":
                updateShapeAlign('hAlign', GC.Spread.Sheets.HorizontalAlign[$element.data("name")]);
                break;
        }
    } else if ($element.parents(".insp-button-group").length > 0) {
        if (!$element.hasClass("no-toggle")) {
            $element.toggleClass("active");
        }

        switch (name) {
            case "bold":
                setStyleFont(sheet, "font-weight", false, ["700", "bold"], "normal");
                break;
            case "labelBold":
                setStyleFont(sheet, "font-weight", true, ["700", "bold"], "normal");
                break;
            case "italic":
                setStyleFont(sheet, "font-style", false, ["italic"], "normal");
                break;
            case "labelItalic":
                setStyleFont(sheet, "font-style", true, ["italic"], "normal");
                break;
            case "underline":
                setTextDecoration(sheet, spreadNS.TextDecorationType.underline);
                var doubleUnderlineElem = $('#cellTab span.font-double-underline');
                if (doubleUnderlineElem.hasClass('active')) {
                    doubleUnderlineElem.removeClass('active');
                    setTextDecoration(sheet, spreadNS.TextDecorationType.doubleUnderline);
                }
                break;
            case "strikethrough":
                setTextDecoration(sheet, spreadNS.TextDecorationType.lineThrough);
                break;
            case "overline":
                setTextDecoration(sheet, spreadNS.TextDecorationType.overline);
                break;
            case "double-underline":
                setTextDecoration(sheet, spreadNS.TextDecorationType.doubleUnderline);
                var underlineElem = $('#cellTab span.font-underline');
                if (underlineElem.hasClass('active')) {
                    underlineElem.removeClass('active');
                    setTextDecoration(sheet, spreadNS.TextDecorationType.underline);
                }
                break;
            case "increaseIndent":
                setTextIndent(sheet, 1);
                break;

            case "decreaseIndent":
                setTextIndent(sheet, -1);
                break;

            case "percentStyle":
                setFormatter(uiResource.cellTab.format.percentValue);
                break;

            case "commaStyle":
                setFormatter(uiResource.cellTab.format.commaValue);
                break;

            case "increaseDecimal":
                increaseDecimal();
                break;

            case "decreaseDecimal":
                decreaseDecimal();
                break;

            case "comment-underline":
            case "comment-overline":
            case "comment-strikethrough":
                setCommentTextDecoration(+$element.data("value"));
                break;
            case "verticalText":
                setVerticalText(sheet);
                break;

            default:
                //console.log("buttonClicked w/o process code for ", name);
                break;
        }
    }
}

function setCommentTextDecoration(flag) {
    if (_activeComment) {
        var textDecoration = _activeComment.textDecoration();
        _activeComment.textDecoration(textDecoration ^ flag);
    }
}

// Increase Decimal related items
function increaseDecimal() {
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    for (var p = 0; p < sheet.getSelections().length; p++) {
        var selectCells = sheet.getSelections()[p];
        var defaultActiveCell = sheet.getCell(selectCells.row, selectCells.col);
        var defaultFormatter = defaultActiveCell.formatter();
        var defaultText = defaultActiveCell.value();
        var i;
        if (defaultText !== undefined && defaultText !== null) {
            zero = "0";
            numberSign = "#";
            decimalPoint = ".";
            zeroPointZero = "0" + decimalPoint + "0";

            var scientificNotationCheckingFormatter = getScientificNotationCheckingFormattter(defaultFormatter);
            if (!defaultFormatter || (defaultFormatter === "General" || (scientificNotationCheckingFormatter && (scientificNotationCheckingFormatter.indexOf("E") >= 0 || scientificNotationCheckingFormatter.indexOf('e') >= 0)))) {
                scientificNotationCheckingFormatter = zeroPointZero;
                if ((!isNaN(defaultText)) && ((defaultText + "").split(".").length > 1)) {
                    var afterPointZero = (defaultText + "").split(".")[1].length;
                    for (var m = 0; m < afterPointZero; m++) {
                        scientificNotationCheckingFormatter = scientificNotationCheckingFormatter + "0";
                    }
                }
            } else {
                formatString = defaultFormatter;
                var formatters = formatString.split(';');
                for (i = 0; i < formatters.length && i < 2; i++) {
                    if (formatters[i] && formatters[i].indexOf("/") < 0 && formatters[i].indexOf(":") < 0 && formatters[i].indexOf("?") < 0) {
                        var indexOfDecimalPoint = formatters[i].lastIndexOf(decimalPoint);
                        if (indexOfDecimalPoint !== -1) {
                            formatters[i] = formatters[i].slice(0, indexOfDecimalPoint + 1) + zero + formatters[i].slice(indexOfDecimalPoint + 1);
                        } else {
                            var indexOfZero = formatters[i].lastIndexOf(zero);
                            var indexOfNumberSign = formatters[i].lastIndexOf(numberSign);
                            var insertIndex = indexOfZero > indexOfNumberSign ? indexOfZero : indexOfNumberSign;
                            if (insertIndex >= 0) {
                                formatters[i] = formatters[i].slice(0, insertIndex + 1) + decimalPoint + zero + formatters[i].slice(insertIndex + 1);
                            }
                        }
                    }
                }
                formatString = formatters.join(";");
                scientificNotationCheckingFormatter = formatString;
            }
            for (var r = selectCells.row; r < selectCells.rowCount + selectCells.row; r++) {
                for (var c = selectCells.col; c < selectCells.colCount + selectCells.col; c++) {
                    var style = sheet.getActualStyle(r, c);
                    style.formatter = scientificNotationCheckingFormatter;
                    sheet.setStyle(r, c, style);
                }
            }
        }
    }
    sheet.resumePaint();
}

//This method is used to get the formatter which not include the string and color
//in order to not misleading with the charactor 'e' / 'E' in scientific notation.
function getScientificNotationCheckingFormattter(formatter) {
    if (!formatter) {
        return formatter;
    }
    var i;
    var signalQuoteSubStrings = getSubStrings(formatter, '\'', '\'');
    for (i = 0; i < signalQuoteSubStrings.length; i++) {
        formatter = formatter.replace(signalQuoteSubStrings[i], '');
    }
    var doubleQuoteSubStrings = getSubStrings(formatter, '\"', '\"');
    for (i = 0; i < doubleQuoteSubStrings.length; i++) {
        formatter = formatter.replace(doubleQuoteSubStrings[i], '');
    }
    var colorStrings = getSubStrings(formatter, '[', ']');
    for (i = 0; i < colorStrings.length; i++) {
        formatter = formatter.replace(colorStrings[i], '');
    }
    return formatter;
}

function getSubStrings(source, beginChar, endChar) {
    if (!source) {
        return [];
    }
    var subStrings = [], tempSubString = '', inSubString = false;
    for (var index = 0; index < source.length; index++) {
        if (!inSubString && source[index] === beginChar) {
            inSubString = true;
            tempSubString = source[index];
            continue;
        }
        if (inSubString) {
            tempSubString += source[index];
            if (source[index] === endChar) {
                subStrings.push(tempSubString);
                tempSubString = "";
                inSubString = false;
            }
        }
    }
    return subStrings;
}
// Increase Decimal related items (end)

// Decrease Decimal related items
function decreaseDecimal() {
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    for (var p = 0; p < sheet.getSelections().length; p++) {
        var selectCells = sheet.getSelections()[p];
        var defaultActiveCell = sheet.getCell(selectCells.row, selectCells.col);
        var defaultFormatter = defaultActiveCell.formatter();
        var defaultText = defaultActiveCell.value();
        decimalPoint = ".";
        var i;
        if (defaultText !== undefined && defaultText !== null) {
            var formatString = null;
            if (!defaultFormatter || defaultFormatter === "General") {
                if (!isNaN(defaultText)) {
                    var result = defaultText.split('.');
                    if (result.length === 2) {
                        result[0] = "0";
                        var isScience = false;
                        var sb = "";
                        for (i = 0; i < result[1].length - 1; i++) {
                            if ((i + 1 < result[1].length) && (result[1].charAt(i + 1) === 'e' || result[1].charAt(i + 1) === 'E')) {
                                isScience = true;
                                break;
                            }
                            sb = sb + ('0');
                        }

                        if (isScience) {
                            sb = sb + ("E+00");
                        }
                        result[1] = sb.toString();
                        formatString = result[0] + (result[1] !== "" ? decimalPoint + result[1] : "");
                    }
                }
            } else {
                formatString = defaultFormatter;
                var formatters = formatString.split(';');
                for (i = 0; i < formatters.length && i < 2; i++) {
                    if (formatters[i] && formatters[i].indexOf("/") < 0 && formatters[i].indexOf(":") < 0 && formatters[i].indexOf("?") < 0) {
                        var indexOfDecimalPoint = formatters[i].lastIndexOf(decimalPoint);
                        if (indexOfDecimalPoint !== -1 && indexOfDecimalPoint + 1 < formatters[i].length) {
                            formatters[i] = formatters[i].slice(0, indexOfDecimalPoint + 1) + formatters[i].slice(indexOfDecimalPoint + 2);
                            var tempString = indexOfDecimalPoint + 1 < formatters[i].length ? formatters[i].substr(indexOfDecimalPoint + 1, 1) : "";
                            if (tempString === "" || tempString !== "0") {
                                formatters[i] = formatters[i].slice(0, indexOfDecimalPoint) + formatters[i].slice(indexOfDecimalPoint + 1);
                            }
                        } else {
                            //do nothing.
                        }
                    }
                }
                formatString = formatters.join(";");
            }
            for (var r = selectCells.row; r < selectCells.rowCount + selectCells.row; r++) {
                for (var c = selectCells.col; c < selectCells.colCount + selectCells.col; c++) {
                    var style = sheet.getActualStyle(r, c);
                    style.formatter = formatString;
                    sheet.setStyle(r, c, style);
                }
            }
        }
    }
    sheet.resumePaint();
}
// Decrease Decimal related items (end)

function setAlignment(sheet, type, value) {
    var sels = sheet.getSelections(),
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount(),
        align;

    value = value.toLowerCase();

    if (value === "middle") {
        value = "center";
    }

    if (type === "hAlign") {
        align = spreadNS.HorizontalAlign[value];
    } else {
        align = spreadNS.VerticalAlign[value];
    }

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount)[type](align);
    }
    sheet.resumePaint();
}

function setTextDecoration(sheet, flag) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount),
            textDecoration = sheet.getCell(sel.row, sel.col).textDecoration();
        if ((textDecoration & flag) === flag) {
            textDecoration = textDecoration - flag;
        } else {
            textDecoration = textDecoration | flag;
        }
        sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount).textDecoration(textDecoration);
    }
    sheet.resumePaint();
}

function setWordWrap(sheet) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount),
            wordWrap = !sheet.getCell(sel.row, sel.col).wordWrap(),
            startRow = sel.row,
            endRow = sel.row + sel.rowCount - 1;

        sheet.getRange(startRow, sel.col, sel.rowCount, sel.colCount).wordWrap(wordWrap);

        for (var row = startRow; row <= endRow; row++) {
            sheet.autoFitRow(row);
        }
    }
    sheet.resumePaint();
}
function setVerticalText(sheet) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount),
            isVerticalText = !sheet.getCell(sel.row, sel.col).isVerticalText(),
            startRow = sel.row,
            endRow = sel.row + sel.rowCount - 1;

        sheet.getRange(startRow, sel.col, sel.rowCount, sel.colCount).isVerticalText(isVerticalText);
    }
    sheet.resumePaint();
}
function setTextIndent(sheet, step) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount),
            indent = sheet.getCell(sel.row, sel.col).textIndent();

        if (isNaN(indent)) {
            indent = 0;
        }

        var value = indent + step;
        if (value < 0) {
            value = 0;
        }
        sheet.getRange(sel.row, sel.col, sel.rowCount, sel.colCount).textIndent(value);
    }
    sheet.resumePaint();
}

function divButtonClicked() {
    var sheet = spread.getActiveSheet(),
        id = this.id;

    spread.suspendPaint();
    switch (id) {
        case "mergeCells":
            mergeCells(sheet);
            updateMergeButtonsState();
            break;

        case "unmergeCells":
            unmergeCells(sheet);
            updateMergeButtonsState();
            break;

        case "freezePane":
            sheet.frozenRowCount(sheet.getActiveRowIndex());
            sheet.frozenColumnCount(sheet.getActiveColumnIndex());
            syncForzenProperties(sheet);
            break;

        case "unfreeze":
            sheet.frozenRowCount(0);
            sheet.frozenColumnCount(0);
            sheet.frozenTrailingRowCount(0);
            sheet.frozenTrailingColumnCount(0);
            syncForzenProperties(sheet);
            break;

        case "sortAZ":
        case "sortZA":
            sortData(sheet, id === "sortAZ");
            break;

        case "filter":
            updateFilter(sheet);
            break;

        case "group":
            addGroup(sheet);
            break;

        case "ungroup":
            removeGroup(sheet);
            break;

        case "showDetail":
            toggleGroupDetail(sheet, true);
            break;

        case "hideDetail":
            toggleGroupDetail(sheet, false);
            break;

        case "groupShape":
            setShapeGroup("group", sheet);
            break;
        case "unGroupShape":
            setShapeGroup("ungroup", sheet);
            break;

        case "add":
        case "remove":

        default:
            //console.log("TODO add code for ", id);
            break;
    }
    spread.resumePaint();
}

function mergeCells(sheet) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        sheet.addSpan(sel.row, sel.col, sel.rowCount, sel.colCount);
    }
}

function unmergeCells(sheet) {
    function removeSpan(range) {
        sheet.removeSpan(range.row, range.col);
    }

    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        sheet.getSpans(sel).forEach(removeSpan);
    }
}

function sortData(sheet, ascending) {
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        sheet.sortRange(sel.row, sel.col, sel.rowCount, sel.colCount, true,
            [
                {index: sel.col, ascending: ascending}
            ]);
    }
}

function updateFilter(sheet) {
    if (sheet.rowFilter()) {
        sheet.rowFilter(null);
    } else {
        var sels = sheet.getSelections();
        if (sels.length > 0) {
            var sel = sels[0];
            sheet.rowFilter(new spreadNS.Filter.HideRowFilter(sel));
        }
    }
}

function setCheckboxEnable($element, enable) {
    if (enable) {
        $element.removeClass("disabled");
        $element.find(".button").addClass("checked");
    } else {
        $element.addClass("disabled");
    }
}

function addGroup(sheet) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1) // row selection
    {
        spread.commandManager().execute({
            cmd: 'outlineRow',
            sheetName: sheet.name(),
            index: sel.row,
            count: sel.rowCount
        });
    }
    else if (sel.row === -1) // column selection
    {
        spread.commandManager().execute({
            cmd: 'outlineColumn',
            sheetName: sheet.name(),
            index: sel.col,
            count: sel.colCount
        });
    }
    else // cell range selection
    {
        alert(getResource("messages.rowColumnRangeRequired"));
    }
}

function removeGroup(sheet) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1 && sel.row === -1) // sheet selection
    {
        sheet.rowOutlines.ungroup(0, sheet.getRowCount());
        sheet.columnOutlines.ungroup(0, sheet.getColumnCount());
    }
    else if (sel.col === -1) // row selection
    {
        spread.commandManager().execute({
            cmd: 'removeRowOutline',
            sheetName: sheet.name(),
            index: sel.row,
            count: sel.rowCount
        });
    }
    else if (sel.row === -1) // column selection
    {
        spread.commandManager().execute({
            cmd: 'removeColumnOutline',
            sheetName: sheet.name(),
            index: sel.col,
            count: sel.colCount
        });
    }
    else // cell range selection
    {
        alert(getResource("messages.rowColumnRangeRequired"));
    }
}

function addMenu(){
    if (isMenuItemExist(spread.contextMenu.menuData, "editRichText")) {
        spread.contextMenu.menuData.forEach(function (item, index) {
            if (item && item.name === "editRichText") {
                spread.contextMenu.menuData.splice(index, 1);
            }
        });
        return;
    }
    var commandManager = spread.commandManager();
    var editRichTextInfo = {
        text: "Edit Rich Text",
        name: "editRichText",
        workArea: "viewport",
        command: "editRichText"
    };
    spread.contextMenu.menuData.push(editRichTextInfo);
    var editRichTextCommand = {
        canUndo: false,
        execute: function (spread, options) {
            var RICHTEXT_DIALOG_WIDTH = 500;
            showModal(uiResource.richTextDialog.editRichText, RICHTEXT_DIALOG_WIDTH, $("#richtextdialog").children(), addRichTextEvent);
        }
    };
    commandManager.register("editRichText", editRichTextCommand, null, false, false, false, false);
    function CustomMenuView() {
    }

    CustomMenuView.prototype = new GC.Spread.Sheets.ContextMenu.MenuView();
    spread.contextMenu.menuView = new CustomMenuView();
}

function isMenuItemExist(menuData, menuItemName) {
    var i = 0, count = menuData.length;
    for (; i < count; i++) {
        if (menuItemName === menuData[i].name) {
            return true;
        }
    }
}

function addRichTextEvent() {
    var spread = $("#ss").data("workbook");
    var sheet = spread.getActiveSheet();
    var richText = getRichText();
    if (richText.length > 0) {
        sheet.setValue(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex(), {
            richText: richText
        });
    }
}

function toggleGroupDetail(sheet, expand) {
    var sels = sheet.getSelections();
    var sel = sels[0];

    if (!sel) return;

    if (sel.col === -1 && sel.row === -1) // sheet selection
    {
    }
    else if (sel.col === -1) // row selection
    {
        for (var i = 0; i < sel.rowCount; i++) {
            var rgi = sheet.rowOutlines.find(sel.row + i, 0);
            if (rgi) {
                sheet.rowOutlines.expand(rgi.level, expand);
            }
        }
    }
    else if (sel.row === -1) // column selection
    {
        for (var i = 0; i < sel.colCount; i++) {
            var rgi = sheet.columnOutlines.find(sel.col + i, 0);
            if (rgi) {
                sheet.columnOutlines.expand(rgi.level, expand);
            }
        }
    }
    else // cell range selection
    {
    }
}

function adjustSpreadSize() {
    var statusBarHeight = $("#statusBar").height();
    var containerBottomBorderWidth = parseInt($("#inner-content-container").css('border-bottom-width'));
    var height = $("#inner-content-container").height() - $("#formulaBar").height() - statusBarHeight + containerBottomBorderWidth,
        spreadHeight = $("#ss").height();

    if (spreadHeight !== height) {
        $("#controlPanel").height(height);
        $("#ss").height(height);
        $("#ss").data("workbook").refresh();
    }
}

function screenAdoption() {
    adjustSpreadSize();

    // adjust toolbar items position
    var $toolbar = $("#toolbar"),
        sectionWidth = Math.floor($toolbar.width() / 3);

    $(".toolbar-left-section", $toolbar).width(sectionWidth);
    var $middle = $(".toolbar-middle-section", $toolbar);
    // + 2 to make sure the right section with enough space to show in same line
    if (sectionWidth > 375 + 5) {  // 340 = (380 + 300) / 2, where 380 is min-width of left section, 300 is the width of right section
        $middle.width(sectionWidth);
        $middle.css("display", "inline-block");
    } else if (sectionWidth < 244) {
        $middle.css("display", "none");
    } else {
        $middle.width("auto");
        $middle.css("display", "inline-block");
    }
    // explicit set formula box' width instead of 100% because it's contained in table
    var width = $("#inner-content-container").width() - $("#positionbox").outerWidth() - 1; // 1: border' width of td contains formulabox (left only)
    $("#formulabox").css({ width: width });
}

function doPrepareWork() {
    /*
     1. expand / collapse .insp-group by checking expanded class
     */
    function processDisplayGroups() {
        $("div.insp-group").each(function () {
            var $group = $(this),
                expanded = $group.hasClass("expanded"),
                $content = $group.find("div.insp-group-content"),
                $state = $group.find("span.group-state");

            if (expanded) {
                $content.show();
                $state.addClass("fa-caret-down");
            } else {
                $content.hide();
                $state.addClass("fa-caret-right");
            }
        });
    }

    function addEventHandlers() {
        $("div.insp-group-title>span").click(toggleState);
        $("div.insp-checkbox").click(checkedChanged);
        $("div.insp-number>input.editor").blur(updateNumberProperty);
        $("div.insp-dropdown-list .dropdown").click(showDropdown);
        $("div.insp-menu .menu-item").click(itemSelected);
        $("div.insp-color-picker .picker").click(showColorPicker);
        $("li.color-cell").click(colorSelected);
        $("#colorpickerTransparency").change(updateColorOpacity);
        $(".insp-button-group span.btn").click(buttonClicked);
        $(".insp-radio-button-group span.btn").click(buttonClicked);
        $(".insp-buttons .btn").click(divButtonClicked);
        $(".insp-text input.editor").blur(updateStringProperty);
    }

    processDisplayGroups();

    addEventHandlers();

    $("input[type='number']:not('.not-min-zero')").attr("min", 0);

    // set default values
    var item = setDropDownValueByIndex($("#conditionalFormatType"), -1);
    processConditionalFormatDetailSetting(item.value, true);
    var cellTypeItem = setDropDownValueByIndex($("#cellTypes"), -1);
    processCellTypeSetting(cellTypeItem.value, true);                     // CellType Setting
    var validationTypeItem = setDropDownValueByIndex($("#validatorType"), 0);
    processDataValidationSetting(validationTypeItem.value);         // Data Validation Setting
    var sparklineTypeItem = setDropDownValueByIndex($("#sparklineExTypeDropdown"), 0);
    processSparklineSetting(sparklineTypeItem.value);               // SparklineEx Setting

    setDropDownValue("numberValidatorComparisonOperator", 0);       // NumberValidator Comparison Operator
    processNumberValidatorComparisonOperatorSetting(0);
    setDropDownValue("dateValidatorComparisonOperator", 0);         // DateValidator Comparison Operator
    processDateValidatorComparisonOperatorSetting(0);
    setDropDownValue("textLengthValidatorComparisonOperator", 0);   // TextLengthValidator Comparison Operator
    processTextLengthValidatorComparisonOperatorSetting(0);
    setDropDownValue("customHighlightStyleType", 0);                // CustomHighlightStyleType Comparison Operator
    processCustomHighlightStyleTypeSetting(0);
    setDropDownValue("dogearPosition", 0);                          // CustomHighlightStyleDogearPosition Comparison Operator
    setDropDownValue("iconPosition", 4);                            // CustomHighlightStyleIconPosition Comparison Operator
    processBorderLineSetting("thin");                               // Border Line Setting
    processArrowStyleSetting('beginArrowStyle','none');
    processArrowStyleSetting('endArrowStyle','none');
    // processShapeBorderLineSetting('solid');

    setDropDownValue("minType", 1);                                 // LowestValue
    setDropDownValue("midType", 4);                                 // Percentile
    setDropDownValue("maxType", 2);                                 // HighestValue
    setDropDownValue("minimumType", 5);                             // Automin
    setDropDownValue("maximumType", 7);                             // Automax
    setDropDownValue("dataBarDirection", 0);                        // Left-to-Right
    setDropDownValue("axisPosition", 0);                            // Automatic
    setDropDownValue("iconSetType", 0);                             // ThreeArrowsColored
    setDropDownValue("checkboxCellTypeTextAlign", 3);               // Right
    setDropDownValue("comboboxCellTypeEditorValueType", 2);         // Value
    setDropDownValue("errorAlert", 0);                              // Data Validation Error Alert Type
    setDropDownValue("zoomSpread", 1);                              // Zoom Value
    setDropDownValueByIndex($("#commomFormatType"), 0);             // Format Setting
    setDropDownValueByIndex($("#boxplotClassType"), 0);             // BoxPlotSparkline Class
    setDropDownValue("boxplotSparklineStyleType", 0);               // BoxPlotSparkline Style
    setDropDownValue("dataOrientationType", 0);                     // CompatibleSparkline DataOrientation
    setDropDownValue("paretoLabelList", 0);                         // ParetoSparkline Label
    setDropDownValue("spreadSparklineStyleType", 4);                // SpreadSparkline Style
    setDropDownValue("stackedSparklineTextOrientation", 0);         // StackedSparkline TextOrientation
    setDropDownValueByIndex($("#spreadTheme"), 1);                  // Spread Theme
    setDropDownValue("resizeZeroIndicator", 1);                     // ResizeZeroIndicator
    setDropDownValueByIndex($("#copyPasteHeaderOptions"), 3);       // CopyPasteHeaderOptins
    setDropDownValueByIndex($("#cellLabelVisibility"), 0);          // CellLabelVisibility
    setDropDownValueByIndex($("#cellLabelAlignment"), 0);           // CellLabelAlignment
    conditionalFormatTexts = uiResource.conditionalFormat.texts;
}

function initSpread() {
    //formulabox
    fbx = new spreadNS.FormulaTextBox.FormulaTextBox(document.getElementById('formulabox'));
    fbx.workbook(spread);

    setCellContent();
    setFormulaContent();
    setConditionalFormatContent();
    setTableContent();
    setSparklineContent();
    setCommentContent();
    setPictureContent();
    setDataContent();
    setSlicerContent();
    addChartContent();
    addBarCodeConent();
    addShapeConent();
}

// Sample Content related items
function setFormulaContent() {
    var sheet = new spreadNS.Worksheet("Formula");
    spread.addSheet(spread.getSheetCount(), sheet);

    sheet.suspendPaint();
    sheet.setColumnCount(50);

    sheet.setColumnWidth(0, 100);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(6, 103);
    sheet.setColumnWidth(8, 150);

    var row = 1, col = 2;                                       // basic function
    sheet.getCell(row, 0).value("Basic Function").font("700 11pt Calibri");
    sheet.getCell(row, col).value("Name");
    sheet.getCell(row, ++col).value("Age");
    row++, col = 2;
    sheet.getCell(row, col).value("Jack").hAlign(spreadNS.HorizontalAlign.right);
    sheet.getCell(row, ++col).value(17);
    row++, col = 2;
    sheet.getCell(row, col).value("Lily").hAlign(spreadNS.HorizontalAlign.right);
    sheet.getCell(row, ++col).value(23);
    row++, col = 2;
    sheet.getCell(row, col).value("Bob").hAlign(spreadNS.HorizontalAlign.right);
    sheet.getCell(row, ++col).value(30);
    row++, col = 2;
    sheet.getCell(row, col).value("Mary").hAlign(spreadNS.HorizontalAlign.right);
    sheet.getCell(row, ++col).value(25);
    row++, col = 2;
    sheet.getCell(row, col).value("Average Age:");
    sheet.getCell(row, ++col).formula("=AVERAGE(D3:D6)");
    row++, col = 2;
    sheet.getCell(row, col).value("Max Age:");
    sheet.getCell(row, ++col).formula("=MAX(D3:D6)");
    row++, col = 2;
    sheet.getCell(row, col).value("Min Age:");
    sheet.getCell(row, ++col).formula("=MIN(D3:D6)");

    row = 1, col = 8;                                           // indirect function
    sheet.getCell(row, 6).value("Indirect Function").font("700 11pt Calibri");
    sheet.getCell(row, col).value("J2");
    sheet.getCell(row, ++col).value(1);
    row++, col = 8;
    sheet.getCell(row, col).value("I");
    sheet.getCell(row, ++col).value(2);
    row++, col = 8;
    sheet.getCell(row, col).value("J");
    sheet.getCell(row, ++col).value(3);
    row = row + 2, col = 8;
    var formulaStr = "=INDIRECT(\"I2\")";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);
    row++, col = 8;
    formulaStr = "=INDIRECT(I2)";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);
    row++, col = 8;
    formulaStr = "=INDIRECT(\"I\"&(1+2))";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);
    row++, col = 8;
    formulaStr = "=INDIRECT(I4&J3)";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);
    row++, col = 8;
    formulaStr = "=INDIRECT(\"" + sheet.name() + "!\"&I2)";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);
    row++, col = 8;
    formulaStr = "=INDIRECT(\"" + sheet.name() + "!I2\")";
    sheet.getCell(row, col).value(formulaStr);
    sheet.getCell(row, ++col).formula(formulaStr);

    row = 12;                                                   // array formula
    sheet.getCell(row, 0).value("Array Formula").font("700 11pt Calibri");
    sheet.addSpan(row, 2, 1, 6);
    sheet.getCell(row, 2).value("Calculation");
    sheet.setArray(13, 2, [
        ["", "Match", "Physical", "Chemistry", "", "Sum"],
        ["Alice", 97, 61, 53],
        ["John", 65, 76, 65],
        ["Bob", 55, 70, 64],
        ["Jack", 89, 77, 73]
    ]);
    sheet.setArrayFormula(14, 7, 4, 1, "=SUBTOTAL(9,OFFSET($D$15,ROW($D$15:$D$18)-ROW($D$15),,1,3))");

    row = 19;
    sheet.addSpan(row, 2, 1, 6);
    sheet.getCell(row, 2).value("Search");
    sheet.setArray(20, 2, [
        ["apple", "apple"],
        ["banana", "pear"],
        ["pear", "potato"],
        ["tomato", "potato"],
        ["potato", "dumpling"],
        ["cake"],
        ["noodel"]
    ]);
    sheet.addSpan(20, 6, 1, 5);
    sheet.getCell(20, 6).value("Find out the first value on D21:D25 that doesn't contain on D21:D27");
    sheet.addSpan(22, 6, 1, 2);
    sheet.getCell(22, 6).value("ArrayFormula Result:");
    sheet.addSpan(23, 6, 1, 2);
    sheet.getCell(23, 6).value("NomalFormula Result:");
    sheet.setArrayFormula(22, 8, 1, 1, "=INDEX(D21:D25,MATCH(TRUE,ISNA(MATCH(D21:D25,C21:C27,0)),0))");
    sheet.setFormula(23, 8, "=INDEX(D21:D25,MATCH(TRUE,ISNA(MATCH(D21:D25,C21:C27,0)),0))");

    row = 28;
    sheet.addSpan(row, 2, 1, 6);
    sheet.getCell(row, 2).value("Statistics");
    sheet.setArray(29, 2, [
        ["Product", "Salesman", "Units Sold"],
        ["Fax", "Brown", 1],
        ["Phone", "Smith", 10],
        ["Fax", "Jones", 20],
        ["Fax", "Smith", 30],
        ["Phone", "Jones", 40],
        ["PC", "Smith", 50],
        ["Fax", "Brown", 60],
        ["Phone", "Davis", 70],
        ["PC", "Jones", 80]
    ]);
    sheet.addSpan(29, 6, 1, 4);
    sheet.getCell(29, 6).value("Summing Sales: Faxes Sold By Brown");
    sheet.setArrayFormula(30, 6, 1, 1, "=SUM((C31:C39=\"Fax\")*(D31:D39=\"Brown\")*(E31:E39))");
    sheet.addSpan(31, 6, 1, 4);
    sheet.getCell(31, 6).value("Logical AND (Faxes And Brown)");
    sheet.setArrayFormula(32, 6, 1, 1, "=SUM((C31:C39=\"Fax\")*(D31:D39=\"Brown\"))");
    sheet.addSpan(33, 6, 1, 4);
    sheet.getCell(33, 6).value("Logical OR (Faxes Or Jones)");
    sheet.setArrayFormula(34, 6, 1, 1, "=SUM(IF((C31:C39=\"Fax\")+(D31:D39=\"Jones\"),1,0))");
    sheet.addSpan(35, 6, 1, 4);
    sheet.getCell(35, 6).value("Logical XOR (Fax Or Jones but not both)");
    sheet.setArrayFormula(36, 6, 1, 1, "=SUM(IF(MOD((C31:C39=\"Fax\")+(D31:D39=\"Jones\"),2),1,0))");
    sheet.addSpan(37, 6, 1, 4);
    sheet.getCell(37, 6).value("Logical NAND (All Sales Except Fax And Jones)");
    sheet.setArrayFormula(38, 6, 1, 1, "=SUM(IF((C31:C39=\"Fax\")+(D31:D39=\"Jones\")<>2,1,0))");

    sheet.resumePaint();
}

function setCellContent() {
    var sheet = new spreadNS.Worksheet("Cell");
    spread.removeSheet(0);
    spread.addSheet(spread.getSheetCount(), sheet);

    sheet.suspendPaint();
    sheet.setColumnCount(50);

    sheet.setColumnWidth(0, 100);
    sheet.setColumnWidth(1, 20);
    for (var col = 2; col < 11; col++) {
        sheet.setColumnWidth(col, 88);
    }

    var Range = spreadNS.Range;
    var row = 1, col = 0;                               // cell background
    sheet.getCell(row, col).value("Background").font("700 11pt Calibri");
    sheet.getCell(row, col + 2).backColor("#1E90FF");
    sheet.getCell(row, col + 4).backColor("#00ff00");

    row = row + 2;                                      // line border
    var borderColor = "red";
    var lineStyle = spreadNS.LineStyle;
    var lineBorder = spreadNS.LineBorder;
    var option = {all: true};
    sheet.getCell(row, 0).value("Border").font("700 11pt Calibri");
    col = 1;
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.empty), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.hair), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dotted), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashDotDot), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashDot), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.dashed), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.thin), option);
    row = row + 2, col = 1;
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashDotDot), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.slantedDashDot), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashDot), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.mediumDashed), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.medium), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.thick), option);
    sheet.getRange(row, ++col, 1, 1).setBorder(new lineBorder(borderColor, lineStyle.double), option);
    row = row + 2, col = 1;
    sheet.getRange(row, ++col, 2, 2).setBorder(new lineBorder("blue", lineStyle.dashed), {all: true});
    sheet.getRange(row, col + 3, 2, 2).setBorder(new lineBorder("yellowgreen", lineStyle.double), {outline: true});
    sheet.getRange(row, col + 6, 2, 2).setBorder(new lineBorder("black", lineStyle.mediumDashed), {innerHorizontal: true});
    sheet.getRange(row, col + 6, 2, 2).setBorder(new lineBorder("black", lineStyle.slantedDashDot), {innerVertical: true});
    row = row + 3, col = 2;
    sheet.getRange(row, col, 3, 2).setBorder(new lineBorder("lightgreen", lineStyle.thick), {outline: true});
    sheet.getRange(row, col, 3, 2).setBorder(new lineBorder("lightgreen", lineStyle.thick), {innerHorizontal: true});
    col = col + 3;
    sheet.getRange(row, col, 3, 3).setBorder(new lineBorder("#CDCD00", lineStyle.thick), {outline: true});
    sheet.getRange(row, col, 3, 3).setBorder(new lineBorder("#CDCD00", lineStyle.thick), {innerVertical: true});

    row = row + 3, col = 1;                             // merge cell
    sheet.getCell(row + 1, 0).value("Span").font("700 11pt Calibri");
    sheet.addSpan(row + 1, ++col, 1, 2);
    sheet.addSpan(row, col + 3, 3, 1);
    sheet.addSpan(row, col + 5, 3, 2);

    row = row + 4, col = 1;                             // font
    var TextDecorationType = spreadNS.TextDecorationType;
    var fontText = "SPREADJS";
    sheet.getCell(row, 0).value("Font").font("700 11pt Calibri");
    sheet.getCell(row, ++col).value(fontText);
    sheet.getCell(row, ++col).value(fontText).font("13pt Calibri");
    sheet.getCell(row, ++col).value(fontText).font("11pt Arial");
    sheet.getCell(row, ++col).value(fontText).font("13pt Times New Roman");
    sheet.getCell(row, ++col).value(fontText).backColor("#FFD700");
    sheet.getCell(row, ++col).value(fontText).foreColor("#436EEE");
    row = row + 2, col = 1;
    sheet.getCell(row, ++col).value(fontText).foreColor("#FFD700").backColor("#436EEE");
    sheet.getCell(row, ++col).value(fontText).font("700 11pt Calibri");
    sheet.getCell(row, ++col).value(fontText).font("italic 11pt Calibri");
    sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.underline);
    sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.lineThrough);
    sheet.getCell(row, ++col).value(fontText).textDecoration(TextDecorationType.overline);

    row = row + 2, col = 1;                             // format
    var number = 0.25;
    sheet.getCell(row, 0).value("Format").font("700 11pt Calibri");
    sheet.getCell(row, ++col).value(number).formatter("0.00");
    sheet.getCell(row, ++col).value(number).formatter("$#,##0.00");
    sheet.getCell(row, ++col).value(number).formatter("$ #,##0.00;$ (#,##0.00);$ \"-\"??;@");
    sheet.getCell(row, ++col).value(number).formatter("0%");
    sheet.getCell(row, ++col).value(number).formatter("# ?/?");
    row = row + 2, col = 1;
    sheet.getCell(row, ++col).value(number).formatter("0.00E+00");
    sheet.getCell(row, ++col).value(number).formatter("@");
    sheet.getCell(row, ++col).value(number).formatter("h:mm:ss AM/PM");
    sheet.getCell(row, ++col).value(number).formatter("m/d/yyyy");
    sheet.getCell(row, ++col).value(number).formatter("dddd, mmmm dd, yyyy");

    row = row + 2, col = 1;                             // text alignment
    var HorizontalAlign = spreadNS.HorizontalAlign;
    var VerticalAlign = spreadNS.VerticalAlign;
    sheet.setRowHeight(row, 60);
    sheet.getCell(row, 0).value("Alignment").font("700 11pt Calibri");
    sheet.getCell(row, ++col).value("Top Left").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.left);
    sheet.getCell(row, ++col).value("Top Center").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.center);
    sheet.getCell(row, ++col).value("Top Right").vAlign(VerticalAlign.top).hAlign(HorizontalAlign.right);
    sheet.getCell(row, ++col).value("Center Left").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.left);
    sheet.getCell(row, ++col).value("Center Center").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.center);
    sheet.getCell(row, ++col).value("Center Right").vAlign(VerticalAlign.center).hAlign(HorizontalAlign.right);
    sheet.getCell(row, ++col).value("Bottom Left").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.left);
    sheet.getCell(row, ++col).value("Bottom Center").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.center);
    sheet.getCell(row, ++col).value("Bottom Right").vAlign(VerticalAlign.bottom).hAlign(HorizontalAlign.right);

    row = row + 2, col = 1;                             // lock cell
    sheet.getCell(row, 0).value("Locked").font("700 11pt Calibri");
    sheet.getCell(row, ++col).value("TRUE").locked(true);
    sheet.getCell(row, ++col).value("FALSE").locked(false);

    row = row + 2, col = 1;                             // word wrap
    sheet.setRowHeight(row, 60);
    sheet.getCell(row, 0).value("WordWrap").font("700 11pt Calibri");
    sheet.getCell(row, ++col).value("ABCDEFGHIJKLMNOPQRSTUVWXYZ").wordWrap(true);
    sheet.getCell(row, ++col).value("ABCDEFGHIJKLMNOPQRSTUVWXYZ").wordWrap(false);

    row = row + 2, col = 1;                             // celltype
    sheet.setRowHeight(row, 25);
    var cellType;
    sheet.getCell(row, 0).value("CellType").font("700 11pt Calibri");
    cellType = new spreadNS.CellTypes.Button();
    cellType.buttonBackColor("#FFFF00");
    cellType.text("I'm a button");
    sheet.getCell(row, ++col).cellType(cellType);

    cellType = new spreadNS.CellTypes.CheckBox();
    cellType.caption("caption");
    cellType.textTrue("true");
    cellType.textFalse("false");
    cellType.textIndeterminate("indeterminate");
    cellType.textAlign(spreadNS.CellTypes.CheckBoxTextAlign.right);
    cellType.isThreeState(true);
    sheet.getCell(row, ++col).cellType(cellType);

    cellType = new spreadNS.CellTypes.ComboBox();
    cellType.items(["apple", "banana", "cat", "dog"]);
    sheet.getCell(row, ++col).cellType(cellType);

    cellType = new spreadNS.CellTypes.HyperLink();
    cellType.linkColor("blue");
    cellType.visitedLinkColor("red");
    cellType.text("SpreadJS");
    cellType.linkToolTip("SpreadJS Web Site");
    sheet.getCell(row, ++col).cellType(cellType).value("http://www.grapecity.com/spreadjs/");

    row = row + 2, col = 1;                             // celltype
    sheet.setRowHeight(row, 100);
    sheet.setColumnWidth(0, 150);
    sheet.getCell(row, 0).value("CellPadding&Label").font("700 11pt Calibri");
    sheet.getCell(row, ++col, GC.Spread.Sheets.SheetArea.viewport).watermark("User ID").cellPadding('20');
    sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
        foreColor: 'red',
        visibility: 2,
        font: 'bold 15px Arial'
    });

    var b = new GC.Spread.Sheets.CellTypes.Button();
    b.text("Click Me!");
    sheet.setColumnWidth(3, 200);
    sheet.setCellType(row, ++col, b, GC.Spread.Sheets.SheetArea.viewport);
    sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).watermark("Button Cell Type").cellPadding('20 20');
    sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
        alignment: 2,
        visibility: 1,
        font: 'bold 15px Arial',
        foreColor: 'grey'
    });

    var c = new GC.Spread.Sheets.CellTypes.CheckBox();
    c.isThreeState(false);
    c.textTrue("Checked!");
    c.textFalse("Check Me!");
    sheet.setColumnWidth(4, 200);
    sheet.setCellType(row, ++col, c, GC.Spread.Sheets.SheetArea.viewport);
    sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).watermark("CheckBox Cell Type").cellPadding('30');
    sheet.getCell(row, col, GC.Spread.Sheets.SheetArea.viewport).labelOptions({
        alignment: 5,
        visibility: 0,
        foreColor: 'green'
    });
    sheet.resumePaint();
}

function setConditionalFormatContent(sheet) {
    var sheet = new spreadNS.Worksheet("Conditional Format");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    sheet.setColumnWidth(0, 20);
    sheet.setColumnWidth(1, 20);
    for (var col = 2; col < 11; col++) {
        sheet.setColumnWidth(col, 83);
    }
    for (var row = 1; row < 16; row++) {
        sheet.addSpan(row, 10, 1, 2);
    }

    var Range = spreadNS.Range;
    var row = 1, col = 1;
    var style = new spreadNS.Style();
    style.backColor = "red";
    var cfs = sheet.conditionalFormats;
    sheet.getCell(row, ++col).value("Cell Value").font("700 11pt Calibri");
    sheet.getCell(row, col + 2).value("Specific Text").font("700 11pt Calibri");
    sheet.getCell(row, col + 4).value("Unique").font("700 11pt Calibri");
    sheet.getCell(row, col + 6).value("Duplicate").font("700 11pt Calibri");
    sheet.getCell(row, col + 8).value("Date Occurring").font("700 11pt Calibri");

    var rowCount = 6;
    row++, col;
    sheet.getCell(row, col).value(0);
    sheet.getCell(row + 1, col).value(1);
    sheet.getCell(row + 2, col).value(2);
    sheet.getCell(row + 3, col).value(3);
    sheet.getCell(row + 4, col).value(4);
    sheet.getCell(row + 5, col).value(5);
    cfs.addCellValueRule(ComparisonOperators.between, 2, 4, style, [new Range(row, col, rowCount, 1)]);

    col = col + 2;
    sheet.getCell(row, col).value("test");
    sheet.getCell(row + 1, col).value("bad");
    sheet.getCell(row + 2, col).value("good");
    sheet.getCell(row + 3, col).value("testing");
    sheet.getCell(row + 4, col).value("tested");
    sheet.getCell(row + 5, col).value("general");
    cfs.addSpecificTextRule(ConditionalFormatting.TextComparisonOperators.contains, "test", style, [new Range(row, col, rowCount, 1)]);

    col = col + 2;
    sheet.getCell(row, col).value(50);
    sheet.getCell(row + 1, col).value(50);
    sheet.getCell(row + 2, col).value(11);
    sheet.getCell(row + 3, col).value(5);
    sheet.getCell(row + 4, col).value(50);
    sheet.getCell(row + 5, col).value(120);
    cfs.addUniqueRule(style, [new Range(row, col, rowCount, 1)]);

    col = col + 2;
    sheet.getCell(row, col).value(50);
    sheet.getCell(row + 1, col).value(50);
    sheet.getCell(row + 2, col).value(11);
    sheet.getCell(row + 3, col).value(5);
    sheet.getCell(row + 4, col).value(50);
    sheet.getCell(row + 5, col).value(120);
    cfs.addDuplicateRule(style, [new Range(row, col, rowCount, 1)]);

    col = col + 2;
    var date = new Date();
    sheet.getCell(row, col).value(date);
    sheet.getCell(row + 1, col).value(new Date(date.setDate(date.getDate() + 1)));
    sheet.getCell(row + 2, col).value(new Date(date.setDate(date.getDate() + 5)));
    sheet.getCell(row + 3, col).value(new Date(date.setDate(date.getDate() + 1)));
    sheet.getCell(row + 4, col).value(new Date(date.setDate(date.getDate() + 7)));
    sheet.getCell(row + 5, col).value(new Date(date.setDate(date.getDate() + 8)));
    cfs.addDateOccurringRule(ConditionalFormatting.DateOccurringType.nextWeek, style, [new Range(row, col, rowCount, 1)]);

    row = row + 7, col = 1;
    sheet.getCell(row, ++col).value("Top/Bottom").font("700 11pt Calibri");
    sheet.getCell(row, col + 2).value("Average").font("700 11pt Calibri");
    sheet.getCell(row, col + 4).value("2-Color Scale").font("700 11pt Calibri");
    sheet.getCell(row, col + 6).value("3-Color Scale").font("700 11pt Calibri");
    sheet.getCell(row, col + 8).value("Data Bar").font("700 11pt Calibri");

    row++;
    sheet.getCell(row, col).value(0);
    sheet.getCell(row + 1, col).value(1);
    sheet.getCell(row + 2, col).value(2);
    sheet.getCell(row + 3, col).value(3);
    sheet.getCell(row + 4, col).value(4);
    sheet.getCell(row + 5, col).value(5);
    cfs.addTop10Rule(ConditionalFormatting.Top10ConditionType.top, 4, style, [new Range(row, col, rowCount, 1)]);

    for (var c = col + 2; c < col + 7; c = c + 2) {
        sheet.getCell(row, c).value(1);
        sheet.getCell(row + 1, c).value(50);
        sheet.getCell(row + 2, c).value(100);
        sheet.getCell(row + 3, c).value(2);
        sheet.getCell(row + 4, c).value(60);
        sheet.getCell(row + 5, c).value(3);
    }
    cfs.addAverageRule(ConditionalFormatting.AverageConditionType.above, style, [new Range(row, col + 2, rowCount, 1)]);
    cfs.add2ScaleRule(1, 1, "red", 2, 100, "yellow", [new Range(row, col + 4, rowCount, 1)]);
    cfs.add3ScaleRule(1, 1, "red", 0, 50, "blue", 2, 100, "yellow", [new Range(row, col + 6, rowCount, 1)]);

    col = col + 8;
    sheet.getCell(row, col).value(1);
    sheet.getCell(row + 1, col).value(15);
    sheet.getCell(row + 2, col).value(25);
    sheet.getCell(row + 3, col).value(-1);
    sheet.getCell(row + 4, col).value(-15);
    sheet.getCell(row + 5, col).value(-25);
    var ScaleValueNumber = ConditionalFormatting.ScaleValueType.number;
    cfs.addDataBarRule(1, null, 2, null, "green", [new Range(row, col, rowCount, 1)]);

    row = row + 8, col = 1;
    sheet.getCell(row, ++col).value("Icon Set").font("700 11pt Calibri");
    sheet.addSpan(row, col, 1, 10);
    sheet.addSpan(row + 6, col, 1, 10);
    row++;
    for (var column = col; column < col + 10; column++) {
        sheet.getCell(row, column).value(-50);
        sheet.getCell(row + 1, column).value(-25);
        sheet.getCell(row + 2, column).value(0);
        sheet.getCell(row + 3, column).value(25);
        sheet.getCell(row + 4, column).value(50);
        sheet.getCell(row + 6, column).value(-50);
        sheet.getCell(row + 7, column).value(-25);
        sheet.getCell(row + 8, column).value(0);
        sheet.getCell(row + 9, column).value(25);
        sheet.getCell(row + 10, column).value(50);
    }
    rowCount = 5;
    cfs.addIconSetRule(0, [new Range(row, col, rowCount, 1)]);
    cfs.addIconSetRule(1, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(2, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(3, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(4, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(5, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(6, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(7, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(8, [new Range(row, ++col, rowCount, 1)]);
    cfs.addIconSetRule(9, [new Range(row, ++col, rowCount, 1)]);
    col = 1;
    cfs.addIconSetRule(10, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(11, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(12, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(13, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(14, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(15, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(16, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(17, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(18, [new Range(row + 6, ++col, rowCount, 1)]);
    cfs.addIconSetRule(19, [new Range(row + 6, ++col, rowCount, 1)]);

    sheet.resumePaint();
}

function getRandomNumber() {
    var num = Math.random();
    if (num - 0.5 > 0) {
        return Math.round(Math.random() * 100);
    }
    else {
        return Math.round(Math.random() * (-100));
    }
}

function setTableContent() {
    var sheet = new spreadNS.Worksheet("Table");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    // table
    var table, rowCount = 5, colCount = 5;
    var row = 0, col = 1;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Table Style - light7").font("700 11pt Calibri");
    sheet.tables.add("sampleTable0", ++row, col, rowCount, colCount, spreadNS.Tables.TableThemes.light7);

    sheet.addSpan(row + 7, col, 1, colCount);
    sheet.getCell(row + 7, col).value("Table Style - medium7").font("700 11pt Calibri");
    sheet.tables.add("sampleTable1", row + 8, col, rowCount, colCount, spreadNS.Tables.TableThemes.medium7);

    sheet.addSpan(row + 15, col, 1, colCount);
    sheet.getCell(row + 15, col).value("Table Style - dark7").font("700 11pt Calibri");
    sheet.tables.add("sampleTable2", row + 16, col, rowCount, colCount, spreadNS.Tables.TableThemes.dark7);

    sheet.addSpan(row + 23, col, 1, colCount);
    sheet.getCell(row + 23, col).value("Hide Filter Button").font("700 11pt Calibri");
    table = sheet.tables.add("sampleTable3", row + 24, col, rowCount, colCount);
    table.filterButtonVisible(false);

    row = 0, col = col + 7;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Hide Header Row").font("700 11pt Calibri");
    table = sheet.tables.add("sampleTable4", ++row, col, rowCount, colCount);
    table.showHeader(false);

    sheet.addSpan(row + 7, col, 1, colCount);
    sheet.getCell(row + 7, col).value("Show Total Row").font("700 11pt Calibri");
    table = sheet.tables.add("sampleTable5", row + 8, col, rowCount, colCount);
    table.showFooter(true);

    sheet.addSpan(row + 15, col, 1, colCount);
    sheet.getCell(row + 15, col).value("Don't display alternating row style").font("700 11pt Calibri");
    table = sheet.tables.add("sampleTable6", row + 16, col, rowCount, colCount);
    table.bandRows(false);

    sheet.addSpan(row + 23, col, 1, colCount);
    sheet.getCell(row + 23, col).value("Display alternating column style").font("700 11pt Calibri");
    table = sheet.tables.add("sampleTable7", row + 24, col, rowCount, colCount);
    table.bandRows(false);
    table.bandColumns(true);

    row = 32, col = 1;
    var data = [
        ["bob", "36", "man", "Beijing", "80"],
        ["Betty", "28", "woman", "Xi'an", "52"],
        ["Gary", "23", "man", "NewYork", "63"],
        ["Hunk", "45", "man", "Beijing", "80"],
        ["Cherry", "37", "woman", "Shanghai", "58"]];
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Highlight first column").font("700 11pt Calibri");
    table = sheet.tables.addFromDataSource("sampleTable8", row + 1, col, data);
    table.highlightFirstColumn(true);
    col = col + 7;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Highlight last column").font("700 11pt Calibri");
    table = sheet.tables.addFromDataSource("sampleTable9", row + 1, col, data);
    table.highlightLastColumn(true);

    sheet.resumePaint();
}

function getHBarFormula(range) {
    return "IF(" + range + ">=0.8,HBARSPARKLINE(" + range + ",\"green\"), " +
        "IF(" + range + ">=0.6,HBARSPARKLINE(" + range + ",\"blue\"), " +
        "IF(" + range + ">=0.4,HBARSPARKLINE(" + range + ",\"yellow\"), " +
        "IF(" + range + ">=0.2,HBARSPARKLINE(" + range + ",\"orange\"), " +
        "IF(" + range + ">=0,HBARSPARKLINE(" + range + ",\"red\"), HBARSPARKLINE(" + range + ",\"red\") " + ")))))";
}

function getVBarFormula(row) {
    return "=IF((Q3:W3>0)=(ROW(Q13:W14)=ROW($Q$13)),VBARSPARKLINE((Q3:W3)/MAX(ABS(Q3:W3)),Q12:W12),\"\")".replace(/(Q|W)3/g, "$1" + row);
}

function setSparklineContent() {
    var sheet = new spreadNS.Worksheet("Sparkline");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    addCompatibleContent(sheet);
    addPieContent(sheet);
    addAreaContent(sheet);
    addScatterContent(sheet);
    addStackedContent(sheet);
    addBulletContent(sheet);
    addBoxPlotContent(sheet);
    addVariContent(sheet);
    addCascadeContent(sheet);
    addSpreadContent(sheet);
    addParetoContent(sheet);
    addHBarContent(sheet);
    addVBarContent(sheet);
    addMonthContent(sheet);
    addYearContent(sheet);
    sheet.resumePaint();
}

function addMonthContent(sheet) {
    sheet.addSpan(51, 3, 4, 2);
    sheet.addSpan(55, 3, 1, 2);
    var day = 1;
    for (var row = 51; row < 82; row++) {
        sheet.setValue(row, 0, new Date(2016, 0, day++));
        sheet.setValue(row, 1, Math.round(Math.random() * 100));
        sheet.setFormatter(row, 0, "MM/DD/YYYY");
    }
    sheet.setFormula(51, 3, '=MONTHSPARKLINE(2016, 1, A52:B82, "lightgray", "lightgreen", "green", "darkgreen")');
    sheet.setFormula(55, 3, '=TEXT(DATE(2016,1, 1),"mmmm")');
}

function  addYearContent(sheet) {
    sheet.addSpan(51, 6, 4, 8);
    sheet.setFormula(51, 6, '=YEARSPARKLINE(2016, A52:B82, "lightgray", "lightgreen", "green", "darkgreen")');
}

function addCompatibleContent(sheet) {
    sheet.addSpan(0, 0, 1, 8);
    sheet.getCell(0, 0).value("The company revenue in 2014").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(1, 2, 1, 2);
    sheet.addSpan(1, 4, 1, 2);
    sheet.addSpan(1, 6, 1, 2);
    sheet.setValue(1, 0, "Month");
    sheet.setValue(1, 1, "Revenue");
    sheet.setValue(1, 2, "Diagram 1");
    sheet.setValue(1, 4, "Diagram 2");
    sheet.setValue(1, 6, "Diagram 3");
    sheet.getRange(1, 0, 1, 7).backColor("Accent 4").foreColor("white");
    for (var i = 2; i < 5; i++) {
        sheet.setValue(i, 0, new Date(2014, i - 1, 1));
        sheet.setFormatter(i, 0, "mm/dd/yyyy");
    }
    sheet.setColumnWidth(0, 80);
    sheet.setValue(2, 1, 30);
    sheet.setValue(3, 1, -60);
    sheet.setValue(4, 1, 80);

    sheet.addSpan(2, 2, 3, 2);
    sheet.setFormula(2, 2, '=LINESPARKLINE(B3:B5,0,A3:A5,0,"{ac:#ffff00,fmc:brown,hmc:red,lastmc:blue,lowmc:green,mc:purple,nc:yellowgreen,sc:pink,dxa:true,sf:true,sh:true,slast:true,slow:true,sn:true,sm:true,lw:3,dh:false,deca:1,rtl:false,minat:1,maxat:1,mmax:5,mmin:-3}")');
    sheet.addSpan(2, 4, 3, 2);
    sheet.setFormula(2, 4, '=COLUMNSPARKLINE(B3:B5,0,A3:A5,0,"{ac:#ffff00,fmc:brown,hmc:red,lastmc:blue,lowmc:green,mc:purple,nc:yellowgreen,sc:pink,dxa:true,sf:true,sh:true,slast:true,slow:true,sn:true,sm:true,lw:3,dh:false,deca:1,rtl:false,minat:1,maxat:1,mmax:5,mmin:-3}")');
    sheet.addSpan(2, 6, 3, 2);
    sheet.setFormula(2, 6, '=WINLOSSSPARKLINE(B3:B5,0,A3:A5,0)');
}

function addPieContent(sheet) {
    sheet.addSpan(6, 0, 1, 5);
    sheet.getCell(6, 0).value("My Assets").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(7, 2, 1, 2);
    sheet.addSpan(8, 2, 3, 2);
    sheet.setValue(7, 0, "Asset Type");
    sheet.setValue(7, 1, "Amount");
    sheet.setValue(7, 2, "Diagram");
    sheet.setValue(7, 4, "Note");
    sheet.setValue(8, 0, "Savings");
    sheet.getRange(7, 0, 1, 5).backColor("Accent 4").foreColor("white");
    sheet.getCell(8, 1).value(25000).formatter("$#,##0");
    sheet.setValue(9, 0, "401k");
    sheet.getCell(9, 1).value(55000).formatter("$#,##0");
    sheet.setValue(10, 0, "Stocks");
    sheet.getCell(10, 1).value(15000).formatter("$#,##0");
    sheet.setFormula(8, 2, '=PIESPARKLINE(B9:B11,"#919F81","#D7913E","CEA722")');
    sheet.getCell(8, 4).backColor("#919F81").formula("=B9/SUM(B9:B11)").formatter("0.00%");
    sheet.getCell(9, 4).backColor("#D7913E").formula("=B10/SUM(B9:B11)").formatter("0.00%");
    sheet.getCell(10, 4).backColor("#CEA722").formula("=B11/SUM(B9:B11)").formatter("0.00%");
}

function addAreaContent(sheet) {
    sheet.addSpan(12, 0, 1, 5);
    sheet.getCell(12, 0).value("Sales by State").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(13, 2, 1, 3);
    sheet.addSpan(14, 2, 4, 3);
    sheet.setValue(13, 0, "State");
    sheet.setValue(13, 1, "Sales");
    sheet.setValue(13, 2, "Diagram");
    sheet.setValue(14, 0, "Idaho");
    sheet.getRange(13, 0, 1, 3).backColor("Accent 4").foreColor("white");
    sheet.getCell(14, 1).value(3500).formatter("$#,##0");
    sheet.setValue(15, 0, "Montana");
    sheet.getCell(15, 1).value(7000).formatter("$#,##0");
    sheet.setValue(16, 0, "Oregon");
    sheet.getCell(16, 1).value(2000).formatter("$#,##0");
    sheet.setValue(17, 0, "Washington");
    sheet.getCell(17, 1).value(5000).formatter("$#,##0");
    sheet.setFormula(14, 2, '=AREASPARKLINE(B15:B18,,,0,6000,"yellowgreen","red")');
}

function addScatterContent(sheet) {
    sheet.addSpan(19, 0, 1, 5);
    sheet.getCell(19, 0).value("Particulate Levels in Rainfall").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(20, 2, 1, 3);
    sheet.addSpan(21, 2, 4, 3);
    sheet.setValue(20, 0, "Daily rainfall");
    sheet.setValue(20, 1, "Particulate level");
    sheet.setValue(20, 2, "Diagram");
    sheet.getRange(20, 0, 1, 3).backColor("Accent 4").foreColor("white");
    sheet.setValue(21, 0, 2.0);
    sheet.setValue(21, 1, 100);
    sheet.setValue(22, 0, 3.0);
    sheet.setValue(22, 1, 130);
    sheet.setValue(23, 0, 4.0);
    sheet.setValue(23, 1, 110);
    sheet.setValue(24, 0, 5.0);
    sheet.setValue(24, 1, 135);
    sheet.setFormula(21, 2, '=SCATTERSPARKLINE(A22:B25,,MIN(A22:A25),MAX(A22:A25),MIN(B22:B25),MAX(B22:B25),AVERAGE(B22:B25),AVERAGE(A22:A25),,,,,TRUE,TRUE,TRUE,"green",,TRUE)');
}

function addStackedContent(sheet) {
    sheet.addSpan(26, 0, 1, 5);
    sheet.getCell(26, 0).value("Sales by State").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(27, 0, 1, 6).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(27, 4, 1, 3);
    sheet.addSpan(28, 4, 1, 3);
    sheet.setRowHeight(28, 30);
    sheet.setValue(27, 0, "State");
    sheet.setValue(27, 1, "Product 1");
    sheet.setValue(27, 2, "Product 2");
    sheet.setValue(27, 3, "Product 3");
    sheet.setValue(27, 4, "Diagram");
    sheet.setValue(28, 0, "Idaho");
    sheet.getCell(28, 1).value(10000).formatter("$#,##0");
    sheet.getCell(28, 2).value(12000).formatter("$#,##0");
    sheet.getCell(28, 3).value(15000).formatter("$#,##0");
    sheet.setValue(29, 1, "orange");
    sheet.setValue(29, 2, "purple");
    sheet.setValue(29, 3, "yellowgreen");
    sheet.setFormula(28, 4, '=STACKEDSPARKLINE(B29:D29,B30:D30,B28:D28,40000)');
}

function addBulletContent(sheet) {
    sheet.addSpan(31, 0, 1, 5);
    sheet.getCell(31, 0).value("Employee KPI").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(32, 0, 1, 4).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.HorizontalAlign.center);
    sheet.addSpan(32, 3, 1, 2);
    sheet.addSpan(33, 3, 1, 2);
    sheet.addSpan(34, 3, 1, 2);
    sheet.addSpan(35, 3, 1, 2);
    sheet.setValue(32, 0, "Name");
    sheet.setValue(32, 1, "Forecast");
    sheet.setValue(32, 2, "Actuality");
    sheet.setValue(32, 3, "Diagram");
    sheet.setValue(33, 0, "Employee 1");
    sheet.setValue(33, 1, 6);
    sheet.setValue(33, 2, 6);
    sheet.setValue(34, 0, "Employee 2");
    sheet.setValue(34, 1, 8);
    sheet.setValue(34, 2, 7);
    sheet.setValue(35, 0, "Employee 3");
    sheet.setValue(35, 1, 6);
    sheet.setValue(35, 2, 4);

    sheet.addSpan(38, 6, 1, 3);
    sheet.setValue(38, 6, "BULLETSPARKLINE Settings:");
    sheet.setValue(39, 6, "target");
    sheet.setValue(39, 7, 7);
    sheet.setValue(40, 6, "maxi");
    sheet.setValue(40, 7, 10);
    sheet.setValue(41, 6, "good");
    sheet.setValue(41, 7, 8);
    sheet.setValue(42, 6, "bad");
    sheet.setValue(42, 7, 5);
    sheet.setValue(43, 6, "color scheme");
    sheet.setValue(43, 7, "gray");

    sheet.setFormula(33, 3, '=BULLETSPARKLINE(C34,H40,H41,H42,H43,H34,1,H44)');
    sheet.setFormula(34, 3, '=BULLETSPARKLINE(C35,H40,H41,H42,H43,H34,1,H44)');
    sheet.setFormula(35, 3, '=BULLETSPARKLINE(C36,H40,H41,H42,H43,H34,1,H44)');
    sheet.setRowHeight(33, 28);
    sheet.setRowHeight(34, 28);
    sheet.setRowHeight(35, 28);
}

function addBoxPlotContent(sheet) {
    sheet.addSpan(31, 6, 1, 8);
    sheet.getCell(31, 6).value("The Company Sales in 2014 (Month)").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(32, 12, 1, 2);
    sheet.addSpan(33, 12, 1, 2);
    sheet.addSpan(34, 12, 1, 2);
    sheet.addSpan(35, 12, 1, 2);
    sheet.setValue(32, 7, 1);
    sheet.setValue(32, 8, 2);
    sheet.setValue(32, 9, 3);
    sheet.setValue(32, 10, 4);
    sheet.setValue(32, 11, 5);
    sheet.setValue(32, 12, "Actual Sales");
    sheet.getRange(32, 7, 1, 7).hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center).wordWrap(true);
    sheet.setValue(32, 6, "Region");
    sheet.setValue(33, 6, "Alabama");
    sheet.setValue(34, 6, "Alaska");
    sheet.setValue(35, 6, "Arizona");
    var data = [[5268, 6281, 8921, 1069, 1239],
        [2837, 5739, 993, 4247, 9514],
        [6661, 4172, 9777, 1282, 9535]];
    sheet.setArray(33, 7, data);
    sheet.addSpan(38, 10, 1, 4);
    sheet.setValue(38, 10, "BOXPLOTSPARKLINE Settings:");
    sheet.setValue(39, 10, "Start scope of the sale:");
    sheet.setValue(40, 10, "End scope of the sale:");
    sheet.setValue(41, 10, "Start scope of expected sale:");
    sheet.setValue(42, 10, "End scope of expected sale:");
    sheet.addSpan(39, 10, 1, 3);
    sheet.addSpan(40, 10, 1, 3);
    sheet.addSpan(41, 10, 1, 3);
    sheet.addSpan(42, 10, 1, 3);
    sheet.setValue(39, 13, 0);
    sheet.setValue(40, 13, 10000);
    sheet.setValue(41, 13, 1000);
    sheet.setValue(42, 13, 8000);

    sheet.getRange(32, 6, 1, 7).backColor("Accent 4").foreColor("white");
    sheet.setFormula(33, 12, '=BOXPLOTSPARKLINE(H34:L34,"5ns",true,N40,N41,N42,N43,"#00FF7F",0,false)');
    sheet.setFormula(34, 12, '=BOXPLOTSPARKLINE(H35:L35,"5ns",true,N40,N41,N42,N43,"#00FF7F",0,false)');
    sheet.setFormula(35, 12, '=BOXPLOTSPARKLINE(H36:L36,"5ns",true,N40,N41,N42,N43,"#00FF7F",0,false)');
}

function addVariContent(sheet) {
    sheet.addSpan(0, 9, 1, 5);
    sheet.getCell(0, 9).value("Mobile Phone Contrast").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(1, 9, 1, 5).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center)
        .vAlign(spreadNS.VerticalAlign.center).wordWrap(true);
    sheet.addSpan(1, 12, 1, 2);
    sheet.addSpan(2, 12, 1, 2);
    sheet.addSpan(3, 12, 1, 2);
    sheet.addSpan(4, 12, 1, 2);
    sheet.setValue(1, 10, "Phone I");
    sheet.setValue(1, 11, "Phone II");
    sheet.setValue(1, 12, "Diagram");
    var data = [["Size(inch)", 5, 4.7],
        ["RAM(G)", 3, 1],
        ["Weight(g)", 149, 129]];
    sheet.setArray(2, 9, data);
    sheet.setFormula(2, 12, '=VARISPARKLINE(ROUND((K3-L3)/K3,2),0,,,,,TRUE)');
    sheet.setFormula(3, 12, '=VARISPARKLINE(ROUND((K4-L4)/K4,2),0,,,,,TRUE)');
    sheet.setFormula(4, 12, '=VARISPARKLINE(ROUND(-1*(K5-L5)/K5,2),0,,,,,TRUE)');
}

function addCascadeContent(sheet) {
    sheet.addSpan(6, 6, 1, 8);
    sheet.getCell(6, 6).value("Checkbook Register").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    for (var r = 7; r < 12; r++) {
        sheet.addSpan(r, 6, 1, 2);
        sheet.addSpan(r, 11, 1, 3);
    }
    sheet.setArray(7, 6, [
        ["", "", "InitialValue", 815.25, "\u03A3"],
        ["12/11/2012", "", "CVS", -200],
        ["12/12/2012", "", "Bank", 1000.12],
        ["12/13/2012", "", "Starbucks", -500.43],
        ["", "", "FinalValue"]
    ]);
    sheet.getRange(8, 6, 3, 1).formatter("MM/dd/yyyy");
    sheet.getRange(7, 9, 5, 1).formatter("#,###.00");
    sheet.getRange(8, 10, 3, 1).formatter("#,###.00");
    sheet.getCell(7, 10).hAlign(spreadNS.HorizontalAlign.center);
    sheet.getRange(7, 8, 1, 2).font("bold 14px Georgia");
    sheet.getRange(11, 8, 1, 2).font("bold 14px Georgia");

    sheet.setFormula(8, 10, "=J8 + J9");
    for (var r = 10; r <= 11; r++) {
        sheet.setFormula(r - 1, 10, "=J" + r + " + K" + (r - 1));
    }
    sheet.setFormula(11, 9, "=K11");
    sheet.getRange(7, 6, 1, 8).setBorder(new spreadNS.LineBorder("black", spreadNS.LineStyle.thin), {bottom: true});
    sheet.getRange(11, 6, 1, 8).setBorder(new spreadNS.LineBorder("black", spreadNS.LineStyle.medium), {top: true});
    sheet.setFormula(7, 11, '=CASCADESPARKLINE(J8:J12,1,I8:I12,,,"#8CBF64","#D6604D",false)');
    sheet.setFormula(8, 11, '=CASCADESPARKLINE(J8:J12,2,I8:I12,,,"#8CBF64","#D6604D",false)');
    sheet.setFormula(9, 11, '=CASCADESPARKLINE(J8:J12,3,I8:I12,,,"#8CBF64","#D6604D",false)');
    sheet.setFormula(10, 11, '=CASCADESPARKLINE(J8:J12,4,I8:I12,,,"#8CBF64","#D6604D",false)');
    sheet.setFormula(11, 11, '=CASCADESPARKLINE(J8:J12,5,I8:I12,,,"#8CBF64","#D6604D",false)');
}

function addSpreadContent(sheet) {
    sheet.addSpan(13, 6, 1, 7);
    sheet.getCell(13, 6).value("Student Grade Statistics").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(14, 6, 1, 8).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.addSpan(15, 6, 2, 1);
    sheet.addSpan(14, 7, 1, 2);
    sheet.addSpan(14, 9, 1, 2);
    sheet.addSpan(14, 11, 1, 2);
    sheet.addSpan(15, 7, 2, 2);
    sheet.addSpan(15, 9, 2, 2);
    sheet.addSpan(15, 11, 2, 2);
    sheet.addSpan(15, 13, 2, 1);
    for (var r = 17; r <= 21; r++) {
        sheet.addSpan(r, 7, 1, 2);
        sheet.addSpan(r, 9, 1, 2);
        sheet.addSpan(r, 11, 1, 2);
    }
    sheet.setArray(14, 6, [["Name", "Chinese", "", "Math", "", "English", "", "Total"]]);
    sheet.setArray(17, 6, [
        ["Student 1", 70, "", 90, "", 51],
        ["Student 2", 99, "", 59, "", 63],
        ["Student 3", 89, "", 128, "", 74],
        ["Student 4", 93, "", 61, "", 53],
        ["Student 5", 106, "", 82, "", 80]
    ]);
    for (var i = 0; i <= 5; i++) {
        r = 17 + i;
        sheet.setFormula(r - 1, 13, "=Sum(H" + r + ":M" + r + ")");
    }
    sheet.setFormula(15, 7, "=SPREADSPARKLINE(H18:I22,TRUE,,,1,\"green\")");
    sheet.setFormula(15, 9, "=SPREADSPARKLINE(J18:K22,TRUE,,,3,\"green\")");
    sheet.setFormula(15, 11, "=SPREADSPARKLINE(L18:M22,TRUE,,,5,\"green\")");
    sheet.setFormula(15, 13, "=SPREADSPARKLINE(N18:N22,TRUE,,,6,\"green\")");
}

function addParetoContent(sheet) {
    sheet.addSpan(23, 8, 1, 6);
    sheet.getCell(23, 8).value("The Reason of Being Late").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(24, 8, 1, 6).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    for (var r = 24; r < 30; r++) {
        sheet.addSpan(r, 11, 1, 3);
    }
    sheet.setArray(24, 8, [
        ["", "Points", "Color", "Diagram"],
        ["Traffic", 20, "#FF1493"],
        ["Child care", 15, "#FFE7BA"],
        ["Weather", 16, "#FFAEB9"],
        ["Overslept", 4, "#FF8C69"],
        ["Emergency", 1, "#FF83FA"]
    ]);
    sheet.addSpan(45, 6, 1, 3);
    sheet.setValue(45, 6, "PARETOSPARKLINE Settings:");
    sheet.setValue(46, 6, "target");
    sheet.setValue(46, 7, 0.5);
    sheet.setValue(47, 6, "target1");
    sheet.setValue(47, 7, 0.8);

    sheet.setFormula(25, 11, '=PARETOSPARKLINE(J26:J30,1,K26:K30,H47,H48,4,2,false)');
    sheet.setFormula(26, 11, '=PARETOSPARKLINE(J26:J30,2,K26:K30,H47,H48,4,2,false)');
    sheet.setFormula(27, 11, '=PARETOSPARKLINE(J26:J30,3,K26:K30,H47,H48,4,2,false)');
    sheet.setFormula(28, 11, '=PARETOSPARKLINE(J26:J30,4,K26:K30,H47,H48,4,2,false)');
    sheet.setFormula(29, 11, '=PARETOSPARKLINE(J26:J30,5,K26:K30,H47,H48,4,2,false)');
}

function addHBarContent(sheet) {
    row = 37, col = 0;
    sheet.addSpan(row, col, 1, 6);
    sheet.getCell(row, col).value("SPRINT 4").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(row + 1, 8, 1, 6).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    for (var r = 38; r < 44; r++) {
        sheet.addSpan(r, 2, 1, 3);
    }
    sheet.getCell(++row, col).value("Name");
    sheet.getCell(++row, col).value("Employee1");
    sheet.getCell(++row, col).value("Employee2");
    sheet.getCell(++row, col).value("Employee3");
    sheet.getCell(++row, col).value("Employee4");
    sheet.getCell(++row, col).value("Employee5");
    row = 38, col++;
    sheet.getCell(row, col).value("Progress");
    sheet.getCell(++row, col).value(0.7);
    sheet.getCell(++row, col).value(0.1);
    sheet.getCell(++row, col).value(0.3);
    sheet.getCell(++row, col).value(1.1);
    sheet.getCell(++row, col).value(0.5);
    row = 38, col++;
    sheet.getCell(row, col).value("Diagram");
    sheet.getRange(38, 0, 1, 3).backColor("Accent 4").foreColor("white");
    sheet.setFormula(++row, col, getHBarFormula("B40"));
    sheet.setFormula(++row, col, getHBarFormula("B41"));
    sheet.setFormula(++row, col, getHBarFormula("B42"));
    sheet.setFormula(++row, col, getHBarFormula("B43"));
    sheet.setFormula(++row, col, getHBarFormula("B44"));
}

function addVBarContent(sheet) {
    sheet.setColumnWidth(15, 60);
    for (var c = 16; c < 23; c++) {
        sheet.setColumnWidth(c, 30);
    }
    sheet.addSpan(0, 15, 1, 8);
    sheet.getCell(0, 15).value("The Temperature Variation").font("20px Arial").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    sheet.getRange(1, 15, 1, 8).foreColor("white").backColor("Accent 4").hAlign(spreadNS.HorizontalAlign.center).vAlign(spreadNS.VerticalAlign.center);
    row = 2;
    sheet.addSpan(row, 15, 3, 1);
    sheet.addSpan(row + 3, 15, 3, 1);
    sheet.addSpan(row + 6, 15, 3, 1);
    sheet.setArray(1, 15, [["City", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"]]);
    var datas = [
        ["Austin", 5, 11, 19, 24, 21, 16, 6],
        ["Buffalo", -8, -3, -1, 3, 14, 6, -4],
        ["Chicago", -9, -2, 2, 18, 12, 5, -6]
    ];
    var colors = ["#0099FF", "#33FFFF", "#9E0142", "#D53E4F", "#F46D43", "#FDAE61", "#FEE08B"];
    sheet.setArray(11, 16, [colors]);
    for (var i = 0; i < datas.length; i++) {
        var row = 2 + 3 * i;
        sheet.setArray(row, 15, [datas[i]]);
        sheet.setArrayFormula(row + 1, 16, 2, 7, getVBarFormula(row + 1));
        sheet.setRowHeight(row + 1, 30);
        sheet.setRowHeight(row + 2, 30);
    }
}

function setCommentContent() {
    var sheet = new spreadNS.Worksheet("Comment");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    for (var col = 2; col < 9; col++) {
        sheet.setColumnWidth(col, 100);
    }

    var Comment = spreadNS.Comments.Comment;
    var DisplayMode = spreadNS.Comments.DisplayMode;
    var commentText = "Hello, world!";
    var rowCount = 5, colCount = 4;
    var row = 2, col = 2;

    sheet.getCell(row, col).value("HoverShown").font("700 11pt Calibri");
    sheet.comments.add(row, col, commentText);
    sheet.getCell(row, col + colCount).value("AlwaysShown").font("700 11pt Calibri");
    sheet.comments.add(row, col + colCount, commentText)
        .displayMode(DisplayMode.alwaysShown);
    row = row + rowCount;
    sheet.getCell(row, col).value("Size").font("700 11pt Calibri");
    sheet.comments.add(row, col, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .height(80)
        .width(200);
    sheet.getCell(row, col + colCount).value("Shadow").font("700 11pt Calibri");
    sheet.comments.add(row, col + colCount, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .showShadow(true);
    row = row + rowCount;
    sheet.getCell(row, col).value("Font").font("700 11pt Calibri");
    sheet.comments.add(row, col, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .fontFamily("Comic Sans MS")
        .fontSize("10pt")
        .fontStyle("italic")
        .fontWeight("bold");
    sheet.getCell(row, col + colCount).value("Color Opacity").font("700 11pt Calibri");
    sheet.comments.add(row, col + colCount, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .foreColor("green")
        .backColor("yellow")
        .opacity(0.8);
    row = row + rowCount;
    sheet.getCell(row, col).value("Border").font("700 11pt Calibri");
    sheet.comments.add(row, col, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .borderColor("green")
        .borderStyle("dotted")
        .borderWidth(2);
    sheet.getCell(row, col + colCount).value("Text Decoration").font("700 11pt Calibri");
    sheet.comments.add(row, col + colCount, commentText)
        .displayMode(DisplayMode.alwaysShown)
        .textDecoration(1)
        .horizontalAlign(1)
        .padding(new spreadNS.Comments.Padding(2));

    sheet.resumePaint();
}

function setPictureContent() {
    var sheet = new spreadNS.Worksheet("Picture");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    sheet.setColumnWidth(0, 20);

    var url = "css/images/logo.png";
    var ImageLayout = spreadNS.ImageLayout;
    var row, col, rowCount = 11, colCount = 5,
        colWidth = sheet.getColumnWidth(1), rowHeight = sheet.getRowHeight(1),
        width = colCount * colWidth, height = rowCount * rowHeight,
        x = sheet.getColumnWidth(0) + colWidth, y = 2 * rowHeight,
        xOffset = (colCount + 2) * colWidth, yOffset = (rowCount + 2) * rowHeight;

    row = 1, col = 2;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Border").font("700 11pt Calibri");
    sheet.pictures.add("border_picture", url, x, y, width, height)
        .backColor("#000000")
        .borderColor("red")
        .borderWidth(4)
        .borderStyle("dotted")
        .borderRadius(5);

    col = col + colCount + 2;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Fixed Position").font("700 11pt Calibri");
    sheet.pictures.add("fixed_picture", url, x + xOffset, y, width, height)
        .backColor("#000000")
        .fixedPosition(true);

    row = row + rowCount + 2, col = 2;
    y += yOffset;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Stretch").font("700 11pt Calibri");
    sheet.pictures.add("stretch_picture", url, x, y, width, height)
        .backColor("#000000");

    col = col + colCount + 2;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Center").font("700 11pt Calibri");
    sheet.pictures.add("center_picture", url, x + xOffset, y, width, height)
        .backColor("#000000")
        .pictureStretch(ImageLayout.center);

    row = row + rowCount + 2, col = 2;
    y += yOffset;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("Zoom").font("700 11pt Calibri");
    sheet.pictures.add("zoom_picture", url, x, y, width, height)
        .backColor("#000000")
        .pictureStretch(ImageLayout.zoom);

    col = col + colCount + 2;
    sheet.addSpan(row, col, 1, colCount);
    sheet.getCell(row, col).value("None").font("700 11pt Calibri");
    sheet.pictures.add("none_picture", url, x + xOffset, y, width, height)
        .backColor("#000000")
        .pictureStretch(ImageLayout.none);

    sheet.resumePaint();
}

function setDataContent() {
    var sheet = new spreadNS.Worksheet("Data");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    for (var col = 1; col < 6; col = col + 2) {
        for (var row = 2; row < 10; row++) {
            sheet.getCell(row, col).value(getRandomNumber());
        }
    }
    var row = 1, col = 1, rowCount = 8, colCount = 1;
    sheet.getCell(row, col).value("Sort A-Z").font("700 11pt Calibri");
    sheet.sortRange(row + 1, col, rowCount, colCount, true, [{index: col, ascending: true}]);

    col = col + 2;
    sheet.getCell(row, col).value("Sort Z-A").font("700 11pt Calibri");
    sheet.sortRange(row + 1, col, rowCount, colCount, true, [{index: col, ascending: false}]);

    col = col + 2;
    sheet.getCell(row, col).value("Filter").font("700 11pt Calibri");
    sheet.rowFilter(new spreadNS.Filter.HideRowFilter(new spreadNS.Range(row + 1, col, rowCount, colCount)));

    sheet.rowOutlines.group(12, 3);
    sheet.columnOutlines.group(8, 5);

    row = 12, col = 1;
    sheet.addSpan(row, col, 1, 9);
    sheet.getCell(row, col).value("Data Validation").vAlign(spreadNS.VerticalAlign.center).hAlign(spreadNS.HorizontalAlign.center).font("700 11pt Calibri");
    row = 13;
    sheet.getCell(row, col).value("List").font("700 11pt Calibri");
    sheet.getCell(row, col + 2).value("Number").font("700 11pt Calibri");
    sheet.getCell(row, col + 4).value("Date").font("700 11pt Calibri");
    sheet.getCell(row, col + 6).value("Formula").font("700 11pt Calibri");
    sheet.getCell(row, col + 8).value("TextLength").font("700 11pt Calibri");

    row = 14;
    var listValidator = DataValidation.createListValidator("Fruit,Vegetable,Food");
    listValidator.inputTitle("Please choose a category:");
    listValidator.inputMessage("Fruit, Vegetable, Food");
    listValidator.highlightStyle({
        type: GC.Spread.Sheets.DataValidation.HighlightType.icon,
        color: "gold",
        position: GC.Spread.Sheets.DataValidation.HighlightPosition.outsideRight,
    });
    sheet.getCell(row + 1, col).value("Vegetable");
    sheet.getCell(row + 2, col).value("Home");
    sheet.getCell(row + 3, col).value("Fruit");
    sheet.getCell(row + 4, col).value("Company");
    sheet.getCell(row + 5, col).value("Food");

    sheet.setDataValidator(row + 1, col, 5, 1, listValidator);

    col = col + 2;
    var numberValidator = DataValidation.createNumberValidator(ComparisonOperators.between, 0, 100, true);
    numberValidator.inputMessage("Value should Between 0 ~ 100");
    numberValidator.inputTitle("Tip");
    numberValidator.highlightStyle({
        type: GC.Spread.Sheets.DataValidation.HighlightType.dogEar,
        color: "green",
        position: GC.Spread.Sheets.DataValidation.HighlightPosition.topRight
    });
    sheet.getCell(row + 1, col).value(-12);
    sheet.getCell(row + 2, col).value(30);
    sheet.getCell(row + 3, col).value(80);
    sheet.getCell(row + 4, col).value(-35);
    sheet.getCell(row + 5, col).value(66);

    sheet.setDataValidator(row + 1, col, 5, 1, numberValidator);

    col = col + 2;
    sheet.setColumnWidth(col, 100);
    var currentDate = new Date().toLocaleDateString().replace(/\u200E/g, ''); // this "replace" is just for IE, the date string contains some special characters
    var dateValidator = DataValidation.createDateValidator(ComparisonOperators.lessThan, currentDate, currentDate);
    dateValidator.inputMessage("Enter a date Less than " + currentDate);
    dateValidator.inputTitle("Tip");
    dateValidator.highlightStyle({
        type: GC.Spread.Sheets.DataValidation.HighlightType.icon,
        color: "yellow",
        position: GC.Spread.Sheets.DataValidation.HighlightPosition.outsideLeft,
        image: "data:image/ico;base64,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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABpH5YAbBiTAG4cjABrGZUCZRGRAGgVkmBzJpr/fzmi/H02of59NqH+fjag/n01oP98NKD/fDSh/3szof98M6D/ezKf/3syoP96MZ//ezGe/3ownv95L5//ejCe/3ownf95L57/eC6d/3gunv94Lp7/eS+e/3Uom/9vH5f/dSib/4tKq/+qe8H/yq7Z/+PU6//38/n//v7+//7+/v/+/v7////////////////////////////////////////////////////////////////////////////////////////////+/v7//v7+//7+/v/59fr/6Nvu/9K53v+3j8v/mGC1/342of9rGZT/YwyO/2QOj/9pFpL/bBuV/2wblf9rGZT/axmU/2sYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPxrGJP/ahmU02oZlBVqGZQAahmUAWkXlQBqGZQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGsYkwBrGJMAahiTAWINjQBlEI8lbx+W5X02of9/OKL9fjeh/303ov59NqH/fjWg/301of98NKD/fDSh/3szoP98M5//ezKf/3oyoP96MZ//ezGe/3own/95L57/ejCe/3kvnf95L57/ejGf/3UpnP9wIpj/gz2l/66BxP/axuT/+PT5///////+/v7//v7+///////+/v7//Pv9//z6/P/7+fz/+/n8//z6/P/9/P3//f3+//7+/v/+/v7//v7+//7+/v/+/v7//v3+//39/v/9+/3//Pr8//v6/P/7+fz/+/r8//z6/f/9+/3//v7+///////+/v7//v7+//7+/v//////9fD4/9vH5f+2jcr/j1Gu/3EimP9jDY7/ZA+P/2oYk/9sG5X/axqU/2sZlP9rGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlPtrGZT/ahmUlmsZkwBrGJMCahmUAGsYlABrGpQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABqGZUAahiTAGwakwB1K50Bax6YAGsZlKh5L57/gDqj+344ov9/OKH+fjeh/303ov9+NqH/fjWg/301of98NKD/fDSh/3szoP98M5//ezKf/3oyoP97MZ//ezCe/3own/95L57/ejGe/3ownv9yJJj/hEGm/7yYz//w6PT///////7+/v/+/v7//f3+//z7/f/8+vz//fz9//7+/v///////v7+//////////////////7+/v///////v7+//79/v/9/P3//fz9//7+/v///////v7+/////////////////////////////v7+///////+/v7//Pv9//z6/P/8+vz//Pv9//7+/v///////v7+///////9/P3/5tnt/7yXzv+NTqz/axmU/2ILjf9oFZL/bBuV/2salP9rGZT/ahiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlP5rGJP+ahmUSmsYkwBrGJMCbBeSAGsYkwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGoYlAB0HY0AaxmUAmUQkABoE5FYdCea/4E6o/1/OKL+fzmj/n44ov9/OKL/fjei/302ov9+NqH/fjWg/301of98NKD/ezSh/3szoP98M6D/ezKg/3oxoP97MZ//ejCe/3sxn/94LZ3/diub/6t9wv/v5vP///////7+/v/9/f7//fz9//38/f/9/f7///////7+/v///////fz9/+7k8v/dyub/y67Z/7qUzf+sf8P/o2+7/5pitf+TWbL/kVSv/49Qrf+OUK3/kVWv/5Vbs/+cZrf/pHK9/6+Cxf+7ls7/y67Z/9rG5P/o3O7/9/T5///////+/v7//v7+//7+/v///////fz+//z7/f/8+/3//fz9//7+/v/+/v7///////v6/P/cyOX/pna//3Uqm/9jDI7/ZxOR/2wblf9rGpT/ahiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU/GsZlP9qGZTOahmUDmsYkwBrGJMAaRqVAGwZkgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABqF5UAahmSAGsZlAFfCIsAZA+PEm4eltV9NqH/gDqk/H84ov9/OKL+fzmj/344ov9/N6H/fjei/302of9+NqH/fTWg/301of98NKD/ezSh/3szoP98Mp//ezKg/3oxn/98M6D/dyyc/343ov/Nsdv///////7+/v/9/P7//v3+//39/v/+/f7//v7+///////z7ff/1Lzg/7GGxv+UWLH/gDuk/3Qnm/9vIJf/bBuV/2sZlP9sG5X/bRyV/20dlf9tHZX/bRyV/20clP9sG5T/axmU/2kYk/9pFpL/aBOR/2gVkv9rGZT/bh6W/3gtnP+EQab/mF60/7CGx//LsNr/5tjs//r3+////////v7+///////+/v7//Pv9//z7/f/9/P3//v7+//7+/v//////7OLx/7SLyf95L57/YwyO/2gVkv9sG5X/axmU/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT+axmU/GsYlP9rGZR7axiTAGsYlAJrGZMAaxmTAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGsZkwBpF5MAbR2VAmcUkQBpFpJ9eC2d/4I8pfx/OaP+fzmj/oA4ov9/OKL/fjmj/384ov9/N6H/fjei/302of9+NqH/fTWg/301of98NKH/ezOg/3wzoP97Mp//fDOg/3gunv+BPKT/2sbk///////9/P3//f3+//7+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pf6EQKb/gz+l/4I/pv+CPqX/gz6k/4I9pf+BPKX/gjyk/4E7o/+BO6T/gDqj/386pP9/OaP/gDmi/385o/9+OaP/gDqj/3csnP/CodP///////38/f/9/f7//////+PU6/9+OKP/ezGf/3wzoP96MqD/ejGf/3sxnv96MJ//eS+f/3ownv95L53/eS+e/3gunf93Lp7/dy2d/3gtnP93LJ3/diud/3crnP92Kpv/diqc/3Upm/90KZz/dCib/3Uomv90J5v/cyab/3Qnmv9zJpn/cyaa/3Ilmf9zJZj/ciSY/3Ejmf9xI5n/cCKZ/3EimP9wIZf/byGY/28gl/9wIJb/bx+X/28flv9uHpf/bR2W/24elv9tHZX/bB2W/2wclf9tG5T/bBuV/2wblf9rGpT/bh2V/2YSkP91KJv/4tLq///////9/P3///////79/v//////9/L5/4lIqv9lD4//bBuV/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/moZlP9rGJP/axiTQWsYkwBrGJMCahiUAAAAAAAAAAAAAAAAAAAAAABrFpMAaxiTAV4EigBgB4saciSZ5oRApv+FQqf9hUGm/oRApf+EQKb/gz+l/4I/pv+DPqX/gz2k/4I9pf+BPKT/gjyj/4E7o/+AO6T/gDqj/4A6pP+AOaP/gDii/385o/+AO6T/dyuc/76b0P///////fz9//38/v//////4NDp/3w1of97M6D/fDOg/3syn/96MqD/ejGf/3swnv96MJ//eS+f/3ownv95L53/eS+e/3gunf94Lp7/eC2d/3gsnP93LJ3/diuc/3crnP92Kpv/diqc/3Upm/90KZz/dCib/3Unmv90J5v/cyab/3Qnmv9zJpn/cyaa/3Ilmf9zJJj/ciSZ/3IjmP9xI5n/cCKZ/3EimP9wIZf/byGY/28gl/9wIJb/bx+X/24elv9uHpf/bR2W/24elf9tHZX/bByW/2wclf9tHJX/bBuV/2salf9rGpT/bh2V/2QPj/+EQKb/8+z2///////+/f7///////38/f//////28jl/2wblf9qF5P/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU+2sZk/9rGZOPahmUAGoYlAJrGZMAAAAAAAAAAAAAAAAAAAAAAGsalABsHJUCZhKQAGcUkVt6MZ//h0So/YRBpv6EQab+hUGm/4RApf+EQKb/gz+l/4I+pv+DPqX/gz2k/4I9pf+BPKT/gjyj/4E7o/+BO6T/gDqj/385o/+AOaP/gDii/4A7pP93LZ3/sIXG///////9/f7//f3+///////k1ev/gDqj/3syoP98NKH/fDOg/3syn/96MqD/ezGf/3swnv96MJ//eS+e/3ownv95L57/eS+e/3gunv93LZ7/eC2d/3gsnP93LJ3/diuc/3crnP92Kpv/diqc/3Upm/90KJz/dSib/3Qnmv90J5v/cyaa/3Qnmv9zJpn/ciaa/3Mlmf9zJZj/ciSZ/3EjmP9xI5n/cCKY/3Eil/9wIZf/byGY/28gl/9wH5b/bx+X/24elv9uHpf/bR2W/24elf9tHZX/bB2W/2wclf9sG5T/bBuV/2salP9rGZP/bh2V/2ILjf+ndb////////79/v/+/v7//v7+//38/f//////p3a//2MMjv9sG5X/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT9axmU/2oZlNRqGZQNahmUAGoZlAEAAAAAAAAAAAAAAABqGpQAaheTAH45ogF2LJwAbRyVp4A6o/+HRKj7hUGm/oVCp/+EQab/hUGl/4RApf+EP6X/gz+m/4I+pv+DPqX/gz2k/4I9pf+BPKT/gjyj/4E7pP+AOqT/gDqk/385pP+AOaP/gDqj/3kxn/+aY7b//fz9///////9/f7///////Dn9P+IRqn/ejCf/3w1of97M6D/fDOg/3syoP96MaD/ezGf/3sxnv96MJ//eS+e/3ownv95L53/eS6d/3gunv93LZ7/eC2d/3gsnP93LJ3/diuc/3crm/92Kpv/dSmc/3Upm/90KJz/dSib/3Qnmv9zJ5v/cyaa/3Qnmf9zJpn/ciWa/3Mlmf9zJJj/ciSZ/3EjmP9xI5n/cCKY/3EimP9wIZj/byCY/3Agl/9wH5b/bx+X/24elv9uHpf/bR2W/24elf9tHZX/bByW/20clf9sG5T/bBuV/2salP9sGpT/axmT/2sZlP/bx+X///////38/f/+/v7//f3+///////m2O3/cyWZ/2kWkv9rGZT/axiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT+axmT/2oYlERrGZMAaxmTAgAAAAAAAAAAAAAAAGQMkABqGJQBWgCJAFwCihJyI5nghEGm/4ZDqP2FQab/hUGm/4VCp/+EQab/hUCl/4RApv+DP6X/gz+m/4I+pf+DPqX/gj2k/4E9pf+BPKT/gjuj/4E7pP+AOqP/gDqk/385o/+AOaP/fjei/4NApv/n2u7///////39/v///////////5tktv93LJ3/fjah/3w0of97M6D/ezKf/3syoP96MZ//ezGf/3ownv96MJ//eS+e/3ownf95L57/eS6d/3gunv93LZ7/eC2d/3csnP92LJ3/diuc/3crnP92Kpv/dSmb/3UpnP90KJv/dSib/3Qnmv9zJpv/dCea/3Qnmv9zJpr/ciWa/3Mlmf9yJJj/ciSZ/3EjmP9xI5n/cCKY/3Ehl/9wIZj/byCY/3Agl/9wH5b/bx+X/24elv9uHpf/bR2W/24elf9tHZb/bByW/20clf9sG5T/axuV/2salP9tHZX/Yw2O/5lhtf///////v7+//7+/v/+/v7//f3+//////+hbbv/YgyO/2wblf9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlPtrGJT/axmUhWsYkwBrGJMDAAAAAAAAAAAAAAAAbB2VAGsalAJjDo8AZRCQRXcsnf+IRaj+hUKn/oVCp/+GQab/hUGm/4RCp/+FQab/hUCl/4RApv+DP6X/gz+m/4I+pf+DPqT/gj2k/4E8pf+CPKT/gjyj/4E7pP+AOqP/gDqk/385o/+BO6T/eC6e/7yYz////////fz9//38/f//////w6HT/3Yqm/9+N6L/fDSg/3szof98M6D/ezKf/3syoP96MZ//ezGe/3ow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xv+tf8P/qnzB/6h4wP+mdb//pXK9/6Ftuv+cZrf/mGC1/5RZsv+OUK3/h0Wo/4M+pf9/OKL/ejCe/3Uom/9wIZj/bh+X/20clv9tG5X/bh6W/3AimP9zJZn/dSia/3Qom/90KJv/cyaZ/3Ikmf9yJJn/cSOY/3Ajmf9xIpj/cSGX/3AhmP9vIJf/cCCW/28flv9uHpf/bh6W/20dl/9uHpb/bh6W/20dlv9tHZb/3Mjl///////9/P3///////7+/v/9/P3//////7SLyf9jDY7/bBuV/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP5rGJP/ahiUVAAAAAAAAAAAahmUAEcAegBJAHsIciWZzIdFqP+JSar8iEep/4lHqP+IRqj/iEap/4dFqf+GRKn/h0So/4ZDp/+GQ6j/hUKn/4ZBpv+FQab/hEGn/4VBpv+EQKX/hECm/4M/pf+CP6b/gj6l/4M+pP+CPaT/gTyl/4I8pP+CO6P/gTuk/4A6o/9/OqT/gDqj/4E7pP98NKD/eS+e/5FUr/+/nND/59rt//37/f///////v7+/////////////////////////////////////////////////////////////v7+///////+/v7//fv9//r4/P/49Pr/8+z2/+vh8f/j0+r/2sbk/9C33f/CodP/s4rI/6Rxvf+TV7D/gz6l/3Yrnf9uHpb/bBqU/20clf9wIpj/cyWZ/3Qnmv9zJZn/cSOZ/3Ejmf9xIpj/cCGX/3AhmP9vIJf/cCCW/28flv9vH5f/bh6X/20dlv9wIZf/ZxOR/49Trv/+/v7///////7+/v///////v7+//39/v//////o3G9/2ILjf9sHJX/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/WoZlP9rGZN0AAAAAGcZlwBqGJQBVwCGAFkAiBh1KZvkikmq/4pJqv2ISKr/iEep/4lHqf+IRqn/h0Wo/4dFqf+GRKj/h0So/4ZDp/+FQ6j/hUKn/4ZBpv+FQqf/hEKn/4VBpv+EQKX/hECm/4M/pf+CP6b/gj6l/4M+pP+CPaX/gTyk/4I8pP+BO6P/gTuk/4A6o/+AOqT/fzmj/4A6o/+AOqP/ezOg/3gtnf+CPaT/m2S3/7mTzP/UveD/5Nbs/+7l8//x6vX/8+32//bx+P/7+fz//f3+//39/v/+/f7//v3+//79/v/+/v7//v7+///////+/v/////////////////////////////////////////////+/v7///////37/f/07/f/49Pr/8yx2v+yiMf/lVqy/301of9vIJf/ahiT/20dlf9xI5j/cyWa/3Ilmv9yJJn/cSKY/3AhmP9wIJf/cCCX/28fl/9vH5f/cCGY/2wblf9sG5T/18Hi///////9/P3//v7+///+///+/v7///////79/v+MTaz/ZA+P/2wblf9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT9axiU/2sZk5cAAAAAAAAAAGsZlAFfB4sAYQmMKXUpm/eKSqv/ikmq/olIqf+ISKr/iEep/4lGqP+IRqn/h0Wo/4dFqP+GRKj/h0So/4ZDp/+FQ6j/hkKn/4ZBpv+FQqf/hEKn/4VBpv+EQKX/hECm/4M/pf+CP6b/gz6l/4M9pP+CPaX/gTyk/4I8pP+BO6P/gTuk/4A6o/+AOqT/gDmj/4A5ov+AOqP/gDuk/384ov96MZ//dCea/3EjmP9+NqH/llyz/7WNyv/Wv+H/6+Hx//n1+v///v///v7+//7+/v/+/v7//v7+//79/v/9/f7//fz9//38/f/8+/3//Pv9//z7/f/8+vz/+/r8//v5/P/7+fz/+/n8//z6/P/8+/3//v7+///////+/v7//v7+///////8+/3/7+bz/9S94P+ugsT/ikip/3Mmmf9qGJP/aReT/2wblf9vH5b/cCGX/3AimP9wIZj/cCCX/20dlf9nE5H/axqV/8Ce0v///////v3+//7+/v/+/v7//v7+//79/v//////6+Dw/3Unmv9oFZL/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlPxrGJT/ahmUtQAAAABlGpkAbRyVAmMNjgBkDo86ezOf/YtMrP+JSar+ikmp/4lIqf+IR6r/iUep/4hGqP+IRqn/h0Wo/4ZFqf+GRKj/h0Sn/4ZDp/+GQ6j/hkKn/4ZBpv+FQqf/hEGm/4VBpv+EQKX/gz+m/4M/pv+DPqX/gz6l/4M9pP+CPaX/gTyk/4I8o/+BO6P/gTuk/4A6pP9/OaT/fzmj/4A6o/+AO6T/ezOg/3kwnv+MTaz/sYfH/9fA4v/w6PT//fz9//7+/v/+/v7//v7+//39/v/9/P3//fz9//7+/v///////v7+///////////////////////////////////////////////////////////////////////+/v7///////7+/v/8+/3//Pr8//39/v///////v7+///////49fr/4M/o/8Ge0v+ga7r/h0ap/3gtnv9vH5f/bRuV/2salP9tHJX/diqc/5Vasv/XwOL///////79/v/+/v7///////7+/v/+/v7//fz9///////FptX/ZRCP/2wblP9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsZlP9rGJTIAAAAAGkZlQBsG5UCZA6PAGUQkEp8M5//jE2s/olJqv6JSar/ikip/4lIqv+IR6r/iUep/4hGqP+IRqn/h0Wo/4ZFqf+GRKj/h0Sn/4ZDqP+FQqj/hkKn/4VBpv+FQqf/hEGm/4VBpf+EQKb/hECm/4M/pv+CPqb/gz6l/4I9pP+CPaX/gTyk/4I8o/+BO6T/gDqj/4A6pP+BPKX/fTWh/3ovnv+TWLH/xabV//Hq9f///////v7+//7+/v///////v7+//7+/v/+/v7///////7+/v/+/v7///////z7/f/y6/b/6d7v/+DQ6f/XwOL/zrPb/8iq1//DotT/wJ7S/7+c0f/Bn9L/w6PU/8ms2P/Qt93/28fl/+XX7P/w5/T/+/n8///////+/v7///////7+/v/7+vz//fz9///////+/v7//v7+///////49Pr/6Nvv/9vH5f/Qtt3/zbLb/9a/4f/p3O///fv9///////+/f7//v7+///////+/v7//v7+//7+/v/+/v7//////5JWsP9lEI//bRyV/2sZk/9qGJT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT9axmU/2sZlNkAAAAAahiUAGwblQNkD5AAZhGRWnsyn/+MTqz9ikqq/olJq/+KSar/ikip/4lIqv+IR6r/iUep/4hGqP+IRqn/h0Wo/4ZFqf+GRKj/h0On/4ZDqP+FQqf/hkKn/4VBpv+EQqf/hEGm/4VApf+EQKb/gz+l/4M/pv+CPqX/gz6l/4I9pP+CPaX/gjyk/4E7o/+BO6T/gj2k/3own/+IR6r/xaXV//j0+v///////v7+//38/f/8+/3//v7+//7+/v/+/v7///////z6/f/l2O3/zbLb/7mSzP+kcr3/lVqy/4pJqv+AOaL/eS+e/3crnP91KJr/cyaZ/3EjmP9wIpj/byGX/3Ahl/9wIZf/cSOZ/3Mlmv91KJv/eS6e/4M+pf+PUq7/o3C9/7uWzv/YwuP/9/P5///////+/f7//v3+//39/v/9/P3//v7+///////+/v7//////////////////v7+///////+/v7//fz9//7+/v///////v7+//7+/v///////fz9///////Wv+H/axmT/2wblf9sG5T/bBqT/2sZk/9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU6QAAAABqGJQAbBuVA2UQkABmEpFiezKf/41OrPyKSqr+ikqq/4lJq/+KSar/ikip/4lIqv+IR6n/iUep/4hGqP+IRqn/h0Wp/4ZEqf+HRKj/h0On/4ZDqP+FQqf/hkGm/4VBpv+EQab/hUGm/4VBpf+EQKb/gz+l/4M/pv+CPqX/gz6k/4I9pP+BPKX/gj2k/4I8pP97MZ//o3C8/+zh8f///////v7+//37/f/9/P3///////7+/v/+/v7//v7+//38/v//////wZ/S/3csnf94Lp7/dCea/3Uom/91KZz/dyyd/3kvnv96MJ7/ejCf/3oxn/97MZ7/ejGf/3oxn/95MJ//eTCe/3kvnv94Lp3/dy2d/3crnP92KZv/cyWZ/3AhmP9uHpb/bR2W/3Ikmf+LS6v/xqfW//r4/P///////v7+//7+/v/+/v7//v7+//39/v/9/P3//fz9//38/f/9/P3//f3+//7+/v/+/v7////////+///+/v7///////7+/v//////+vj7/45Qrf9mEpD/bh2W/2sblf9rGpT/bBqT/2sZlP9qGJT/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTxAAAAAGsYkwBsG5UDZhCQAGcSkWd7Mp//jU+s/ItLqv6LSqr/ikqr/4lJq/+KSar/iUip/4lIqv+IR6n/iUeo/4hGqP+HRaj/h0Wp/4ZEqP+HRKj/hkOn/4ZDqP+FQqf/hkGm/4VBpv+EQqf/hUGm/4VApf+EQKb/gz+l/4I/pv+CPqX/gj2k/4I+pf+CPaX/fTSg/7iQy//7+fz///////38/f/9/P3//v7+///////t4/L/7OLx///////+/f7//v3+///////k1ev/iEap/3csnf9/OqP/fjah/300oP97M6D/ezKg/3sxn/96MZ//ejCf/3ownv96MJ3/eS+e/3gunv94Lp7/eC2d/3gtnf93LJ3/dyyd/3csnP93LJz/dyyc/3csnf92K53/dSmc/3Ahl/9rGZT/om68//r3+////////v7+///////////////////+///+/v7//v7+//7+/v///////////////////////v7+///////+/v7//Pv9//////+1jcr/ZhKR/24elv9tHJX/bBuU/2wblf9sGpT/bBmT/2sZlP9qGJT/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPcAAAAAahiUAGwblQNmEZAAZxORbXwzn/+NT638i0uq/otLqv+KSqr/ikqr/4lJq/+KSar/iUip/4hIqv+IR6n/iUep/4hGqf+IRqj/h0Wo/4ZEqP+HRKj/hkOn/4VDqP+FQqf/hkKn/4VCp/+EQqf/hUGm/4VApf+EQKb/gz+l/4I/pv+DP6X/gz+l/301of++mtD///////79/v/9/f7//fz+///////49Pr/wJ3R/4ZEqP+OUa7/7OLx///////9/P3//v3+///////o2+7/lFmx/3Qnm/97M6D/fTah/3wzoP97MqD/ejGf/3sxnv96MJ//ejCf/3ownv96MJ3/eS+e/3gunf93Lp7/dy2d/3gtnP93LJz/diud/3crnP92Kpv/diqc/3UpnP91KZz/diuc/3gsnf9qGJP/0rne///////9/P3//v7+//////////////////////////////////////////////////7+/v///////v7+//z7/f//////zLDa/20dlf9tHZX/bR2W/2wclf9tHJT/bBuV/2salf9sGpT/axmT/2sZlP9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU/QAAAABrGZMAbBuUA2YRkABnE5FrfDOf/45PrfyLS6v+i0uq/4tLq/+KSqr/ikqr/4lJqv+KSan/iUip/4hIqv+JR6n/iUao/4hGqf+HRaj/h0Wp/4ZEqP+HRKf/hkOn/4ZDqP+GQqf/hkGm/4VCp/+EQqf/hUGm/4RApf+EQKb/gz+l/4RBp/98NKH/tIvJ///////+/f7//v3+//39/v//////49Tr/5hftf94Lp7/fjei/3kvnv+WXbP/8+z2///////9/P7//f3+///////18Pj/s4jH/301of90KJz/ezGf/301of98M6D/ezKg/3sxn/96MJ//eS+f/3ownv95L53/eS+e/3gunf93Lp7/dy2d/3gsnP93LJ3/dyyd/3crnP93K5z/dyyd/3csnf90KJv/axmU/4VCp//s4vH///////79/v///v////////////////////////////////////////7+/v///////v7+//39/v/9/P3//////9C23P9yJJn/axqV/28flv9tHZX/bB2W/2wclf9tG5T/bBuV/2salf9sGpT/axmT/2oYlP9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT7AAAAAGsZkwBtG5QDZhCPAGcSkGV8M5//jk+t/ItLq/6LS6v/jEyq/4tLq/+KSqr/ikqr/4lJqv+KSan/iUip/4hHqv+JR6n/iUao/4hGqf+HRaj/hkWp/4ZEqP+HRKf/hkOn/4VCqP+GQqf/hkGm/4VCp/+EQab/hUGm/4RApf+FQqb/fjei/5xmuP/49fr///////79/v/9/f7//////9K53v+FQaf/ezKg/4E8pf+AOqP/gTuk/3gvnv+eabn/9e/4///////9/P3//v3+//7+/v//////4M/p/6Z1v/+AOqP/dCea/3YrnP95MJ//ezKg/3wzoP98M6D/ezKg/3syn/96MZ7/ejGe/3ownv95MJ//eTCe/3kvnv94Lp3/dyud/3Qnmv9wIZf/bh6W/3gtnf+hbrz/6t7v///////+/f7//v7+//////////////////////////////////7+/v///////v7+//79/v/9/P3//v7+//7+/v/BntL/cCGX/2wblf9vIJf/bR2W/24elf9tHZX/bByW/20clf9tG5T/bBuV/2salf9sGpT/axmT/2oYlP9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPUAAAAAbBiTAG0blQNlEI8AZxKQYHw0oP+OUK38jEyr/otLq/+MTKv/jEyq/4tLq/+KSqr/iUqr/4lJqv+KSan/iUiq/4hHqv+JR6n/iEao/4hGqf+HRaj/hkWp/4ZEqP+HRKj/hkOo/4VCqP+GQqf/hUGm/4VCp/+EQab/hUGm/4Q/pf+EQab/4M/p///////9/P3//f3+///////Ptd3/gDmi/384ov+CPaX/gDqk/385pP+AOaP/gTuk/3gvnv+cZ7j/8ur1///////9/P3//f3+//79/v/+/v7//////+jb7v+6lM3/lFmx/384ov91Kpz/cyWZ/3Mkmf9zJZr/dCaa/3Qnmv90J5r/cyWZ/3Ekmf9vIZj/byCX/3EjmP91Kpz/gjyk/5hetP+6lM3/49Tr//7+/v/+/v7//v7+///////+/v7///////////////////////7+/v/+/v7//v7+//38/f/9/f7//v7+///////r4PD/n2q6/2oYk/9uHpb/cCCX/24elv9uHpf/bR2W/24elf9tHZb/bByW/20clf9tG5T/bBuV/2salP9sGpP/axmU/2oYlP9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU7wAAAABqGZQAbBuVA2QPjwBmEZBWfDSg/45Qrf2MTKv+jEys/4tLq/+MTKv/i0uq/4tLq/+KSqr/ikqr/4pJqv+JSKn/iUiq/4hHqv+JR6n/iEao/4hGqf+HRaj/hkWp/4dEqP+HQ6f/hkOo/4VCqP+GQqf/hUGm/4RCp/+GQ6f/fjah/6+Dxf///////v3+//38/v//////3szn/4M+pf+AOaP/gz6k/4E7o/+AOqP/gDqk/385o/+AOaP/gTuk/3gvnv+RVbD/4tLq///////9/f7//fz9//38/f/+/f7//v7+///////8+vz/5tjt/8+13P+6lM3/q33C/6Ftu/+bY7f/l160/5hfs/+aYrX/oW27/6l6wf+2jsr/x6jW/9rG5P/u5vP//f3+/////////////v3+//39/v///////v7+//7+///+/v7//v7+//7+/v/+/f7//fz9//39/v/+/v7//v7+///////u5fP/t4/L/3syn/9pFpL/cCKY/3Ahl/9vH5b/bx+X/24elv9tHZf/bR2W/24elf9tHZb/bByV/20clf9sG5T/axuV/2salP9sGpP/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTkAAAAAGoYlABtHJUCYw2PAGQPkEV9NaH/jlCu/oxNrP6MTKv/jEys/4tLq/+MTKv/i0uq/4tLq/+KSqr/iUmr/4pJqv+JSKn/iUiq/4hHqf+JR6n/iEao/4dGqf+HRaj/h0So/4dEqP+HQ6f/hkOo/4VCp/+GQab/hkKn/4RBpv+FQab/4dDp///////9/P3///////Xw+P+WW7L/fTah/4M/pv+BPKT/gTuj/4E7pP+AOqP/gDqk/385o/+AOKL/gTuk/3oyn/+CPqX/xKTU//v4/P///////v3+//38/f/9/P3//fz9//7+/v///////v7+///////+/v7///////7+/v/+/f7//v3+//7+/v///////v7+/////////////v7+///////+/v7//fz9//39/v/+/v7///////7+/v/////////////////////////////////+/v7///////v5/P/g0Oj/roHE/301of9pF5P/bh6W/3Ikmf9wIZj/byCX/3Agl/9vH5b/bh6X/24elv9tHZf/bh6W/24elf9tHZb/bByV/20clP9sG5T/axuV/2salP9sGpP/axmU/2oYlP9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT9axmU/2sZlNMAAAAAAAAAAG0clgJhC44AYgyPNX01ofyPUK7/jE2s/o1NrP+MTKv/jEys/4tLq/+MTKr/i0uq/4pKq/+KSqv/iUmr/4pJqv+JSKn/iUiq/4hHqf+JR6j/iEao/4hGqP+HRan/hkSp/4dEqP+GQ6f/hkOo/4VCp/+HQ6f/gDqj/55ouf/9/P3///////38/f//////wZ/T/3szoP+FQab/gj2k/4E8pf+CPKT/gTuj/4E7pP+AOqP/fzqk/385o/9/OKL/gDuj/303ov95L53/m2S2/9vH5f/+/f7///////7+/v/9/P3//fz9//38/f/9/P3//fz9//39/v/+/v7//v////7+/v/+/v7///////7+/v/9/f7//fz9//38/f/9/P3//v3+//7+/v/+/v7///////////////////////7+/v/+/v7//v7+//z7/f/28Pj/4tLq/8mr2P+pesH/ikqr/3Qnmf9sG5T/cCCX/3Mmmv9xJJn/cCKY/3Ehl/9wIZj/byCX/3Aglv9vH5b/bx+X/24el/9tHZf/bh6W/20dlf9tHZb/bByV/20clP9sG5T/axqV/2walP9rGZP/axmU/2oYlP9rGZT/ahmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP1rGZT/axmTwgAAAAB0IooAaxmUAVwCiQBeBYsjdyud8o5Prf+NTqz+jE2s/41NrP+MTKv/i0us/4tLq/+MTKr/i0uq/4tKq/+KSqv/iUmr/4pJqv+JSKn/iEiq/4hHqf+JR6j/iEap/4dFqP+HRan/hkSo/4dEp/+GQ6f/hkOo/4dEqP9/N6H/vZjP///////9/P3///////bx+P+UWbL/fzmj/4NApv+CPaT/gj2l/4E8pf+CPKT/gTuj/4E7pP+AOqP/fzqk/4A5o/9/OKL/fzmj/4A6o/95MJ7/fTWg/6NwvP/XwuL/+vj7///////+/v7//v7+//38/f/9/P3//v3+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v///////////////////////////////////////v7+//7+/v/+/v7/+fX6/+jc7//Ostv/rYDD/5BSrv95MJ7/bh6W/20blP9xIpf/cyaa/3Ikmf9wIpj/cSKY/3Ahl/9wIZj/byCX/3Agl/9vH5f/bh6W/24el/9tHZf/bh6W/20dlf9tHZb/bByV/20clP9sG5X/axqV/2walP9sGZP/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GsZk/9qGJSsAAAAAGgWlwBqGJMBTwB/AFIAgRN2K5zbjU6t/45Prf2NTqz/jE2s/41NrP+MTKv/jEys/4xMq/+MTKv/i0ur/4pKqv+KSqv/iUmq/4pJqv+JSKn/iEiq/4hHqf+JRqj/iEap/4dFqP+HRan/hkSo/4dEp/+GQ6j/hkSo/4E7pP/Wv+H///////z6/f//////2sXk/4E8pP+EQKb/gz+m/4M+pf+DPaT/gj2l/4E8pf+CPKT/gTuj/4E7pP+AOqP/fzqk/4A5o/+AOKL/fzmj/4A6o/9/OKL/eC6d/3syoP+XXbP/wJ7R/+fa7f/9/P3///////7+/v///////v7+//39/v/9/P3//fz+//79/v/+/v7//v7+//7+/v/+/v7//v7+//7+/v/+/v7///////////////////////////////////////////////////////7+/v/+/v7///////Hp9P/Qt93/pXK9/3w0n/9qGJP/bh+W/3Mmmv9xI5n/cCGX/3Ahl/9wIZj/byCX/3Aflv9vH5f/bh6W/24el/9tHZb/bh6W/20dlf9sHZb/bRyV/20clP9sG5X/axqV/2walP9rGZP/ahiU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT9axmT/2oZlIsAAAAAAAAAAGoYlAAjAF8AFwBYA3IkmcSKSqv/jlCu/I1OrP+MTa3/jE2s/41NrP+MTKv/i0us/4xMq/+MTKr/i0ur/4pKqv+KSqv/iUmq/4pJqv+JSKr/iEeq/4lHqf+JRqj/iEap/4dFqP+HRan/hkSo/4dEqP+FQqf/iEap/+XW7P///////Pr9///////AntH/fTWg/4VCp/+DP6b/gj6m/4M+pf+CPaT/gj2l/4E8pP+CPKP/gTuj/4A6pP+AOqP/fzmk/4A5o/9/OKL/fzmj/385ov+AOqL/fzii/3kvnv93K5z/gj2k/5pitv+7lc3/2sbk//Do9P/+/f7///////7+/v///////v7+///////+/v7//v3+//39/v/9/P3//fz9//38/f/9/P7//f3+//79/v/+/f7//v7+//7+/v/+/v7//v7+//7+/v/+/f7//fz+//39/v/+/v7//v7+//7+/v//////7OLx/7iRy/9/OaL/ahiT/3AimP9yJJn/cSGX/3Ahl/9vIJj/cCCX/3Aflv9vH5f/bh6W/24el/9tHZb/bh6V/20dlf9sHJb/bRyV/20blP9sG5X/axqU/2walP9rGZP/ahiU/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP5qGZT/axmTaAAAAAAAAAAAaheTAI1PrwGBO6UAbh6WqYhHqf+PUa78jU6t/o1OrP+MTq3/jE2s/41NrP+MTKz/i0ur/4xMq/+LS6r/ikqr/4pKqv+KSqv/iUmq/4pIqf+JSKr/iEeq/4lHqf+JRqj/iEap/4dFqP+GRan/h0Wp/4VApv+OUK3/7eTy///////9/P3//////7GGxv9+NqH/hkOn/4M/pv+DP6b/gj6l/4M+pf+CPaT/gT2l/4E8pP+CPKP/gTuk/4E6o/+AOqT/fzmj/4A5o/9/OKL/fzmj/344ov9/OKL/fzmi/385o/98NKD/eC2d/3Upm/96MZ//h0ap/5xmuP+0isj/yazY/97M5//s4vH/9vH4//79/v///////v7+/////////////////////////////////////////////v7+///////+/v7//f3+//38/f/9/f7//v7+//7+/v/+/f7//fz9//79/v/+/v7//////+/n8/+tgcT/ciWZ/2wclf9yJZr/cSGX/3Ahl/9vIJj/cCCX/3Aflv9vH5f/bh6W/20dl/9tHZb/bh6W/20dlv9sHJb/bRyV/2wblP9sG5X/axqU/2walP9rGZP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlP1qGJRJAAAAAAAAAABpGZMAcySaAmwalQBrGZR/hECm/5BTr/yNT63+jk+s/41OrP+MTq3/jU2s/41Mq/+MTKz/i0us/4xMq/+LS6r/i0ur/4pKqv+KSqr/ikmq/4pIqf+JSKr/iEep/4lHqf+IRqj/iEap/4dFqP+HRqn/hECm/5FUr//x6vX///////39/v//////q33C/344ov+GQ6f/hECm/4M/pf+DP6b/gj6l/4M+pf+CPaT/gT2l/4I8pP+CO6P/gTuk/4A6o/+AOqT/fzmj/4A4ov9/OKL/fzmj/344ov9/OKH/fjei/343ov9/OKL/fzii/301of95MJ//diqc/3Qnmv91KZv/ejGf/4I9pf+OT63/m2O2/6h5wP+2jcr/w6LT/86z2//WwOH/387o/+fa7v/t5PL/8en1//Tu9//38vn/+/j8//7+/v///////v7+///////+/v7//fz9//39/v/+/v7//v7+//39/v/9/P3//v7+///////axeT/hEGm/2oYlP9yJZn/cSKY/3AhmP9vIJj/cCCX/28flv9vH5f/bh6W/24el/9tHZb/bh6V/20dlv9sHJb/bRyV/2wblP9sG5X/axqU/2wak/9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP1rGZT/axmU62sZkyUAAAAAAAAAAGodkABtHJUDZRCPAGYSkFZ+N6L/kVSv/Y1Prf6NT63/jk+s/41OrP+MTa3/jU2s/4xMq/+MTKz/i0ur/4xMq/+LS6v/i0ur/4pKq/+JSav/ikmq/4pIqf+JSKr/iEep/4lHqP+IRqj/iEap/4hGqf+EQaf/kFKu/+/n8////////fz+//////+ugsT/fjii/4ZDp/+FQKX/hECm/4M/pf+DP6b/gj6l/4M+pP+CPaT/gTyl/4I8pP+CO6P/gTuk/4A6o/+AOqT/fzmj/4A4ov9/OKL/fjmj/384ov9/N6H/fjei/302of9+NqH/fTah/302of99NqH/fTai/301of97Mp//eS+e/3Yrnf90J5r/cyWZ/3Ikmf9yJJj/dCea/3crnP95L57/ezOg/343ov+CPqX/h0Wo/4tMq/+RVK//mmK2/6t9wv/FpdX/4tPq//v5/P///////v7+//39/v/+/f7///////7+/v/+/f7//fz9///////x6fT/lFqy/2oYlP9yJZn/cSGX/3AhmP9vIJf/cCCX/28fl/9vH5f/bh6W/20dlv9uHpb/bR2V/20dlv9sHJX/bRyU/2wblP9rG5X/axqU/2wZk/9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsZlP9rGZPRaxmTCAAAAAAAAAAAagSPAGwalAJfBosAYQiMLnownviQUq7/jlCt/o5Qrv+OT63/jk+t/41Orf+MTa3/jU2s/4xMq/+MTKz/i0ur/4xMq/+LS6r/i0qr/4pKq/+JSav/ikmq/4lIqf+ISKr/iEep/4lHqP+IRqj/iEap/4ZDqP+LS6z/6Nvu///////8+/3//////7mRzP9+NqH/hkSo/4VBpv+FQaX/hECm/4M/pf+DP6b/gj6l/4M9pP+CPaT/gTyl/4I8pP+CO6P/gTuk/4A6o/9/OqT/fzmj/4A5o/9/OKL/fjmj/384ov9+N6H/fjei/302of9+NqH/fTWg/301of98NKH/ezSh/3wzoP98M6D/fDOg/3wzoP98M6D/fDOg/3syoP97MZ//ejCe/3gunf93LZ3/diuc/3UpnP90J5r/cyaZ/3Ekmf9wIJf/bx+W/28gl/94LZ3/j1Kv/7qVzv/r3/D///////7+/v/9/f7//v7+///////+/v7//fz9///////38/n/mWG1/2kYk/9yJZn/cCGX/3AhmP9vIJf/cCCX/28flv9vH5f/bh6X/20dl/9uHpb/bR2V/20dlv9sHJX/bRyU/2wblP9sG5X/bBqU/2wZk/9rGZT/ahiU/2sZlP9qGJT/axmU/2sZlP5rGZT7axmU/2sYlKlqGZQAAAAAAAAAAABpHZQAaxmTAEoAfABLAH0Mdimb1Y5PrP+PUa79jlCt/41Qrv+NT63/jk6s/41Orf+MTa3/jU2s/4xMq/+MTKz/i0ur/4xMqv+LS6v/ikqq/4pKq/+JSav/ikmq/4lIqf+ISKr/iEep/4lHqf+IRqn/iEap/4RApv/axuX///////z6/P//////y6/a/4A4ov+GQ6f/hEKn/4VBpv+FQKX/hECm/4M/pf+CP6b/gj6l/4M+pP+CPaX/gTyl/4I8pP+BO6P/gTuk/4A6o/9/OqT/gDmj/4A4ov9/OaP/fjij/384ov9+N6H/fjei/302of9+NqD/fTWg/3w0oP98NKH/ezOg/3wzoP97Mp//ezKg/3oxn/97MZ7/ejCf/3own/96MJ7/ejCe/3kvnv95L57/eC+e/3gunv95Lp3/eC6d/3gunv94LZ3/eCyc/3Upm/9wIZj/bR2W/300of+whcb/8Of0///////9/f7//v7+/////////////fz9///////07vf/jU+t/2salP9yJJn/cCGX/3AhmP9vIJf/cCCW/28flv9vH5b/bh6X/20dl/9uHpb/bR2V/2wdlv9sHJX/bRyU/2wblP9rGpX/bBqU/2wZk/9rGZT/ahiU/2sZlP9qGJT/axmU/msZlPxrGJT/axmTdmoYlAAAAAAAAAAAAGkRjgBqGJMAl1yzAYtKqgBxI5inikqq/5BTrvuOUK3+jlCt/41Prf+OT63/jU6s/41Orf+MTa3/jU2s/4xMq/+MTKz/i0ur/4tLqv+LS6v/ikqq/4pKq/+JSar/ikmq/4lIqf+ISKr/iEep/4lHqP+JSKr/gTuk/8Oi1P///////Pv9///////l1+z/h0Wo/4VBpv+FQqf/hEGm/4VBpv+EQKX/gz+m/4M/pf+DP6b/gz6l/4M9pP+CPaX/gTyl/4I8pP+BO6P/gTuk/4A6o/+AOqP/gDmj/4A4ov9/OaP/fjmj/384ov9+N6H/fTei/302of9+NqD/fTWh/3w0oP98NKH/ezOg/3wzn/97Mp//ejGf/3oxn/97MZ7/ejCf/3kvnv96MJ7/eS+d/3kvnv94Lp3/eC6e/3ctnf93LZ3/dyyd/3Yrnf93K5z/dyuc/3csnP93LJ3/cyaa/2wblf+DPaX/z7Tc///////9/f7//v7+/////////////fz9///////k1uz/eC6e/28fl/9xI5j/cCGX/28hmP9vIJf/cCCX/28fl/9uHpb/bh6X/20dlv9uHpb/bR2V/2wdlv9tHJX/bBuU/2wblf9rG5X/bBqU/2sZk/9rGZT/ahiT/2sZlP9rGZT+axmU/moYlP9rGZNDahiUAAAAAAAAAAAAAAAAAGsYkwBwIZgDaRaTAGkWk2+DP6X/klWv/I9Rrf6PUa7/jlCu/41Prv+OT63/jU6s/41Orf+MTaz/jU2r/4xMq/+LS6z/jEyr/4tLqv+LS6v/i0uq/4pKq/+JSar/ikmp/4lIqv+JR6n/iUep/4pIqf+CPaX/pHO+///////+/v7///////z7/f+farr/gTqj/4dDp/+FQqf/hEGm/4VBpf+EQKX/hECm/4M/pf+CPqb/gz6l/4I9pP+CPaX/gTyk/4I8o/+BO6P/gTuk/4A6pP9/OaP/gDmj/4A4ov9/OaP/fjii/384of9+N6H/fjei/342of9+NaD/fTWh/3w0oP98NKH/ezOg/3wzn/97Mp//ejKg/3sxn/97MJ7/ejCf/3kvnv96MJ7/eS+d/3kvnv94Lp3/dy6e/3gtnf93LJz/dyyd/3YrnP93K5z/diqb/3YqnP92Kpz/dyyd/3EimP9wIZf/t4/L//7+/v/+/v7//v7+///////+/v7//fz9//////+6lM3/ahiU/3Ilmf9xIpf/cCGX/28hmP9wIJf/bx+W/28fl/9uHpb/bh6X/20dlv9uHpX/bR2V/2wdlv9tHJX/bRuU/2wblf9rG5X/bBqU/2sZk/9qGJT/axiU/msZlP1qGZT/axiT5GsYkxhrGJMAAAAAAAAAAAAAAAAAbBqTAGwalAJfBYsAYQmMM3oxn/iRVK//j1Gt/o9Rrf6OUK3/jlCu/41Prf+OT63/jU6s/4xOrf+MTaz/jUyr/4xMq/+LS6z/jEyr/4tLqv+LS6v/ikqq/4lJq/+JSar/ikmp/4lIqv+IR6r/iUep/4hFqP+LS6v/59ru///////8+/3//////8qt2f9/OaP/h0So/4VBpv+FQqf/hEGm/4VBpf+EQKb/gz+m/4M/pv+CPqX/gz6l/4I9pP+BPaX/gTyk/4I8o/+BO6P/gDqk/4A6o/9/OaP/gDmj/384ov9/OaP/fjii/384of9+N6L/fTai/342of9+NaD/fTWh/3w0oP98NKH/ezOg/3wzoP97MqD/ejGg/3sxn/96MJ7/eS+f/3kvnv96MJ3/eS+d/3gunv94Lp3/dy2e/3gtnf93LJz/dyyd/3YrnP93K5v/diqc/3Upm/91KZv/diud/3Uom/9sG5X/tIvJ///////+/f7//v7+//7+/v/+/f7///////Pt9v+EQKb/bR2W/3Ekmf9xIpf/cCGX/28gmP9wIJf/bx+W/28fl/9uHpb/bR2X/20dlv9uHpX/bR2V/2wclv9tHJX/bBuU/2wblf9rGpT/bBqT/2sZlP9qGJT+axmU+2sZk/9qGJSwaxiTAGwZkgEAAAAAAAAAAAAAAABsGJEAahqUADYAbwA2AHAHdSibz41Prf+QU6/8j1Gt/o9Rrv+OUK3/jlCu/41Prf+OT6z/jU6s/4xOrf+NTaz/jU2s/4xMrP+LS6z/jEyr/4tLqv+LS6v/ikqr/4lJq/+KSar/ikip/4lIqv+IR6n/ikmp/4I8pP+7lc3///////38/f//////9vD4/5ddtP+CPaX/h0Oo/4VBpv+EQqf/hUGm/4RApf+EQKb/gz+l/4M/pv+CPqX/gz6l/4I9pP+BPaX/gTyk/4I8pP+BO6T/gDqj/4A6pP9/OaP/gDii/384ov9+OaP/fjii/383of9+N6L/fTai/342of99NaD/fTWh/3w0oP98NKH/fDOg/3wyn/97MqD/ejGg/3sxn/96MJ7/eS+f/3kvnv95L53/eS+d/3kunf94Lp7/dy2d/3gtnf93LJz/diyd/3YrnP93Kpv/diqc/3Upm/91KZv/dSmc/3Uom/9vH5b/yq3Z///////9/P7//v7+//7+/v/9/P3//////7SLyP9pF5P/ciWa/3EimP9xIZf/cCGY/28gmP9wIJf/bx+W/28fl/9uHpb/bR2X/24elv9uHpX/bR2W/2wclv9tHJX/bBuU/2sblf9rGpT/bBmT/2sZlP5qGJT8axmU/2oZlG5rGZMAaxmTAwAAAAAAAAAAAAAAAGkYlgBqGZQAeC2dAnEimABsHJWPhkSo/5FVsPuPUa7+kFGt/49Rrv+OUK3/jlCu/41Prf+OT6z/jU6s/4xOrf+NTaz/jUyr/4xMrP+LS6v/jEyr/4tLqv+KSqv/ikqr/4lJq/+KSar/ikip/4lIqv+JSKr/h0Sn/5BSrv/s4fH///////z7/f//////0rne/4E8pP+HRKj/hkKn/4VBpv+EQqf/hUGm/4VApf+EQKb/gz+l/4M/pv+CPqX/gz6k/4I9pP+BPaX/gjyk/4I7o/+BO6T/gDqj/4A6pP9/OaP/gDii/384ov9/OaP/fzii/383of9+N6L/fTah/342oP99NaD/fTWh/3w0oP97M6D/fDOg/3wyn/97MqD/ejGf/3sxn/96MJ7/eS+f/3kvnv96MJ3/eS+e/3gunf94Lp7/dy2d/3gtnP93LJz/diyd/3crnP93Kpv/diqc/3Upm/91KZz/diqc/3EimP+CPaT/8Of0///////+/f7//v7+//38/f//////3Mjl/3Ikmf9xI5j/cSOZ/3EimP9xIZf/cCGY/28gmP9wIJf/bx+W/28fl/9uHpb/bR2X/24elv9uHpX/bR2W/2wclf9tHJT/bBuU/2sblf9rGpT+bBqT/msZlP9qGJT1axiUL2oYlABqGJQCAAAAAAAAAAAAAAAAAAAAAG0flQBsG5UCYQqNAGMNjkN+OKL/klaw/o9Srv6QUq7+j1Gt/49Rrv+OUK3/jlCu/41Prf+OT63/jU6s/4xNrP+NTaz/jEyr/4xMrP+LS6v/jEyr/4tLqv+LS6v/ikqr/4lJq/+KSar/iUip/4lIqv+KSar/gj2k/7SKyP///////f3+//7+/v//////r4LF/384ov+HRan/hUGm/4VBpv+EQqf/hUGm/4VApf+EQKb/gz+l/4I/pv+CPqX/gz6l/4I9pf+BPKT/gjyk/4I7o/+BO6T/gDqj/386pP9/OaP/gDii/384ov9+OaP/fzii/343of9+N6L/fTah/342oP99NaD/fDWh/3w0of97M6H/fDOg/3syn/97MqD/ejGf/3sxnv96MJ7/ejCf/3ownv96MJ3/eS+e/3gunf94Lp7/dy2d/3gtnP93LJ3/diyd/3crnP93Kpv/diqc/3Upm/91KZz/diud/20clf+8l87///////38/f/+/v7//v7+///////07vf/gTyk/24elv9yJJn/cCKZ/3EimP9xIZf/cCGY/28gl/9wIJb/bx+W/28fl/9uHpf/bR2X/24elv9uHpX/bR2W/2wclf9tHJT/bBuU/2sblf5sGpT8axmT/2sZlMBmE5YAaBaUAGoYlAAAAAAAAAAAAAAAAAAAAAAAaRiUAGoXkwBEAHgARQB5CXQnmtKOUK3/kFSv/Y9Srv6QUq7/j1Gt/49Rrv+OUK3/jlCu/45Prf+OTqz/jU6t/4xNrP+NTaz/jEyr/4xMrP+LS6v/jEyq/4tLq/+LS6r/ikqr/4lJq/+KSar/iUip/4lIqv+JR6n/hkOn/9jC4v///////Pv9///////07vf/mmK2/4A6pP+HRan/hUGm/4VCp/+EQqf/hUGm/4RApf+EQKb/gz+l/4I/pv+CPqX/gz2k/4I9pf+BPKT/gjyk/4I8o/+BO6T/gDqj/386pP9/OaP/gDmi/385o/9+OKL/fzii/343of99N6L/fTah/342oP99NaD/fDSg/3w0of97M6H/fDOg/3syn/96MqD/ejGf/3sxnv96MJ//eS+f/3ownv95L53/eS+e/3gunf94Lp7/dy2d/3gtnP93LJ3/diuc/3crnP92Kpv/diqc/3Upm/92K53/bx+X/5NXsf/9+/3///////7+/v/+/v7///////38/f+QUq7/bRyV/3Mlmf9xI5j/cCKZ/3EimP9wIZf/cCGY/28gl/9wIJf/bx+X/24elv9uHpf/bR2W/24elv9tHZX/bR2W/2wclf9tHJT/bBuV/msalfxrGpT/axmTdGsYlABrGZQDaxmTAAAAAAAAAAAAAAAAAAAAAABrGpQAaheSAHYsnQJwIZgAbByVhoZEqP+SVrD8j1Ku/o9Srv+QUq7/j1Gt/49Rrv+OUK3/jU+t/45Prf+OTqz/jU6t/4xNrP+NTav/jEyr/4tLrP+LS6v/jEyq/4tLq/+KSqr/ikqr/4lJqv+KSar/iUip/4lJqv+GQ6f/lVmx/+3k8v///////Pv9///////q3/D/kVWw/4E8pP+HRaj/hUGm/4VCp/+EQab/hUGm/4RApf+EQKb/gz+l/4I+pf+DPqX/gz2k/4I9pf+BPKT/gjyk/4E7o/+BO6T/gDqj/386pP+AOaP/gDii/385o/9+OKL/fzii/343of99N6L/fTah/342of99NaH/fDSg/3w0of97M6D/fDOg/3syn/96MqD/ejGf/3swnv96MJ//eS+f/3ownv95L53/eS+e/3gunf93LZ7/eC2d/3csnP93LJ3/diuc/3crnP92Kpv/diqc/3YqnP9yJZr/fjei/+7l8////////v3+//7+/v///////v7//5Zcs/9sG5T/cyaa/3EjmP9xI5n/cCKY/3Eil/9wIZf/byGY/28gl/9vH5b/bx+X/24elv9uHpf/bR2W/24elf9tHZX/bB2W/2wclf5tG5T+bBuV/2salPFrGZMqaxmTAGsZkwFqGJQAAAAAAAAAAAAAAAAAAAAAAAAAAABpFpMAbRuVAl4EigBhCIwxfDSg9pJWsP+QU67+kFOv/o9Srv+QUq3/j1Gt/45Qrf+OUK7/jU+u/45Prf+NTqz/jU6t/4xNrP+NTav/jEyr/4xMrP+MTKv/jEyq/4tLq/+KSqr/ikqr/4lJqv+KSar/iUip/4pKq/+DPqX/o3G8//fy+f///////Pv9///////m2Oz/kVWw/4A6pP+HRaj/hkKm/4VCp/+EQab/hUGl/4RApf+EQKb/gz+m/4I+pv+DPqX/gj2k/4I9pf+BPKT/gjyj/4E7o/+BO6T/gDqj/385o/+AOaP/fzii/345o/9+OKL/fzii/343of99N6L/fjah/301oP99NaH/fDSg/3w0of97M6D/fDOf/3syn/96MaD/ezGf/3swnv96MJ//eS+e/3ownv95L53/eS+e/3gunv93Lp7/eC2d/3gsnP93LJ3/diuc/3crm/92Kpv/diqc/3Qom/93LZ3/5tnt///////9/f7//v7+///////9/P3/kVWw/20dlf90Jpn/ciSZ/3EjmP9xI5n/cCKY/3Eil/9wIZf/cCGX/3Agl/9wIJb/bx+X/24elv9uHpf/bR2W/24elf9tHZX/bByW/m0clfxsG5T/axqUrG8flQBwIJUBahmUAGsZkwAAAAAAAAAAAAAAAAAAAAAAAAAAAGkakgBqF5MAyq/bAKl6wgBxI5i2i0yr/5JWsPyQU67+kFOv/49Srv+QUq3/j1Gu/45Qrf+OUK7/jU+u/45Prf+NTqz/jE6t/4xNrP+NTaz/jEys/4tLrP+MTKv/i0uq/4tLq/+KSqr/iUqr/4lJqv+KSKn/iUip/4pKq/+CPaT/rYDD//r4+////////Pv9///////p3u//mWG2/384o/+HRKj/hkOn/4RBp/+EQab/hUGm/4RApf+EQKX/gz+m/4I+pv+DPqX/gj2k/4E9pf+BPKT/gjyj/4E7pP+AOqT/gDqk/385o/+AOaP/fzii/345o/9+OKL/fzih/343ov99NqH/fjah/301oP99NaH/fDSg/3w0of98M6D/fDOf/3syoP96MaD/ezGf/3swnv96MJ//eS+e/3ownf95L53/eS+e/3gunv93LZ7/eC2d/3csnP92LJ3/diuc/3crm/92K5z/dCib/3own//o3O////////39/v/+/v7///////bx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qf+JSKr/iEep/4lGqP+IRqn/iEao/4dFqf+HRqn/iEap/4RApv9/OKP/gjyk/5JVr/+vgsX/zrTc/+re7//8+vz///////7+/v/+/v7///////7+/v/+/f7//v7+///////+/v7//////////////////////////////////////////////////v7+///////+/v7//Pv9//z7/f/9/P3//v7+///////+/v7////+//7+/v/t5PL/zbHb/6Ftu/9/OKL/ciSY/3crnP96MZ//eC6e/3ctnf94LZ3/dyyd/3Yrnf93K5z/diqb/3YqnP91KZv/dSmc/3Qom/90J5r/dCeb/3Mmm/90J5r/cyaZ/3Mmmv9yJZn+cyWY/nIkmfxxIpj/bh2WZXAhmABwIZgCXQiLAG0clQBtFIwAAAAAAAAAAAAAAAAAAAAAAAAAAABgB4wweC2d8ZVbs/+VW7P+lVyz/pVbsv+TV7H/kVWw/5FVsP+RVK//kVSv/5FUr/+TV7H/k1ix/5JWsf+RVbD/klWw/5FUr/+RVbD/j1Ku/45Prf+OUK3/jU+t/41Orf+NTqz/jU2s/4xNrP+MTaz/jEys/4xNrf+OUK7/jlCt/41OrP+MTaz/jU+t/4hGqf+HRan/jE2r/4tLqv+LS6v/ikqq/4pKq/+JSar/ikmp/4lIqf+IR6r/iUep/4lGqP+IRqn/h0Wo/4dFqf+HRan/h0So/4hFqf+GRKj/gz6l/383ov9/OaP/ikqq/5xmt/+1jMn/zbHb/+LT6v/x6fX//Pv9///////+/v7////////////////////////////////////////////////////////////////////////////+/v7//v7+///////59fr/6d3v/9W+4f+8l8//nWe4/4Q/pv91KZv/cyWa/3gtnv97Mp//ejCe/3kvnv94Lp3/dy6e/3gtnf94LJz/dyyd/3YrnP93K5z/diqb/3YqnP91KZz/dCib/3Uom/91KJr/dCeb/3Mmmv90J5r/cyaZ/nImmv9zJZn7ciSY/24el6WTWq0AfDWeAWoWlABvIJUAaxqVAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGQOj1x9NqH/mF+0/pVbs/6RVbD/kVWw/5lhtf+dZ7j/nmi5/55ouP+dZ7f/mWG1/5FT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AAAAAAAAAAAAAP///////jwAAAHj//////////////PAAAAAHP/////////////OAAAAAAOf///////////+cAAAAAAAZ///////////+YAAAAAAABn///////////IAAAAAAAAGf//////////IAAAAAAAAAZ//////////IAAAAAAAAABP/////////kAAAAAAAAAAE/////////kAAAAAAAAAAA3////////yAAAAAAAAAAAC////////6AAAAAAAAAAAAT///////9AAAAAAAAAAAABf//////8gAAAAAAAAAAAAL//////+QAAAAAAAAAAAABf//////QAAAAAAAAAAAAAL//////oAAAAAAAAAAAAABf/////0AAAAAAAAAAAAAAL/////6AAAAAAAAAAAAAABf////9AAAAAAAAAAAAAAAL/////QAAAAAAAAAAAAAABf////oAAAAAAAAAAAAAAAL////0AAAAAAAAAAAAAAABf///6AAAAAAAAAAAAAAAAf///9AAAAAAAAAAAAAAAAC////QAAAAAAAAAAAAAAAAX///oAAAAAAAAAAAAAAAAC///0AAAAAAAAAAAAAAAAAv//9AAAAAAAAAAAAAAAAAF//+gAAAAAAAAAAAAAAAAA///QAAAAAAAAAAAAAAAAAL//0AAAAAAAAAAAAAAAAABf/6AAAAAAAAAAAAAAAAAAX/+gAAAAAAAAAAAAAAAAAC//QAAAAAAAAAAAAAAAAAAv/0AAAAAAAAAAAAAAAAAAF/6AAAAAAAAAAAAAAAAAABf+gAAAAAAAAAAAAAAAAAAL/wAAAAAAAAAAAAAAAAAAC/0AAAAAAAAAAAAAAAAAAAf9AAAAAAAAAAAAAAAAAAAF+gAAAAAAAAAAAAAAAAAABfoAAAAAAAAAAAAAAAAAAAL6AAAAAAAAAAAAAAAAAAAC9AAAAAAAAAAAAAAAAAAAAvQAAAAAAAAAAAAAAAAAAAH0AAAAAAAAAAAAAAAAAAAB/AAAAAAAAAAAAAAAAAAAAegAAAAAAAAAAAAAAAAAAAHoAAAAAAAAAAAAAAAAAAAB6AAAAAAAAAAAAAAAAAAAAOgAAAAAAAAAAAAAAAAAAAD4AAAAAAAAAAAAAAAAAAAA8AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADQAAAAAAAAAAAAAAAAAAAA0AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADQAAAAAAAAAAAAAAAAAAAA0AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADQAAAAAAAAAAAAAAAAAAAA0AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADQAAAAAAAAAAAAAAAAAAAA0AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADQAAAAAAAAAAAAAAAAAAAA0AAAAAAAAAAAAAAAAAAAANAAAAAAAAAAAAAAAAAAAADwAAAAAAAAAAAAAAAAAAAA6AAAAAAAAAAAAAAAAAAAAOgAAAAAAAAAAAAAAAAAAADoAAAAAAAAAAAAAAAAAAAA6AAAAAAAAAAAAAAAAAAAAfgAAAAAAAAAAAAAAAAAAAH0AAAAAAAAAAAAAAAAAAAB9AAAAAAAAAAAAAAAAAAAAfQAAAAAAAAAAAAAAAAAAAL8AAAAAAAAAAAAAAAAAAAC+gAAAAAAAAAAAAAAAAAAAvoAAAAAAAAAAAAAAAAAAAf+AAAAAAAAAAAAAAAAAAAF/QAAAAAAAAAAAAAAAAAABf0AAAAAAAAAAAAAAAAAAAv/gAAAAAAAAAAAAAAAAAAL/oAAAAAAAAAAAAAAAAAAD/+AAAAAAAAAAAAAAAAAABf/QAAAAAAAAAAAAAAAAAAX+EAAAAAAAAAAAAAAAAAAL+egAAAAAAAAAAAAAAAAAC/YAAAAAAAAAAAAAAAAAABfoAAAAAAAAAAAAAAAAAAAf0AAAAAAAAAAAAAAAAAAAL+AAAAAAAAAAAAAAAAAAAF/gAAAAAAAAAAAAAAAAAABfwAAAAAAAAAAAAAAAAAAAv8AAAAAAAAAAAAAAAAAAAX/AAAAAAAAAAAAAAAAAAAF/wAAAAAAAAAAAAAAAAAAC/8AAAAAAAAAAAAAAAAAABf/AAAAAAAAAAAAAAAAAAAv/wAAAAAAAAAAAAAAAAAAL/8AAAAAAAAAAAAAAAAAAF//AAAAAAAAAAAAAAAAAAC//wAAAAAAAAAAAAAAAAABf/8AAAAAAAAAAAAAAAAAAv//AAAAAAAAAAAAAAAAAAX//wAAAAAAAAAAAAAAAAAL//8AAAAAAAAAAAAAAAAAF///AAAAAAAAAAAAAAAAAE///wAAAAAAAAAAAAAAAACf//8AAAAAAAAAAAAAAAABf///AAAAAAAAAAAAAAAABP///wAAAAAAAAAAAAAAAAn///8AAAAAAAAAAAAAAAAn////AAAAAAAAAAAAAAAAT////wAAAAAEAAAAAAAAAT////8AAAAABQAAAAAAAAT/////AAAAAATAAAAAAAA7/////wAAAAAFOAAAAAAAz/////8AAAAABecAAAAABz//////AAAAAAf54AAAADn//////4AAAAAL/x8AAAfP//////+AAAAAC//x///4f///////QAAAABf//4AAD////////6AAAAAv///////////////IAAAA3///////////////KAAAABAAAAAgAAAAAQAgAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAEUAdwAuAGYAdyydA28flgBuHpUWbx+XgW0cldRrGZT3axmU/msZlO1qGJO4ahmUUGsZkwBqGJQBaxmUAWoZlACUV7IAejCeAnYqnABzJppSdSmb6XIkmf9tHJX9ahiT/mkXk/9sG5X+bBuV/msZlP9qGJO1axmUF2sZkwBrGZMBezGfA3swnwB3K5xmeTCe/28gl/1vIJf7eS+e/Xkunv5wIpj+Yw2O/mAJjPxrGZT7axqU/2oYk9lrGZQYaxmTAHkvngB3LZ0zfjeh+3YqnP2eaLj8xKTU/7iSzP63kMv/u5XO/7qUzf6cZbf/ZxOR/mgVkvlrGpT/ahiTsWoYkwCQUK8AgDqjsHszoP+TWLH71L3g/4M+pf51KZv/cSOY/3Mlmf9/OKL/roLF/8eo1v5uHZb+aRaS/GsZlP9qGJNSeS6eJYRBpvN8NKD/m2S3/sOi1P5rGZT/cCGX/2wblf9pF5P/ZxOR/1kAh/+sfsP/upTN/l8Hi/xsG5X/axmUs4E7pFmGRKj/hUKn/X43ov6ibrz/t4/L/7GGxv+ugsX/pXO+/5BTr/9pFpL/mmK2/9K53v5hCo3+bBuV/2sZlOmEQKZyiUip/4hGqfx/OaL+nGW3/9vH5f/i0ur/pXS+/5detP+7lc3/6d7v//Tu9/+farn/YwyO/mwblf9rGZT8hUKnaYxMq/+HRKj8j1Gu/suu2f9/OaP/pHG9/8Sk1P/Nsdr/6Nzv//////+3kMv/YguO/24elv5rGZP/ahiU9YRApkGNTq3/iEep/pVbsv7Bn9L/eS+e/3wz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PaX/zbHa///////9/f7//v7+///////8+/3///////Tv9/9+OKL/aBSR/2wblP9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU/HIkmABwIZdHh0So/4xNq/yKSar+iUiq/4hHqf+HRaj/hkSo/4ZCp/+FQqb/hUGm/384ov/Sut///////8603P9+N6L/rH7D///////q3/D/n2q6/3own/9zJZr/cyWZ/3Mmmf9xI5n/byCX/24elv9sG5X/to3K///////7+fz//Pv9//39/v//////9/L5/5JWsP9mEpD/bh+W/2walP9rGZT/ahiT/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU/XMmmQBxI5hAiEap/41NrP2LS6r+ikmq/4lIqf+IRqn/h0Wo/4ZDqP+HRaj/fTah/6x+w///////y67Z/3gtnf9/OKP/eC2d/6p8wv/59vv//////+fa7v+/m9D/qnzC/6Rxvf+mdb7/sYfH/8ip1//r4PD///////7+/v/+/v7///////j0+v/PtNz/gj2l/2gUkf9wIZj/bR2W/2wclf9sGpT/axmT/2oYlP9rGZT/axiU/2sZlP5rGZT/axmU93EimABvIJcwiEap/41Orf6LS6v+i0ur/4lJqv+JSKn/iEap/4dFqP+HRaj/gDmj/9/N6P/69/v/h0Wo/4A6o/+DPqX/gTyl/3gtnf+NTqz/zrLb//v5/P///////v7+/////////////v7+///////+/v7//Pr9//38/f/38vn/zrPb/49Srv9hCo3/aBSS/3Qnmv9vH5b/bh6W/20dlv9sHJX/bBqU/2sZk/9qGJT/axmU/2oZlP5rGZT/ahmU6GMMjgBgCYwZh0So845Qrf+MTKz+jEur/4pLq/+JSar/iUep/4hHqf+FQqf/iUiq//j1+v/ZxOP/fDSg/4RBp/+CPKT/gTuk/4I8pf97M6D/diuc/4pJqv+sfsL/zbHb/+XX7P/z7fb//Pr9///////+/v7//v7+//7+/v/+/v7//v7+//Ls9v/Nstv/iEap/2gVkv9xI5j/bx+X/24elv9tHZX/bBuV/2salP9rGZP/ahiU/msZlP1rGZT/axmUygAANgAAAAAChECm149Rrv+NTaz8jEys/otLq/+KSqr/iUmq/4lIqf+GQ6f/jU6t//37/f/S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oNL/wqDS/8eo1v/Qtt3/3Mnm/+vg8P/6+Pz///////7+/v/+/v7///////Tu9//Ostv/mF+0/2salP9hCo3/axmU/2wblf9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlPxrGZT/ahmUTmsZkwBrGJMCbBiSAGsZkwAAAAAAAAAAAGsalABdBYoAgDqjAoE7pABxI5iXfzij/4A5o/t/OKL/fjei/n42of99NaH/fDSh/3w0oP97Mp//eCyd/8mr2P///////v3+//7+/v/g0On/sITG/45Prf96MZ//cSKY/20blf9rGZT/aReT/2kWkv9pFpL/aReT/20clf90J5r/gTyk/5ddtP+1jMn/18Li//by+f///////v7+///////k1uz/om67/2gVkv9kD4//bRyV/2sZlP9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT8axiU/2sZk9JrGZMJaxmTAGsZkwBqGJQAAAAAAGUYmQBsGpUAbh6WAmcUkQBmE5EyezKf+oI9pP9/OaP+fzmj/n84ov9+N6L/fjah/301oP99NqL/dCib/8Kg0v///////v7+//Ps9v+mdb7/dCib/24elv9xI5j/dCia/3UqnP92Kpz/dimc/3Upm/91KJv/dCea/3Ilmv9xI5j/bh+W/2sZlP9nEpH/ZA6P/2kWkv98NaH/pHK9/9jC4v/9/P3//v7+///////fzej/h0Wo/2EJjf9sG5X/axmU/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPtrGZT/axiTcGoZlABqGZQDbBiSAAAAAABrGZQAZhGQAIlIqQKLS6sAdCibpYI8pP+CPKT7gTuj/4A5o/5/OaP/fzii/343ov9/OKL/eC2d/5Rasv/8+/3///////Pt9v+PUK3/bh6W/3kwnv96MJ//eC6d/3csnP92Kpz/dSmb/3Qnm/90J5r/cyaa/3Mlmf9yJJj/cSOZ/3EimP9xIpj/cSKY/3EimP9vH5f/aRiT/2MNjv9nE5H/jk+t/9fB4v////////////v6/P+qe8H/YgyO/2sZlP9rGZT/ahiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT9axmU/2sYk95rGZMPaxmTAGsZkwEAAAAAbRyVAGwblQJkDo8AZA6OKnw0oPeEQKb/gTyk/oE8pP6AOqT/gDmj/385o/9/OKL/gDqj/3Uom//BntH///////////+tgMT/bx+X/301of95L57/eC6e/3ctnf93LJz/diuc/3UpnP91KJv/dCeb/3Qnmv9zJpr/ciWZ/3IkmP9xI5j/cCKY/3Agl/9vH5f/bh6X/24fl/9vIJf/bh6W/2UPj/9lEI//p3e///n2+v///////////7qVzf9jDY7/axqU/2sZlP9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlPtrGZT/ahmUaGsYkwBrGJMDahaTAGgUkgB7MZ8DejGfAHIlmYeCPaT/g0Cm+4I+pf6CPaT/gTuk/4A6pP+AOaP/fzmj/4A6o/94Lp3/2MPj///////18Pj/hkOo/3gunv97MZ//eS+e/3kvnf94Lp7/dy2d/3csnP92Kpz/dSmc/3Uom/90J5v/cyaa/3Mmmf9yJZn/ciSY/3EimP9wIZj/cCCX/28flv9uHpf/bR2W/20dlv9uH5b/bR2V/18Hi/+PUK7/9vH4//39/v//////sofH/2IKjf9sHJX/ahiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT8axiT/2oZlMhsGZIAbRiRAGoZlABrGZMAQAB0ADwAcQl6MZ/chUKn/4M/pf2DP6X+gj2l/4I8pP+BO6P/gDqk/4A5o/+AOqP/ejGf/93L5v//////6+Dw/385o/96MZ//ezKf/3own/95L57/eS+d/3gunf93LJ3/dyuc/3YqnP91KZz/dSib/3Qnmv9zJpr/cyaZ/3Ilmf9xI5n/cSKY/3AhmP9wIJf/bx+W/24el/9tHZb/bR2V/20clf9vIJf/YAiL/5NYsf/9/f7//fz+//////+QU6//Yw2O/2wblf9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/moZlP5rGJP/axiTOGsYkwBrGJMAcCGYAm0clQBrGpRFgTqj/4ZDp/2EQKb+hECm/4M+pf+CPaX/gjyk/4E7o/+AOqT/gTuk/3gtnf/PtNz//////+/m8/+CPqX/ejGf/3szoP96MZ//ejCe/3kvnv95L53/eC2d/3csnf93K5z/diqb/3UpnP90KJv/dCea/3Mmmv9zJpn/ciSZ/3Ejmf9xIpj/cCGY/28gl/9vH5b/bh6W/20dlv9tHZX/bBuV/28flv9gB4v/u5XN////////////383o/2oYk/9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT7axmU/2oZlIVrGJMAaRaTAH00oAN8M6AAdCeajIRApv+GQqf7hUGm/oRBpv+DP6b/gz6l/4I9pf+BPKT/gTuj/4E8pf94LZ3/qnzC///////+/v7/llyz/3csnf99NqH/ezKf/3oxn/96MJ7/eS+e/3gunv94LZ3/dyyd/3crnP92Kpv/dSmc/3Qom/90J5r/cyaa/3Mlmf9yJJj/cSOZ/3EimP9wIZf/cCCX/28flv9uHpb/bR2W/20clf9tHZX/aBWS/3kvnv/28vn///////////+UWbH/YwyO/2wblf9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GsZlP9rGZTJaheUAGsZlAD///8A////AHoxn8qHRaj/hkKn/IVCp/6FQab/hECm/4M/pv+DPqX/gj2k/4E8pP+BO6T/fzmj/4A5o//h0en//////9C33f90J5r/fDSg/343ov98M6D/ezGf/3ownv95L57/eC6e/3gtnf93LJ3/diuc/3Yqm/91KZv/dCeb/3Qnmv9zJpr/cyWZ/3IkmP9xI5n/cCKY/3Ahl/9vIJf/bh6X/24elv9tHZb/bRyV/24dlv9lEI//1L3g////////////v5zR/2EKjf9tHJX/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmT9GsZkydtHJUBXgWKAF0Diht/OKLyiEep/4ZDqP6GQ6f/hUKn/4VBpv+EQKX/gz+m/4I+pf+CPaT/gTyk/4I9pP97MqD/k1ew//Pt9v//////wqDS/3kvnv9zJpr/eC2d/3oxn/97Mp//ezKf/3syn/96MZ//eTCe/3kvnv95Lp3/eC2d/3csnf92Kpz/dSmb/3Qomv9zJpn/ciSY/3Ejmf9wIpj/cCGX/28fl/9uHpf/bh6W/20dlv9vH5b/Yw2O/8Sj1P///////////9fB4v9lEZD/bBqU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/WsZlP9rGZNbcCGXA2salABqGJNAgj2k/4lIqf2GRKj+hkOo/4ZCp/+FQqf/hUGm/4RApf+DP6b/gj6l/4I9pP+BO6T/gj2l/3own/+SVbD/5dfs///////p3e//rYDE/4lJqv97MqD/diqc/3Qmmv9yJJn/ciOY/3AimP9vIJf/bh2W/20clf9sG5X/bByV/24elv9wIZj/ciWZ/3Qnmv90J5r/ciWZ/3EimP9wIJf/bx+X/24el/9tHZb/byCX/2URkP/Qtt3////////////dyub/aBSR/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlPxrGZP/axiTjXYrnAN1KZsAciSZZoVCp/+JSKn7h0Wo/odEqP+GQ6j/hkKn/4VCp/+EQab/hECl/4M/pf+CPqX/gj2k/4E7o/+CPaX/ezKg/385o/+yiMf/7ePy///////9/f7/7+bz/93K5v/VvuD/0rne/9C23P/Mr9r/xqfW/7+d0f+3j8v/rH7D/59ruv+RVK//gj2l/3Uom/9rGpT/aRaS/20clf9xIpj/ciSZ/3Ikmf9xIpj/cSOZ/2oZlP91Kpv/8+32////////////07vf/2QPj/9sGpT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT8axmU/2oYlLd9NaEEfTWhAHYqm4aIRqj/iUiq+4hGqf6HRan/h0So/4ZDqP+GQqf/hUGm/4RBpv+EQKX/gz+l/4I9pf+CPKT/gTuj/4E8pP9/OKP/eC6e/3w0oP+MTaz/rYDD/9W+4f/28fj///////7+/v/+/v7////////////////////////////+/v7///////v5/P/p3e//zbLb/6x+w/+KSan/cSOY/2gVkv9nEpH/ZxOR/2YRkP9lEZD/wqHT///////7+fz//////7mSzP9hCYz/bRyV/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsYlP9rGZPXh0WoA4dEpwB5L56diUiq/4pJqvuJR6n/iEap/4dFqf+HRKj/hkOn/4VCp/+FQab/hEGm/4M/pf+DP6X/gj2l/4I8pP+BO6T/gj2l/3oxn/+CPqX/qnzC/9W+4f/w5/T/+/j8//7+/v//////+/j8//Hp9P/o2+7/4M/p/9zJ5v/byOX/383o/+fZ7f/z7Pb//v7+//7+/v///////Pr8/+XX7P/FptX/rH7D/6Ftu/+pecD/2MPj///////+/f7//f3+//////+NT63/ZA6P/2wblf9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU645PrAONTasAeC6drIpKqv+KSqv7iUiq/4lHqf+IRqj/h0Wp/4dEqP+GQ6f/hUKn/4VBpv+EQKb/gz+m/4I+pf+CPaX/gj2l/3ownv+mdb//7ePy///////+/v7//v3+///////Gp9b/nGW3/5FVsP+DPqX/ezKg/3csnP90KJv/cyaa/3Qom/93K5z/fTSg/4lIqv+eaLn/0Lbd//7+/v///////v7+//////////////////7+/v/+/f7//Pr9///////Tu9//aBSR/2wblf9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPeSVa8CkFKuAHgtnbKLSqr/i0ur+4pJqv+JSKr/iUep/4hGqP+HRan/hkSo/4ZDp/+FQqf/hUGm/4RApv+DP6b/hECm/3szoP+9mM/////////////n2u7/vJfO//bx+P//////zrPb/3YqnP9vIJf/eS+e/3wzoP98M6D/ezKf/3oxn/95L57/eC2d/3YrnP9zJpr/byCY/2QNjv+1jcr///////v4/P/8+/3//fz9//z7/f/8+/3//Pr9///////y6vX/gDqj/2gVkv9tHZX/axqU/2sZk/9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT9klavApBSrgB4Lp2yi0ur/4xMq/uKSqv/ikmq/4lIqv+JR6n/iEao/4dFqf+GRKj/hkKn/4VCp/+FQab/hkOn/3wzoP+yiMj////////////Os9z/hD+m/3EjmP+rfcL//v3+///////o2+7/om+8/3syn/9yJJn/cyaa/3Uom/91KZv/dCib/3Ilmf9wIpj/bx6W/3AimP+BO6T/07rf///////9/P7//v7+//38/f/8+/3//f3+///////v5vP/iUiq/2cTkf9vIJf/bByV/2wblf9rGpT/axmT/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP5rGZT/axmU/I9QrQONTqwAeS+eqoxLq/+MTKv7i0uq/4pKq/+KSar/iUip/4hHqf+IRqj/h0Wo/4ZDqP+GQqf/hkOn/4I9pP+QU67/9O73///////MsNr/ei+e/384ov+DP6b/dyyd/6d4wP/49fr///////79/v/k1uz/vJfO/6Fsu/+SVrD/jE2s/4xNq/+QVK//nGa3/7GGx//RuN7/9vH4///////+/v7///7////////+/v7///////v5/P/Hqdf/eTCe/2kVkv9wIZj/bR2W/20dlv9sHJX/bBuU/2salP9rGZP/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlPWJR6kDiEapAHoxn5qMTKz/jE2s+4xLq/+LS6r/ikqr/4lJqv+JSKn/iEep/4hGqf+HRKj/hkOo/4dFqP99NaH/vpvQ///////p3e//hECm/4A7o/+DPaX/gDqj/4E8pP93LZ3/kVWv/9rG5P/+/v///v7+//7+/v///////fz9//r3+//7+fz//v7+///////+/v7///////7+/v/9/f7///////7+/v/z7fb/1sDh/66BxP+CPaT/ahiT/28flv9xI5j/bx+W/24el/9tHZb/bR2V/2wclf9sG5T/axmU/2sZk/9qGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTofziiBH84ogB3LJ2Bi0ur/41OrfuMTKz+jEyr/4tLqv+KSqv/iUmq/4lIqf+IRqn/h0Wp/4dEqP+HRKj/gz2l/+DP6P//////uJHM/3oxn/+FQab/gTyk/4E7pP+AOqP/gTyk/3syn/96MJ7/n2q6/9K53v/07/f///////7+/v/+/v7//v7+///////9/f7/+/n8//v5/P/8+v3//Pv9//z7/f/8+v3/7eTy/8yw2v+kcb3/fzii/2salP9uHpb/ciSZ/28gl/9vH5b/bh6X/20dlv9tHZX/bByV/2walP9rGZT/ahiT/2sZlP9rGZT/axmU/2sZlP1rGZT/ahmU0nctnQN2KpwAciWZYIlJqv+OUK37jE2s/oxMrP+LS6v/i0uq/4pKqv+JSar/iUep/4hGqf+IRqn/hUKn/4tLq//w6PT//////59quf9+N6L/hECm/4I9pf+CPKT/gTuj/4A6o/+BO6T/gDqj/3gtnf93LJz/iEap/6V0vv/Do9T/3crm/+7l8//59/v///////7+/v/+/v7////////////////////////////+/v7///////r3+//JrNj/gTuk/2gWkv9yJJn/byCX/28flv9uHpb/bR2W/20dlf9sG5X/bBqU/2sZlP9qGJP/axmU/2sZlP5rGZT8axmU/2sZk7BxI5gCaxmUAGoYkzqHRaj/j1Ku/o1OrP6MTaz/jEyr/4tLq/+LS6r/ikmq/4lIqv+JR6n/iUep/4VCp/+PUa7/9e/3//////+aY7b/gDmi/4RBpv+DPqX/gj2l/4E8pP+BO6P/gDqj/4A5o/+AO6P/gDqj/3syn/92KZv/cyea/3csnf9/OaP/ikqr/5dds/+ib7z/rH/D/7WMyf+7ls7/wqDS/86z2//m2O3//Pr9///////+/f7///////by+f+earn/aBWS/3Ijmf9vIJf/bx+W/24elv9tHZb/bRyV/2wblf9sGpT/axmT/2oYlP9rGZT+axmU/GsZlP9rGZOFbRyVAVUAgwBTAIIVhECl7JBTr/+NT639jU6s/4xNrP+MTKv/i0ur/4pKq/+KSar/iUiq/4lIqf+HRaj/iUmq/+zi8f//////pHG9/384ov+GQ6f/gz+m/4M+pf+CPaX/gjyk/4E7o/+AOqP/fzmj/384ov9/OKL/fzmi/384ov99NaH/ejGf/3csnf90J5r/ciSZ/3Eil/9vIJf/bh+W/24elv9vH5b/diqc/45Qrv/GptX/+/j8//79/v/9/P3//////6JwvP9oFZL/ciSZ/28gl/9vHpb/bh6W/20dlv9tHJX/bBuV/2salP9rGZP/ahiU/msZlP1rGZT/axmTUmsZlADw5/IB2cTiAIA5osCQUq7/jlCt/I1Prf6NTqz/jE2s/4xMq/+LS6v/ikqr/4pJqv+JSKr/ikmq/4I9pP/VvuH//////7+c0f99NKH/h0Wo/4RApf+DP6b/gj6l/4I9pP+BPKT/gTuj/4A5o/9/OaP/fzii/343ov9+NqH/fTWh/3w0of98NKD/fDSg/3wzoP98M6D/ezOf/3oxn/95MJ7/eS+e/3YrnP9wIZj/axmU/5BTr//r4PH///////79/v/49fr/hkSo/2walP9xIpj/byCX/24el/9uHpb/bR2W/2wclf9sG5T/axqU/2sZk/5qGJT/axmU7moYkx9oFJEAgTykA4A7pAB3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/////v7+///////+/v///v7+///////9/P7/+ff7//Pt9v/o3O7/28bk/8eo1v+whMX/lVuy/301of9rGZP/YguN/2MMjv9oFZL/bBuV/2wblf9rGZT/ahiT/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlPtrGJT/ahiUbWsYkwBrGJMDbBeSAGoZlAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAahmVAGsalABmEpAAhkKoAZRYswBtHZasfDSg/385ovt9NqH/fjah/n01oP98NKH/fDSh/3wzoP97Mp//ejGf/3sxn/96MJ//eS+e/3kvnf96MZ//dy2d/28hl/99NaH/p3a//9O73//z7Pb///////7+/v/+/v7///////////////////////////////////////////////////////////////////////////////////////7+/v/+/v7///////bx+P/fzuj/vZnP/5Vbsv90J5r/YwyO/2MNjv9qGJP/bRyV/2sZlP9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/2sZlP1rGJP/ahmU82oZlC9qGZQAahmUAmsZkwBrGZMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAahiUAG0fmABtHZUDaRaSAGkWkl13LJ3/gDuk/H44ov5+N6H+fjeh/342of99NaD/fDSh/3w0oP98M6D/ezKg/3oxn/97MZ7/ejCf/3syn/90J5r/ezKf/7OJyP/t5PL///////7+/v/+/v7//v3+///////+/v7//v7+///////9/P3/9vH4/+/m8//q3/D/6Nvu/+fZ7f/o3O7/6+Hx//Hp9P/49Pr//f3+///////+/v7//v7//////////////v3+//7+/v///////v7+///////v5vP/xqbV/5JWsP9sGpT/YQqN/2kWkv9tHJX/axmU/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9r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PaT/gjyk/4E7o/+AOqT/gDqk/4A5o/+AOqP/ezKg/5BTrv/6+Pv//v7+///////n2u7/fjei/3syoP98M6D/ezKg/3sxn/96MJ7/eTCf/3ownv95L53/eC6e/3ctnf94LZ3/dyyc/3YrnP93K5z/diqb/3UpnP90KJv/dSib/3Qnm/9zJpr/cyaZ/3Immv9zJZn/ciSY/3Ejmf9xI5n/cSKY/3Ahl/9vIJf/cCCX/28flv9uHpf/bR2W/24elv9tHZX/bByV/20clP9sG5X/bRyV/2sYk/9qF5P/2sbk///////9/P3//fz9//////+dZ7j/YgyO/2wclf9qGJP/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlPtrGZT/axmTjGsYkwBrGJMDAAAAAAAAAABqGJMAcSOZA24elgBsG5V3gDqj/4dDp/uEQab+hEGm/4RApf+DP6b/gz+m/4I+pf+CPaT/gj2l/4I8pP+BO6P/gDqk/386pP+AOaP/fjei/4M/pv/s4vH////////////w6PT/hUKn/3oxn/98NKH/fDKf/3syoP97MZ//ejCe/3kvnv96MJ7/eS+d/3gunv94LZ3/eC2d/3csnf92K5z/diqb/3Yqm/91KZz/dCib/3Qnm/90J5v/dCea/3Mmmf9yJZr/cyWZ/3IkmP9xI5n/cSKY/3EimP9wIZj/byCX/3Agl/9vH5b/bh6X/24elv9uHpb/bR2W/2wclf9sG5T/axqU/24dlf9lD4//iUep//38/f///////fz9///////k1uv/bh6W/2oXk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlPxrGZT/ahmU1moZlAdqGJQAAAAAAGsakwBpGJQAqXa+AaBotwByJJm7hECm/4ZDp/yFQab+hUKn/4RBpv+EQKX/gz+m/4M/pv+DPqX/gj2k/4E8pf+CPKT/gTuj/4A6pP+AOqP/gTuk/3gunv/LsNr///////38/f//////mWC0/3csnf99NqL/ezOg/3syn/96MaD/ezGf/3own/95L57/ejCd/3kvnf94Lp7/dy2d/3gtnf93LJ3/diuc/3Yqm/92Kpz/dSmc/3Uom/90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P6b/gj6l/4I9pP+CPaX/gjyk/4E7o/+AOqP/gTyl/3w0oP+CPaX/yq7Z//79/v/+/v7//////+fZ7f+4kcz/lVuz/4I8pf94Lp3/dCib/3Ikmf9yI5j/cSKY/3Eil/9wIpf/cCGX/28hmP9vIJf/byCX/28gl/9wIZf/cSKY/3Ikmf9zJZr/dCea/3Upm/91KZz/dSmb/3Qom/9zJpn/ciSZ/3Ejmf9xI5n/cSKY/3Ahl/9vIJf/byCX/28flv9uHpf/bR2X/24elv9uH5b/aBaS/4E8pP/59fr///////79/v///////v7+/4tKqv9kD4//bBuV/2oYk/9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+ahmU/WsYk/9rGJNaAAAAAGsalAFcAYkAXAGJIHoxn/KKSqv/iEao/YhGqP+HRan/h0So/4dEqP+GQ6j/hUKn/4ZBpv+FQqf/hEGm/4RApf+DP6b/gz+m/4M+pf+CPaT/gj2l/4I8pP+BO6P/gDqj/4E7pP9+N6L/eS+e/59quv/bx+X//Pv9///////+/v7//v7+/+7m8//dyub/z7Tc/8Wl1f/BntL/vprQ/7yWzv+6k8z/tY3J/7CGxv+sfsP/pXO9/51ouP+XXbP/jk+t/4VBpv98NKH/dCib/28flv9rGZT/axmU/24elf9xI5j/dCea/3Qnmv9yJZn/cSOZ/3EimP9wIZj/byCX/3Aglv9vH5b/bh6X/20dlv9wIpf/ZA+P/6Z2v////////fz9//7+/v//////+vf7/4I9pf9lEZD/bBqU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GsZlP9qGJSDaBqWAG0clQJkDo8AZA+POX84ov+LS6v/iEep/olHqf+IRqj/h0Wp/4ZEqP+HRKj/hkOo/4VCp/+FQab/hUKn/4VBpv+EQKX/gz+m/4M/pv+DPqX/gj2k/4E8pP+CPKT/gTuj/4A6o/+AO6T/gTuk/3kvnv97M6D/mWG1/8Ge0v/h0en/+PT6///////+/v7//v7+/////////////////////////////////////////////v7+//7+/v//////+/n8//by+P/t4/L/4dHp/9K53v++m9D/p3a//49Rrv96MJ7/bByV/2oXkv9tHJX/cSOY/3Ilmv9yJZn/cSOZ/3AimP9wIZf/cCCX/3AimP9sG5X/bR2V/+LS6v///////f3+//79/v//////7+bz/3Qnmv9oFZL/axqU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GoYlP9rGZOmAAAAAG4dlgNoFJEAaBSRVIA6o/+LTKv8iUep/ohHqf+IRqj/iEao/4dFqf+GRKj/h0So/4ZDqP+FQqf/hUGm/4RCp/+FQab/hECl/4M/pv+DP6b/gz6l/4I9pP+BPKT/gTyk/4E7o/+AOqT/gDmj/4E7pP+AOqP/ezKf/3csnP93LJ3/gz6l/51nuP+/m9H/383o//Ps9v/8+v3//Pr9//v5/P/7+fz/+/n8//v5/P/8+v3//fz9//39/v/+/v7///////7+/v/////////////////+/v7///////z7/f/s4fH/0rrf/6+Dxf+JSKn/ciSY/2kWk/9oFZL/axmT/2wblf9tHJX/bBmU/2cSkf9oFZL/wJ7S///////9/P7//v7+//38/f//////1b7g/2cSkP9rGpT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsYlP9qGZTHZhuYAHMlmQNvIJcAbR2VaIRApv+LTKv7iUip/olIqv+IR6n/iEao/4hGqP+HRan/hkSo/4ZDp/+GQ6j/hkKn/4VBpv+EQab/hUGm/4RApf+DP6b/gz+l/4M+pf+CPaX/gTyk/4E7o/+BO6P/gDqj/4A6pP+BO6P/ejGf/302of+bY7b/xabV/+bY7P/38/n//v3+//7+/v/+/v7///////7+/v///////////////////////v7+///////+/v7//v3+//7+/v////////7+/////////////////////////////v7+///////6+Pz/49Pq/8Kh0/+jcb3/jU6t/4A6o/98NaH/gjyk/5pitf/XweL///////79/v/+/v7//v7+//38/f//////qXrB/2ILjf9tHJX/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsZlP9rGJTaahiUAHMlmQNwIJcAbRyVeoRBpv+LTKz7iUmq/olJqv+JSKr/iUep/4hGqP+HRaj/h0Wp/4dEqP+GQ6f/hUOo/4ZCp/+FQab/hEGm/4RApf+EQKX/gz+m/4I+pf+DPqX/gj2l/4I8pP+BO6P/gj2k/343ov96MJ//oGu6/9zK5v/9/P3///////7+/v///////v7+///////9/f7/8On0/97M5//Nstv/vJfP/66BxP+kcb3/m2S2/5Vbsv+RVbD/kFOv/5JWsP+WXLT/nmi5/6h4wP+3j8v/yq3Z/+DO6P/17/f///////7+/v/+/v7///////7+/v///////f3+//bx+P/y6/X/+PX6///////+/v7//v3+//7+/v///////v3+///////w5/T/eS6e/2gWkv9sG5T/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTqahiUAHIlmQNvIJcAbByVhoRBpv+MTaz7iUmq/olJqv+JSKn/iUiq/4lHqf+IRqj/h0ap/4dFqf+HRKj/hkOn/4VCp/+GQqf/hUKn/4RBpv+EQKb/hECl/4M/pv+CPqX/gz2k/4E8pP+DPqX/fTWg/4VCp//Nstv//v7+//7+/v/+/v7///////79/v/9/P3//////+/m8/+UWrL/gz+m/3oxn/90J5r/ciWZ/3IjmP9yJJn/cyaZ/3Mmmf9zJpr/ciWa/3Ikmf9xI5j/byCX/24dlf9tHJX/bh6W/3Qnm/+CPaT/qnzC/+rf8P///////fz9//z7/f/9/P7//v7+///////+/v7///////39/v/9/P3//v7+///////+/v7//Pv9//////+vg8X/ZA6O/24elv9rGpT/axmT/2sZlP9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT1axiUAHMlmQNwIJcAbRyVjIVBpv+NTqz7ikqq/opKq/+KSar/iUip/4hHqf+JR6n/iEao/4dFqf+HRaj/h0So/4ZDqP+FQqf/hkKn/4VCp/+EQab/hECl/4RApf+DP6b/gj6l/4RApv99NaH/kFKu/+bZ7f///////Pv9///////8+v3/1L3g/+7l8v///////v7+//z6/P+mdb7/cCKY/3w0oP9/OKL/fTWh/3w0oP98M6D/ezKf/3syn/96MZ7/eTCe/3kvnv95L57/eS+e/3gunv94Lp3/dyyd/3UpnP9yJJn/bBqU/3wzoP/fzej///////39/v/+/v7//v7+//7+/v/+/f7//v7+//7+/v////////////7+/v/8+/3//////9jC4v9tHZb/bBuV/20clf9sG5X/bBqU/2sZk/9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT6axmTAHMmmQNwIZcAbRyVkIVBpv+NTqz7i0uq/opKq/+KSqv/ikmq/4lIqf+IR6r/iEep/4hGqP+HRan/h0Wp/4dEqP+GQ6j/hUKn/4VBpv+FQqf/hEGm/4RApf+EP6X/hECm/385o/+OT63/6+Dx///////7+fz//////+bZ7f+eaLn/dSmb/5pjtv/49Pr///////79/v/9/f7/vZnP/3szoP90KJv/ezKf/301of98M6D/ezGf/3own/96MJ7/eS+e/3gunv94Lp7/eC2d/3csnf93LJ3/eCyc/3gtnf94LZ7/diuc/2QOj//CoNP///////37/f/+/v7///////////////////7///7+/v///////v3+//z7/f//////5dbs/3kvnv9qGJP/bh+W/2wclf9sG5T/axuV/2walP9rGZP/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT9axmTAHMmmQNwIZcAbRyVi4VCpv+NTq37i0ur/otLqv+KSqv/ikqq/4pJqv+JSKr/iEep/4lHqP+IRqj/h0Wp/4ZEqP+HRKf/hkOo/4VCp/+FQab/hUKn/4VBpv+FQab/gz+l/4I9pf/dyub///////v5/P//////0Lfd/4I9pf95MJ//gj2l/3ctnf+jcr3/+/j8//7+/v/8+v3//////+PT6v+jcL3/fTWh/3Mlmf91KZz/eC6d/3own/96MZ//ezGe/3oxnv95MJ7/eC+e/3gtnf91KZv/ciSZ/28elv9uHZb/ejGf/7WNyv/8+v3///////7+/v////////////7+/v/+/v///v7+//79/v/9/P3//f3+///////axeT/ejCe/2kXk/9vIJf/bR2V/20dlv9tHJX/bBuU/2wblf9sGpT/axmT/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZT5axiTAHMmmQNwIZcAbRyVhIZCp/+OT637i0ur/oxMq/+LS6r/ikqr/4lJq/+KSar/iUiq/4hHqf+IRqn/iEao/4dFqf+GRKj/hkOo/4ZDqP+FQqf/hUGm/4RBpv+GQ6f/fTSg/7ePy///////+/n8///////Os9v/fDSg/385ov+CPaX/fzmj/4E7pP93LZ3/om+7//bx+P///////Pr8//7+/v//////59ru/7qUzf+XXbP/gz+m/3kvnv91KZv/cyaa/3Mlmf9zJZn/cyea/3YqnP9+NqH/i0ur/6NwvP/FpdX/7eTy///////+/v7//v7+///////+/v7//v7+//39/v/9+/3//f3+//7+/v//////9vH4/7SLyf9vH5f/bBqU/3AhmP9uHZb/bR2W/24elf9tHZb/bRyV/2wblP9rGpT/bBqU/2sZk/9qGJT/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTzahmUAHQmmgNwIZcAbR2WdoZDp/+OT637jEyr/otLq/+MTKv/i0uq/4pKq/+JSar/ikmq/4lIqv+IR6n/iEao/4hGqP+HRan/h0So/4ZDp/+GQ6j/hkKn/4ZCp/+DP6X/ikqq/+3j8v///////////+LS6v+BPKT/gDqj/4M+pP+BO6P/gDqk/385o/+BO6T/eC6d/5JXsP/j1Ov///////38/v/8+/3//v7+///////9/f7/8en0/+LS6v/XwOH/0Lbd/82y2//Qtdz/1sDh/+DP6f/t4/L/+vj7///////+/v7///////38/v/+/v7///////7+/v////////////7+/v/+/v7///////Hq9f+9mM//gj2l/2gUkf9vIJf/cSKY/28flv9uHpf/bh6X/24elv9tHZX/bRyV/20clf9sG5X/axqU/2sZk/9rGZT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/msZlP9rGZTmZA2bAHQmmgNwIZgAbh2WY4ZDp/+OUK77jEyr/oxMrP+LS6v/jEyq/4tLqv+KSqv/iUmq/4lIqf+JSKr/iUep/4hGqP+IRqj/h0Wp/4dEqP+GQ6f/hkOo/4dEqP9+N6H/roHE///////9/f7//v7+/55puf99NKD/hECm/4E8pP+BO6P/gTqj/4A6pP9/OaP/gTuk/3oxn/9/OaL/uZPM//Tu9////////v7+//z7/f/9/P3///////7+/v///v/////////////////////////////+/v7///////39/v/9+/3//v3+//7+/v///////v7+///////9/P3/9O73/97L5/++m9D/mmO2/3syn/9rGZT/bh6X/3Ilmf9xIpj/cCGX/3Agl/9vH5b/bh6X/20dl/9uHpb/bR2V/2wclf9tHJX/bBuV/2salP9rGZP/axmU/2oYlP9rGZT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT/axmU/WsZlP9rGZPWAAAAAG4elgNnE5IAZxSSTYM+pf+PUa78jE2s/o1NrP+MTKz/i0ur/4tLqv+LS6r/ikqr/4pJqv+JSKn/iUiq/4lHqf+IRqj/h0Wp/4dFqf+HRKj/hkOo/4dEqP9/OKL/0Lbd////////////3crm/385o/+DQKb/gz6l/4I9pf+BPKT/gTuj/4E7pP+AOqT/fzmj/4E7o/9+OKL/dyyc/4pJqv+8l8//7OLx///////+/v7///////79/v/7+vz/+/n8//v5/P/8+v3//Pv9//38/f/9/f7//v7+//7+/v///////v7////////+/v7//v7+///////8+vz/7OLx/82y2/+qfMH/ikqr/3Qnmv9rGpT/bx+W/3Ilmv9xI5j/cCGX/28gmP9wIJf/bx+W/24el/9uHpf/bh6W/20dlf9sHJX/bRyV/2wblf9rGpT/axmT/2sZlP9qGJT/axmU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/WsZk/9qGZTAaBaWAG0clQJiC40AYguNMoE7pPyPUq7/jU2s/oxNrP+NTaz/jEys/4xMq/+LS6r/i0uq/4pKq/+JSar/iUip/4hHqf+JR6n/iEao/4dFqf+HRan/h0So/4ZDp/+FQqf/5dfs////////////u5XO/3w0oP+FQqf/gj6l/4I9pP+CPaX/gjyk/4E7o/+AOqT/gDqk/385o/+AOaP/gDuk/3w0oP92Kpz/gj2l/6NwvP/Krdn/6t/w//38/f///////v7+///////+/v7///////7+/v/9/f7//fz9//37/f/9/P3//fz9//38/f/9/P3//fz+//39/v/+/v7///////7+/v/+/v7//v7+/+re7/+6lMz/hECl/2kXkv9vIJf/ciSZ/3Ahl/9vIJf/cCCW/28flv9uHpf/bR2W/24elv9tHZb/bByV/2wblP9sG5X/axqU/2sZk/9qGJT/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GsZlP9qGZSdAAAAAGsZlAFVAIMAVQCDGn01oe2PUq//jU6s/Y1Orf+MTaz/jEyr/4tLrP+MTKv/i0uq/4pKq/+KSqv/iUmq/4lIqf+JR6n/iUep/4hGqP+HRan/h0Wp/4VBpv+NTqz/7+f0////////////qHjA/342of+FQqf/gz+m/4M+pf+CPaT/gTyk/4I8pP+BO6P/gDqk/4A6pP+AOaP/fzii/386o/+AOqP/fTWh/3csnP92Kpv/gDqj/5Rasv+vg8X/yKrX/97L5//u5PL/+PT6///////+/v7//v7+//////////////////////////////////7+/v///////f3+//v5/P/7+vz//v7+//7+/v//////9/L5/7yYz/94LZ3/ahmU/3Mlmf9wIZf/byCX/3Agl/9vH5b/bh6X/20dlv9uHpb/bR2W/20clf9sG5T/bBuV/2walP9rGZP/ahiU/2sZlP9rGZT/axmU/2sZlP9rGZT+axmU/GsZlP9rGJN6AAAAAGoZlAAAAEQAAAAqA3kvns+OUK3/jk+t/I1OrP+NTq3/jU2s/4xMq/+LS6z/jEyr/4tLqv+KSqv/ikmq/4pJqv+JSKn/iEep/4lHqf+IRqj/iEap/4RBp/+QUq7/8+z2////////////onC8/384ov+FQqb/gz+l/4M/pv+DPqX/gj2k/4E8pf+CPKT/gTuj/4A6pP9/OaP/gDmj/385o/9+OKL/fzii/385o/9/OKL/fDSg/3ctnf90J5v/dCea/3kunf+AO6T/jEyr/5lgtf+mdL7/s4nI/76az//Hqdb/0Lfd/9fB4v/cyeb/4dHp/+nc7v/07vf//v7+/////////////f3+//z6/f/9/P3//v3+///////q3u//kFOv/2kWk/9yJZn/cCGY/28gl/9wH5b/bx+W/24el/9tHZb/bh2V/20dlv9tHJX/bBuU/2salf9sGpT/axmT/2oYlP9rGZT/axmU/2sZlP9rGZT+axmU/msZlP9rGZNPAAAAAGoXkgCKSqsChkSoAHMmmqmMTKv/jlGu+41Prf6NTqz/jU6t/41NrP+MTKv/jEyr/4xMq/+LS6r/ikqr/4lJqv+KSar/iUiq/4hHqf+JRqj/iEep/4VCp/+NT63/7+fz////////////qHfA/344ov+GQ6f/hECl/4M/pv+DP6X/gz6l/4I9pP+BPKT/gjyk/4E7o/+AOqT/fzmj/4A5o/9/OaP/fjii/343of9+N6L/fjeh/343of9+N6L/fTai/3w0oP96MJ7/dyud/3Qnmv9yJJn/cSKY/3Eil/9xI5j/cSOZ/3Ilmf91KJv/dyyc/3ownv+DPqX/l120/7iRzP/j1Ov//v3+///////9/f7//fz+//38/f//////+vf7/55quf9pFpP/ciSZ/3Ahl/9vIJf/bx+W/28flv9uHpf/bR2W/20dlf9tHZb/bRyV/2wblP9rGpT/bBqU/2sZk/9qGJT/axmU/2oZlP9rGZT+axmU/2sYlPJrGZMnAAAAAGsXkgB1KZsDciSYAG8flnWJR6n/kFOv+45Qrf6NT63/jU6s/41Orf+NTaz/jEys/4tLq/+MTKv/i0uq/4pKq/+JSar/iUmq/4lIqv+IR6n/iUep/4hGqP+GRKj/5Nbs////////////uJDL/342of+GRKj/hUGm/4RApf+DP6b/gj6m/4M+pf+CPaX/gTyk/4I8pP+BO6P/gDqk/385o/9/OKL/fzmj/344ov9+N6H/fTei/342of99NaD/fDSh/3w0of98M6D/fDSg/3wzoP98M6D/fDOg/3syoP97Mp//ejGf/3kwnv94Lp3/dyyc/3YqnP90Jpr/cCCX/20clf91KZz/mmK2/93K5v///////f3+//7+/v/9/P3//v7+//v6/P+XXrT/aReT/3Ikmf9wIZj/byCX/28flv9vH5f/bh6X/24elv9tHZX/bByW/20clf9sG5T/axqU/2walP9rGZP/ahiU/2sZlP9qGJT8axmU/2sZlNFrGZMGAAAAAGkZkQBuH5YCZhKQAGYSkECEQab/kVSv/o5Qrf6OUK7/jk+t/41OrP+MTq3/jU2s/4xMq/+LS6v/i0uq/4tLqv+KSqv/ikmq/4pIqf+JSKr/iUep/4pJqf+BO6T/zbHb////////////1L3g/4A5o/+GRKf/hEGm/4VBpv+EQKX/gz+m/4I+pv+DPqX/gj2l/4E8pP+BO6P/gTuj/4A6pP9/OaP/fzii/385o/9/OKL/fjeh/302ov9+NqH/fTWg/3w0of97M6D/fDOg/3syn/96MZ//ejCe/3own/96MJ7/eS+d/3kvnv94Lp7/eC6d/3gtnf93LZ3/eC2d/3gtnf91KZv/bR2W/3Ijmf+wg8X/+vf7///////+/v7//fz9///////v5/T/fjei/20dlv9xI5j/cCGX/3Agl/9vH5b/bh6X/24el/9uHpb/bR2V/2wclv9tHJX/bBuV/2salP9rGZT/axmT/2oYlP5rGZT7axmU/2sYk6BqGJQAAAAAAGkSkABqF5MBSwB8AEoAfBF8NKDjkVSv/49Rrf2OUK7+jVCu/45Prf+NTqz/jE2s/41NrP+MTKz/i0yr/4tLqv+LS6r/ikqr/4lJqv+KSan/iUiq/4pJqv+CPKT/qnzC////////////9fD4/5FUsP+DPqX/hkOn/4RBpv+FQKX/hECl/4M/pv+CPqX/gj2k/4I9pf+CPKT/gTuj/4A7pP+AOqT/fzmj/384ov9/OKP/fzii/343ov99NqH/fjah/301oP98NKH/ezOg/3wzoP97MqD/ejGf/3ownv95L5//ejCe/3kvnf94Lp7/dy6e/3gtnf93LJz/diuc/3crnP92Kpz/dyyd/3UpnP9qF5P/mmK2//f0+f///////v7+//z7/f//////xKTU/2kXk/9yJJn/cCGX/28gmP9wIJf/bx+W/24el/9tHZf/bh2W/20dlf9sHJX/bBuU/2wblf9rGpT/axmT/2sZlP5qGJP8axmU/2oYlGJqGZQAAAAAAGwZkwBqGZQAklWwAo9QrQB1KZuqjU+t/5BTrvuPUa3+jlCu/45Prf+OT63/jU6t/4xNrP+MTaz/jEys/4xMq/+LS6r/ikqq/4pKq/+JSar/iUip/4lIqv+IRaj/i0ur/+nd7////////////7+c0f9+NqL/iEWo/4VBpv+FQab/hECl/4RApv+DP6b/gj6l/4I9pP+CPaX/gjyk/4E7o/+AOqT/gDqk/4A5o/9/OKL/fjii/384of9+N6L/fTah/342oP99NaD/fDSh/3szoP97Mp//ezKg/3sxn/96MJ7/eS+f/3ownv95L53/eC6e/3gtnf94LZ3/dyyc/3YrnP92Kpz/diqb/3UqnP93LZ3/aheT/6Ryvf///////v7+//79/v//////9fD4/4I9pP9tHZb/cSOZ/3Ahl/9vIJf/byCX/28flv9uHpf/bR2X/24elv9tHZb/bByV/2wblP9sG5X/bBqU/msZk/5qGJT/axmU92oZlCdqGZQAAAAAAAAAAABqGZQAcyWZA20dlQBsG5Vjh0ap/5FVsPyPUa3+j1Gt/45Qrv+OT63/jk+t/41Orf+MTaz/jEyr/4tMrP+MTKv/i0uq/4pKq/+KSqv/ikmq/4lIqf+KSqv/gTyj/7WMyf///////v7+//by+P+VW7P/gTyk/4dEqP+FQab/hUGm/4RApf+EQKb/gz+m/4I+pf+CPaT/gj2l/4I8pP+BO6P/gDqk/4A6o/+AOaP/fzii/344ov9/OKH/fjei/302of99NaD/fTWh/3w0of97M6D/ezKf/3syoP97MZ//ejCe/3kvn/96MJ7/eS+d/3gunv93LZ3/eC2d/3csnf92K5z/diqb/3Yqm/91KZz/diuc/24dlv/RuN7///////38/f/9/P3//////6d3v/9pFpL/ciWa/3EimP9wIZf/byCX/28gl/9vH5b/bh6X/20dlv9uHpX/bR2W/2wclf9sG5T/bBuV/mwalPxrGZP/ahiUxWcgkgBUAJ8AAAAAAAAAAABqF5IAaxqUAVYAhQBXAIYefzii7pJWsP+PUq79j1Gu/o9Rrf+OUK7/jU+t/41OrP+NTq3/jE2s/4xMq/+MTKz/jEyr/4tLqv+KSqv/ikqr/4pJqv+JSKr/iEep/4hFqP/fzuf////////////dy+b/hECm/4VBp/+GQ6f/hEGm/4VBpv+EQKX/gz+m/4M/pv+DPqX/gj2k/4E8pP+CPKT/gTuj/4A6pP9/OaT/gDmj/385o/9+OKL/fjeh/343ov9+NqH/fTWg/3w0of98NKH/ezOg/3syn/96MaD/ezGf/3own/95L57/ejCe/3kvnf94Lp7/dy2d/3gtnP93LJ3/diuc/3Yqm/92Kpv/dyyd/24dlv+aYbX///////79/v/8+/3//////8Oj1P9qF5L/cyaa/3AimP9xIpj/cCGY/28gl/9wIJb/bx+W/24el/9tHZb/bh6V/20dlv9sHJX/bBuU/mwalPtrGpT/axmTeGsZkwBrGZQDAAAAAAAAAABpGJQAaBWRAJ9rugGbZbcAdSqbrY5Qrf+QVK/7j1Ku/o9Rrf+PUa3/jlCu/41Prf+NTqz/jU6t/41NrP+MTKv/i0ur/4xMq/+LS6r/ikqr/4pKq/+KSar/ikqq/4RBpv+aYbX/9fD4//7+/v//////zLDa/4A6o/+GQqf/hkOn/4RBpv+FQab/hECl/4M/pv+DPqX/gz6l/4I9pP+BPKT/gjyk/4E7o/+AOqT/fzmj/4A5o/9/OaP/fjii/343of99N6L/fjah/301oP98NKD/fDSh/3wzoP97Mp//ejGf/3sxnv96MJ//eS+e/3kvnf95L53/eC6e/3ctnf93LZ3/dyyd/3crnP92Kpv/diuc/3Ilmv9/OaP/8uv1///////8+/3//////82z2/9sG5T/cyaZ/3Ejmf9wIpj/cSGX/3AhmP9vIJf/bx+W/28fl/9uHpf/bR2W/20dlf9tHZb+bRyV/mwblP9rGpT3axmTK2sZkwBrGZQCAAAAAAAAAABrGpQAahiTAHEimAJqF5MAaRaTUYZEqP+TV7D9j1Ku/o9Srv6PUa3/j1Gu/45Qrv+NT63/jU6s/4xOrf+NTaz/jEys/4tLq/+MTKv/i0uq/4pKq/+JSar/iUip/4pKq/+CPaT/q3zC//38/f/9/P3//////8mr1/+BO6T/hD+m/4dEqP+EQab/hUGm/4RApf+DP6b/gj6l/4M+pf+CPaX/gTyk/4E7o/+BO6P/gDqk/385o/9/OKL/fzmj/384ov9+N6H/fTei/302of99NaD/fDSh/3w0oP98M6D/ezKf/3oxn/97MZ7/ejCf/3kwnv95L53/eS+d/3gunv93LZ3/dyyc/3csnf93K5z/dyuc/3Qom/94Lp7/6t7w///////8+vz//////8mr2P9rGpT/dCea/3Ikmf9xI5n/cCKY/3Ahl/9wIZj/cCCX/28flv9vH5f/bh6X/20dlv9tHZX+bByW/G0clf9sGpS0cB+VAHAglgFqGZQAAAAAAAAAAAAAAAAAahiTAGoXkwA1AGsANwBsCnszoNWRVbD/kFOu/ZBTr/6QUq7/j1Gt/45Qrv+OUK7/jk+t/41OrP+MTa3/jU2s/4xMrP+LS6v/i0uq/4tLqv+KSqv/iUmq/4lIqf+KSqv/gTyk/7OIx//+/f7//f3+///////Tu9//iEep/4A5o/+HRaj/hUKn/4VApf+EQKX/gz+m/4I+pf+DPaT/gj2l/4E8pP+BO6P/gTuj/4A6pP+AOaP/fzii/385o/9/OKL/fjeh/302of9+NqH/fTWg/3w0of97M6H/fDOg/3syoP96MZ//ezGf/3own/96MJ7/eS+d/3gunf94Lp7/eC2d/3csnP92K53/dyyc/3Qnmv9+N6L/8Oj0///////8+v3//////7SKyP9qGJP/dCia/3IkmP9yJJn/cSOZ/3EimP9wIZf/byGY/3Agl/9vH5b/bh6X/24el/5uHZb+bR2V/Gwclf9sGpRWbBuUAGwblANsF5QAAAAAAAAAAAAAAAAAaRWSAGgVkgB4Lp0DdCiaAG8gl3SKSqr/klaw+5BTrv6PUq7+j1Ku/49Rrf+OUK3/jU+t/45Prf+NTqz/jE2s/41Nq/+MTKz/jEyr/4tLqv+LS6r/ikqr/4lJqv+KSKn/ikqr/4E8pP+vgsT/+vf7////////////59ru/55ouP9+NaH/gz+l/4ZEqP+FQab/hECl/4M/pv+CPqX/gz2k/4I9pf+CPKT/gTuj/4E7o/+AOqT/gDmj/384ov9+OKL/fzii/343ov99NqH/fjWg/301oP98NKH/fDOg/3syn/97MqD/ejGf/3ownv96MJ//ejCe/3kvnf94Lp7/eC6d/3gtnf93LJz/eS+e/28flv+ZYbX///////79/v/+/v7//fz+/5FTr/9uHpX/dCia/3Mlmf9yJJj/cSOY/3Ejmf9xIpj/cCGX/28hmP9wIJf/bx+W/24el/5uHpb8bh6W/2wcldFoE5EIZxKRAGsalABsGZMAaReVAGsakwBhCYoAaBWRAG4flgFzJpoEAAAaAFUAhBl9NqHnklex/5FUr/2RVK//kFOv/5BTrv+QUq7/j1Gu/45Qrv+OT63/jU6t/41NrP+MTKv/jEur/4tLq/+LS6r/ikqq/4pKq/+JSar/iUiq/4pKq/+CPKT/oW27/+3k8v////////////r3+//FpNX/jEur/302of+DP6X/hkOn/4VBpv+DP6b/gj6l/4M9pP+CPaX/gjyk/4E7o/+AOqT/gDqk/4A5o/9/OKL/fjii/384ov9+N6L/fTah/341oP99NaD/fDSh/3szoP98Mp//ezKg/3sxn/96MJ7/ejCf/3ownv95L53/eC6e/3ctnf95L57/dyuc/3Ilmf/ZxeP///////v5/P//////2MPj/3Ejmf90J5r/dCea/3Mmmv9zJZn/ciSY/3Ejmf9xI5n/cSKY/3Ahl/9vIZj/byCX/m8flv5uHpf7bR2W/2wblWxtHZUAbR2VA2oWkwBpGZQAfTahAHkvngBvH5cCikqqAW8flgBqF5IAdyydA3IlmQBrGZRziEep/5FVr/yOUK3+jlCt/41Prf+NTqz/jU6s/4xNrP+MTaz/jE2s/4xNrP+NTqz/jU6s/41NrP+MTav/i0uq/4pKqv+KSqv/ikmq/4lIqf+KSqv/gz6l/49Rrf/Tu9////////7+/v//////7uXz/7iRy/+KS6v/fTWh/4A6o/+EQab/hEKn/4M/pv+DPqX/gj2l/4I8pP+BO6P/gDqk/386pP+AOaP/fzii/344ov9/OKH/fjei/302of99NaD/fTWg/3w0of97M6D/ezKf/3syoP97MZ//ejCe/3kwnv96MJ7/ejGe/3oyn/9zJ5v/ciSY/8Oj1P///////Pv9//7+/v/7+Pz/kVWw/28flv91KZz/dCaa/3Mmmf9yJZr/cyWZ/3IkmP9xI5n/cSKY/3EimP9wIZf/byCX/m8fl/1vH5f/bR2W12gVkwxnFJIAaxqVAG0bkwAAAAAAeS6fAGoekwHDSN4A/3X/AFkIhxxwIJhLeS+ec3szoIx9NaGwgDqj/4M/pf6EQab+i0qq/4lIqv+JSKn/ikmq/4VCp/+BO6T/gj6l/4RBpv+DP6X/hEGm/4ZDp/+IRqn/i0ur/4xMq/+KSqr/ikqr/4pJqv+JSKn/ikqq/4ZDp/+CPaX/rH7D/+vg8P///////v3+///////u5fP/wZ/S/5dds/+AOqP/fDSg/384o/+CPqX/hECm/4M/pv+DPqT/gjyk/4E7pP+AOqT/gDmj/385o/9+OKL/fzeh/343ov9+NqH/fTWg/301of98NKH/fDOg/3wzoP98M6D/fDOg/3wzoP95L57/dCia/3Ail/+IR6n/07zf///////8+/3//f3+//////+tgMP/bR2W/3YrnP90J5r/cyeb/3Qnmv9zJpn/ciWa/3Mlmf9yJJj/cSOZ/3AimP9xIpj+cCGY/nAgl/xvH5b/bByVZm0dlgBuHZYDahaUAGoYlAAAAAAAbBuVAWIMjgBiCo0sdiucrIVCp/GNTaz/j1Gu/pBSr/6SVbD/k1aw/5FVsP+RVbD/kVWv/5FUr/+QVK//kFOu/5BSrv+PUa7/jlCu/41Prf+MTKz/i0qr/4hHqf+GQ6j/hD+m/4hHqP+MTKv/ikqq/4lJq/+KSar/iUip/4lIqv+JSKn/gTuk/4pKq/+7ls7/7+f0///////+/v7///////n1+v/bx+X/to3K/5Zcs/+DP6b/ezOg/3sxn/98NKD/fzei/4A5o/+BO6T/gTuk/4E7pP+AO6T/gDuj/4A6o/9/OaP/fzii/383of99NqH/ezOg/3kvnv92Kpz/cyaa/3Ikmf96L57/klaw/8Oi0//18Pj///////z6/f///////Pv9/66BxP9vIJb/diuc/3UpnP90KJv/dCea/3Mmm/90Jpr/cyaZ/3Ilmf9yJJn/ciSY/3Ejmf5wIpj/cSGX/HAhmP9uHpbHYweMBFcAfgBsGpUAbR2WAGsXkwAAAAAAYAeMAGMMjjJ7MqDtkVWw/5Vbs/6VWrH9lFix+5NXsfuRVrH8kVWw/5FUr/6RVLD+kFSv/5BTrv+QU67/j1Ku/49Rrf+PUa3/jlCt/45Qrv+OUK7/jlCt/45Prf+OT63/jEyr/4VApv+JR6j/i0yr/4pKqv+KSar/ikmq/4lIqv+JR6n/ikmq/4dEqP9/OaP/jU6t/7mSzP/n2u7//v3+///+/v/+/v7///////n2+//k1uz/y6/Z/7KHx/+dZ7j/jk+t/4VApv9+NqH/ejGf/3kvnv94LZ3/dyyd/3csnP93LJz/dyyc/3csnP95L57/fjei/4ZDqP+UWLH/p3fA/8Oi1P/j0+r/+/n8///////+/v7//v7+///////l1+z/l12z/28gl/94LZ3/diqc/3Upm/91KZz/dSib/3Qnmv9zJ5r/dCea/3Mmmv9yJZn/ciSY/3Ikmf5xI5n+cSKY/XAhl/5sHJZFbR2WAG4elgJrGpQAahmTAAAAAAAAAAAA////AHEimLyQU6//ll2z+ZNYsf2TWLH+k1ew/pNXsf6SVrH/kVWw/5JVsP+RVK//kVWw/5BUr/+RVK7/kFOv/5BSrv+QUq7/j1Gu/45Qrv+OUK7/jk+t/41OrP+NTaz/jU6t/4xMq/+GRKf/jEyr/4tLqv+KSqv/iUmq/4lIqf+JSKr/iEep/4lHqf+JSKr/hUKn/384o/+IRaj/pna//9C23f/x6fT///////7+/v/+/v7//v7+///////9/P3/9O73/+nd7//ezOf/1Lzg/8uv2v/Gp9b/w6LU/8Kh0//FpdX/yq3Y/9G53v/bx+X/5tjt//Hq9f/7+fz///////7+/v///////v7+//7+/v//////59rt/66CxP96MZ//cSSZ/3kvnv93LJ3/dyuc/3Yqm/91KZz/dCmc/3Uom/90J5v/cyaa/3Qnmf9zJpr/ciWZ/nIkmP9yJJn7cSOY/24flpdyJpsAciSaAmYSjQBuHpcAaRqVAAAAAAAAAAAAXwaLN4E8pPiXXrT/lVqy/pVbsv6UWbL/lFix/5RYsf+TWLH/lFmy/5JXsf+SVbD/kVSv/5BVsP+RVa//kVSv/5BTr/+QU67/kFKu/49Rrv+PUa7/jlGu/49Rrv+OUK3/jE2s/41Orf+JR6n/ikiq/4xMq/+LSqr/ikqr/4pJqv+JSKn/iUiq/4lHqf+IRqj/iEap/4hHqv+GQ6j/gDqj/384o/+NTqz/qXnA/8mr2P/l1uz/9/T5///////+/v7//v7+/////////////////////////////////////////////////////////////////////////////v7+///////8+/3/6d3v/8mr2P+eaLn/fDOf/3Eil/94LZ3/eTCf/3ctnf93LJz/diuc/3crnP92Kpv/dSmc/3Qom/91KJv/dCeb/3Mmmv9zJpn+cyaa/3IlmfxyJJn/cCGY2moXkxVqF5MAbR2VAXAwqgBsF5IAAAAAAAAAAAAAAAAAaRaSc4lIqf+YYLX8kFOv/o5Qrf+QUq7/kFKv/5BSrv+OT63/i0ur/49Rrv+SV7H/kVSw/5FVsP+NT63/jE2s/41NrP+MTKz/jEyr/4tLqv+KSqr/ikmq/4ZFqf+KSar/jk+t/41Orf+JSar/iEWp/4xNq/+LS6r/i0ur/4pKq/+K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AAAX/QAAAAAAAAAAAAAL+gAAAAAAAAAAAAAL+gAAAAAAAAAAAAAF9AAAAAAAAAAAAAAF9AAAAAAAAAAAAAAH/AAAAAAAAAAAAAAC6AAAAAAAAAAAAAAC6AAAAAAAAAAAAAAB6AAAAAAAAAAAAAAB0AAAAAAAAAAAAAAB0AAAAAAAAAAAAAAB0AAAAAAAAAAAAAAA0AAAAAAAAAAAAAAA4AAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAAoAAAAAAAAAAAAAAA4AAAAAAAAAAAAAAA0AAAAAAAAAAAAAAA0AAAAAAAAAAAAAAA0AAAAAAAAAAAAAAB0AAAAAAAAAAAAAAB6AAAAAAAAAAAAAAB6AAAAAAAAAAAAAAD6AAAAAAAAAAAAAAC9AAAAAAAAAAAAAAC9AAAAAAAAAAAAAAF/AAAAAAAAAAAAAAF+gAAAAAAAAAAAAAH8gAAAAAAAAAAAAALzQAAAAAAAAAAAAAPsAAAAAAAAAAAAAAXQAAAAAAAAAAAAAAfgAAAAAAAAAAAAAAvgAAAAAAAAAAAAABfAAAAAAAAAAAAAABfAAAAAAAAAAAAAAC/AAAAAAAAAAAAAAF/AAAAAAAAAAAAAAH/AAAAAAAAAAAAAAL/AAAAAAAAAAAAAAX/AAAAAAAAAAAAAAv/AAAAAAAAAAAAABf/AAAAAAAAAAAAAC//AAAAAAAAAAAAAF//AAAAAAAAAAAAAL//AAAAAAAAAAAAAn//AAAAAAAAAAAABP//AAAAAAAAAAAAC///AAAAAAAAAAAAJ///AAAAAAAAAAAAT///AAAACAAAAAABP///AAAACgAAAAAE////AAAACYAAAAAz////AAAACnAAAAHP////AAAAF88AAB4/////gAAAF/j//+H/////QAAAL/+AAD//////oAAAX///////////"
    });
    sheet.getCell(row + 1, col).value(new Date(2014, 8, 20)).formatter('m/dd/yyyy');
    sheet.getCell(row + 2, col).value(new Date(2050, 10, 12)).formatter('m/dd/yyyy');
    sheet.getCell(row + 3, col).value(new Date(1993, 5, 23)).formatter('m/dd/yyyy');
    sheet.getCell(row + 4, col).value(new Date(2020, 1, 2)).formatter('m/dd/yyyy');
    sheet.getCell(row + 5, col).value(new Date(2015, 10, 20)).formatter('m/dd/yyyy');

    sheet.setDataValidator(row + 1, col, 5, 1, dateValidator);

    col = col + 2;
    var formula = getCellPositionString(sheet, row + 6, col + 1) + "<100";
    var formulaValidator = DataValidation.createFormulaValidator(formula);
    formulaValidator.inputTitle("Tip");
    sheet.getCell(row + 1, col).value(20);
    sheet.getCell(row + 2, col).value(300);
    sheet.getCell(row + 3, col).value(2);
    sheet.getCell(row + 4, col).value(-35);
    var sumFormula = "=SUM(" + getCellPositionString(sheet, row + 2, col + 1)
        + ":" + getCellPositionString(sheet, row + 5, col + 1) + ")";
    sheet.getCell(row + 5, col).formula(sumFormula);
    formulaValidator.inputMessage("Be sure " + sumFormula.substr(1) + " less than 100");

    sheet.setDataValidator(row + 5, col, formulaValidator);

    col = col + 2;
    sheet.setColumnWidth(col, 120);
    var textLengthValidator = DataValidation.createTextLengthValidator(ComparisonOperators.lessThan, 6, 6);
    textLengthValidator.inputMessage("Text length should Less than 6");
    textLengthValidator.inputTitle("Tip");
    sheet.getCell(row + 1, col).value("Hello, SpreadJS");
    sheet.getCell(row + 2, col).value("God");
    sheet.getCell(row + 3, col).value("Word");
    sheet.getCell(row + 4, col).value("Warning");
    sheet.getCell(row + 5, col).value("Boy");

    sheet.setDataValidator(row + 1, col, 5, 1, textLengthValidator);

    spread.options.highlightInvalidData = true;
    sheet.resumePaint();
}

function setSlicerContent() {
    var sheet = new spreadNS.Worksheet("Slicer");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);

    var tableName = "slicerTable";
    var dataColumns = ["Name", "Age", "Sex", "Address", "Weight", "Height"];
    var data = [
        ["bob", "36", "man", "Beijing", "80", "180"],
        ["Betty", "28", "woman", "Xi'an", "52", "168"],
        ["Gary", "23", "man", "NewYork", "63", "175"],
        ["Hunk", "45", "man", "Beijing", "80", "171"],
        ["Cherry", "37", "woman", "Shanghai", "58", "161"],
        ["Eva", "30", "woman", "NewYork", "63", "180"]];
    sheet.tables.addFromDataSource(tableName, 6, 3, data);
    var table = sheet.tables.findByName(tableName);
    table.setColumnName(0, dataColumns[0]);
    table.setColumnName(1, dataColumns[1]);
    table.setColumnName(2, dataColumns[2]);
    table.setColumnName(3, dataColumns[3]);
    table.setColumnName(4, dataColumns[4]);
    table.setColumnName(5, dataColumns[5]);

    var slicer0 = sheet.slicers.add("slicer1", tableName, "Name");
    slicer0.position(new spreadNS.Point(50, 300));

    var slicer1 = sheet.slicers.add("slicer2", tableName, "Sex");
    slicer1.position(new spreadNS.Point(275, 300));

    var slicer2 = sheet.slicers.add("slicer3", tableName, "Height");
    slicer2.position(new spreadNS.Point(500, 300));

    sheet.resumePaint();
}

function addChartContent() {
    var sheet = new spreadNS.Worksheet("Chart");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(80);
    var dataArray = [
        ["", 'Mon', 'Tues', 'Wed', 'Thur', 'Fri', 'Sat', 'Sun'],
        ["BUS", 320, 302, 301, 334, 390, 330, 320],
        ["UBER", 120, 132, 101, 134, 90, 230, 210],
        ["TAXI", 220, 182, 191, 234, 290, 330, 310],
        ["SUBWAY", 820, 832, 901, 934, 1290, 1330, 1320]
    ];

    var sunburstDataArray = [
        ['Region', 'Subregion', 'country', 'Population'],
        ['Asia', 'Southern', 'India', 1354051854],
        [, , 'Pakistan', 200813818],
        [, 'Eastern', 'China', 1415045928],
        [, , 'Japan', 127185332],
        [, 'South-Eastern', , 655636576],
        [, 'Western', , 272298399],
        ['Africa', 'Eastern', , 433643132],
        [, 'Western', , 381980688],
        [, 'Northern', , 237784677],
        [, 'Others', , 234512021],
        ['Europe', , , 742648010]
    ];

    sheet.setArray(25,0, sunburstDataArray)
    sheet.setArray(0, 0, dataArray);
    sheet.resumePaint();
}

function addBarCodeConent() {
    var sheet = new spreadNS.Worksheet("Barcode");
    spread.addSheet(spread.getSheetCount(), sheet);
    sheet.suspendPaint();
    sheet.setColumnCount(50);
    sheet.getRange(1,1,11,1).font('bold normal 12px normal Arial');
    sheet.getRange(1,1,11,2).hAlign(GC.Spread.Sheets.HorizontalAlign.center).vAlign(GC.Spread.Sheets.VerticalAlign.center);
    for(var col = 1; col<4; col++){
        sheet.setColumnWidth(col,130);
    }

    for(var row = 1; row<12; row++){
        sheet.setRowHeight(row,60);
    }
    var dataArray = [
        ["QRCode", 123545346],
        ["DataMatrix", 4254534],
        ["Codabar", 1325143],
        ["PDF417", 43564364],
        ["EAN8", 1425775],
        ["EAN13", 456987123594],
        ["Code39", 423535645],
        ["Code49", 578554745],
        ["Code93", 45245325],
        ["Code128", 5246456],
        ["GS1_128", 15343566383],
    ];
    sheet.setArray(1, 1, dataArray);

    var formulaList = ["QRCODE","DATAMATRIX","CODABAR","PDF417","EAN8","EAN13","CODE39","CODE49","CODE93","CODE128","GS1_128"];

    for(var row = 1; row<12; row++){
        sheet.setFormula(row,3,'=BC_'+formulaList[row-1]+'(C'+ (row+1) +')');
    }
    sheet.resumePaint();
}

function addShapeConent(){
    var sheet = new spreadNS.Worksheet("Shape");
    spread.addSheet(spread.getSheetCount(), sheet);
    var autoTypes = GC.Spread.Sheets.Shapes.AutoShapeType;
    var names = [
        {name: "smileyFace", value: autoTypes.smileyFace, bgColor: 'orange'},
        {name: "noSymbol", value: autoTypes.noSymbol},
        {name: "heart", value: autoTypes.heart, bgColor: 'red'},
        {name: "sun", value: autoTypes.sun, bgColor: 'yellow'},
        {name: "stripedRightArrow", value: autoTypes.stripedRightArrow}
    ];
    sheet.suspendPaint();
    var left = 50, top = 50, tempX = 0, tempY = 240, tempShape = null, name , autoType;
    for(var i =  0, len = names.length ; i < len; i++) {
        name = names[i].name;
        autoType = names[i].value;
        bgColor = names[i].bgColor;
        if(name === "none") {
            continue;
        }
        tempShape = sheet.shapes.add(name, autoType, left + tempX * 240, top + tempY, 150, 150);
        tempShape.text(name);
        var style = tempShape.style();

        style.textEffect.color = 'black';
        style.textFrame.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
        style.textFrame.vAlign = GC.Spread.Sheets.VerticalAlign.center;
        if(bgColor) {
            style.fill.color = bgColor;
        }

        tempShape.style(style);

        tempX ++;
        if(tempX === 5) {
            tempX = 0;
            tempY += 240;
        }
    }

    // add connector
    var _setConnectorStyle = function(connector) {
        var connectorSltyle = connector.style();
        connectorSltyle.line.capType = 2;
        connectorSltyle.line.dashType = 0;
        connectorSltyle.line.joinType = 0;
        connectorSltyle.line.transparency = 0;
        connectorSltyle.line.color = 'rgb(91,155,213)';
        connectorSltyle.line.width = 4;
        connector.style(connectorSltyle);
    };
    var elbow = sheet.shapes.addConnector("elbow", GC.Spread.Sheets.Shapes.ConnectorType.elbow, 50, 50, 200, 200);
    _setConnectorStyle(elbow);
    var straight = sheet.shapes.addConnector("straight", GC.Spread.Sheets.Shapes.ConnectorType.straight, 300, 50 ,400, 200);
    _setConnectorStyle(straight);

    // add group shape
    var groupShapeItem1 = sheet.shapes.add("shape1", GC.Spread.Sheets.Shapes.AutoShapeType.heart, 700, 50, 150, 150);
    var groupShapeItem2 = sheet.shapes.addConnector("shape2", GC.Spread.Sheets.Shapes.ConnectorType.elbow, 900, 50, 1050, 200);
    _setConnectorStyle(groupShapeItem2);
    var shapes = [groupShapeItem1, groupShapeItem2];
    sheet.shapes.group(shapes)

    sheet.resumePaint();
}

// Sample Content related items (end)

function getCellInfo(sheet, row, column) {
    var result = {type: ""}, object;

    if ((object = sheet.comments.get(row, column))) {
        result.type = "comment";
    } else if ((object = sheet.tables.find(row, column))) {
        result.type = "table";
    }

    result.object = object;

    return result;
}

var specialTabNames = ["table", "picture", "comment", "sparklineEx", "chartEx", "slicer", "shapeEx"];
var specialTabRefs = specialTabNames.map(function (name) {
    return "#" + name + "Tab";
});

function isSpecialTabSelected() {
    var href = $(".insp-container ul.nav-tabs li.active a").attr("href");

    return specialTabRefs.indexOf(href) !== -1;
}

function getTabItem(tabName) {
    return $(".insp-container ul.nav-tabs a[href='#" + tabName + "Tab']").parent();
}

function setActiveTab(tabName) {
    // show / hide tabs
    var $target = getTabItem(tabName),
        $spreadTab = getTabItem("spread");

    if (specialTabNames.indexOf(tabName) >= 0) {
        if ($target.hasClass("hidden")) {
            hideSpecialTabs(false);

            $target.removeClass("hidden");
            $spreadTab.addClass("hidden");
            $("a", $target).tab("show");
        }
    } else {
        if ($spreadTab.hasClass("hidden")) {
            $spreadTab.removeClass("hidden");
            hideSpecialTabs(true);
        }
        if (!$target.hasClass("active")) {
            // do not switch from Data to Cell tab
            if (!(tabName === "cell" && getTabItem("data").hasClass("active"))) {
                $("a", $target).tab("show");
            }
        }
    }
}

function hideSpecialTabs(clearCache) {
    specialTabNames.forEach(function (name) {
        getTabItem(name).addClass("hidden");
    });

    if (clearCache) {
        clearCachedItems();
    }
}

function getActualRange(range, maxRowCount, maxColCount) {
    var row = range.row < 0 ? 0 : range.row;
    var col = range.col < 0 ? 0 : range.col;
    var rowCount = range.rowCount < 0 ? maxRowCount : range.rowCount;
    var colCount = range.colCount < 0 ? maxColCount : range.colCount;

    return new spreadNS.Range(row, col, rowCount, colCount);
}

function getActualCellRange(sheet, cellRange, rowCount, columnCount) {
    if (cellRange.row === -1 && cellRange.col === -1) {
        return new spreadNS.CellRange(sheet, 0, 0, rowCount, columnCount);
    }
    else if (cellRange.row === -1) {
        return new spreadNS.CellRange(sheet, 0, cellRange.col, rowCount, cellRange.colCount);
    }
    else if (cellRange.col === -1) {
        return new spreadNS.CellRange(sheet, cellRange.row, 0, cellRange.rowCount, columnCount);
    }
    return new spreadNS.CellRange(sheet, cellRange.row, cellRange.col, cellRange.rowCount, cellRange.colCount);
}

function setStyleFont(sheet, prop, isLabelStyle, optionValue1, optionValue2) {
    var styleEle = document.getElementById("setfontstyle"),
        selections = sheet.getSelections(),
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount(),
        defaultStyle = sheet.getDefaultStyle();

    function updateStyleFont(style) {
        if (!style.font) {
            style.font = defaultStyle.font || "11pt Calibri";
        }
        styleEle.style.font = style.font;
        var styleFont = $(styleEle).css(prop);
        if (styleFont === optionValue1[0] || styleFont === optionValue1[1]) {
            if (defaultStyle.font) {
                styleEle.style.font = defaultStyle.font;
                var defaultFontProp = $(styleEle).css(prop);
                styleEle.style.font = style.font;
                $(styleEle).css(prop, defaultFontProp);
            }
            else {
                $(styleEle).css(prop, optionValue2);
            }
        } else {
            $(styleEle).css(prop, optionValue1[0]);
        }
        style.font = styleEle.style.font;
    }

    sheet.suspendPaint();
    for (var n = 0; n < selections.length; n++) {
        var sel = getActualCellRange(sheet, selections[n], rowCount, columnCount);
        for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
            for (var c = sel.col; c < sel.col + sel.colCount; c++) {
                var style = sheet.getStyle(r, c);
                if (!style) {
                    style = new spreadNS.Style();
                }
                // reset themeFont to make sure font be used
                style.themeFont = undefined;
                if (isLabelStyle) {
                    if (!style.labelOptions) {
                        style.labelOptions = {};
                    }
                    updateStyleFont(style.labelOptions);
                } else {
                    updateStyleFont(style)
                }
                sheet.setStyle(r, c, style);
            }
        }
    }
    sheet.resumePaint();
}


function attachEvents() {
    attachToolbarItemEvents();
    attachSpreadEvents();
    attachConditionalFormatEvents();
    attachDataValidationEvents();
    attachOtherEvents();
    attachCellTypeEvents();
    attachLockCellsEvent();
    attachBorderTypeClickEvents();
    attachSparklineSettingEvents();
    attachChartItemEvents();
    attachShapeEvents();
}

// Border Type related items
function syncDisabledBorderType() {
    var sheet = spread.getActiveSheet();
    var selections = sheet.getSelections(), selectionsLength = selections.length;
    var isDisabledInsideBorder = true;
    var isDisabledHorizontalBorder = true;
    var isDisabledVerticalBorder = true;
    for (var i = 0; i < selectionsLength; i++) {
        var selection = selections[i];
        var col = selection.col, row = selection.row,
            rowCount = selection.rowCount, colCount = selection.colCount;
        if (isDisabledHorizontalBorder) {
            isDisabledHorizontalBorder = rowCount === 1;
        }
        if (isDisabledVerticalBorder) {
            isDisabledVerticalBorder = colCount === 1;
        }
        if (isDisabledInsideBorder) {
            isDisabledInsideBorder = rowCount === 1 || colCount === 1;
        }
    }
    [isDisabledInsideBorder, isDisabledVerticalBorder, isDisabledHorizontalBorder].forEach(function (value, index) {
        var $item = $("div.group-item:eq(" + (index * 3 + 1) + ")");
        if (value) {
            $item.addClass("disable");
        } else {
            $item.removeClass("disable");
        }
    });
}

function getBorderSettings(borderType, borderStyle) {
    var result = [];

    switch (borderType) {
        case "outside":
            result.push({lineStyle: borderStyle, options: {outline: true}});
            break;

        case "inside":
            result.push({lineStyle: borderStyle, options: {innerHorizontal: true}});
            result.push({lineStyle: borderStyle, options: {innerVertical: true}});
            break;

        case "all":
        case "none":
            result.push({lineStyle: borderStyle, options: {all: true}});
            break;

        case "left":
            result.push({lineStyle: borderStyle, options: {left: true}});
            break;

        case "innerVertical":
            result.push({lineStyle: borderStyle, options: {innerVertical: true}});
            break;

        case "right":
            result.push({lineStyle: borderStyle, options: {right: true}});
            break;

        case "top":
            result.push({lineStyle: borderStyle, options: {top: true}});
            break;

        case "innerHorizontal":
            result.push({lineStyle: borderStyle, options: {innerHorizontal: true}});
            break;

        case "bottom":
            result.push({lineStyle: borderStyle, options: {bottom: true}});
            break;
        case "diagonalUp":
            result.push({lineStyle: borderStyle, options: {up: true}});
            break;
        case "diagonalDown":
            result.push({lineStyle: borderStyle, options: {down: true}});

            break;
    }

    return result;
}

function setBorderlines(sheet, borderType, borderStyle, borderColor) {
    function setSheetBorder(setting) {
        var lineBorder = new spreadNS.LineBorder(borderColor, setting.lineStyle);
        var options = setting.options;
        if(options.up) {
            sel.diagonalUp(lineBorder);
        } else if (options.down) {
            sel.diagonalDown(lineBorder);
        } else {
            sel.setBorder(lineBorder, setting.options);
            setRangeBorder(sheet, sel, setting.options, lineBorder);
        }
    }

    var settings = getBorderSettings(borderType, borderStyle);
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    var sels = sheet.getSelections();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        settings.forEach(setSheetBorder);
    }
    sheet.resumePaint();
}

function attachBorderTypeClickEvents() {
    var $groupItems = $(".group-item>div");
    $groupItems.bind("mousedown", function () {
        if ($(this).parent().hasClass("disable")) {
            return;
        }
        var name = $(this).data("name").split("Border")[0];
        applyBorderSetting(name);
    });
}

function applyBorderSetting(name) {
    var sheet = spread.getActiveSheet();
    var borderLine = getBorderLineType($("#border-line-type").attr("class"));
    var borderColor = getBackgroundColor("borderColor");
    setBorderlines(sheet, name, borderLine, borderColor);
}

function setDiagonalLines(sheet, name, borderLine, borderColor) {
    var lineBorder = new spreadNS.LineBorder(borderColor, borderLine);
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    sheet.suspendPaint();
    var sels = sheet.getSelections();

    for (var n = 0; n < sels.length; n++) {
        var sel = getActualCellRange(sheet, sels[n], rowCount, columnCount);
        name === "diagonalUpLine" ? sel.diagonalUp(name) : sel.diagonalDown(name);
    }
    sheet.resumePaint();
}

function getBorderLineType(className) {
    switch (className) {
        case "no-border":
            return spreadNS.LineStyle.empty;

        case "line-style-hair":
            return spreadNS.LineStyle.hair;

        case "line-style-dotted":
            return spreadNS.LineStyle.dotted;

        case "line-style-dash-dot-dot":
            return spreadNS.LineStyle.dashDotDot;

        case "line-style-dash-dot":
            return spreadNS.LineStyle.dashDot;

        case "line-style-dashed":
            return spreadNS.LineStyle.dashed;

        case "line-style-thin":
            return spreadNS.LineStyle.thin;

        case "line-style-medium-dash-dot-dot":
            return spreadNS.LineStyle.mediumDashDotDot;

        case "line-style-slanted-dash-dot":
            return spreadNS.LineStyle.slantedDashDot;

        case "line-style-medium-dash-dot":
            return spreadNS.LineStyle.mediumDashDot;

        case "line-style-medium-dashed":
            return spreadNS.LineStyle.mediumDashed;

        case "line-style-medium":
            return spreadNS.LineStyle.medium;

        case "line-style-thick":
            return spreadNS.LineStyle.thick;

        case "line-style-double":
            return spreadNS.LineStyle.double;
    }
}

function getArrowStyleType(className) {
    switch (className) {
        case "begin-arrow-style-none":
        case "end-arrow-style-none":
            return spreadNS.Shapes.ArrowheadStyle.none;

        case "begin-arrow-style-triangle":
        case "end-arrow-style-triangle":
            return spreadNS.Shapes.ArrowheadStyle.triangle;

        case "begin-arrow-style-stealth":
        case "end-arrow-style-stealth":
            return spreadNS.Shapes.ArrowheadStyle.stealth;

        case "begin-arrow-style-diamond":
        case "end-arrow-style-diamond":
            return spreadNS.Shapes.ArrowheadStyle.diamond;

        case "begin-arrow-style-oval":
        case "end-arrow-style-oval":
            return spreadNS.Shapes.ArrowheadStyle.oval;

        case "begin-arrow-style-open":
        case "end-arrow-style-open":
            return spreadNS.Shapes.ArrowheadStyle.open;
    }
}

function processArrowStyleSetting(name,nameValue){
    var $arrowStyleType,prefix;
    if(name == "beginArrowStyle"){
        $arrowStyleType = $('#begin-arrow-style-type');
        prefix = "begin-";
    }else{
        $arrowStyleType = $('#end-arrow-style-type');
        prefix = "end-";
    }
    $arrowStyleType.text("");
    $arrowStyleType.removeClass();
    switch (nameValue) {
        case 'none':
            $arrowStyleType.addClass(prefix + "arrow-style-none");
            return;

        case 'triangle':
            $arrowStyleType.addClass(prefix + "arrow-style-triangle");
            break;

        case 'stealth':
            $arrowStyleType.addClass(prefix + "arrow-style-stealth");
            break;

        case 'diamond':
            $arrowStyleType.addClass(prefix + "arrow-style-diamond");
            break;

        case 'oval':
            $arrowStyleType.addClass(prefix + "arrow-style-oval");
            break;

        case 'open':
            $arrowStyleType.addClass(prefix + "arrow-style-open");
            break;

        default:
            //console.log("processArrowStyleSetting not add for ", name);
            break;
    }
}

function processBorderLineSetting(name) {
    var $borderLineType = $('#border-line-type');
    $borderLineType.text("");
    $borderLineType.removeClass();
    switch (name) {
        case "none":
        case "0":
            $('#border-line-type').text(getResource("cellTab.border.noBorder"));
            $('#border-line-type').addClass("no-border");
            return;

        case "hair":
            $('#border-line-type').addClass("line-style-hair");
            break;

        case "dotted":
            $('#border-line-type').addClass("line-style-dotted");
            break;

        case "dash-dot-dot":
            $('#border-line-type').addClass("line-style-dash-dot-dot");
            break;

        case "dash-dot":
            $('#border-line-type').addClass("line-style-dash-dot");
            break;

        case "dashed":
            $('#border-line-type').addClass("line-style-dashed");
            break;

        case "thin":
            $('#border-line-type').addClass("line-style-thin");
            break;

        case "medium-dash-dot-dot":
            $('#border-line-type').addClass("line-style-medium-dash-dot-dot");
            break;

        case "slanted-dash-dot":
            $('#border-line-type').addClass("line-style-slanted-dash-dot");
            break;

        case "medium-dash-dot":
            $('#border-line-type').addClass("line-style-medium-dash-dot");
            break;

        case "medium-dashed":
            $('#border-line-type').addClass("line-style-medium-dashed");
            break;

        case "medium":
            $('#border-line-type').addClass("line-style-medium");
            break;

        case "thick":
            $('#border-line-type').addClass("line-style-thick");
            break;

        case "double":
            $('#border-line-type').addClass("line-style-double");
            break;

        default:
            //console.log("processBorderLineSetting not add for ", name);
            break;
    }
}

function processShapeBorderLineSetting(value) {
    var $shapeBorderLineType = $('#shape-border-line-type');
    $shapeBorderLineType.text("");
    $shapeBorderLineType.removeClass();
    var borderStyleMap = {
        solid: 'shape-border-style-solid',
        squareDot: 'shape-border-style-square-dot',
        dash: 'shape-border-style-dash',
        longDash: 'shape-border-style-long-dash',
        dashDot: 'shape-border-style-dash-dot',
        longDashDot: 'shape-border-style-long-dash-dot',
        longDashDotDot: 'shape-border-style-long-dash-dot-dot',
        sysDash: 'shape-border-style-sys-dash',
        sysDot: 'shape-border-style-sys-dot',
        sysDashDot: 'shape-border-style-sys-dash-dot',
        dashDotDot: 'shape-border-style-dash-dot-dot'
    };
    if(borderStyleMap[value]) {
        $shapeBorderLineType.addClass(borderStyleMap[value]);
        $shapeBorderLineType.data("value", value);
    }
}

function setRangeBorder(sheet, range, options) {
    var outline = options.all || options.outline,
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount(),
        startRow = range.row, endRow = startRow + range.rowCount - 1,
        startCol = range.col, endCol = startCol + range.colCount - 1;

    // update related borders for all cells arround the range

    // left side
    if ((startCol > 0) && (outline || options.left)) {
        sheet.getRange(startRow, startCol - 1, range.rowCount, 1).borderRight(undefined);
    }
    // top side
    if ((startRow > 0) && (outline || options.top)) {
        sheet.getRange(startRow - 1, startCol, 1, range.colCount).borderBottom(undefined);
    }
    // right side
    if ((endCol < columnCount - 1) && (outline || options.right)) {
        sheet.getRange(startRow, endCol + 1, range.rowCount, 1).borderLeft(undefined);
    }
    // bottom side
    if ((endRow < rowCount - 1) && (outline || options.bottom)) {
        sheet.getRange(endRow + 1, startCol, 1, range.colCount).borderTop(undefined);
    }
}
// Border Type related items (end)

function attachOtherEvents() {
    $("div.table-format-item").click(changeTableStyle);
    $("div.slicer-format-item").click(changeSlicerStyle);
    $("#fileSelector").change(processFileSelected);
    $("#sparklineextypes button").click(processAddSparklineEx);
    $("#chartContainer button").click(processAddChartEx);
    $("#connectorShapeTypeContainer button").click(processAddConnectorShapeEx);
    var shapeContainers = ['shapeRectanglesContainer', 'shapeBasicsContainer' ,'shapeBlockArrowsContainer','shapeEquationsContainer','shapeFlowchartContainer','shapeStarsAndBannersContainer','shapeCalloutsContainer'];
    shapeContainers.forEach(function(container) {
        $('#' + container + ' button').click(processAddShapeEx);
    });
}

function processFileSelected() {
    var file = this.files[0],
        action = $(this).data("action");

    if (!file) return false;

    // clear to make sure change event occures even when same file selected again
    $("#fileSelector").val("");

    if (action === "doImport") {
        return importFile(file);
    }

    if (!/image\/\w+/.test(file.type)) {
        alert(getResource("messages.imageFileRequired"));
        return false;
    }
    var reader = new FileReader();
    reader.onload = function () {
        switch (action) {
            case "addpicture":
                addPicture(this.result);
                break;
            case "iconUpload":
                dataValidationIconUpload(this.result);
                break;
        }
    };
    reader.readAsDataURL(file);
}

var PICTURE_ROWCOUNT = 16, PICTURE_COLUMNCOUNT = 10;
function addPicture(pictureUrl) {
    var sheet = spread.getActiveSheet();
    var defaults = sheet.defaults, rowHeight = defaults.rowHeight, colWidth = defaults.colWidth;
    var sel = sheet.getSelections()[0];
    if (pictureUrl !== "" && sel) {
        sheet.suspendPaint();

        var cr = getActualRange(sel, sheet.getRowCount(), sheet.getColumnCount());
        var name = "Picture" + pictureIndex;
        pictureIndex++;

        // prepare and adjust the range for add picture
        var row = cr.row, col = cr.col,
            endRow = row + PICTURE_ROWCOUNT,
            endColumn = col + PICTURE_COLUMNCOUNT,
            rowCount = sheet.getRowCount(),
            columnCount = sheet.getColumnCount();

        if (endRow > rowCount) {
            endRow = rowCount - 1;
            row = endRow - PICTURE_ROWCOUNT;
        }

        if (endColumn > columnCount) {
            endColumn = columnCount - 1;
            col = endColumn - PICTURE_COLUMNCOUNT;
        }

        var picture = sheet.pictures.add(name, pictureUrl, col * colWidth, row * rowHeight, (endColumn - col) * colWidth, (endRow - row) * rowHeight)
            .backColor("#FFFFFF").borderColor("#000000")
            .borderStyle("solid").borderWidth(1).borderRadius(3);
        sheet.resumePaint();

        spread.focus();
        picture.isSelected(true);
    }
}

function dataValidationIconUpload(icon) {
    $('#iconUpload').data('icon', icon);
    $('#iconUploadPreview').show();
    $('#iconUploadPreview').attr('src', icon);
}

function updatePositionBox(sheet) {
    var selection = sheet.getSelections().slice(-1)[0];
    if (selection) {
        var position;
        if (!isShiftKey) {
            position = getCellPositionString(sheet,
                sheet.getActiveRowIndex() + 1,
                sheet.getActiveColumnIndex() + 1, selection);
        }
        else {
            position = getSelectedRangeString(sheet, selection);
        }

        $("#positionbox").val(position);
    }
}

function syncCellRelatedItems() {
    updateMergeButtonsState();
    syncDisabledLockCells();
    syncDisabledBorderType();

    // reset conditional format setting
    var item = setDropDownValueByIndex($("#conditionalFormatType"), -1);
    processConditionalFormatDetailSetting(item.value, true);
    // sync cell type related information
    syncCellTypeInfo();
}

function syncCellTypeInfo() {
    function updateButtonCellTypeInfo(cellType) {
        setNumberValue("buttonCellTypeMarginTop", cellType.marginTop());
        setNumberValue("buttonCellTypeMarginRight", cellType.marginRight());
        setNumberValue("buttonCellTypeMarginBottom", cellType.marginBottom());
        setNumberValue("buttonCellTypeMarginLeft", cellType.marginLeft());
        setTextValue("buttonCellTypeText", cellType.text());
        setColorValue("buttonCellTypeBackColor", cellType.buttonBackColor());
    }

    function updateCheckBoxCellTypeInfo(cellType) {
        setTextValue("checkboxCellTypeCaption", cellType.caption());
        setTextValue("checkboxCellTypeTextTrue", cellType.textTrue());
        setTextValue("checkboxCellTypeTextIndeterminate", cellType.textIndeterminate());
        setTextValue("checkboxCellTypeTextFalse", cellType.textFalse());
        setDropDownValue("checkboxCellTypeTextAlign", cellType.textAlign());
        setCheckValue("checkboxCellTypeIsThreeState", cellType.isThreeState());
    }

    function updateComboBoxCellTypeInfo(cellType) {
        setDropDownValue("comboboxCellTypeEditorValueType", cellType.editorValueType());
        var items = cellType.items(),
            texts = items.map(function (item) {
                return item.text || item;
            }).join(","),
            values = items.map(function (item) {
                return item.value || item;
            }).join(",");

        setTextValue("comboboxCellTypeItemsText", texts);
        setTextValue("comboboxCellTypeItemsValue", values);
    }

    function updateHyperLinkCellTypeInfo(cellType) {
        setColorValue("hyperlinkCellTypeLinkColor", cellType.linkColor());
        setColorValue("hyperlinkCellTypeVisitedLinkColor", cellType.visitedLinkColor());
        setTextValue("hyperlinkCellTypeText", cellType.text());
        setTextValue("hyperlinkCellTypeLinkToolTip", cellType.linkToolTip());
    }

    var sheet = spread.getActiveSheet(),
        index,
        cellType = sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()).cellType();

    if (cellType instanceof spreadNS.CellTypes.Button) {
        index = 0;
        updateButtonCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.CellTypes.CheckBox) {
        index = 1;
        updateCheckBoxCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.CellTypes.ComboBox) {
        index = 2;
        updateComboBoxCellTypeInfo(cellType);
    } else if (cellType instanceof spreadNS.CellTypes.HyperLink) {
        index = 3;
        updateHyperLinkCellTypeInfo(cellType);
    } else {
        index = -1;
    }
    var cellTypeItem = setDropDownValueByIndex($("#cellTypes"), index);
    processCellTypeSetting(cellTypeItem.value, true);

    if (index >= 0) {
        var $group = $("#groupCellType");
        if ($group.find(".group-state").hasClass("fa-caret-right")) {
            $group.click();
        }
    }
}

function onCellSelected() {
    $("#addslicer").addClass("hidden");
    var sheet = spread.getActiveSheet(),
        row = sheet.getActiveRowIndex(),
        column = sheet.getActiveColumnIndex();

    if (showSparklineSetting(row, column)) {
        setActiveTab("sparklineEx");
        return;
    }
    var cellInfo = getCellInfo(sheet, row, column),
        cellType = cellInfo.type;

    syncCellRelatedItems();
    updatePositionBox(sheet);
    updateCellStyleState(sheet, row, column);

    var tabType = "cell";

    clearCachedItems();

    // add map from cell type to tab type here
    if (cellType === "table") {
        tabType = "table";
        syncTablePropertyValues(sheet, cellInfo.object);
        $("#addslicer").removeClass("hidden");
    } else if (cellType === "comment") {
        tabType = "comment";
        syncCommentPropertyValues(sheet, cellInfo.object);
    }

    setActiveTab(tabType);
}

var _activeComment;

function syncCommentPropertyValues(sheet, comment) {
    _activeComment = comment;

    // General
    setCheckValue("commentDynamicSize", comment.dynamicSize());
    setCheckValue("commentDynamicMove", comment.dynamicMove());
    setCheckValue("commentLockText", comment.lockText());
    setCheckValue("commentShowShadow", comment.showShadow());

    // Font
    setDropDownText($("#commentTab div.insp-dropdown-list[data-name='commentFontFamily']"), comment.fontFamily());
    setDropDownText($("#commentTab div.insp-dropdown-list[data-name='commentFontSize']"), parseFloat(comment.fontSize()));
    setDropDownText($("#commentTab div.insp-dropdown-list[data-name='commentFontStyle']"), comment.fontStyle());
    setDropDownText($("#commentTab div.insp-dropdown-list[data-name='commentFontWeight']"), comment.fontWeight());
    var textDecoration = comment.textDecoration();
    var TextDecorationType = spreadNS.TextDecorationType;
    setFontStyleButtonActive("comment-underline", (textDecoration & TextDecorationType.underline) === TextDecorationType.underline);
    setFontStyleButtonActive("comment-overline", (textDecoration & TextDecorationType.overline) === TextDecorationType.overline);
    setFontStyleButtonActive("comment-strikethrough", (textDecoration & TextDecorationType.lineThrough) === TextDecorationType.lineThrough);

    // Border
    setNumberValue("commentBorderWidth", comment.borderWidth());
    setDropDownText($("#commentTab div.insp-dropdown-list[data-name='commentBorderStyle']"), comment.borderStyle());
    setColorValue("commentBorderColor", comment.borderColor());

    // Appearance
    setDropDownValue($("#commentTab div.insp-dropdown-list[data-name='commentHorizontalAlign']"), comment.horizontalAlign());
    setDropDownValue($("#commentTab div.insp-dropdown-list[data-name='commentDisplayMode']"), comment.displayMode());
    setColorValue("commentForeColor", comment.foreColor());
    setColorValue("commentBackColor", comment.backColor());
    setTextValue("commentPadding", getPaddingString(comment.padding()));
    setNumberValue("commentOpacity", comment.opacity() * 100);
}

function getPaddingString(padding) {
    if (!padding) return "";

    return [padding.top, padding.right, padding.bottom, padding.left].join(", ");
}

function clearCachedItems() {
    _activePicture = null;
    _activeComment = null;
    _activeTable = null;
}

var _activeTable;
function syncTablePropertyValues(sheet, table) {
    _activeTable = table;

    setCheckValue("tableFilterButton", table.filterButtonVisible());

    setCheckValue("tableHeaderRow", table.showHeader());
    setCheckValue("tableTotalRow", table.showFooter());

    setCheckValue("tableFirstColumn", table.highlightFirstColumn());
    setCheckValue("tableLastColumn", table.highlightLastColumn());
    setCheckValue("tableBandedRows", table.bandRows());
    setCheckValue("tableBandedColumns", table.bandColumns());
    var tableStyle = table.style(),
        styleName = tableStyle && table.style().name();

    $("#tableStyles .table-format-item").removeClass("table-format-item-selected");
    if (styleName) {
        $("#tableStyles .table-format-item div[data-name='" + styleName.toLowerCase() + "']").parent().addClass("table-format-item-selected");
    }
    setTextValue("tableName", table.name());
}

function changeTableStyle() {
    if (_activeTable) {
        spread.suspendPaint();

        var styleName = $(">div", this).data("name");

        _activeTable.style(spreadNS.Tables.TableThemes[styleName]);

        $("#tableStyles .table-format-item").removeClass("table-format-item-selected");
        $(this).addClass("table-format-item-selected");

        spread.resumePaint();
    }
}

var _activePicture;
function syncPicturePropertyValues(sheet, picture) {
    _activePicture = picture;

    // General
    if (picture.dynamicMove()) {
        if (picture.dynamicSize()) {
            setRadioItemChecked("pictureMoveAndSize", "picture-move-size");
        }
        else {
            setRadioItemChecked("pictureMoveAndSize", "picture-move-nosize");
        }
    }
    else {
        setRadioItemChecked("pictureMoveAndSize", "picture-nomove-size");
    }
    setCheckValue("pictureFixedPosition", picture.fixedPosition());

    // Border
    setNumberValue("pictureBorderWidth", picture.borderWidth());
    setNumberValue("pictureBorderRadius", picture.borderRadius());
    setDropDownText($("#pictureTab div.insp-dropdown-list[data-name='pictureBorderStyle']"), picture.borderStyle());
    setColorValue("pictureBorderColor", picture.borderColor());

    // Appearance
    setDropDownValue($("#pictureTab div.insp-dropdown-list[data-name='pictureStretch']"), picture.pictureStretch());
    setColorValue("pictureBackColor", picture.backColor());

    $("#positionbox").val(picture.name());
}

var _floatInspector = false;

function adjustInspectorDisplay() {
    var $inspectorContainer = $(".insp-container"),
        $contentContainer = $("#inner-content-container"),
        toggleInspectorClasses;

    if (_floatInspector) {
        $inspectorContainer.draggable("enable");
        $inspectorContainer.addClass("float-inspector");
        $contentContainer.addClass("float-inspector");
        toggleInspectorClasses = ["fa-angle-down", "fa-angle-up"];
        $("#inner-content-container").addClass("hide-inspector");
    } else {
        $inspectorContainer.draggable("disable");
        $inspectorContainer.removeClass("float-inspector");
        $inspectorContainer.css({left: "auto", top: 0});
        $contentContainer.removeClass("float-inspector");
        toggleInspectorClasses = ["fa-angle-left", "fa-angle-right"];
    }

    // update toggleInspector
    var classIndex = ($(".insp-container:visible").length > 0) ? 1 : 0;
    $("#toggleInspector > span")
        .removeClass("fa-angle-left fa-angle-right fa-angle-up fa-angle-down")
        .addClass(toggleInspectorClasses[classIndex]);
}
function processMediaQueryResponse(mql) {
    if (mql.matches) {
        if (!_floatInspector) {
            _floatInspector = true;
            adjustInspectorDisplay();
        }
    } else {
        if (_floatInspector) {
            _floatInspector = false;
            adjustInspectorDisplay();
        }
    }
}

function checkMediaSize() {
    var mql = window.matchMedia("screen and (max-width: 768px)");
    processMediaQueryResponse(mql);
    adjustInspectorDisplay();
    mql.addListener(processMediaQueryResponse);
}

function toggleInspector() {
    if ($(".insp-container:visible").length > 0) {
        $(".insp-container").hide();
        if (!_floatInspector) {
            $("#inner-content-container").addClass("hide-inspector");
            $("span", this).removeClass("fa-angle-right fa-angle-up fa-angle-down").addClass("fa-angle-left");
        } else {
            $("#inner-content-container").addClass("hide-inspector");
            $("span", this).removeClass("fa-angle-right fa-angle-left fa-angle-up").addClass("fa-angle-down");
        }

        $(this).attr("title", uiResource.toolBar.showInspector);
    } else {
        $(".insp-container").show();
        if (!_floatInspector) {
            $("#inner-content-container").removeClass("hide-inspector");
            $("span", this).removeClass("fa-angle-left fa-angle-up fa-angle-down").addClass("fa-angle-right");
        } else {
            $("#inner-content-container").addClass("hide-inspector");
            $("span", this).removeClass("fa-angle-right fa-angle-left fa-angle-down").addClass("fa-angle-up");
        }

        $(this).attr("title", uiResource.toolBar.hideInspector);
    }
    spread.refresh();
}

function attachToolbarItemEvents() {
    $("#addtable").click(function () {
        var sheet = spread.getActiveSheet(),
            row = sheet.getActiveRowIndex(),
            column = sheet.getActiveColumnIndex(),
            name = "Table" + tableIndex,
            rowCount = 1,
            colCount = 1;

        tableIndex++;

        var selections = sheet.getSelections();

        if (selections.length > 0) {
            var range = selections[0],
                r = range.row,
                c = range.col;

            rowCount = range.rowCount,
                colCount = range.colCount;

            // update row / column for whole column / row was selected
            if (r >= 0) {
                row = r;
            }
            if (c >= 0) {
                column = c;
            }
        }

        sheet.suspendPaint();
        try {
            // handle exception if the specified range intersect with other table etc.
            sheet.tables.add(name, row, column, rowCount, colCount, spreadNS.Tables.TableThemes.light2);
        } catch (e) {
            alert(e.message);
        }
        sheet.resumePaint();

        spread.focus();

        onCellSelected();
    });

    $("#addcomment").click(function () {
        var sheet = spread.getActiveSheet(),
            row = sheet.getActiveRowIndex(),
            column = sheet.getActiveColumnIndex(),
            comment;

        sheet.suspendPaint();
        comment = sheet.comments.add(row, column, new Date().toLocaleString());
        sheet.resumePaint();

        comment.commentState(spreadNS.Comments.CommentState.edit);
    });

    $("#addpicture, #doImport").click(function () {
        $("#fileSelector").data("action", this.id);
        $("#fileSelector").click();
    });

    $("#toggleInspector").click(toggleInspector);

    $("#doClear").click(function () {
        var $dropdown = $("#clearActionList"),
            $this = $(this),
            offset = $this.offset();

        $dropdown.css({left: offset.left, top: offset.top + $this.outerHeight()});
        $dropdown.show();
        processEventListenerHandleClosePopup(true);
    });

    $("#doExport").click(function () {
        var $dropdown = $("#exportActionList"),
            $this = $(this),
            offset = $this.offset();

        $dropdown.css({left: offset.left, top: offset.top + $this.outerHeight()});
        $dropdown.show();
        processEventListenerHandleClosePopup(true);
    });

    $("#addslicer").click(processAddSlicer);
}

// Protect Sheet related items
function getCurrentSheetProtectionOption(sheet) {
    var options = sheet.options.protectionOptions;
    if (options.allowSelectLockedCells || options.allowSelectLockedCells === undefined) {
        setCheckValue("checkboxSelectLockedCells", true);
    }
    else {
        setCheckValue("checkboxSelectLockedCells", false);
    }
    if (options.allowSelectUnlockedCells || options.allowSelectUnlockedCells === undefined) {
        setCheckValue("checkboxSelectUnlockedCells", true);
    }
    else {
        setCheckValue("checkboxSelectUnlockedCells", false);
    }
    if (options.allowSort) {
        setCheckValue("checkboxSort", true);
    }
    else {
        setCheckValue("checkboxSort", false);
    }
    if (options.allowFilter) {
        setCheckValue("checkboxUseAutoFilter", true);
    }
    else {
        setCheckValue("checkboxUseAutoFilter", false);
    }
    if (options.allowResizeRows) {
        setCheckValue("checkboxResizeRows", true);
    }
    else {
        setCheckValue("checkboxResizeRows", false);
    }
    if (options.allowResizeColumns) {
        setCheckValue("checkboxResizeColumns", true);
    }
    else {
        setCheckValue("checkboxResizeColumns", false);
    }
    if (options.allowEditObjects) {
        setCheckValue("checkboxEditObjects", true);
    }
    else {
        setCheckValue("checkboxEditObjects", false);
    }
}

function setProtectionOption(sheet, optionItem, value) {
    var options = sheet.options.protectionOptions;
    switch (optionItem) {
        case "allowSelectLockedCells":
            options.allowSelectLockedCells = value;
            break;
        case "allowSelectUnlockedCells":
            options.allowSelectUnlockedCells = value;
            break;
        case "allowSort":
            options.allowSort = value;
            break;
        case "allowFilter":
            options.allowFilter = value;
            break;
        case "allowResizeRows":
            options.allowResizeRows = value;
            break;
        case "allowResizeColumns":
            options.allowResizeColumns = value;
            break;
        case "allowEditObjects":
            options.allowEditObjects = value;
            break;
        case "allowDragInsertRows":
            options.allowDragInsertRows = value;
            break;
        case "allowDragInsertColumns":
            options.allowDragInsertColumns = value;
            break;
        case "allowInsertRows":
            options.allowInsertRows = value;
            break;
        case "allowInsertColumns":
            options.allowInsertColumns = value;
            break;
        case "allowDeleteRows":
            options.allowDeleteRows = value;
            break;
        case "allowDeleteColumns":
            options.allowDeleteColumns = value;
            break;
        default:
            //console.log("There is no protection option:", optionItem);
            break;
    }
    setActiveTab("sheet");
}

function syncSheetProtectionText(isProtected) {
    var $protectSheetText = $("#protectSheetText");
    if (isProtected) {
        $protectSheetText.text(uiResource.cellTab.protection.sheetIsProtected);
    }
    else {
        $protectSheetText.text(uiResource.cellTab.protection.sheetIsUnprotected);
    }
}

function syncProtectSheetRelatedItems(sheet, value) {
    sheet.options.isProtected = value;
    syncSheetProtectionText(value);

    if (isAllSelectedSlicersLocked(sheet)) {
        setActiveTab("sheet");
    }
}

function isAllSelectedSlicersLocked(sheet) {
    var selectedSlicers = getSelectedSlicers(sheet);
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return null;
    }
    var allLocked = true;
    for (var item in selectedSlicers) {
        allLocked = allLocked && selectedSlicers[item].isLocked();
        if (!allLocked) {
            break;
        }
    }
    return allLocked;
}
// Protect Sheet related items (end)

// Lock Cell related items
function getCellsLockedState() {
    var isLocked = false;
    var sheet = spread.getActiveSheet();
    var selections = sheet.getSelections(), selectionsLength = selections.length;
    var cell;
    var row, col, rowCount, colCount;
    if (selectionsLength > 0) {
        for (var i = 0; i < selectionsLength; i++) {
            var range = selections[i];
            row = range.row;
            rowCount = range.rowCount;
            colCount = range.colCount;
            if (row < 0) {
                row = 0;
            }
            for (row; row < range.row + rowCount; row++) {
                col = range.col;
                if (col < 0) {
                    col = 0;
                }
                for (col; col < range.col + colCount; col++) {
                    cell = sheet.getCell(row, col);
                    isLocked = isLocked || cell.locked();
                    if (isLocked) {
                        return isLocked;
                    }
                }
            }
        }
        return false;
    } else {
        return sheet.getCell(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex()).locked();
    }
}

function syncDisabledLockCells() {
    var cellsLockedState = getCellsLockedState();
    setCheckValue("checkboxLockCell", cellsLockedState);
}

function attachLockCellsEvent() {
    $("#lockCells").click(function () {
        var value = getCheckValue("checkboxLockCell");
        setSelectedCellsLock(value);
    });
}

function setSelectedCellsLock(value) {
    var sheet = spread.getActiveSheet();
    var selections = sheet.getSelections();
    var row, col, rowCount, colCount;
    for (var i = 0; i < selections.length; i++) {
        var range = selections[i];
        row = range.row;
        col = range.col;
        rowCount = range.rowCount;
        colCount = range.colCount;
        if (row < 0 && col < 0) {
            sheet.getDefaultStyle().locked = value;
        }
        else if (row < 0) {
            sheet.getRange(-1, col, -1, colCount).locked(value);
        }
        else if (col < 0) {
            sheet.getRange(row, -1, rowCount, -1).locked(value);
        }
        else {
            sheet.getRange(row, col, rowCount, colCount).locked(value);
        }
    }
}
// Lock Cell related items (end)

function attachSpreadEvents(rebind) {
    spread.bind(spreadNS.Events.EnterCell, onCellSelected);

    spread.bind(spreadNS.Events.ValueChanged, function (sender, args) {
        var row = args.row, col = args.col, sheet = args.sheet;

        if (sheet.getCell(row, col).wordWrap()) {
            sheet.autoFitRow(row);
        }
    });

    function shouldAutofitRow(sheet, row, col, colCount) {
        for (var c = 0; c < colCount; c++) {
            if (sheet.getCell(row, col++).wordWrap()) {
                return true;
            }
        }

        return false;
    }

    spread.bind(spreadNS.Events.RangeChanged, function (sender, args) {
        var sheet = args.sheet, row = args.row, rowCount = args.rowCount;

        if (args.action === spreadNS.RangeChangedAction.paste) {
            var col = args.col, colCount = args.colCount;
            for (var i = 0; i < rowCount; i++) {
                if (shouldAutofitRow(sheet, row, col, colCount)) {
                    sheet.autoFitRow(row);
                }
                row++;
            }
        }
    });

    spread.bind(spreadNS.Events.ActiveSheetChanged, function () {
        setActiveTab("sheet");
        syncSheetPropertyValues();
        syncCellRelatedItems();

        var sheet = spread.getActiveSheet(),
            picture,
            chart,
            shape;
        var slicers = sheet.slicers.all();
        for (var item in slicers) {
            slicers[item].isSelected(false);
        }

        if (sheet.getSelections().length === 0) {
            sheet.pictures.all().forEach(function (pic) {
                if (!picture && pic.isSelected()) {
                    picture = pic;
                }
            });

            sheet.charts.all().forEach(function (cha) {
                if(!chart && cha.isSelected()){
                    chart = cha;
                }
            })

            sheet.shapes.all().forEach(function (sha) {
                if (!shape && sha.isSelected()) {
                    shape = sha;
                }
            });
            // fix bug, make sure selection was shown after unselect slicer
            if (!picture || !chart || !shape) {
                sheet.setSelection(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex(), 1, 1);
            }
        }
        if (picture) {
            syncPicturePropertyValues(sheet, picture);
            setActiveTab("picture");
        } else if (chart) {
            //syncChartPropertyValues(sheet, chart)
            showChartPanel(chart);
        } else if (shape) {
            showShapePanel(shape);
        } else{
            onCellSelected();
        }

        var value = $("div.button", $("div[data-name='allowOverflow']")).hasClass("checked");
        if (sheet.options.allowCellOverflow !== value) {
            sheet.options.allowCellOverflow = value;
        }
    });

    spread.bind(spreadNS.Events.SelectionChanging, function () {
        var sheet = spread.getActiveSheet();
        var selection = sheet.getSelections().slice(-1)[0];
        if (selection) {
            var position = getSelectedRangeString(sheet, selection);
            $("#positionbox").val(position);
        }
        syncDisabledBorderType();
    });

    spread.bind(spreadNS.Events.SelectionChanged, function () {
        syncCellRelatedItems();

        updatePositionBox(spread.getActiveSheet());
    });

    spread.bind(spreadNS.Events.PictureSelectionChanged, function (event, args) {
        var sheet = args.sheet, picture = args.picture;

        if (picture && picture.isSelected()) {
            syncPicturePropertyValues(sheet, picture);
            setActiveTab("picture");
        }
    });

    spread.bind(spreadNS.Events.FloatingObjectChanged, function (event, args) {
        var floatingObject = args.floatingObject;
        if (floatingObject && floatingObject instanceof spreadNS.Charts.Chart) {
            showChartPanel(floatingObject);
        }
    });

    // spread.bind(spreadNS.Events.ChartClicked, function (event, args) {
    //     var sheet = args.sheet, chart = args.chart;
    //     showChartPanel(chart);
    // });

    spread.bind(spreadNS.Events.ShapeChanged, function (event, args) {
        var sheet = args.sheet, shape = args.shape;
        showShapePanel(shape, true);
    });

    spread.bind(spreadNS.Events.ShapeSelectionChanged, function (event, args) {
        var sheet = args.sheet, shape = args.shape;
        showShapePanel(shape);
    });

    spread.bind(spreadNS.Events.CommentChanged, function (event, args) {
        var sheet = args.sheet, comment = args.comment, propertyName = args.propertyName;

        if (propertyName === "commentState" && comment) {
            if (comment.commentState() === spreadNS.Comments.CommentState.edit) {
                syncCommentPropertyValues(sheet, comment);
                setActiveTab("comment");
            }
        }
    });

    spread.bind(spreadNS.Events.ValidationError, function (event, data) {
        var dv = data.validator;
        if (dv) {
            alert(dv.errorMessage() || dv.inputMessage());
        }
    });

    spread.bind(spreadNS.Events.SlicerChanged, function (event, args) {
        bindSlicerEvents(args.sheet, args.slicer, args.propertyName);
    });

    spread.bind(spreadNS.Events.ActiveSheetChanged, function (event, args) {
        var newSheet = args.newSheet;
        if(newSheet.name() === 'Chart'){
            newSheet.setColumnWidth(1, 100, GC.Spread.Sheets.SheetArea.viewport);
            newSheet.setColumnWidth(3, 100, GC.Spread.Sheets.SheetArea.viewport);
            if(isFirstChart){
                var chartCount = newSheet.charts.all().length || 0;
                var columnType = GC.Spread.Sheets.Charts.ChartType.columnClustered;
                var sunburstType = GC.Spread.Sheets.Charts.ChartType.sunburst;
                var lineType = GC.Spread.Sheets.Charts.ChartType.line;
                var lineChart = newSheet.charts.add(('ChartLine' + chartCount), lineType, 550, 130, 450, 300, "Chart!$A$1:$H$5");
                var columnChart = newSheet.charts.add(('ChartColumn' + chartCount), columnType, 30, 130, 450, 300, "Chart!$A$1:$H$5");
                var sunburstChart = newSheet.charts.add(('ChartSunburst' + chartCount), sunburstType, 550, 500, 450, 300, "Chart!$A$26:$D$37");
                var allCharts = newSheet.charts.all();
                allCharts.forEach(function(chart){
                    var chartType = getChartGroupString(chart.chartType());
                    if(chartType === "ColumnGroup" || chartType === "BarGroup" || chartType ===  "LineGroup" || chartType ===  "PieGroup"){
                        chart.useAnimation(true);
                    }
                })
                addChartEvent(columnChart);
            }
            isFirstChart = false;
        }
    })

    $(document).bind("keydown", function (event) {
        if (event.shiftKey) {
            isShiftKey = true;
        }
    });
    $(document).bind("keyup", function (event) {
        if (!event.shiftKey) {
            isShiftKey = false;

            var sheet = spread.getActiveSheet(),
                position = getCellPositionString(sheet, sheet.getActiveRowIndex() + 1, sheet.getActiveColumnIndex() + 1);
            $("#positionbox").val(position);
        }
    });

}

function setConditionalFormatSettingGroupVisible(groupName) {
    var $groupItems = $("#conditionalFormatSettingContainer .settingGroup .groupitem");

    $groupItems.hide();
    $groupItems.filter("[data-group='" + groupName + "']").show();
}

function processConditionalFormatSetting(groupName, listRef, rule) {
    $("#conditionalFormatSettingContainer div.details").show();
    setConditionalFormatSettingGroupVisible(groupName);

    var $ruleType = $("#highlightCellsRule"),
        $setButton = $("#setConditionalFormat");
    if (listRef) {
        $ruleType.data("list-ref", listRef);
        $setButton.data("rule-type", rule);
        var item = setDropDownValueByIndex($ruleType, 0);
        updateEnumTypeOfCF(item.value);
    } else {
        $setButton.data("rule-type", groupName);
    }
}

function processConditionalFormatDetailSetting(name, noAction) {
    switch (name) {
        case "highlight-cells-rules":
            $("#formatSetting").show();
            processConditionalFormatSetting("normal", "highlightCellsRulesList", 0);
            break;

        case "top-bottom-rules":
            $("#formatSetting").show();
            processConditionalFormatSetting("normal", "topBottomRulesList", 4);
            break;

        case "color-scales":
            $("#formatSetting").hide();
            processConditionalFormatSetting("normal", "colorScaleList", 8);
            break;

        case "data-bars":
            processConditionalFormatSetting("databar");
            break;

        case "icon-sets":
            processConditionalFormatSetting("iconset");
            updateIconCriteriaItems(0);
            break;

        case "remove-conditional-formats":
            $("#conditionalFormatSettingContainer div.details").hide();
            if (!noAction) {
                removeConditionFormats();
            }
            break;

        default:
            //console.log("processConditionalFormatSetting not add for ", name);
            break;
    }
}

function getBackgroundColor(name) {
    return $("div.insp-color-picker[data-name='" + name + "'] div.color-view").css("background-color");
}

function addCondionalFormaterRule(rule) {
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();
    var style = new spreadNS.Style();

    if (getCheckValue("useFormatBackColor")) {
        style.backColor = getBackgroundColor("formatBackColor");
    }
    if (getCheckValue("useFormatForeColor")) {
        style.foreColor = getBackgroundColor("formatForeColor");
    }
    if (getCheckValue("useFormatBorder")) {
        var lineBorder = new spreadNS.LineBorder(getBackgroundColor("formatBorderColor"), spreadNS.LineStyle.thin);
        style.borderTop = style.borderRight = style.borderBottom = style.borderLeft = lineBorder;
    }
    var value1 = $("#value1").val();
    var value2 = $("#value2").val();
    var cfs = sheet.conditionalFormats;
    var operator = +getDropDownValue("comparisonOperator");

    var minType = +getDropDownValue("minType");
    var midType = +getDropDownValue("midType");
    var maxType = +getDropDownValue("maxType");
    var midColor = getBackgroundColor("midColor");
    var minColor = getBackgroundColor("minColor");
    var maxColor = getBackgroundColor("maxColor");
    var midValue = getNumberValue("midValue");
    var maxValue = getNumberValue("maxValue");
    var minValue = getNumberValue("minValue");

    switch (rule) {
        case "0":
            var doubleValue1 = parseFloat(value1);
            var doubleValue2 = parseFloat(value2);
            cfs.addCellValueRule(operator, isNaN(doubleValue1) ? value1 : doubleValue1, isNaN(doubleValue2) ? value2 : doubleValue2, style, sels);
            break;
        case "1":
            cfs.addSpecificTextRule(operator, value1, style, sels);
            break;
        case "2":
            cfs.addDateOccurringRule(operator, style, sels);
            break;
        case "4":
            cfs.addTop10Rule(operator, parseInt(value1, 10), style, sels);
            break;
        case "5":
            cfs.addUniqueRule(style, sels);
            break;
        case "6":
            cfs.addDuplicateRule(style, sels);
            break;
        case "7":
            cfs.addAverageRule(operator, style, sels);
            break;
        case "8":
            cfs.add2ScaleRule(minType, minValue, minColor, maxType, maxValue, maxColor, sels);
            break;
        case "9":
            cfs.add3ScaleRule(minType, minValue, minColor, midType, midValue, midColor, maxType, maxValue, maxColor, sels);
            break;
        default:
            var doubleValue1 = parseFloat(value1);
            var doubleValue2 = parseFloat(value2);
            cfs.addCellValueRule(operator, isNaN(doubleValue1) ? value1 : doubleValue1, isNaN(doubleValue2) ? value2 : doubleValue2, style, sels);
            break;
    }
    sheet.repaint();
}

function addDataBarRule() {
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();

    var selections = sheet.getSelections();
    if (selections.length > 0) {
        var ranges = [];
        $.each(selections, function (i, v) {
            ranges.push(new spreadNS.Range(v.row, v.col, v.rowCount, v.colCount));
        });
        var cfs = sheet.conditionalFormats;
        var dataBarRule = new ConditionalFormatting.DataBarRule();
        dataBarRule.ranges(ranges);
        dataBarRule.minType(+getDropDownValue("minimumType"));
        dataBarRule.minValue(getNumberValue("minimumValue"));
        dataBarRule.maxType(+getDropDownValue("maximumType"));
        dataBarRule.maxValue(getNumberValue("maximumValue"));
        dataBarRule.gradient(getCheckValue("gradient"));
        dataBarRule.color(getBackgroundColor("gradientColor"));
        dataBarRule.showBorder(getCheckValue("showBorder"));
        dataBarRule.borderColor(getBackgroundColor("barBorderColor"));
        dataBarRule.dataBarDirection(+getDropDownValue("dataBarDirection"));
        dataBarRule.negativeFillColor(getBackgroundColor("negativeFillColor"));
        dataBarRule.useNegativeFillColor(getCheckValue("useNegativeFillColor"));
        dataBarRule.negativeBorderColor(getBackgroundColor("negativeBorderColor"));
        dataBarRule.useNegativeBorderColor(getCheckValue("useNegativeBorderColor"));
        dataBarRule.axisPosition(+getDropDownValue("axisPosition"));
        dataBarRule.axisColor(getBackgroundColor("barAxisColor"));
        dataBarRule.showBarOnly(getCheckValue("showBarOnly"));
        cfs.addRule(dataBarRule);
    }

    sheet.resumePaint();
}

function addIconSetRule() {
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();

    var selections = sheet.getSelections();
    if (selections.length > 0) {
        var ranges = [];
        $.each(selections, function (i, v) {
            ranges.push(new spreadNS.Range(v.row, v.col, v.rowCount, v.colCount));
        });
        var cfs = sheet.conditionalFormats;
        var iconSetRule = new ConditionalFormatting.IconSetRule();
        iconSetRule.ranges(ranges);
        iconSetRule.iconSetType(+getDropDownValue("iconSetType"));
        var $divs = $("#iconCriteriaSetting .settinggroup:visible");
        var iconCriteria = iconSetRule.iconCriteria();

        var icons = iconSetRule.icons();
        $.each($divs, function (i, $div) {
            var v = $(".selections", $div)[0];
            var iconInfo = v.getAttribute("name").split("-");
            icons[i] = {
                iconSetType: parseInt(iconInfo[0]),
                iconIndex: parseInt(iconInfo[1])
            };
            if(i < $divs.length) {
                var suffix = i + 1,
                    isGreaterThanOrEqualTo = +getDropDownValue("iconSetCriteriaOperator" + suffix, this) === 1,
                    iconValueType = +getDropDownValue("iconSetCriteriaType" + suffix, this),
                    iconValue = $("input.editor", this).val();
                if (iconValueType !== ConditionalFormatting.IconValueType.formula) {
                    iconValue = +iconValue;
                }
                iconCriteria[i] = new ConditionalFormatting.IconCriterion(isGreaterThanOrEqualTo, iconValueType, iconValue);
            }
        });
        iconSetRule.reverseIconOrder(getCheckValue("reverseIconOrder"));
        iconSetRule.showIconOnly(getCheckValue("showIconOnly"));
        cfs.addRule(iconSetRule);
    }

    sheet.resumePaint();
}

function removeConditionFormats() {
    var sheet = spread.getActiveSheet();
    var cfs = sheet.conditionalFormats;
    var row = sheet.getActiveRowIndex(), col = sheet.getActiveColumnIndex();
    var rules = cfs.getRules(row, col);
    sheet.suspendPaint();
    $.each(rules, function (i, v) {
        cfs.removeRule(v);
    });
    sheet.resumePaint();
}

// Cell Type related items
function attachCellTypeEvents() {
    $("#setCellTypeBtn").click(function () {
        var currentCellType = getDropDownValue("cellTypes");
        applyCellType(currentCellType);
    });
}

function processCellTypeSetting(name, noAction) {
    $("#cellTypeSettingContainer").show();
    switch (name) {
        case "button-celltype":
            $("#celltype-button").show();
            $("#celltype-checkbox").hide();
            $("#celltype-combobox").hide();
            $("#celltype-hyperlink").hide();
            break;

        case "checkbox-celltype":
            $("#celltype-button").hide();
            $("#celltype-checkbox").show();
            $("#celltype-combobox").hide();
            $("#celltype-hyperlink").hide();
            break;

        case "combobox-celltype":
            $("#celltype-button").hide();
            $("#celltype-checkbox").hide();
            $("#celltype-combobox").show();
            $("#celltype-hyperlink").hide();
            break;

        case "hyperlink-celltype":
            $("#celltype-button").hide();
            $("#celltype-checkbox").hide();
            $("#celltype-combobox").hide();
            $("#celltype-hyperlink").show();
            break;

        case "clear-celltype":
            if (!noAction) {
                clearCellType();
            }
            $("#cellTypeSettingContainer").hide();
            return;

        default:
            //console.log("processCellTypeSetting not process with ", name);
            return;
    }
}

function applyCellType(name) {
    var sheet = spread.getActiveSheet();
    var cellType;
    switch (name) {
        case "button-celltype":
            cellType = new spreadNS.CellTypes.Button();
            cellType.marginTop(getNumberValue("buttonCellTypeMarginTop"));
            cellType.marginRight(getNumberValue("buttonCellTypeMarginRight"));
            cellType.marginBottom(getNumberValue("buttonCellTypeMarginBottom"));
            cellType.marginLeft(getNumberValue("buttonCellTypeMarginLeft"));
            cellType.text(getTextValue("buttonCellTypeText"));
            cellType.buttonBackColor(getBackgroundColor("buttonCellTypeBackColor"));
            break;

        case "checkbox-celltype":
            cellType = new spreadNS.CellTypes.CheckBox();
            cellType.caption(getTextValue("checkboxCellTypeCaption"));
            cellType.textTrue(getTextValue("checkboxCellTypeTextTrue"));
            cellType.textIndeterminate(getTextValue("checkboxCellTypeTextIndeterminate"));
            cellType.textFalse(getTextValue("checkboxCellTypeTextFalse"));
            cellType.textAlign(getDropDownValue("checkboxCellTypeTextAlign"));
            cellType.isThreeState(getCheckValue("checkboxCellTypeIsThreeState"));
            break;

        case "combobox-celltype":
            cellType = new spreadNS.CellTypes.ComboBox();
            cellType.editorValueType(getDropDownValue("comboboxCellTypeEditorValueType"));
            var comboboxItemsText = getTextValue("comboboxCellTypeItemsText");
            var comboboxItemsValue = getTextValue("comboboxCellTypeItemsValue");
            var itemsText = comboboxItemsText.split(",");
            var itemsValue = comboboxItemsValue.split(",");
            var itemsLength = itemsText.length > itemsValue.length ? itemsText.length : itemsValue.length;
            var items = [];
            for (var count = 0; count < itemsLength; count++) {
                var t = itemsText.length > count && itemsText[0] !== "" ? itemsText[count] : undefined;
                var v = itemsValue.length > count && itemsValue[0] !== "" ? itemsValue[count] : undefined;
                if (t !== undefined && v !== undefined) {
                    items[count] = {text: t, value: v};
                }
                else if (t !== undefined) {
                    items[count] = {text: t};
                } else if (v !== undefined) {
                    items[count] = {value: v};
                }
            }
            cellType.items(items);
            break;

        case "hyperlink-celltype":
            cellType = new spreadNS.CellTypes.HyperLink();
            cellType.linkColor(getBackgroundColor("hyperlinkCellTypeLinkColor"));
            cellType.visitedLinkColor(getBackgroundColor("hyperlinkCellTypeVisitedLinkColor"));
            cellType.text(getTextValue("hyperlinkCellTypeText"));
            cellType.linkToolTip(getTextValue("hyperlinkCellTypeLinkToolTip"));
            break;
    }
    sheet.suspendPaint();
    sheet.suspendEvent();
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var i = 0; i < sels.length; i++) {
        var sel = getActualCellRange(sheet, sels[i], rowCount, columnCount);
        for (var r = 0; r < sel.rowCount; r++) {
            for (var c = 0; c < sel.colCount; c++) {
                sheet.setCellType(sel.row + r, sel.col + c, cellType, spreadNS.SheetArea.viewport);
            }
        }
    }
    sheet.resumeEvent();
    sheet.resumePaint();
}

function clearCellType() {
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();
    sheet.suspendPaint();
    for (var i = 0; i < sels.length; i++) {
        var sel = getActualCellRange(sheet, sels[i], rowCount, columnCount);
        sheet.clear(sel.row, sel.col, sel.rowCount, sel.colCount, spreadNS.SheetArea.viewport, spreadNS.StorageType.style);
    }
    sheet.resumePaint();
}

function processComparisonOperator(value) {
    if ($("#ComparisonOperator").data("list-ref") === "cellValueOperatorList") {
        // between (6) and not between ( 7) with two values
        if (value === 6 || value === 7) {
            $("#andtext").show();
            $("#value2").show();
        }
    }
}

function updateEnumTypeOfCF(itemType) {
    var $operator = $("#ComparisonOperator"),
        $setButton = $("#setConditionalFormat");

    $setButton.data("rule-type", itemType);

    switch ("" + itemType) {
        case "0":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "cellValueOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "1":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "specificTextOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "2":
            $("#ruletext").text(conditionalFormatTexts.cells);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "dateOccurringOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "4":
            $("#ruletext").text(conditionalFormatTexts.rankIn);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").show();
            $("#value1").val("10");
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "top10OperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "5":
        case "6":
            $("#ruletext").text(conditionalFormatTexts.all);
            $("#andtext").hide();
            $("#formattext").show();
            $("#formattext").text(conditionalFormatTexts.inRange);
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.hide();
            break;
        case "7":
            $("#ruletext").text(conditionalFormatTexts.values);
            $("#andtext").hide();
            $("#formattext").show();
            $("#formattext").text(conditionalFormatTexts.average);
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").hide();
            $operator.show();
            $operator.data("list-ref", "averageOperatorList");
            setDropDownValueByIndex($operator, 0);
            break;
        case "8":
            $("#ruletext").text(conditionalFormatTexts.allValuesBased);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").show();
            $("#midpoint").hide();
            $("#minType").val("1");
            $("#maxType").val("2");
            $("#minValue").val("");
            $("#maxValue").val("");
            $("#minColor").css("background", "#F8696B");
            $("#maxColor").css("background", "#63BE7B");
            $operator.hide();
            break;
        case "9":
            $("#ruletext").text(conditionalFormatTexts.allValuesBased);
            $("#andtext").hide();
            $("#formattext").hide();
            $("#value1").hide();
            $("#value2").hide();
            $("#colorScale").show();
            $("#midpoint").show();
            $("#minType").val("1");
            $("#midType").val("4");
            $("#maxType").val("2");
            $("#minValue").val("");
            $("#midValue").val("50");
            $("#maxValue").val("");
            $("#minColor").css("background-color", "#F8696B");
            $("#midColor").css("background-color", "#FFEB84");
            $("#maxColor").css("background-color", "#63BE7B");
            $operator.hide();
            break;
        default:
            break;
    }
}

function attachConditionalFormatEvents() {
    $("#setConditionalFormat").click(function () {
        var ruleType = $(this).data("rule-type");

        switch (ruleType) {
            case "databar":
                addDataBarRule();
                break;

            case "iconset":
                addIconSetRule();
                break;

            default:
                addCondionalFormaterRule("" + ruleType);
                break;
        }
    });
}

// Data Validation related items
function processDataValidationSetting(name, title) {
    $("#dataValidationErrorAlertMessage").val("");
    $("#dataValidationErrorAlertTitle").val("");
    $("#dataValidationInputTitle").val("");
    $("#dataValidationInputMessage").val("");
    switch (name) {
        case "anyvalue-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();
            break;

        case "number-validator":
            $("#validatorNumberType").show();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();
            processNumberValidatorComparisonOperatorSetting(getDropDownValue("numberValidatorComparisonOperator"));

            setTextValue("numberMinimum", 0);
            setTextValue("numberMaximum", 0);
            setTextValue("numberValue", 0);
            break;

        case "list-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").show();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("listSource", "1,2,3");
            break;

        case "formulalist-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").show();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("formulaListFormula", "E5:I5");
            break;

        case "date-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").show();
            $("#validatorTextLengthType").hide();
            processDateValidatorComparisonOperatorSetting(getDropDownValue("dateValidatorComparisonOperator"));

            var date = getCurrentTime();
            setTextValue("startDate", date);
            setTextValue("endDate", date);
            setTextValue("dateValue", date);
            break;

        case "textlength-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").hide();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").show();
            processTextLengthValidatorComparisonOperatorSetting(getDropDownValue("textLengthValidatorComparisonOperator"));

            setNumberValue("textLengthMinimum", 0);
            setNumberValue("textLengthMaximum", 0);
            setNumberValue("textLengthValue", 0);
            break;

        case "formula-validator":
            $("#validatorNumberType").hide();
            $("#validatorListType").hide();
            $("#validatorFormulaListType").show();
            $("#validatorDateType").hide();
            $("#validatorTextLengthType").hide();

            setTextValue("formulaListFormula", "=ISERROR(FIND(\" \",A1))");
            break;

        default:
            //console.log("processDataValidationSetting not process with ", name, title);
            break;
    }
}

function processNumberValidatorComparisonOperatorSetting(value) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#numberValue").hide();
        $("#numberBetweenOperator").show();
    }
    else {
        $("#numberBetweenOperator").hide();
        $("#numberValue").show();
    }
}

function processDateValidatorComparisonOperatorSetting(value) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#dateValue").hide();
        $("#dateBetweenOperator").show();
    }
    else {
        $("#dateBetweenOperator").hide();
        $("#dateValue").show();
    }
}

function processTextLengthValidatorComparisonOperatorSetting(value) {
    if (value === ComparisonOperators.between || value === ComparisonOperators.notBetween) {
        $("#textLengthValue").hide();
        $("#textLengthBetweenOperator").show();
    }
    else {
        $("#textLengthBetweenOperator").hide();
        $("#textLengthValue").show();
    }
}

function processCustomHighlightStyleTypeSetting(value) {
    var type = DataValidation.HighlightType;
    switch (value) {
        case type.circle:
            $("#dogearPosition").hide();
            $("#iconPosition").hide();
            $("#iconUpload").hide();
            $('#iconUploadPreview').hide();
            break;
        case type.dogEar:
            $("#dogearPosition").show();
            $("#iconPosition").hide();
            $("#iconUpload").hide();
            $('#iconUploadPreview').hide();
            break;
        case type.icon:
            $("#dogearPosition").hide();
            $("#iconPosition").show();
            $("#iconUpload").show();
            $('#iconUploadPreview').show();
            break;
    }
}

function setDataValidator() {
    var validatorType = getDropDownValue("validatorType");
    var currentDataValidator = null;
    var dropDownValue; // for kinds of validatorType comparison operator
    var highlightStyleType = getDropDownValue("customHighlightStyleType");
    var highlightStyleColor = getBackgroundColor("customHighlightStyleColor");
    var highlightStyle = {
        type: highlightStyleType,
        color: highlightStyleColor
    };

    var formulaListFormula = getTextValue("formulaListFormula");

    switch (validatorType) {
        case "anyvalue-validator":
            currentDataValidator = new spreadNS.DataValidation.DefaultDataValidator();
            break;
        case "number-validator":
            var numberMinimum = getTextValue("numberMinimum");
            var numberMaximum = getTextValue("numberMaximum");
            var numberValue = getTextValue("numberValue");
            var isInteger = getCheckValue("isInteger");
            dropDownValue = getDropDownValue("numberValidatorComparisonOperator");
            if (dropDownValue !== ComparisonOperators.between && dropDownValue !== ComparisonOperators.notBetween) {
                numberMinimum = numberValue;
            }
            if (isInteger) {
                currentDataValidator = DataValidation.createNumberValidator(dropDownValue,
                    isNaN(numberMinimum) ? numberMinimum : parseInt(numberMinimum, 10),
                    isNaN(numberMaximum) ? numberMaximum : parseInt(numberMaximum, 10),
                    true);
            } else {
                currentDataValidator = DataValidation.createNumberValidator(dropDownValue,
                    isNaN(numberMinimum) ? numberMinimum : parseFloat(numberMinimum, 10),
                    isNaN(numberMaximum) ? numberMaximum : parseFloat(numberMaximum, 10),
                    false);
            }
            break;
        case "list-validator":
            var listSource = getTextValue("listSource");
            currentDataValidator = DataValidation.createListValidator(listSource);
            break;
        case "formulalist-validator":
            currentDataValidator = DataValidation.createFormulaListValidator(formulaListFormula);
            break;
        case "date-validator":
            var startDate = getTextValue("startDate");
            var endDate = getTextValue("endDate");
            var dateValue = getTextValue("dateValue");
            var isTime = getCheckValue("isTime");
            dropDownValue = getDropDownValue("dateValidatorComparisonOperator");
            if (dropDownValue !== ComparisonOperators.between && dropDownValue !== ComparisonOperators.notBetween) {
                startDate = dateValue;
            }
            if (isTime) {
                currentDataValidator = DataValidation.createDateValidator(dropDownValue,
                    isNaN(startDate) ? startDate : new Date(startDate),
                    isNaN(endDate) ? endDate : new Date(endDate),
                    true);
            } else {
                currentDataValidator = DataValidation.createDateValidator(dropDownValue,
                    isNaN(startDate) ? startDate : new Date(startDate),
                    isNaN(endDate) ? endDate : new Date(endDate),
                    false);
            }
            break;
        case "textlength-validator":
            var textLengthMinimum = getNumberValue("textLengthMinimum");
            var textLengthMaximum = getNumberValue("textLengthMaximum");
            var textLengthValue = getNumberValue("textLengthValue");
            dropDownValue = getDropDownValue("textLengthValidatorComparisonOperator");
            if (dropDownValue !== ComparisonOperators.between && dropDownValue !== ComparisonOperators.notBetween) {
                textLengthMinimum = textLengthValue;
            }
            currentDataValidator = DataValidation.createTextLengthValidator(dropDownValue, textLengthMinimum, textLengthMaximum);
            break;
        case "formula-validator":
            currentDataValidator = DataValidation.createFormulaValidator(formulaListFormula);
            break;
    }

    if (currentDataValidator) {
        currentDataValidator.errorMessage($("#dataValidationErrorAlertMessage").val());
        currentDataValidator.errorStyle(getDropDownValue("errorAlert"));
        currentDataValidator.errorTitle($("#dataValidationErrorAlertTitle").val());
        currentDataValidator.showErrorMessage(getCheckValue("showErrorAlert"));
        currentDataValidator.ignoreBlank(getCheckValue("ignoreBlank"));
        var showInputMessage = getCheckValue("showInputMessage");
        if (showInputMessage) {
            currentDataValidator.inputTitle($("#dataValidationInputTitle").val());
            currentDataValidator.inputMessage($("#dataValidationInputMessage").val());
        }
        if (highlightStyleType === DataValidation.HighlightType.dogEar) {
            highlightStyle.position = getDropDownValue("dogearPosition");
        } else if (highlightStyleType === DataValidation.HighlightType.icon) {
            highlightStyle.position = getDropDownValue("iconPosition");
            highlightStyle.image = $('#iconUpload').data('icon');
        }
        currentDataValidator.highlightStyle(highlightStyle);

        setDataValidatorInRange(currentDataValidator);
    }
}

function setDataValidatorInRange(dataValidator) {
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();
    var sels = sheet.getSelections();
    var rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();

    for (var i = 0; i < sels.length; i++) {
        var sel = getActualCellRange(sheet, sels[i], rowCount, columnCount);
        sheet.setDataValidator(sel.row, sel.col, sel.rowCount, sel.colCount, dataValidator);
    }
    sheet.resumePaint();
}

function getCurrentTime() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();

    var strDate = year + "-";
    if (month < 10)
        strDate += "0";
    strDate += month + "-";
    if (day < 10)
        strDate += "0";
    strDate += day;

    return strDate;
}

function attachDataValidationEvents() {
    $("#setDataValidator").click(function () {
        setDataValidator();
    });
    $("#clearDataValidatorSettings").click(function () {
        // reset to default
        var validationTypeItem = setDropDownValueByIndex($("#validatorType"), 0);
        processDataValidationSetting(validationTypeItem.value);
        setDropDownValue("errorAlert", 0);
        setCheckValue("showInputMessage", true);
        setCheckValue("showErrorAlert", true);
    });
    $("#iconUpload").click(function () {
        $("#fileSelector").data("action", this.id);
        $("#fileSelector").click();
    });
}
// Data Validation related items (end)

function applyIconSetInfos(iconStyleType, IconSetType) {
    var result = {}, values, iconImages, iconInfos;
    if (iconStyleType <= IconSetType.threeSymbolsUncircled) {
        values = [33, 67];
        if(iconStyleType === IconSetType.threeArrowsColored) {
            iconImages = ["down-arrow-red", "right-arrow-yellow", "up-arrow-green"];
            iconInfos = ["0-0", "0-1", "0-2"];
        } else if (iconStyleType === IconSetType.threeArrowsGray) {
            iconImages = ["down-arrow-gray", "right-arrow-gray", "up-arrow-gray"];
            iconInfos = ["1-0", "1-1", "1-2"];
        } else if (iconStyleType === IconSetType.threeTriangles) {
            iconImages = ["up-triangle-red", "minus-yellow", "up-triangle-green"];
            iconInfos = ["2-0", "2-1", "2-2"];
        } else if (iconStyleType === IconSetType.threeStars) {
            iconImages = ["star-hollow", "star-half", "star-solid"];
            iconInfos = ["3-0", "3-1", "3-2"];
        } else if (iconStyleType === IconSetType.threeFlags) {
            iconImages = ["flag-red", "flag-yellow", "flag-green"];
            iconInfos = ["4-0", "4-1", "4-2"];
        } else if (iconStyleType === IconSetType.threeTrafficLightsUnrimmed) {
            iconImages = ["traffic-light-red", "traffic-light-yellow", "traffic-light-green"];
            iconInfos = ["5-0", "5-1", "5-2"];
        } else if (iconStyleType === IconSetType.threeTrafficLightsRimmed) {
            iconImages = ["traffic-light-rimmed-red", "traffic-light-rimmed-yellow", "traffic-light-rimmed-green"];
            iconInfos = ["6-0", "6-1", "6-2"];
        } else if (iconStyleType === IconSetType.threeSigns) {
            iconImages = ["down-rhombus-red", "up-triangle-yellow", "traffic-light-green"];
            iconInfos = ["7-0", "7-1", "5-2"];
        } else if (iconStyleType === IconSetType.threeSymbolsCircled) {
            iconImages = ["close-circled-red", "notice-circled-yellow", "check-circled-green"];
            iconInfos = ["8-0", "8-1", "8-2"];
        } else {
            iconImages = ["close-uncircled-red", "notice-uncircled-yellow", "check-uncircled-green"];
            iconInfos = ["9-0", "9-1", "9-2"];
        }
    } else if (iconStyleType <= IconSetType.fourTrafficLights) {
        values = [25, 50, 75];
        if(iconStyleType === IconSetType.fourArrowsColored) {
            iconImages = ["down-arrow-red", "right-down-arrow-yellow", "right-up-arrow-yellow", "up-arrow-green"];
            iconInfos = ["0-0", "10-1", "10-2", "0-2"];
        } else if (iconStyleType === IconSetType.fourArrowsGray) {
            iconImages = ["down-arrow-gray", "right-down-arrow-gray", "right-up-arrow-gray", "up-arrow-gray"];
            iconInfos = ["1-0", "11-1", "11-2", "1-2"];
        } else if (iconStyleType === IconSetType.fourRedToBlack) {
            iconImages = ["ball-black", "ball-gray", "ball-pink", "ball-red"];
            iconInfos = ["12-0", "12-1", "12-2", "12-3"];
        } else if (iconStyleType === IconSetType.fourRatings) {
            iconImages = ["rating-1", "rating-2", "rating-3", "rating-4"];
            iconInfos = ["17-1", "17-2", "17-3", "17-4"];
        } else {
            iconImages = ["traffic-light-black","traffic-light-red", "traffic-light-yellow", "traffic-light-green"];
            iconInfos = ["14-0", "5-0", "5-1", "5-2"];
        }
    } else {
        values = [20, 40, 60, 80];
        if(iconStyleType === IconSetType.fiveArrowsColored) {
            iconImages = ["down-arrow-red", "right-down-arrow-yellow", "right-arrow-yellow", "right-up-arrow-yellow", "up-arrow-green"];
            iconInfos = ["0-0", "10-1", "0-1", "10-2", "0-2"];
        } else if (iconStyleType === IconSetType.fiveArrowsGray) {
            iconImages = ["down-arrow-gray", "right-down-arrow-gray", "right-arrow-gray", "right-up-arrow-gray", "up-arrow-gray"];
            iconInfos = ["1-0", "11-1", "11-1", "11-2", "1-2"];
        } else if (iconStyleType === IconSetType.fiveRatings) {
            iconImages = ["rating-0", "rating-1", "rating-2", "rating-3", "rating-4"];
            iconInfos = ["17-0", "17-1", "17-2", "17-3", "17-4"];
        } else if (iconStyleType === IconSetType.fiveQuarters) {
            iconImages = ["quarters-0", "quarters-1", "quarters-2", "quarters-3", "quarters-4"];
            iconInfos = ["18-0", "18-1", "18-2", "18-3", "18-4"];
        } else {
            iconImages = ["box-0", "box-1", "box-2", "box-3", "box-4"];
            iconInfos = ["19-0", "19-1", "19-2", "19-3", "19-4"];
        }
    }
    result.values = values;
    result.iconImages = iconImages;
    result.iconInfos = iconInfos;
    return result;
}

function updateIconCriteriaItems(iconStyleType) {
    var IconSetType = ConditionalFormatting.IconSetType,
        items = $("#iconCriteriaSetting .settinggroup");
    var result = applyIconSetInfos(iconStyleType, IconSetType);
    var values = result.values;
    var iconImages = result.iconImages;
    var iconInfos = result.iconInfos;

    items.each(function (index) {
        var value = values[index], $item = $(this), suffix = index + 1;
        var image = iconImages[index], info = iconInfos[index];
        var commonCss = "ui-icon iconSetsIcons";

        if (value) {
            $item.show();
            var $span = $(".iconSetsIcons", $item);
            $span.removeClass();
            $span.addClass(commonCss);
            $span.addClass(image);
            $(".selections", $item).attr('name', info);
            setDropDownValue("iconSetCriteriaOperator" + suffix, 1, this);
            setDropDownValue("iconSetCriteriaType" + suffix, 4, this);
            $("input.editor", this).val(value);
        } else {
            $item.hide();
        }
    });
    var item = items[items.length - 1];
    $(item).show();
    var $span = $(".iconSetsIcons", $(item));
    $span.removeClass();
    $span.addClass("ui-icon iconSetsIcons");
    $span.addClass(iconImages[iconImages.length - 1]);
    $(".selections", $(item)).attr('name', iconInfos[iconInfos.length - 1]);

    // var iconPicker = $(".icons-popup-dialog");
    var activeSelection;
    $(".selections").click(function(e) {
        activeSelection = e.currentTarget;
        // iconPicker.toggle();
    });


    $(".icons-popup-dialog .iconSetsIcons").click(function(e) {
        var needRemoveClassNamesForDestSpan = "ui-icon iconSetsIcons ";
        var classNames = e.currentTarget.className;
        var imageClassName = classNames.substring(needRemoveClassNamesForDestSpan.length, classNames.length);
        var name = e.currentTarget.getAttribute('name').split(',');
        $(activeSelection).attr('name', iconNameToIconSetType(name[0]) + '-' + name[1]);
        $($('span', activeSelection)[0]).removeClass();
        $($('span', activeSelection)[0]).addClass(needRemoveClassNamesForDestSpan);
        $($('span', activeSelection)[0]).addClass(imageClassName);
        // iconPicker.hide();
        if (_dropdownitem) {
            $(_dropdownitem).removeClass("show");
            _dropdownitem = null;
        }
        processEventListenerHandleClosePopup(false);
    });
}

function iconNameToIconSetType(iconName) {
    var iconSetType;
    switch (iconName) {
        case "3-arrows-icon-set":
            iconSetType = 0 /* ThreeArrowsColored */ ;
            break;
        case "3-arrows-gray-icon-set":
            iconSetType = 1 /* ThreeArrowsGray */ ;
            break;
        case "3-triangles-icon-set":
            iconSetType = 2 /* ThreeTriangles */ ;
            break;
        case "3-traffic-lights-unrimmed-icon-set":
            iconSetType = 5 /* ThreeTrafficLightsUnrimmed */ ;
            break;
        case "3-traffic-lights-rimmed-icon-set":
            iconSetType = 6 /* ThreeTrafficLightsRimmed */ ;
            break;
        case "3-signs-icon-set":
            iconSetType = 7 /* ThreeSigns */ ;
            break;
        case "3-symbols-circled-icon-set":
            iconSetType = 8 /* ThreeSymbolsCircled */ ;
            break;
        case "3-symbols-uncircled-icon-set":
            iconSetType = 9 /* ThreeSymbolsUncircled */ ;
            break;
        case "3-flags-icon-set":
            iconSetType = 4 /* ThreeFlags */ ;
            break;
        case "3-stars-icon-set":
            iconSetType = 3 /* ThreeStars */ ;
            break;
        case "4-arrows-gray-icon-set":
            iconSetType = 11 /* FourArrowsGray */ ;
            break;
        case "4-arrows-icon-set":
            iconSetType = 10 /* FourArrowsColored */ ;
            break;
        case "4-traffic-lights-icon-set":
            iconSetType = 14 /* FourTrafficLights */ ;
            break;
        case "red-to-black-icon-set":
            iconSetType = 12 /* FourRedToBlack */ ;
            break;
        case "4-ratings-icon-set":
            iconSetType = 13 /* FourRatings */ ;
            break;
        case "5-arrows-gray-icon-set":
            iconSetType = 16 /* FiveArrowsGray */ ;
            break;
        case "5-arrows-icon-set":
            iconSetType = 15 /* FiveArrowsColored */ ;
            break;
        case "5-quarters-icon-set":
            iconSetType = 18 /* FiveQuarters */ ;
            break;
        case "5-ratings-icon-set":
            iconSetType = 17 /* FiveRatings */ ;
            break;
        case "5-boxes-icon-set":
            iconSetType = 19 /* FiveBoxes */ ;
            break;
        case "noIcons":
            iconSetType = 20 /* No Cell Icon */ ;
            break;
    }
    return iconSetType;
}

function processMinItems(type, name) {
    var value = "";
    switch (type) {
        case 0: // Number
        case 3: // Percent
            value = "0";
            break;
        case 4: // Percentile
            value = "10";
            break;
        default:
            value = "";
            break;
    }
    setTextValue(name, value);
}

function processMidItems(type, name) {
    var value = "";
    switch (type) {
        case 0: // Number
            value = "0";
            break;
        case 3: // Percent
        case 4: // Percentile
            value = "50";
            break;
        default:
            value = "";
            break;
    }
    setTextValue(name, value);
}

function processMaxItems(type, name) {
    var value = "";
    switch (type) {
        case 0: // Number
            value = "0";
            break;
        case 3: // Percent
            value = "100";
            break;
        case 4: // Percentile
            value = "90";
            break;
        default:
            value = "";
            break;
    }
    setTextValue(name, value);
}

// Sparkline related items
function processAddSparklineEx() {
    var sheet = spread.getActiveSheet();
    var selection = sheet.getSelections()[0];
    if (!selection) {
        return;
    }

    var id = this.id,
        sparklineType = id.toUpperCase(),
        isBarCode = $(this).hasClass('btn-barcode'),
        $typeInfo;
    if(isBarCode){
        $(".menu-item.common-item").hide();
        $(".menu-item.barcode-item").show();
        sparklineType = 'QRCODE';
        $typeInfo = $(".menu-item.barcode-item>div.text[data-value='" + sparklineType + "']");

    }else{
        $(".menu-item.barcode-item").hide();
        $(".menu-item.common-item").show();
        $typeInfo = $(".menu-item.common-item>div.text[data-value='" + sparklineType + "']");
    }

    if ($typeInfo.length > 0) {
        setDropDownValue("sparklineExType", sparklineType);
        processSparklineSetting(sparklineType);
    }
    else {
        processSparklineSetting(getDropDownValue("sparklineExType"));
    }
    setTextValue("txtLineDataRange", parseRangeToExpString(selection));
    setTextValue("txtLineLocationRange", "");

    var SPARKLINE_DIALOG_WIDTH = 360;               // sprakline dialog width
    showModal(uiResource.sparklineDialog.title, SPARKLINE_DIALOG_WIDTH, $("#sparklineexdialog").children(), addSparklineEvent);
}

function setActiveShape(shape) {
    var sheet = spread.getActiveSheet();
    var shapesArray= sheet.shapes.all();
    shapesArray.forEach(function(shapeItem){
        shapeItem.isSelected(false);
    });
    shape.isSelected(true);
}

function getNewShapeName(){
    var sheet = spread.getActiveSheet();
    return 'shape' + sheet.shapes.all().length;
}

function processAddShapeEx(){
    var sheet = spread.getActiveSheet();
    var shapeExType = this.id;
    var shapeType = setShapeType(shapeExType);
    var shapeWidth = 120;
    var shapeHeight = 120;
    var longShapes = ['leftRightArrow', 'leftRightArrowCallout'];
    var heightShapes = ['upDownArrow', 'upDownArrowCallout'];
    if(longShapes.indexOf(shapeExType) >= 0) {
        shapeWidth = 180;
    }
    if(heightShapes.indexOf(shapeExType) >= 0) {
        shapeHeight = 180;
    }
    shape = sheet.shapes.add(getNewShapeName(), shapeType, 400, 100, shapeWidth, shapeHeight);
    setActiveShape(shape);
    addShapeEvent(shape);
}

function processAddConnectorShapeEx(){
    var sheet = spread.getActiveSheet();
    var connectorShapeExType = this.id;
    var connectorShapeType = setConnectorShapeType(connectorShapeExType);

    var connectorShape = sheet.shapes.addConnector(getNewShapeName(), connectorShapeType, 400, 400, 520, 520);
    addShapeEvent(connectorShape);
    setActiveShape(connectorShape);

    // setting shape style
    var shapeStyle = connectorShape.style();
    var arrowHeadStyle = GC.Spread.Sheets.Shapes.ArrowheadStyle.triangle;
    if(connectorShapeExType.toLowerCase().indexOf('begin') >= 0) {
        shapeStyle.line.beginArrowheadStyle = arrowHeadStyle;
    }
    if(connectorShapeExType.toLowerCase().indexOf('end') >= 0) {
        shapeStyle.line.endArrowheadStyle = arrowHeadStyle;
    }
    connectorShape.style(shapeStyle);
}

function processAddChartEx() {
    var sheet = spread.getActiveSheet();
    var selection = sheet.getSelections()[0];
    if(!selection || (selection.rowCount === 1 && selection.colCount === 1)) {
        return;
    }
    var formula = GC.Spread.Sheets.CalcEngine.rangeToFormula(selection);
    var chartExType = this.id;
    var chartType = setChartType(chartExType);
    var chartCount = sheet.charts.all().length || 0;
    var chart = null;
    if(formula){
        if(chartType > 0){
            try{
                chart = sheet.charts.add((chartExType + chartCount), chartType, 0, 100, 400, 300, formula);
                var chartGroup = getChartGroupString(chartType);
                if(chartGroup === "ColumnGroup" || chartGroup === "BarGroup" || chartGroup ===  "LineGroup" || chartGroup ===  "PieGroup"){
                    chart.useAnimation(true);
                }
            }catch (e){
                alert(e.message);
                return;
            }

        }else{
            chart = createComboChart(formula,('Chart' + chartCount),GC.Spread.Sheets.Charts.ChartType.columnClustered,GC.Spread.Sheets.Charts.ChartType.line);
        }
        var chartsArray= sheet.charts.all();
        for(var i = 0; i < chartsArray.length; i++){
            var chartItem = chartsArray[i];
            chartItem.isSelected(false);
        }
        chart.isSelected(true);
        addChartEvent(chart);
    }

}
function unParseFormula(expr, row, col) {
    if (!expr) {
        return "";
    }
    var sheet = spread.getActiveSheet();
    if (!sheet) {
        return null;
    }
    var calcService = sheet.getCalcService();
    return calcService.unparse(null, expr, row, col);
}

function processSparklineSetting(name, title) {
    //Show only when data range is illegal.
    $("#dataRangeError").hide();
    $("#singleDataRangeError").hide();
    //Show only when location range is illegal.
    $("#locationRangeError").hide();

    switch (name) {
        case "LINESPARKLINE":
        case "COLUMNSPARKLINE":
        case "WINLOSSSPARKLINE":
        case "PIESPARKLINE":
        case "AREASPARKLINE":
        case "SCATTERSPARKLINE":
        case "SPREADSPARKLINE":
        case "STACKEDSPARKLINE":
        case "BOXPLOTSPARKLINE":
        case "CASCADESPARKLINE":
        case "PARETOSPARKLINE":
        case 'EAN8':
        case 'GS1_128':
        case 'EAN13':
        case 'CODE93':
        case 'CODE39':
        case 'CODE128':
        case 'CODE49':
        case 'DATAMATRIX':
        case 'PDF417':
        case 'CODABAR':
        case 'QRCODE':
            $("#lineContainer").show();
            $("#bulletContainer").hide();
            $("#monthContainer").hide();
            $("#hbarContainer").hide();
            $("#yearContainer").hide();
            $("#varianceContainer").hide();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();
            break;

        case "BULLETSPARKLINE":
            $("#lineContainer").hide();
            $("#monthContainer").hide();
            $("#bulletContainer").show();
            $("#hbarContainer").hide();
            $("#yearContainer").hide();
            $("#varianceContainer").hide();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();

            setTextValue("txtBulletMeasure", "");
            setTextValue("txtBulletTarget", "");
            setTextValue("txtBulletMaxi", "");
            setTextValue("txtBulletGood", "");
            setTextValue("txtBulletBad", "");
            setTextValue("txtBulletForecast", "");
            setTextValue("txtBulletTickunit", "");
            setCheckValue("checkboxBulletVertial", false);
            break;

        case "HBARSPARKLINE":
        case "VBARSPARKLINE":
            $("#lineContainer").hide();
            $("#bulletContainer").hide();
            $("#monthContainer").hide();
            $("#hbarContainer").show();
            $("#yearContainer").hide();
            $("#varianceContainer").hide();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();

            setTextValue("txtHbarValue", "");
            break;

        case "VARISPARKLINE":
            $("#lineContainer").hide();
            $("#bulletContainer").hide();
            $("#hbarContainer").hide();
            $("#monthContainer").hide();
            $("#yearContainer").hide();
            $("#varianceContainer").show();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();

            setTextValue("txtVariance", "");
            setTextValue("txtVarianceReference", "");
            setTextValue("txtVarianceMini", "");
            setTextValue("txtVarianceMaxi", "");
            setTextValue("txtVarianceMark", "");
            setTextValue("txtVarianceTickUnit", "");
            setCheckValue("checkboxVarianceLegend", false);
            setCheckValue("checkboxVarianceVertical", false);
            break;

        case "MONTHSPARKLINE":
            $("#lineContainer").show();
            $("#bulletContainer").hide();
            $("#hbarContainer").hide();
            $("#varianceContainer").hide();
            $("#yearContainer").show();
            $("#monthContainer").show();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();

            setTextValue("txtYearValue", "");
            setTextValue("txtMonthValue", "");
            setTextValue("txtEmptyColorValue", "");
            setTextValue("txtStartColorValue", "");
            setTextValue("txtMiddleColorValue", "");
            setTextValue("txtEndColorValue", "");
            setTextValue("txtColorRangeValue", "");
            break;

        case "YEARSPARKLINE":
            $("#lineContainer").show();
            $("#bulletContainer").hide();
            $("#hbarContainer").hide();
            $("#varianceContainer").hide();
            $("#monthContainer").hide();
            $("#yearContainer").show();
            $("#ean8Container").hide();
            $("#gs1128Container").hide();
            $("#ean13Container").hide();
            $("#qrCodeContainer").hide();
            $("#dataMatrixContainer").hide();
            $("#pdf417Container").hide();
            $("#code93Container").hide();
            $("#code39Container").hide();
            $("#code49Container").hide();
            $("#code128Container").hide();
            $("#codaBarContainer").hide();

            setTextValue("txtYearValue", "");
            setTextValue("txtEmptyColorValue", "");
            setTextValue("txtStartColorValue", "");
            setTextValue("txtMiddleColorValue", "");
            setTextValue("txtEndColorValue", "");
            setTextValue("txtColorRangeValue", "");
            break;

        default:
            //console.log("processSparklineSetting not process with ", name, title);
            break;
    }
}

function addSparklineEvent() {
    var sheet = spread.getActiveSheet(),
        selection = sheet.getSelections()[0],
        isValid = true;

    var sparklineExType = getDropDownValue("sparklineExType");

    if (selection) {
        var range = getActualRange(selection, sheet.getRowCount(), sheet.getColumnCount());
        var formulaStr = '', row = range.row, col = range.col, direction = 0;

        switch (sparklineExType) {
            case "BULLETSPARKLINE":
                var measure = getTextValue("txtBulletMeasure"),
                    target = getTextValue("txtBulletTarget"),
                    maxi = getTextValue("txtBulletMaxi"),
                    good = getTextValue("txtBulletGood"),
                    bad = getTextValue("txtBulletBad"),
                    forecast = getTextValue("txtBulletForecast"),
                    tickunit = getTextValue("txtBulletTickunit"),
                    colorScheme = getBackgroundColor("colorBulletColorScheme"),
                    vertical = getCheckValue("checkboxBulletVertial");
                formulaStr = '=' + sparklineExType + '(' + measure + ',' + target + ',' + maxi + ',' + good + ',' + bad + ',' + forecast + ',' + tickunit + ',' + '"' + colorScheme + '"' + ',' + vertical + ')';
                sheet.setFormula(row, col, formulaStr);
                break;
            case "HBARSPARKLINE":
                var value = getTextValue("txtHbarValue"),
                    colorScheme = getBackgroundColor("colorHbarColorScheme");
                formulaStr = '=' + sparklineExType + '(' + value + ',' + '"' + colorScheme + '"' + ')';
                sheet.setFormula(row, col, formulaStr);
                break;
            case "VBARSPARKLINE":
                var value = getTextValue("txtHbarValue"),
                    colorScheme = getBackgroundColor("colorHbarColorScheme");
                formulaStr = '=' + sparklineExType + '(' + value + ',' + '"' + colorScheme + '"' + ')';
                sheet.setFormula(row, col, formulaStr);
                break;
            case "VARISPARKLINE":
                var variance = getTextValue("txtVariance"),
                    reference = getTextValue("txtVarianceReference"),
                    mini = getTextValue("txtVarianceMini"),
                    maxi = getTextValue("txtVarianceMaxi"),
                    mark = getTextValue("txtVarianceMark"),
                    tickunit = getTextValue("txtVarianceTickUnit"),
                    colorPositive = getBackgroundColor("colorVariancePositive"),
                    colorNegative = getBackgroundColor("colorVarianceNegative"),
                    legend = getCheckValue("checkboxVarianceLegend"),
                    vertical = getCheckValue("checkboxVarianceVertical");
                formulaStr = '=' + sparklineExType + '(' + variance + ',' + reference + ',' + mini + ',' + maxi + ',' + mark + ',' + tickunit + ',' + legend + ',' + '"' + colorPositive + '"' + ',' + '"' + colorNegative + '"' + ',' + vertical + ')';
                sheet.setFormula(row, col, formulaStr);
                break;
            case "CASCADESPARKLINE":
            case "PARETOSPARKLINE":
                var dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    vertical = false,
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }
                if (locationRange && locationRange.rowCount < locationRange.colCount) {
                    vertical = true;
                }
                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError").show();
                }
                if (isValid) {
                    var pointCount = dataRange.rowCount * dataRange.colCount,
                        i = 1;
                    for (var r = locationRange.row; r < locationRange.row + locationRange.rowCount; r++) {
                        for (var c = locationRange.col; c < locationRange.col + locationRange.colCount; c++) {
                            if (i <= pointCount) {
                                formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ',' + i + ',,,,,,' + vertical + ')';
                                sheet.setFormula(r, c, formulaStr);
                                sheet.setActiveCell(r, c);
                                i++;
                            }
                        }
                    }
                }
                break;
            case "MONTHSPARKLINE":
                var year = getTextValue("txtYearValue"),
                    month = getTextValue("txtMonthValue"),
                    emptyColor = getBackgroundColor("emptyColorValue"),
                    startColor = getBackgroundColor("startColorValue"),
                    middleColor = getBackgroundColor("middleColorValue"),
                    endColor = getBackgroundColor("endColorValue"),
                    dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    colorRangeStr = getTextValue("txtColorRangeValue"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }
                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError").show();
                }
                var row = locationRange.row, col = locationRange.col;
                if (isValid) {
                    if (!colorRangeStr) {
                        formulaStr = "=" + sparklineExType + "(" + year + "," + month + "," + dataRangeStr + "," + parseSparklineColorOptions(emptyColor) + "," + parseSparklineColorOptions(startColor) + "," + parseSparklineColorOptions(middleColor) + "," + parseSparklineColorOptions(endColor) + ")";
                    } else {
                        formulaStr = "=" + sparklineExType + "(" + year + "," + month + "," + dataRangeStr + "," + colorRangeStr + ")";
                    }
                    sheet.setFormula(row, col, formulaStr);
                }
                break;
            case "YEARSPARKLINE":
                var year = getTextValue("txtYearValue"),
                    emptyColor = getBackgroundColor("emptyColorValue"),
                    startColor = getBackgroundColor("startColorValue"),
                    middleColor = getBackgroundColor("middleColorValue"),
                    endColor = getBackgroundColor("endColorValue"),
                    dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    colorRangeStr = getTextValue("txtColorRangeValue"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }
                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError").show();
                }
                var row = locationRange.row, col = locationRange.col;
                if (isValid) {
                    if (!colorRangeStr) {
                        formulaStr = "=" + sparklineExType + "(" + year + "," + dataRangeStr + "," + parseSparklineColorOptions(emptyColor) + "," + parseSparklineColorOptions(startColor) + "," + parseSparklineColorOptions(middleColor) + "," + parseSparklineColorOptions(endColor) + ")";
                    } else {
                        formulaStr = "=" + sparklineExType + "(" + year + "," + dataRangeStr + "," + colorRangeStr + ")";
                    }
                    sheet.setFormula(row, col, formulaStr);
                }
                break;

            case 'EAN8':
            case 'GS1_128':
            case 'EAN13':
            case 'CODE93':
            case 'CODE39':
            case 'CODE128':
            case 'CODE49':
            case 'QRCODE':
            case 'PDF417':
            case 'CODABAR':
            case 'DATAMATRIX':
                var dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    type = 'BC_' + sparklineExType,
                    dataRange, locationRange;

                    if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                        dataRange = dataRangeObj[0].range;
                    }
                    if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                        locationRange = locationRangeObj[0].range;
                    }
                    if (!dataRange) {
                        isValid = false;
                        $("#dataRangeError").show();
                    }
                    if (!locationRange) {
                        isValid = false;
                        $("#locationRangeError").show();
                    }
                    if (isValid) {
                        var row = locationRange.row, col = locationRange.col;
                        formulaStr = '=' + type + '(' + dataRangeStr + ')';
                        sheet.setFormula(row, col, formulaStr);
                        sheet.setActiveCell(row, col);
                    }
                    break;

            default:
                var dataRangeStr = getTextValue("txtLineDataRange"),
                    locationRangeStr = getTextValue("txtLineLocationRange"),
                    dataRangeObj = parseStringToExternalRanges(dataRangeStr, sheet),
                    locationRangeObj = parseStringToExternalRanges(locationRangeStr, sheet),
                    dataRange, locationRange;
                if (dataRangeObj && dataRangeObj.length > 0 && dataRangeObj[0].range) {
                    dataRange = dataRangeObj[0].range;
                }
                if (locationRangeObj && locationRangeObj.length > 0 && locationRangeObj[0].range) {
                    locationRange = locationRangeObj[0].range;
                }

                if (!dataRange) {
                    isValid = false;
                    $("#dataRangeError").show();
                }
                if (!locationRange) {
                    isValid = false;
                    $("#locationRangeError").show();
                }
                if (isValid) {
                    if (["LINESPARKLINE", "COLUMNSPARKLINE", "WINLOSSSPARKLINE"].indexOf(sparklineExType) >= 0) {
                        if (dataRange.rowCount === 1) {
                            direction = 1;
                        }
                        else if (dataRange.colCount === 1) {
                            direction = 0;
                        }
                        else {
                            $("#singleDataRangeError").show();
                            isValid = false;
                        }
                        if (isValid) {
                            formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ',' + direction + ')';
                        }
                    }
                    else {
                        formulaStr = '=' + sparklineExType + '(' + dataRangeStr + ')';
                    }
                    if (isValid) {
                        row = locationRange.row;
                        col = locationRange.col;
                        sheet.setFormula(row, col, formulaStr);
                        sheet.setActiveCell(row, col);
                    }
                }
                break;
        }
    }

    if (!isValid) {
        return {canceled: true};
    }
    else {
        if (showSparklineSetting(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex())) {
            updateFormulaBar();
            setActiveTab("sparklineEx");
            return;
        }
        //console.log("Added sparkline", sparklineExType);
    }
}

function addChartEvent(chart) {
    var sheet = spread.getActiveSheet();
    showChartPanel(chart);
}

function addShapeEvent(shape) {
    var sheet = spread.getActiveSheet();
    showShapePanel(shape);
}

function setChartType(chartExType) {
    var chartType;
    switch (chartExType) {
        case "columnClusteredChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.columnClustered;
            break;
        case "columnStackedChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.columnStacked;
            break;
        case "columnStacked100Chart":
            chartType = GC.Spread.Sheets.Charts.ChartType.columnStacked100;
            break;
        case "lineChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.line;
            break;
        case "lineStackedChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.lineStacked;
            break;
        case "lineStacked100Chart":
            chartType = GC.Spread.Sheets.Charts.ChartType.lineStacked100;
            break;
        case "lineMarkersChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.lineMarkers;
            break;
        case "lineMarkersStackedChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.lineMarkersStacked;
            break;
        case "lineMarkersStacked100Chart":
            chartType = GC.Spread.Sheets.Charts.ChartType.lineMarkersStacked100;
            break;
        case "pieChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.pie;
            break;
        case "doughnutChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.doughnut;
            break;
        case "barClusteredChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.barClustered;
            break;
        case "barStackedChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.barStacked;
            break;
        case "barStacked100Chart":
            chartType = GC.Spread.Sheets.Charts.ChartType.barStacked100;
            break;
        case "areaChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.area;
            break;
        case "areaStackedChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.areaStacked;
            break;
        case "areaStacked100Chart":
            chartType = GC.Spread.Sheets.Charts.ChartType.areaStacked100;
            break;
        case "xyScatterChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.xyScatter;
            break;
        case "xyScatterSmoothChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.xyScatterSmooth;
            break;
        case "xyScatterSmoothNoMarkersChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.xyScatterSmoothNoMarkers;
            break;
        case "xyScatterLinesChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.xyScatterLines;
            break;
        case "xyScatterLinesNoMarkersChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.xyScatterLinesNoMarkers;
            break;
        case "bubbleChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.bubble;
            break;
        case "stockHLCChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.stockHLC;
            break;
        case "stockOHLCChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.stockOHLC;
            break;
        case "stockVHLCChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.stockVHLC;
            break;
        case "stockVOHLCChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.stockVOHLC;
            break;
        case "comboChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.combo;
            break;
        case "radarChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.radar;
            break;
        case "radarMarkersChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.radarMarkers;
            break;
        case "radarFilledChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.radarFilled;
            break;
        case "sunburstChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.sunburst;
            break;
        case "treemapChart":
            chartType = GC.Spread.Sheets.Charts.ChartType.treemap;
            break;
    }

    return chartType;
}

function setShapeType(shapeExType) {
    var shapeType = GC.Spread.Sheets.Shapes.AutoShapeType[shapeExType];
    return shapeType;
}

function setConnectorShapeType(connectorShapeExType) {
    var type = 'straight';
    if(connectorShapeExType.toLowerCase().indexOf('elbow')>=0) {
        type = 'elbow';
    }
    return connectorShapeType = GC.Spread.Sheets.Shapes.ConnectorType[type];
}

function parseSparklineColorOptions(str) {
    return '"' + str + '"';
}

function unparseSparklineColorOptions(str){
    return str = str.replace(/\"/g, "");;
}

function unparseBraceOptions(str){
    return str = str.substring(1,str.length-1);
}

function parseRangeToExpString(range) {
    return SheetsCalc.rangeToFormula(range, 0, 0, SheetsCalc.RangeReferenceRelative.allRelative);
}

function parseStringToExternalRanges(expString, sheet) {
    var results = [];
    var exps = expString.split(",");
    try {
        for (var i = 0; i < exps.length; i++) {
            var range = SheetsCalc.formulaToRange(sheet, exps[i]);
            results.push({"range": range});
        }
    }
    catch (e) {
        return null;
    }
    return results;
}

function parseFormulaSparkline(row, col) {
    var sheet = spread.getActiveSheet();
    if (!sheet) {
        return null;
    }
    var formula = sheet.getFormula(row, col);
    if (!formula) {
        return null;
    }
    var calcService = sheet.getCalcService();
    try {
        var expr = calcService.parse(null, formula, row, col);
        if (expr.type === ExpressionType.function) {
            var fnName = expr.functionName;
            if (fnName && spread.getSparklineEx(fnName)) {
                return expr;
            }
        }
    }
    catch (ex) {
        //console.log("parse failed:", ex);
    }
    return null;
}

function parseColorExpression(colorExpression, row, col) {
    if (!colorExpression) {
        return null;
    }
    var sheet = spread.getActiveSheet();
    if (colorExpression.type === ExpressionType.string) {
        return colorExpression.value;
    }
    else if (colorExpression.type === ExpressionType.missingArgument) {
        return null;
    }
    else {
        var formula = null;
        try {
            formula = unParseFormula(colorExpression, row, col);
        }
        catch (ex) {
        }
        return SheetsCalc.evaluateFormula(sheet, formula, row, col);
    }
}

function getAreaSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorPositive: "#787878", colorNegative: "#CB0000"};
    if (formulaArgs[0]) {
        setTextValue("areaSparklinePoints", unParseFormula(formulaArgs[0], row, col));
    }
    else {
        setTextValue("areaSparklinePoints", "");
    }
    var inputList = ["areaSparklineMinimumValue", "areaSparklineMaximumValue", "areaSparklineLine1", "areaSparklineLine2"];
    var len = inputList.length;
    for (var i = 1; i <= len; i++) {
        if (formulaArgs[i]) {
            setNumberValue(inputList[i - 1], unParseFormula(formulaArgs[i], row, col));
        }
        else {
            setNumberValue(inputList[i - 1], "");
        }
    }
    var positiveColor = parseColorExpression(formulaArgs[5], row, col);
    if (positiveColor) {
        setColorValue("areaSparklinePositiveColor", positiveColor);
    }
    else {
        setColorValue("areaSparklinePositiveColor", defaultValue.colorPositive);
    }
    var negativeColor = parseColorExpression(formulaArgs[6], row, col);
    if (negativeColor) {
        setColorValue("areaSparklineNegativeColor", negativeColor);
    }
    else {
        setColorValue("areaSparklineNegativeColor", defaultValue.colorNegative);
    }
}

function getBoxPlotSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {boxplotClass: "5ns", style: 0, colorScheme: "#D2D2D2", vertical: false, showAverage: false};
    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var boxPlotClassValue = formulaArgs[1] && (formulaArgs[1].type === ExpressionType.string ? formulaArgs[1].value : null);
        var showAverageValue = formulaArgs[2] && (formulaArgs[2].type === ExpressionType.boolean ? formulaArgs[2].value : null);
        var scaleStartValue = unParseFormula(formulaArgs[3], row, col);
        var scaleEndValue = unParseFormula(formulaArgs[4], row, col);
        var acceptableStartValue = unParseFormula(formulaArgs[5], row, col);
        var acceptableEndValue = unParseFormula(formulaArgs[6], row, col);
        var colorValue = parseColorExpression(formulaArgs[7], row, col);
        var styleValue = formulaArgs[8] ? unParseFormula(formulaArgs[8], row, col) : null;
        var verticalValue = formulaArgs[9] && (formulaArgs[9].type === ExpressionType.boolean ? formulaArgs[9].value : null);

        setTextValue("boxplotSparklinePoints", pointsValue);
        setDropDownValue("boxplotClassType", boxPlotClassValue === null ? defaultValue.boxplotClass : boxPlotClassValue);
        setTextValue("boxplotSparklineScaleStart", scaleStartValue);
        setTextValue("boxplotSparklineScaleEnd", scaleEndValue);
        setTextValue("boxplotSparklineAcceptableStart", acceptableStartValue);
        setTextValue("boxplotSparklineAcceptableEnd", acceptableEndValue);
        setColorValue("boxplotSparklineColorScheme", colorValue === null ? defaultValue.colorScheme : colorValue);
        setDropDownValue("boxplotSparklineStyleType", styleValue === null ? defaultValue.style : styleValue);
        setCheckValue("boxplotSparklineShowAverage", showAverageValue === null ? defaultValue.showAverage : showAverageValue);
        setCheckValue("boxplotSparklineVertical", verticalValue === null ? defaultValue.vertical : verticalValue);
    }
    else {
        setTextValue("boxplotSparklinePoints", "");
        setDropDownValue("boxplotClassType", defaultValue.boxplotClass);
        setTextValue("boxplotSparklineScaleStart", "");
        setTextValue("boxplotSparklineScaleEnd", "");
        setTextValue("boxplotSparklineAcceptableStart", "");
        setTextValue("boxplotSparklineAcceptableEnd", "");
        setColorValue("boxplotSparklineColorScheme", defaultValue.colorScheme);
        setDropDownValue("boxplotSparklineStyleType", defaultValue.style);
        setCheckValue("boxplotSparklineShowAverage", defaultValue.showAverage);
        setCheckValue("boxplotSparklineVertical", defaultValue.vertical);
    }
}

function getBulletSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {vertical: false, colorScheme: "#A0A0A0"};
    if (formulaArgs && formulaArgs.length > 0) {
        var measureValue = unParseFormula(formulaArgs[0], row, col);
        var targetValue = unParseFormula(formulaArgs[1], row, col);
        var maxiValue = unParseFormula(formulaArgs[2], row, col);
        var goodValue = unParseFormula(formulaArgs[3], row, col);
        var badValue = unParseFormula(formulaArgs[4], row, col);
        var forecastValue = unParseFormula(formulaArgs[5], row, col);
        var tickunitValue = unParseFormula(formulaArgs[6], row, col);
        var colorSchemeValue = parseColorExpression(formulaArgs[7], row, col);
        var verticalValue = formulaArgs[8] && (formulaArgs[8].type === ExpressionType.boolean ? formulaArgs[8].value : null);

        setTextValue("bulletSparklineMeasure", measureValue);
        setTextValue("bulletSparklineTarget", targetValue);
        setTextValue("bulletSparklineMaxi", maxiValue);
        setTextValue("bulletSparklineForecast", forecastValue);
        setTextValue("bulletSparklineGood", goodValue);
        setTextValue("bulletSparklineBad", badValue);
        setTextValue("bulletSparklineTickUnit", tickunitValue);
        setColorValue("bulletSparklineColorScheme", colorSchemeValue ? colorSchemeValue : defaultValue.colorScheme);
        setCheckValue("bulletSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("bulletSparklineMeasure", "");
        setTextValue("bulletSparklineTarget", "");
        setTextValue("bulletSparklineMaxi", "");
        setTextValue("bulletSparklineForecast", "");
        setTextValue("bulletSparklineGood", "");
        setTextValue("bulletSparklineBad", "");
        setTextValue("bulletSparklineTickUnit", "");
        setColorValue("bulletSparklineColorScheme", defaultValue.colorScheme);
        setCheckValue("bulletSparklineVertical", defaultValue.vertical);
    }
}

function getCascadeSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorPositive: "#8CBF64", colorNegative: "#D6604D", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsRangeValue = unParseFormula(formulaArgs[0], row, col);
        var pointIndexValue = unParseFormula(formulaArgs[1], row, col);
        var labelsRangeValue = unParseFormula(formulaArgs[2], row, col);
        var minimumValue = unParseFormula(formulaArgs[3], row, col);
        var maximumValue = unParseFormula(formulaArgs[4], row, col);
        var colorPositiveValue = parseColorExpression(formulaArgs[5], row, col);
        var colorNegativeValue = parseColorExpression(formulaArgs[6], row, col);
        var verticalValue = formulaArgs[7] && (formulaArgs[7].type === ExpressionType.boolean ? formulaArgs[7].value : null);

        setTextValue("cascadeSparklinePointsRange", pointsRangeValue);
        setTextValue("cascadeSparklinePointIndex", pointIndexValue);
        setTextValue("cascadeSparklineLabelsRange", labelsRangeValue);
        setTextValue("cascadeSparklineMinimum", minimumValue);
        setTextValue("cascadeSparklineMaximum", maximumValue);
        setColorValue("cascadeSparklinePositiveColor", colorPositiveValue ? colorPositiveValue : defaultValue.colorPositive);
        setColorValue("cascadeSparklineNegativeColor", colorNegativeValue ? colorNegativeValue : defaultValue.colorNegative);
        setCheckValue("cascadeSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("cascadeSparklinePointsRange", "");
        setTextValue("cascadeSparklinePointIndex", "");
        setTextValue("cascadeSparklineLabelsRange", "");
        setTextValue("cascadeSparklineMinimum", "");
        setTextValue("cascadeSparklineMaximum", "");
        setColorValue("cascadeSparklinePositiveColor", defaultValue.colorPositive);
        setColorValue("cascadeSparklineNegativeColor", defaultValue.colorNegative);
        setCheckValue("cascadeSparklineVertical", defaultValue.vertical);
    }
}

function parseSetting(jsonSetting) {
    var setting = {}, inBracket = false, inProperty = true, property = "", value = "";
    if (jsonSetting) {
        jsonSetting = jsonSetting.substr(1, jsonSetting.length - 2);
        for (var i = 0, len = jsonSetting.length; i < len; i++) {
            var char = jsonSetting.charAt(i);
            if (char === ":") {
                inProperty = false;
            }
            else if (char === "," && !inBracket) {
                setting[property] = value;
                property = "";
                value = "";
                inProperty = true;
            }
            else if (char === "\'" || char === "\"") {
                // discard
            }
            else {
                if (char === "(") {
                    inBracket = true;
                }
                else if (char === ")") {
                    inBracket = false;
                }
                if (inProperty) {
                    property += char;
                }
                else {
                    value += char;
                }
            }
        }
        if (property) {
            setting[property] = value;
        }
        for (var p in setting) {
            var v = setting[p];
            if (v !== null && typeof (v) !== "undefined") {
                if (v.toUpperCase() === "TRUE") {
                    setting[p] = true;
                } else if (v.toUpperCase() === "FALSE") {
                    setting[p] = false;
                } else if (!isNaN(v) && isFinite(v)) {
                    setting[p] = parseFloat(v);
                }
            }
        }
    }
    return setting;
}

function updateManual(type, inputDataName) {
    var $manualDiv = $("div.insp-text[data-name='" + inputDataName + "']");
    var $manualInput = $manualDiv.find("input");
    if (type !== "custom") {
        $manualInput.attr("disabled", "disabled");
        $manualDiv.addClass("manual-disable");
    }
    else {
        $manualInput.removeAttr("disabled");
        $manualDiv.removeClass("manual-disable");
    }
}

function updateStyleSetting(settings) {
    var defaultValue = {
        negativePoints: "#A52A2A", markers: "#244062", highPoint: "#0000FF",
        lowPoint: "#0000FF", firstPoint: "#95B3D7", lastPoint: "#95B3D7",
        series: "#244062", axis: "#000000"
    };
    setColorValue("compatibleSparklineNegativeColor", settings.negativeColor ? settings.negativeColor : defaultValue.negativePoints);
    setColorValue("compatibleSparklineMarkersColor", settings.markersColor ? settings.markersColor : defaultValue.markers);
    setColorValue("compatibleSparklineAxisColor", settings.axisColor ? settings.axisColor : defaultValue.axis);
    setColorValue("compatibleSparklineSeriesColor", settings.seriesColor ? settings.seriesColor : defaultValue.series);
    setColorValue("compatibleSparklineHighMarkerColor", settings.highMarkerColor ? settings.highMarkerColor : defaultValue.highPoint);
    setColorValue("compatibleSparklineLowMarkerColor", settings.lowMarkerColor ? settings.lowMarkerColor : defaultValue.lowPoint);
    setColorValue("compatibleSparklineFirstMarkerColor", settings.firstMarkerColor ? settings.firstMarkerColor : defaultValue.firstPoint);
    setColorValue("compatibleSparklineLastMarkerColor", settings.lastMarkerColor ? settings.lastMarkerColor : defaultValue.lastPoint);
    setTextValue("compatibleSparklineLastLineWeight", settings.lineWeight || settings.lw);
}

function updateSparklineSetting(setting) {
    if (!setting) {
        return;
    }
    var defaultSetting = {
        rightToLeft: false,
        displayHidden: false,
        displayXAxis: false,
        showFirst: false,
        showHigh: false,
        showLast: false,
        showLow: false,
        showNegative: false,
        showMarkers: false
    };

    setDropDownValue("emptyCellDisplayType", setting.displayEmptyCellsAs ? setting.displayEmptyCellsAs : -1);
    setCheckValue("showDataInHiddenRowOrColumn", setting.displayHidden ? setting.displayHidden : defaultSetting.displayHidden);
    setCheckValue("compatibleSparklineShowFirst", setting.showFirst ? setting.showFirst : defaultSetting.showFirst);
    setCheckValue("compatibleSparklineShowLast", setting.showLast ? setting.showLast : defaultSetting.showLast);
    setCheckValue("compatibleSparklineShowHigh", setting.showHigh ? setting.showHigh : defaultSetting.showHigh);
    setCheckValue("compatibleSparklineShowLow", setting.showLow ? setting.showLow : defaultSetting.showLow);
    setCheckValue("compatibleSparklineShowNegative", setting.showNegative ? setting.showNegative : defaultSetting.showNegative);
    setCheckValue("compatibleSparklineShowMarkers", setting.showMarkers ? setting.showMarkers : defaultSetting.showMarkers);
    var minAxisType = Sparklines.SparklineAxisMinMax[setting.minAxisType];
    setDropDownValue("minAxisType", minAxisType ? minAxisType : -1);
    setTextValue("manualMin", setting.manualMin ? setting.manualMin : "");
    var maxAxisType = Sparklines.SparklineAxisMinMax[setting.maxAxisType];
    setDropDownValue("maxAxisType", maxAxisType ? maxAxisType : -1);
    setTextValue("manualMax", setting.manualMax ? setting.manualMax : "");
    setCheckValue("rightToLeft", setting.rightToLeft ? setting.rightToLeft : defaultSetting.rightToLeft);
    setCheckValue("displayXAxis", setting.displayXAxis ? setting.displayXAxis : defaultSetting.displayXAxis);

    var type = getDropDownValue("minAxisType");
    updateManual(type, "manualMin");
    type = getDropDownValue("maxAxisType");
    updateManual(type, "manualMax");
}

function getCompatibleSparklineSetting(formulaArgs, row, col) {
    var sparklineSetting = {};

    setTextValue("compatibleSparklineData", unParseFormula(formulaArgs[0], row, col));
    setDropDownValue("dataOrientationType", formulaArgs[1].value);
    if (formulaArgs[2]) {
        setTextValue("compatibleSparklineDateAxisData", unParseFormula(formulaArgs[2], row, col));
    }
    else {
        setTextValue("compatibleSparklineDateAxisData", "");
    }
    if (formulaArgs[3]) {
        setDropDownValue("dateAxisOrientationType", formulaArgs[3].value);
    }
    else {
        setDropDownValue("dateAxisOrientationType", -1);
    }
    var colorExpression = parseColorExpression(formulaArgs[4], row, col);
    if (colorExpression) {
        sparklineSetting = parseSetting(colorExpression);
    }
    updateSparklineSetting(sparklineSetting);
    updateStyleSetting(sparklineSetting);
}

function getScatterSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {
        tags: false,
        drawSymbol: true,
        drawLines: false,
        dash: false,
        color1: "#969696",
        color2: "#CB0000"
    };
    var inputList = ["scatterSparklinePoints1", "scatterSparklinePoints2", "scatterSparklineMinX", "scatterSparklineMaxX",
        "scatterSparklineMinY", "scatterSparklineMaxY", "scatterSparklineHLine", "scatterSparklineVLine",
        "scatterSparklineXMinZone", "scatterSparklineXMaxZone", "scatterSparklineYMinZone", "scatterSparklineYMaxZone"];
    for (var i = 0; i < inputList.length; i++) {
        var formula = "";
        if (formulaArgs[i]) {
            formula = unParseFormula(formulaArgs[i], row, col);
        }
        setTextValue(inputList[i], formula);
    }

    var color1 = parseColorExpression(formulaArgs[15], row, col);
    var color2 = parseColorExpression(formulaArgs[16], row, col);
    var tags = formulaArgs[12] && (formulaArgs[12].type === ExpressionType.boolean ? formulaArgs[12].value : null);
    var drawSymbol = formulaArgs[13] && (formulaArgs[13].type === ExpressionType.boolean ? formulaArgs[13].value : null);
    var drawLines = formulaArgs[14] && (formulaArgs[14].type === ExpressionType.boolean ? formulaArgs[14].value : null);
    var dashLine = formulaArgs[17] && (formulaArgs[17].type === ExpressionType.boolean ? formulaArgs[17].value : null);

    setColorValue("scatterSparklineColor1", (color1 !== null) ? color1 : defaultValue.color1);
    setColorValue("scatterSparklineColor2", (color2 !== null) ? color2 : defaultValue.color2);
    setCheckValue("scatterSparklineTags", tags !== null ? tags : defaultValue.tags);
    setCheckValue("scatterSparklineDrawSymbol", drawSymbol !== null ? drawSymbol : defaultValue.drawSymbol);
    setCheckValue("scatterSparklineDrawLines", drawLines !== null ? drawLines : defaultValue.drawLines);
    setCheckValue("scatterSparklineDashLine", dashLine !== null ? dashLine : defaultValue.dash);
}

function getHBarSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorScheme: "#969696"};

    var value = unParseFormula(formulaArgs[0], row, col);
    var colorScheme = parseColorExpression(formulaArgs[1], row, col);

    setTextValue("hbarSparklineValue", value);
    setColorValue("hbarSparklineColorScheme", (colorScheme !== null) ? colorScheme : defaultValue.colorScheme);
}

function getVBarSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {colorScheme: "#969696"};

    var value = unParseFormula(formulaArgs[0], row, col);
    var colorScheme = parseColorExpression(formulaArgs[1], row, col);

    setTextValue("vbarSparklineValue", value);
    setColorValue("vbarSparklineColorScheme", (colorScheme !== null) ? colorScheme : defaultValue.colorScheme);
}

function getParetoSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {label: 0, vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsRangeValue = unParseFormula(formulaArgs[0], row, col);
        var pointIndexValue = unParseFormula(formulaArgs[1], row, col);
        var colorRangeValue = unParseFormula(formulaArgs[2], row, col);
        var targetValue = unParseFormula(formulaArgs[3], row, col);
        var target2Value = unParseFormula(formulaArgs[4], row, col);
        var highlightPositionValue = unParseFormula(formulaArgs[5], row, col);
        var labelValue = formulaArgs[6] && (formulaArgs[6].type === ExpressionType.number ? formulaArgs[6].value : null);
        var verticalValue = formulaArgs[7] && (formulaArgs[7].type === ExpressionType.boolean ? formulaArgs[7].value : null);

        setTextValue("paretoSparklinePoints", pointsRangeValue);
        setTextValue("paretoSparklinePointIndex", pointIndexValue);
        setTextValue("paretoSparklineColorRange", colorRangeValue);
        setTextValue("paretoSparklineHighlightPosition", highlightPositionValue);
        setTextValue("paretoSparklineTarget", targetValue);
        setTextValue("paretoSparklineTarget2", target2Value);
        setDropDownValue("paretoLabelType", labelValue === null ? defaultValue.label : labelValue);
        setCheckValue("paretoSparklineVertical", verticalValue === null ? defaultValue.vertical : verticalValue);
    }
    else {
        setTextValue("paretoSparklinePoints", "");
        setTextValue("paretoSparklinePointIndex", "");
        setTextValue("paretoSparklineColorRange", "");
        setTextValue("paretoSparklineHighlightPosition", "");
        setTextValue("paretoSparklineTarget", "");
        setTextValue("paretoSparklineTarget2", "");
        setDropDownValue("paretoLabelType", defaultValue.label);
        setCheckValue("paretoSparklineVertical", defaultValue.vertical);
    }
}

function getSpreadSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {showAverage: false, style: 4, colorScheme: "#646464", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var showAverageValue = formulaArgs[1] && (formulaArgs[1].type === ExpressionType.boolean ? formulaArgs[1].value : null);
        var scaleStartValue = unParseFormula(formulaArgs[2], row, col);
        var scaleEndValue = unParseFormula(formulaArgs[3], row, col);
        var styleValue = formulaArgs[4] ? unParseFormula(formulaArgs[4], row, col) : null;
        var colorSchemeValue = parseColorExpression(formulaArgs[5], row, col);
        var verticalValue = formulaArgs[6] && (formulaArgs[6].type === ExpressionType.boolean ? formulaArgs[6].value : null);

        setTextValue("spreadSparklinePoints", pointsValue);
        setCheckValue("spreadSparklineShowAverage", showAverageValue ? showAverageValue : defaultValue.showAverage);
        setTextValue("spreadSparklineScaleStart", scaleStartValue);
        setTextValue("spreadSparklineScaleEnd", scaleEndValue);
        setDropDownValue("spreadSparklineStyleType", styleValue ? styleValue : defaultValue.style);
        setColorValue("spreadSparklineColorScheme", colorSchemeValue ? colorSchemeValue : defaultValue.colorScheme);
        setCheckValue("spreadSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
    }
    else {
        setTextValue("spreadSparklinePoints", "");
        setCheckValue("spreadSparklineShowAverage", defaultValue.showAverage);
        setTextValue("spreadSparklineScaleStart", "");
        setTextValue("spreadSparklineScaleEnd", "");
        setDropDownValue("spreadSparklineStyleType", defaultValue.style);
        setColorValue("spreadSparklineColorScheme", defaultValue.colorScheme);
        setCheckValue("spreadSparklineVertical", defaultValue.vertical);
    }
}

function getStackedSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {color: "#646464", vertical: false, textOrientation: 0};

    if (formulaArgs && formulaArgs.length > 0) {
        var pointsValue = unParseFormula(formulaArgs[0], row, col);
        var colorRangeValue = unParseFormula(formulaArgs[1], row, col);
        var labelRangeValue = unParseFormula(formulaArgs[2], row, col);
        var maximumValue = unParseFormula(formulaArgs[3], row, col);
        var targetRedValue = unParseFormula(formulaArgs[4], row, col);
        var targetGreenValue = unParseFormula(formulaArgs[5], row, col);
        var targetBlueValue = unParseFormula(formulaArgs[6], row, col);
        var targetYellowValue = unParseFormula(formulaArgs[7], row, col);
        var colorValue = parseColorExpression(formulaArgs[8], row, col);
        var highlightPositionValue = unParseFormula(formulaArgs[9], row, col);
        var verticalValue = formulaArgs[10] && (formulaArgs[10].type === ExpressionType.boolean ? formulaArgs[10].value : null);
        var textOrientationValue = unParseFormula(formulaArgs[11], row, col);
        var textSizeValue = unParseFormula(formulaArgs[12], row, col);

        setTextValue("stackedSparklinePoints", pointsValue);
        setTextValue("stackedSparklineColorRange", colorRangeValue);
        setTextValue("stackedSparklineLabelRange", labelRangeValue);
        setNumberValue("stackedSparklineMaximum", maximumValue);
        setNumberValue("stackedSparklineTargetRed", targetRedValue);
        setNumberValue("stackedSparklineTargetGreen", targetGreenValue);
        setNumberValue("stackedSparklineTargetBlue", targetBlueValue);
        setNumberValue("stackedSparklineTargetYellow", targetYellowValue);
        setColorValue("stackedSparklineColor", "stacked-color-span", colorValue ? colorValue : defaultValue.color);
        setNumberValue("stackedSparklineHighlightPosition", highlightPositionValue);
        setCheckValue("stackedSparklineVertical", verticalValue ? verticalValue : defaultValue.vertical);
        setDropDownValue("stackedSparklineTextOrientation", textOrientationValue ? textOrientationValue : defaultValue.textOrientation);
        setNumberValue("stackedSparklineTextSize", textSizeValue);
    }
    else {
        setTextValue("stackedSparklinePoints", "");
        setTextValue("stackedSparklineColorRange", "");
        setTextValue("stackedSparklineLabelRange", "");
        setNumberValue("stackedSparklineMaximum", "");
        setNumberValue("stackedSparklineTargetRed", "");
        setNumberValue("stackedSparklineTargetGreen", "");
        setNumberValue("stackedSparklineTargetBlue", "");
        setNumberValue("stackedSparklineTargetYellow", "");
        setColorValue("stackedSparklineColor", "stacked-color-span", defaultValue.color);
        setNumberValue("stackedSparklineHighlightPosition", "");
        setCheckValue("stackedSparklineVertical", defaultValue.vertical);
        setDropDownValue("stackedSparklineTextOrientation", defaultValue.textOrientation);
        setNumberValue("stackedSparklineTextSize", "");
    }
}

function getVariSparklineSetting(formulaArgs, row, col) {
    var defaultValue = {legend: false, colorPositive: "green", colorNegative: "red", vertical: false};

    if (formulaArgs && formulaArgs.length > 0) {
        var varianceValue = unParseFormula(formulaArgs[0], row, col);
        var referenceValue = unParseFormula(formulaArgs[1], row, col);
        var miniValue = unParseFormula(formulaArgs[2], row, col);
        var maxiValue = unParseFormula(formulaArgs[3], row, col);
        var markValue = unParseFormula(formulaArgs[4], row, col);
        var tickunitValue = unParseFormula(formulaArgs[5], row, col);
        var legendValue = formulaArgs[6] && (formulaArgs[6].type === ExpressionType.boolean ? formulaArgs[6].value : null);
        var colorPositiveValue = parseColorExpression(formulaArgs[7], row, col);
        var colorNegativeValue = parseColorExpression(formulaArgs[8], row, col);
        var verticalValue = formulaArgs[9] && (formulaArgs[9].type === ExpressionType.boolean ? formulaArgs[9].value : null);

        setTextValue("variSparklineVariance", varianceValue);
        setTextValue("variSparklineReference", referenceValue);
        setTextValue("variSparklineMini", miniValue);
        setTextValue("variSparklineMaxi", maxiValue);
        setTextValue("variSparklineMark", markValue);
        setTextValue("variSparklineTickUnit", tickunitValue);
        setColorValue("variSparklineColorPositive", colorPositiveValue ? colorPositiveValue : defaultValue.colorPositive);
        setColorValue("variSparklineColorNegative", colorNegativeValue ? colorNegativeValue : defaultValue.colorNegative);
        setCheckValue("variSparklineLegend", legendValue);
        setCheckValue("variSparklineVertical", verticalValue);
    }
    else {
        setTextValue("variSparklineVariance", "");
        setTextValue("variSparklineReference", "");
        setTextValue("variSparklineMini", "");
        setTextValue("variSparklineMaxi", "");
        setTextValue("variSparklineMark", "");
        setTextValue("variSparklineTickUnit", "");
        setColorValue("variSparklineColorPositive", defaultValue.colorPositive);
        setColorValue("variSparklineColorNegative", defaultValue.colorNegative);
        setCheckValue("variSparklineLegend", defaultValue.legend);
        setCheckValue("variSparklineVertical", defaultValue.vertical);
    }
}

function getMonthSparklineSetting(formulaArgs, row, col) {
    var year = "", month = "", dataRangeStr = "", emptyColor = "lightgray", startColor = "lightgreen", middleColor = "green", endColor = "darkgreen", colorRangeStr = "";
    if (formulaArgs) {
        if (formulaArgs.length === 7) {
            year = unParseFormula(formulaArgs[0], row, col);
            month = unParseFormula(formulaArgs[1], row, col);
            dataRangestr = unParseFormula(formulaArgs[2], row, col);
            emptyColor = parseColorExpression(formulaArgs[3], row, col);
            startColor = parseColorExpression(formulaArgs[4], row, col);
            middleColor = parseColorExpression(formulaArgs[5], row, col);
            endColor = parseColorExpression(formulaArgs[6], row, col);
            setTextValue("monthSparklineYear", year);
            setTextValue("monthSparklineMonth", month);
            setTextValue("monthSparklineData", dataRangestr);
            setColorValue("monthSparklineEmptyColor", emptyColor);
            setColorValue("monthSparklineStartColor", startColor);
            setColorValue("monthSparklineMiddleColor", middleColor);
            setColorValue("monthSparklineEndColor", endColor);
            setTextValue("monthSparklineColorRange", "");
        } else {
            year = unParseFormula(formulaArgs[0], row, col);
            month = unParseFormula(formulaArgs[1], row, col);
            dataRangestr = unParseFormula(formulaArgs[2], row, col);
            colorRangeStr = unParseFormula(formulaArgs[3], row, col);
            setTextValue("monthSparklineYear", year);
            setTextValue("monthSparklineMonth", month);
            setTextValue("monthSparklineData", dataRangestr);
            setColorValue("monthSparklineEmptyColor", emptyColor);
            setColorValue("monthSparklineStartColor", startColor);
            setColorValue("monthSparklineMiddleColor", middleColor);
            setColorValue("monthSparklineEndColor", endColor);
            setTextValue("monthSparklineColorRange", colorRangeStr);
        }
    } else {
        setTextValue("monthSparklineYear", year);
        setTextValue("monthSparklineMonth", month);
        setTextValue("monthSparklineData", dataRangestr);
        setColorValue("monthSparklineEmptyColor", emptyColor);
        setColorValue("monthSparklineStartColor", startColor);
        setColorValue("monthSparklineMiddleColor", middleColor);
        setColorValue("monthSparklineEndColor", endColor);
        setTextValue("monthSparklineColorRange", colorRangeStr);
    }
}

function getYearSparklineSetting(formulaArgs, row, col) {
    var year = "", month = "", dataRangeStr = "", emptyColor = "lightgray", startColor = "lightgreen", middleColor = "green", endColor = "darkgreen", colorRangeStr = "";
    if (formulaArgs) {
        if (formulaArgs.length === 6) {
            year = unParseFormula(formulaArgs[0], row, col);
            dataRangestr = unParseFormula(formulaArgs[1], row, col);
            emptyColor = parseColorExpression(formulaArgs[2], row, col);
            startColor = parseColorExpression(formulaArgs[3], row, col);
            middleColor = parseColorExpression(formulaArgs[4], row, col);
            endColor = parseColorExpression(formulaArgs[5], row, col);
            setTextValue("yearSparklineYear", year);
            setTextValue("yearSparklineData", dataRangestr);
            setColorValue("yearSparklineEmptyColor", emptyColor);
            setColorValue("yearSparklineStartColor", startColor);
            setColorValue("yearSparklineMiddleColor", middleColor);
            setColorValue("yearSparklineEndColor", endColor);
            setTextValue("yearSparklineColorRange", "");
        } else {
            year = unParseFormula(formulaArgs[0], row, col);
            dataRangestr = unParseFormula(formulaArgs[1], row, col);
            colorRangeStr = unParseFormula(formulaArgs[2], row, col);
            setTextValue("yearSparklineYear", year);
            setTextValue("yearSparklineData", dataRangestr);
            setColorValue("yearSparklineEmptyColor", emptyColor);
            setColorValue("yearSparklineStartColor", startColor);
            setColorValue("yearSparklineMiddleColor", middleColor);
            setColorValue("yearSparklineEndColor", endColor);
            setTextValue("yearSparklineColorRange", colorRangeStr);
        }
    } else {
        setTextValue("yearSparklineYear", year);
        setTextValue("yearSparklineData", dataRangestr);
        setColorValue("yearSparklineEmptyColor", emptyColor);
        setColorValue("yearSparklineStartColor", startColor);
        setColorValue("yearSparklineMiddleColor", middleColor);
        setColorValue("yearSparklineEndColor", endColor);
        setTextValue("yearSparklineColorRange", colorRangeStr);
    }
}

function getQRCodeSparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", level: "L", model: 2, version: "auto", mask: "auto", connection: false, connectionNo: 0, charCode: "", charset: "UTF-8", quietZoneLeft: 4, quietZoneRight: 4, quietZoneTop: 4, quietZoneBottom: 4};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            level = unParseFormula(formulaArgs[3], row, col),
            model = unParseFormula(formulaArgs[4], row, col),
            version = unParseFormula(formulaArgs[5], row, col),
            mask = unParseFormula(formulaArgs[6], row, col),
            connection = formulaArgs[7]  && formulaArgs[7].value,
            connectionNo = unParseFormula(formulaArgs[8], row, col),
            charCode = unParseFormula(formulaArgs[9], row, col),
            charset = unParseFormula(formulaArgs[10], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[11], row, col),
            quietZoneRight = unParseFormula(formulaArgs[12], row, col),
            quietZoneTop = unParseFormula(formulaArgs[13], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[14], row, col);

        setTextValue("qrCodeSparklineData", dataRangestr);
        setColorValue("qrCodeSparklineColor", color ? color : defaultValue.color);
        setColorValue("qrCodeSparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setDropDownValue("errorCorrectionLevelType", level ? unparseSparklineColorOptions(level) : defaultValue.level);
        setDropDownValue("qrCodeSparklineModel", model ? model : defaultValue.model);
        setDropDownValue("qrCodeSparklineVersion", version ? unparseSparklineColorOptions(version) : defaultValue.version);
        setDropDownValue("qrCodeSparklineMask", mask ? unparseSparklineColorOptions(mask) : defaultValue.mask);
        setCheckValue("checkboxQRCodeSparklineConnection", connection);
        setDropDownValue("qrCodeSparklineConnectionNo", connectionNo ? connectionNo : defaultValue.connectionNo);
        setDropDownValue("qrCodeCharsetType", charset ? unparseSparklineColorOptions(charset) : defaultValue.charset);
        setTextValue("qrCodeSparklineCharCode", charCode ? unparseBraceOptions(charCode) : defaultValue.charCode);
        setNumberValue("qrCodeSparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("qrCodeSparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("qrCodeSparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("qrCodeSparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("qrCodeSparklineData", "");
        setColorValue("qrCodeSparklineColor", defaultValue.color);
        setColorValue("qrCodeSparklineBackgroundColor", defaultValue.backgroundColor);
        setDropDownValue("errorCorrectionLevelType", defaultValue.level);
        setDropDownValue("qrCodeSparklineModel", defaultValue.model);
        setDropDownValue("qrCodeSparklineVersion", defaultValue.version);
        setDropDownValue("qrCodeSparklineMask", defaultValue.mask);
        setCheckValue("checkboxQRCodeSparklineConnection", defaultValue.connection);
        setDropDownValue("qrCodeSparklineConnectionNo", defaultValue.connectionNo);
        setDropDownValue("qrCodeCharsetType", defaultValue.charCode);
        setTextValue("qrCodeSparklineCharCode", defaultValue.charset);
        setNumberValue("qrCodeSparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("qrCodeSparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("qrCodeSparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("qrCodeSparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
}

function getEAN8SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 11, quietZoneRight:7, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            fontFamily = unParseFormula(formulaArgs[5], row, col),
            fontStyle = unParseFormula(formulaArgs[6], row, col),
            fontWeight = unParseFormula(formulaArgs[7], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[8], row, col),
            fontTextAlign = unParseFormula(formulaArgs[9], row, col),
            fontSize = unParseFormula(formulaArgs[10], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[11], row, col),
            quietZoneRight = unParseFormula(formulaArgs[12], row, col),
            quietZoneTop = unParseFormula(formulaArgs[13], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[14], row, col);

        setTextValue("ean8SparklineData", dataRangestr);
        setColorValue("ean8SparklineColor", color ? color : defaultValue.color);
        setColorValue("ean8SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxEAN8SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("ean8SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("ean8SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("ean8SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("ean8SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("ean8SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("ean8SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("ean8SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("ean8SparklineData", "");
        setColorValue("ean8SparklineColor", defaultValue.color);
        setColorValue("ean8SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("ean8SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("ean8SparklineLabelPosition",  defaultValue.labelPosition);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("ean8SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("ean8SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean8SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("ean8SparklineLeft", efaultValue.quietZoneLeft);
        setNumberValue("ean8SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("ean8SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("ean8SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getEAN13SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", addOn: "", addOnLabelPosition: "bottom", fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 11, quietZoneRight: 7, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            addOn = unParseFormula(formulaArgs[5], row, col),
            addOnLabelPosition = unParseFormula(formulaArgs[6], row, col),
            fontFamily = unParseFormula(formulaArgs[7], row, col),
            fontStyle = unParseFormula(formulaArgs[8], row, col),
            fontWeight = unParseFormula(formulaArgs[9], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[10], row, col),
            fontTextAlign = unParseFormula(formulaArgs[11], row, col),
            fontSize = unParseFormula(formulaArgs[12], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[13], row, col),
            quietZoneRight = unParseFormula(formulaArgs[14], row, col),
            quietZoneTop = unParseFormula(formulaArgs[15], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[16], row, col);

        setTextValue("ean13SparklineData", dataRangestr);
        setColorValue("ean13SparklineColor", color ? color : defaultValue.color);
        setColorValue("ean13SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxEAN13SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("ean13SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) :defaultValue.labelPosition);
        setTextValue("ean13SparklineAddOn", addOn ? unparseSparklineColorOptions(addOn) : defaultValue.addOn);
        setDropDownValue("ean13SparklineAddOnLabelPosition", addOnLabelPosition ? unparseSparklineColorOptions(addOnLabelPosition) : defaultValue.addOnLabelPosition);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("ean13SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("ean13SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("ean13SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("ean13SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("ean13SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("ean13SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("ean13SparklineData", "");
        setColorValue("ean13SparklineColor", defaultValue.color);
        setColorValue("ean13SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("ean13SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("ean13SparklineLabelPosition",  defaultValue.labelPosition);
        setTextValue("ean13SparklineAddOn",  defaultValue.addOn);
        setDropDownValue("ean13SparklineAddOnLabelPosition",  defaultValue.addOnLabelPosition);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("ean13SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("ean13SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='ean13SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("ean13SparklineLeft", efaultValue.quietZoneLeft);
        setNumberValue("ean13SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("ean13SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("ean13SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getGS1SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 10, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            fontFamily = unParseFormula(formulaArgs[5], row, col),
            fontStyle = unParseFormula(formulaArgs[6], row, col),
            fontWeight = unParseFormula(formulaArgs[7], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[8], row, col),
            fontTextAlign = unParseFormula(formulaArgs[9], row, col),
            fontSize = unParseFormula(formulaArgs[10], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[11], row, col),
            quietZoneRight = unParseFormula(formulaArgs[12], row, col),
            quietZoneTop = unParseFormula(formulaArgs[13], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[14], row, col);

            setTextValue("gs1SparklineData", dataRangestr);
            setColorValue("gs1SparklineColor", color ? color : defaultValue.color);
            setColorValue("gs1SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
            setCheckValue("checkboxGS1SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
            setDropDownValue("gs1SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
            setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
            setDropDownValue("gs1SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
            setDropDownValue("gs1SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
            setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
            setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
            setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
            setNumberValue("gs1SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
            setNumberValue("gs1SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
            setNumberValue("gs1SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
            setNumberValue("gs1SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("gs1SparklineData", "");
        setColorValue("gs1SparklineColor", defaultValue.color);
        setColorValue("gs1SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("gs1SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("gs1SparklineLabelPosition",  defaultValue.labelPosition);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("gs1SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("gs1SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='gs1SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("gs1SparklineLeft", efaultValue.quietZoneLeft);
        setNumberValue("gs1SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("gs1SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("gs1SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getCodabarSparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", checkDigit: false, nwRatio: 3, fontFamily: "Sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 10, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            checkDigit = formulaArgs[5] && formulaArgs[5].value,
            nwRatio = unParseFormula(formulaArgs[6], row, col),
            fontFamily = unParseFormula(formulaArgs[7], row, col),
            fontStyle = unParseFormula(formulaArgs[8], row, col),
            fontWeight = unParseFormula(formulaArgs[9], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[10], row, col),
            fontTextAlign = unParseFormula(formulaArgs[11], row, col),
            fontSize = unParseFormula(formulaArgs[12], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[13], row, col),
            quietZoneRight = unParseFormula(formulaArgs[14], row, col),
            quietZoneTop = unParseFormula(formulaArgs[15], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[16], row, col);

        setTextValue("codabarSparklineData", dataRangestr);
        setColorValue("codabarSparklineColor", color ? color : defaultValue.color);
        setColorValue("codabarSparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxCodabarSparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("codabarSparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setCheckValue("checkboxCodabarSparklineCheckDigit", checkDigit  ? checkDigit : defaultValue.checkDigit);
        setDropDownValue("codabarNWRatio", nwRatio ? nwRatio : defaultValue.nwRatio);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("codabarSparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("codabarSparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("codabarSparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("codabarSparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("codabarSparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("codabarSparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);

    }
    else {
        setTextValue("codabarSparklineData", "");
        setColorValue("codabarSparklineColor", defaultValue.color);
        setColorValue("codabarSparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("codabarSparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("codabarSparklineLabelPosition",  defaultValue.labelPosition);
        setCheckValue("checkboxCodabarSparklineCheckDigit", defaultValue.checkDigit);
        setDropDownValue("codabarNWRatio", defaultValue.nwRatio);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("codabarSparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("codabarSparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='codabarSparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("codabarSparklineLeft", efaultValue.quietZoneLeft);
        setNumberValue("codabarSparklineRight", defaultValue.quietZoneRight);
        setNumberValue("codabarSparklineTop", defaultValue.quietZoneTop);
        setNumberValue("codabarSparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getCode93SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", checkDigit: false, fullASCII: false, fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 10, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            checkDigit = formulaArgs[5] && formulaArgs[5].value,
            fullASCII = formulaArgs[6] && formulaArgs[6].value,
            fontFamily = unParseFormula(formulaArgs[7], row, col),
            fontStyle = unParseFormula(formulaArgs[8], row, col),
            fontWeight = unParseFormula(formulaArgs[9], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[10], row, col),
            fontTextAlign = unParseFormula(formulaArgs[11], row, col),
            fontSize = unParseFormula(formulaArgs[12], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[13], row, col),
            quietZoneRight = unParseFormula(formulaArgs[14], row, col),
            quietZoneTop = unParseFormula(formulaArgs[15], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[16], row, col);


        setTextValue("code93SparklineData", dataRangestr);
        setColorValue("code93SparklineColor", color ? color : defaultValue.color);
        setColorValue("code93SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxCode93SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("code93SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setCheckValue("checkboxCode93SparklineCheckDigit", checkDigit ? checkDigit : defaultValue.checkDigit);
        setCheckValue("checkCode93SparklineFullASCII", fullASCII ? fullASCII : defaultValue.fullASCII);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("code93SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("code93SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("code93SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("code93SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("code93SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("code93SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);

    }
    else {
        setTextValue("code93SparklineData", "");
        setColorValue("code93SparklineColor", defaultValue.color);
        setColorValue("code93SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("code93SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("code93SparklineLabelPosition",  defaultValue.labelPosition);
        setCheckValue("checkboxCode93SparklineCheckDigit", defaultValue.checkDigit);
        setCheckValue("checkCode93SparklineFullASCII", defaultValue.fullASCII);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("code93SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("code93SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code93SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("code93SparklineLeft", efaultValue.quietZoneLeft);
        setNumberValue("code93SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("code93SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("code93SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getCode39SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", checkDigit: false, fullASCII: false, labelWithStartAndStopCharacter: false, nwRatio: 3, fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 10, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            character = formulaArgs[5] && formulaArgs[5].value,
            nwRatio = unParseFormula(formulaArgs[7], row, col),
            checkDigit = formulaArgs[6] && formulaArgs[6].value,
            fullASCII = formulaArgs[8] && formulaArgs[8].value,
            fontFamily = unParseFormula(formulaArgs[9], row, col),
            fontStyle = unParseFormula(formulaArgs[10], row, col),
            fontWeight = unParseFormula(formulaArgs[11], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[12], row, col),
            fontTextAlign = unParseFormula(formulaArgs[13], row, col),
            fontSize = unParseFormula(formulaArgs[14], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[15], row, col),
            quietZoneRight = unParseFormula(formulaArgs[16], row, col),
            quietZoneTop = unParseFormula(formulaArgs[17], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[18], row, col);

        setTextValue("code39SparklineData", dataRangestr);
        setColorValue("code39SparklineColor", color ? color : defaultValue.color);
        setColorValue("code39SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxCode39SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("code39SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setCheckValue("checkboxCode39SparklineCharacter", character ? character : defaultValue.character);
        setDropDownValue("code39SparklineNWRatio", nwRatio ? nwRatio : defaultValue.nwRatio);
        setCheckValue("checkboxCode39SparklineCheckDigit", checkDigit ? checkDigit : defaultValue.checkDigit);
        setCheckValue("checkCode39SparklineFullASCII", fullASCII ? fullASCII : defaultValue.fullASCII);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("code39SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("code39SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("code39SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("code39SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("code39SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("code39SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);

    }
    else {
        setTextValue("code39SparklineData", "");
        setColorValue("code39SparklineColor", defaultValue.color);
        setColorValue("code39SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("code39SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("code39SparklineLabelPosition",  defaultValue.labelPosition);
        setCheckValue("checkboxCode39SparklineCharacter", defaultValue.character);
        setDropDownValue("code39SparklineNWRatio",defaultValue.nwRatio);
        setCheckValue("checkboxCode39SparklineCheckDigit", defaultValue.checkDigit);
        setCheckValue("checkCode39SparklineFullASCII", defaultValue.fullASCII);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("code39SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("code39SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code39SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("code39SparklineLeft", defaultValue.quietZoneLeft);
        setNumberValue("code39SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("code39SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("code39SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getShowLabelValue(showLabel,defaultValue){
    switch(showLabel){
        case undefined:
            showLabel = Boolean(defaultValue.showLabel);
            break;
        case 0:
        case 1:
            showLabel = Boolean(showLabel);
            break;
    }
    return showLabel;
}

function getCode49SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", grouping: false, groupNoValue: 0, fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 1, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            grouping = formulaArgs[5] && formulaArgs[5].value,
            groupNoValue = unParseFormula(formulaArgs[6], row, col),
            fontFamily = unParseFormula(formulaArgs[7], row, col),
            fontStyle = unParseFormula(formulaArgs[8], row, col),
            fontWeight = unParseFormula(formulaArgs[9], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[10], row, col),
            fontTextAlign = unParseFormula(formulaArgs[11], row, col),
            fontSize = unParseFormula(formulaArgs[12], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[13], row, col),
            quietZoneRight = unParseFormula(formulaArgs[14], row, col),
            quietZoneTop = unParseFormula(formulaArgs[15], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[16], row, col);

        setTextValue("code49SparklineData", dataRangestr);
        setColorValue("code49SparklineColor", color ? color : defaultValue.color);
        setColorValue("code49SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxCode49SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("code49SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setCheckValue("checkboxCode49SparklineGrouping", grouping ? grouping : defaultValue.grouping);
        setNumberValue("code49SparklineGroupNo", groupNoValue ? groupNoValue : defaultValue.groupNoValue);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("code49SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("code49SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("code49SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("code49SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("code49SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("code49SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("code49SparklineData", "");
        setColorValue("code49SparklineColor", defaultValue.color);
        setColorValue("code49SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("code49SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("code49SparklineLabelPosition",  defaultValue.labelPosition);
        setCheckValue("checkboxCode49SparklineGrouping", defaultValue.grouping);
        setNumberValue("code49SparklineGroupNo", defaultValue.groupNoValue);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("code49SparklineFontStyle", defaultValue.fontStyle);
        setDropDownValue("code49SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code49SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("code49SparklineLeft", defaultValue.quietZoneLeft);
        setNumberValue("code49SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("code49SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("code49SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getCode128SparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", showLabel: 1, labelPosition: "bottom", codeset: "auto", fontFamily: "sans-serif", fontStyle: "normal", fontWeight: "normal", fontTextDecoration: "none", fontTextAlign: "center", fontSize: 12, quietZoneLeft: 10, quietZoneRight: 10, quietZoneTop: 0, quietZoneBottom: 0};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            showLabel = formulaArgs[3] && formulaArgs[3].value,
            labelPosition = unParseFormula(formulaArgs[4], row, col),
            codeset = unParseFormula(formulaArgs[5], row, col),
            fontFamily = unParseFormula(formulaArgs[6], row, col),
            fontStyle = unParseFormula(formulaArgs[7], row, col),
            fontWeight = unParseFormula(formulaArgs[8], row, col),
            fontTextDecoration = unParseFormula(formulaArgs[9], row, col),
            fontTextAlign = unParseFormula(formulaArgs[10], row, col),
            fontSize = unParseFormula(formulaArgs[11], row, col),
            quietZoneLeft = unParseFormula(formulaArgs[12], row, col),
            quietZoneRight = unParseFormula(formulaArgs[13], row, col),
            quietZoneTop = unParseFormula(formulaArgs[14], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[15], row, col);

        setTextValue("code128SparklineData", dataRangestr);
        setColorValue("code128SparklineColor", color ? color : defaultValue.color);
        setColorValue("code128SparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setCheckValue("checkboxCode128SparklineShowLabel", getShowLabelValue(showLabel,defaultValue));
        setDropDownValue("code128SparklineLabelPosition", labelPosition ? unparseSparklineColorOptions(labelPosition) : defaultValue.labelPosition);
        setDropDownValue("code128Codeset", codeset ? unparseSparklineColorOptions(codeset) : defaultValue.codeset);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontFamily']"), fontFamily ? unparseSparklineColorOptions(fontFamily) : defaultValue.fontFamily);
        setDropDownValue("code128SparklineFontStyle", fontStyle ? unparseSparklineColorOptions(fontStyle) : defaultValue.fontStyle);
        setDropDownValue("code128SparklineFontWeight", fontWeight ? unparseSparklineColorOptions(fontWeight) : defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontTextDecoration']"), fontTextDecoration ? unparseSparklineColorOptions(fontTextDecoration) : defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontTextAlign']"), fontTextAlign ? unparseSparklineColorOptions(fontTextAlign) : defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontSize']"), fontSize ? fontSize : defaultValue.fontSize);
        setNumberValue("code128SparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("code128SparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("code128SparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("code128SparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("code128SparklineData", "");
        setColorValue("code128SparklineColor", defaultValue.color);
        setColorValue("code128SparklineBackgroundColor", defaultValue.backgroundColor);
        setCheckValue("code128SparklineShowLabel", Boolean(defaultValue.showLabel));
        setDropDownValue("code128SparklineLabelPosition",  defaultValue.labelPosition);
        setDropDownValue("code128Codeset", defaultValue.codeset);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontFamily']"), defaultValue.fontFamily);
        setDropDownValue("code128SparklineFontStyle",defaultValue.fontStyle);
        setDropDownValue("code128SparklineFontWeight", defaultValue.fontWeight);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontTextDecoration']"), defaultValue.fontTextDecoration);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontTextAlign']"), defaultValue.fontTextAlign);
        setDropDownText($("#sparklineExTab div.insp-dropdown-list[data-name='code128SparklineFontSize']"), defaultValue.fontSize);
        setNumberValue("code128SparklineLeft", defaultValue.quietZoneLeft);
        setNumberValue("code128SparklineRight", defaultValue.quietZoneRight);
        setNumberValue("code128SparklineTop", defaultValue.quietZoneTop);
        setNumberValue("code128SparklineBottom", defaultValue.quietZoneBottom);
    }
}

function getPDFSparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", errorCorrectionLevel: "auto", rows: "auto", columns: "auto", compact: false, quietZoneLeft: 2, quietZoneRight: 2, quietZoneTop: 2, quietZoneBottom: 2};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            errorCorrectionLevel = unParseFormula(formulaArgs[3], row, col),
            rows = unParseFormula(formulaArgs[4], row, col),
            columns = unParseFormula(formulaArgs[5], row, col),
            compact = formulaArgs[6] && formulaArgs[6].value,
            quietZoneLeft = unParseFormula(formulaArgs[7], row, col),
            quietZoneRight = unParseFormula(formulaArgs[8], row, col),
            quietZoneTop = unParseFormula(formulaArgs[9], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[10], row, col);

        setTextValue("pdfSparklineData", dataRangestr);
        setColorValue("pdfSparklineColor", color ? color : defaultValue.color);
        setColorValue("pdfSparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setDropDownValue("pdfSparklineLevel", errorCorrectionLevel ? unparseSparklineColorOptions(errorCorrectionLevel) : defaultValue.errorCorrectionLevel);
        setDropDownValue("pdfSparklineRows", rows ? unparseSparklineColorOptions(rows) : defaultValue.rows);
        setDropDownValue("pdfSparklineColumns", columns ? unparseSparklineColorOptions(columns) : defaultValue.columns);
        setCheckValue("checkboxPDFSparklineCompact", compact ? compact : defaultValue.compact);
        setNumberValue("pdfSparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("pdfSparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("pdfSparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("pdfSparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("pdfSparklineData", "");
        setColorValue("pdfSparklineColor", defaultValue.color);
        setColorValue("pdfSparklineBackgroundColor", defaultValue.backgroundColor);
        setDropDownValue("pdfSparklineLevel", defaultValue.errorCorrectionLevel);
        setDropDownValue("pdfSparklineRows", defaultValue.rows);
        setDropDownValue("pdfSparklineColumns", defaultValue.columns);
        setCheckValue("checkboxPDFSparklineCompact", defaultValue.compact);
        setNumberValue("pdfSparklineLeft", defaultValue.quietZoneLeft);
        setNumberValue("pdfSparklineRight", defaultValue.quietZoneRight);
        setNumberValue("pdfSparklineTop", defaultValue.quietZoneTop);
        setNumberValue("pdfSparklineBottom",defaultValue.quietZoneBottom);
    }
}

function getDataMatrixSparklineSetting(formulaArgs, row, col){
    var defaultValue = {color: "#000", backgroundColor: "#fff", eccMode: "ECC000", ecc200SymbolSize: "squareAuto", ecc200EndcodingMode: "auto", ecc00_140Symbole: "auto", structureAppend: false, structureNumber: '0', fileIdentifier: 0, quietZoneLeft: 4, quietZoneRight: 4, quietZoneTop: 4, quietZoneBottom: 4};
    if(formulaArgs){
        var dataRangestr = unParseFormula(formulaArgs[0], row, col),
            color = parseColorExpression(formulaArgs[1], row, col),
            backgroundColor = parseColorExpression(formulaArgs[2], row, col),
            eccMode = unParseFormula(formulaArgs[3], row, col),
            ecc200SymbolSize = unParseFormula(formulaArgs[4], row, col),
            ecc200EndcodingMode = unParseFormula(formulaArgs[5], row, col),
            ecc00_140Symbole =unParseFormula(formulaArgs[6], row, col),
            structureNumber = unParseFormula(formulaArgs[8], row, col),
            fileIdentifier = unParseFormula(formulaArgs[9], row, col),
            structureAppend = formulaArgs[7] && formulaArgs[7].value,
            quietZoneLeft = unParseFormula(formulaArgs[10], row, col),
            quietZoneRight = unParseFormula(formulaArgs[11], row, col),
            quietZoneTop = unParseFormula(formulaArgs[12], row, col),
            quietZoneBottom = unParseFormula(formulaArgs[13], row, col);

        setTextValue("dataMatrixSparklineData", dataRangestr);
        setColorValue("dataMatrixSparklineColor", color ? color : defaultValue.color);
        setColorValue("dataMatrixSparklineBackgroundColor", backgroundColor ? backgroundColor : defaultValue.backgroundColor);
        setDropDownValue("dataMatrixSparklineEccMode", eccMode ? unparseSparklineColorOptions(eccMode) : defaultValue.eccMode);
        setTextValue("dataMatrixSparklineSize", ecc200SymbolSize ? unparseSparklineColorOptions(ecc200SymbolSize) : defaultValue.ecc200SymbolSize);
        setTextValue("dataMatrixSparklineEndcodingMode", ecc200EndcodingMode ? unparseSparklineColorOptions(ecc200EndcodingMode) : defaultValue.ecc200EndcodingMode);
        setTextValue("dataMatrixSparklineSymbole", ecc00_140Symbole ? unparseSparklineColorOptions(ecc00_140Symbole) : defaultValue.ecc00_140Symbole);
        setDropDownValue("dataMatrixSparklineStructureNumber", structureNumber ? structureNumber : defaultValue.structureNumber);
        setNumberValue("dataMatrixSparklineFileIdentifier", fileIdentifier ? fileIdentifier : defaultValue.fileIdentifier);
        setCheckValue("checkboxPDFSparklineStructureAppend", structureAppend ? structureAppend : defaultValue.structureAppend);
        setNumberValue("dataMatrixSparklineLeft", quietZoneLeft ? quietZoneLeft : defaultValue.quietZoneLeft);
        setNumberValue("dataMatrixSparklineRight", quietZoneRight ? quietZoneRight : defaultValue.quietZoneRight);
        setNumberValue("dataMatrixSparklineTop", quietZoneTop ? quietZoneTop : defaultValue.quietZoneTop);
        setNumberValue("dataMatrixSparklineBottom", quietZoneBottom ? quietZoneBottom : defaultValue.quietZoneBottom);
    }
    else {
        setTextValue("dataMatrixSparklineData", "");
        setColorValue("dataMatrixSparklineColor", defaultValue.color);
        setColorValue("dataMatrixSparklineBackgroundColor", defaultValue.backgroundColor);
        setDropDownValue("dataMatrixSparklineEccMode", defaultValue.eccMode);
        setTextValue("dataMatrixSparklineSize", defaultValue.ecc200SymbolSize);
        setTextValue("dataMatrixSparklineEndcodingMode", defaultValue.ecc200EndcodingMode);
        setTextValue("dataMatrixSparklineSymbole", defaultValue.ecc00_140Symbole);
        setDropDownValue("dataMatrixSparklineStructureNumber", defaultValue.structureNumber);
        setNumberValue("dataMatrixSparklineFileIdentifier", defaultValue.fileIdentifier);
        setCheckValue("checkboxPDFSparklineStructureAppend", defaultValue.structureAppend);
        setNumberValue("dataMatrixSparklineLeft", defaultValue.quietZoneLeft);
        setNumberValue("dataMatrixSparklineRight", defaultValue.quietZoneRight);
        setNumberValue("dataMatrixSparklineTop", defaultValue.quietZoneTop);
        setNumberValue("dataMatrixSparklineBottom", defaultValue.quietZoneBottom);
    }
}

function addPieSparklineColor(count, color, isMinusSymbol) {
    var defaultColor = "rgb(237, 237, 237)";
    color = color ? color : defaultColor;
    var symbolFunClass, symbolClass;
    if (isMinusSymbol) {
        symbolFunClass = "remove-pie-color";
        symbolClass = "ui-pie-sparkline-icon-minus";
    }
    else {
        symbolFunClass = "add-pie-color";
        symbolClass = "ui-pie-sparkline-icon-plus";
    }
    var $pieSparklineColorContainer = $("#pieSparklineColorContainer");
    var pieColorDataName = "pieColorName";
    var $colorDiv = $("<div>" +
        "<div class=\"insp-row\">" +
        "<div>" +
        "<div class=\"insp-color-picker insp-inline-row\" data-name=\"" + pieColorDataName + count + "\">" +
        "<div class=\"title insp-inline-row-item insp-col-6 localize\">" + uiResource.sparklineExTab.pieSparkline.values.color + count + "</div>" +
        "<div class=\"picker insp-inline-row-item insp-col-4\">" +
        "<div style=\"width: 100%; height: 100%\">" +
        "<div class=\"color-view\" style=\"background-color: " + color + ";\"></div>" +
        "</div>" +
        "</div>" +
        "<div class=\"" + symbolFunClass + " insp-inline-row-item insp-col-2\"><span class=\"ui-pie-sparkline-icon " + symbolClass + "\"></span></div>" +
        "</div>" +
        "</div>" +
        "</div>" +
        "</div>");
    $colorDiv.appendTo($pieSparklineColorContainer);
}

function addPieColor(count, color, isMinusSymbol) {
    var $colorSpanDiv = $(".add-pie-color");
    $colorSpanDiv.addClass("remove-pie-color").removeClass("add-pie-color");
    $colorSpanDiv.find("span").addClass("ui-pie-sparkline-icon-minus").removeClass("ui-pie-sparkline-icon-plus");
    addPieSparklineColor(count, color, isMinusSymbol);
    $(".add-pie-color").unbind("click");
    $(".add-pie-color").bind("click", function (evt) {
        var count = $("#pieSparklineColorContainer").find("span").length;
        addPieColor(count + 1);
    });
    $(".remove-pie-color").unbind("click");
    $(".remove-pie-color").bind("click", function (evt) {
        resetPieColor($(evt.target));
    });
    $("div.insp-color-picker .picker").click(showColorPicker);
}

function resetPieColor($colorSpanDiv) {
    if (!$colorSpanDiv.hasClass("ui-pie-sparkline-icon")) {
        return;
    }
    $colorDiv = $colorSpanDiv.parent().parent().parent().parent().parent();
    $colorDiv.remove();
    var $pieSparklineColorContainer = $("#pieSparklineColorContainer");
    var colorArray = [];
    $pieSparklineColorContainer.find(".color-view").each(function () {
        colorArray.push($(this).css("background-color"));
    });
    $pieSparklineColorContainer.empty();
    addMultiPieColor(colorArray);
}

function addMultiPieColor(colorArray) {
    if (!colorArray || colorArray.length === 0) {
        return;
    }
    var length = colorArray.length;
    var i = 0;
    for (i; i < length - 1; i++) {
        addPieSparklineColor(i + 1, colorArray[i], true);
    }
    addPieColor(i + 1, colorArray[i]);
}

function getPieSparklineSetting(formulaArgs, row, col) {
    var agrsLength = formulaArgs.length;
    if (formulaArgs && agrsLength > 0) {
        var range = unParseFormula(formulaArgs[0], row, col);
        setTextValue("pieSparklinePercentage", range);

        var actualLen = agrsLength - 1;
        if (actualLen === 0) {
            addPieColor(1);
        }
        else {
            var colorArray = [];
            for (var i = 1; i <= actualLen; i++) {
                var colorItem = null;
                var color = parseColorExpression(formulaArgs[i], row, col);
                colorArray.push(color);
            }
            addMultiPieColor(colorArray);
        }
    }
}

var sparklineName;
function showSparklineSetting(row, col) {
    var expr = parseFormulaSparkline(row, col);
    if (!expr || !expr.arguments) {
        return false;
    }
    var formulaSparkline = spread.getSparklineEx(expr.functionName);

    if (formulaSparkline) {
        var $sparklineSettingDiv = $("#sparklineExTab>div>div");
        var formulaArgs = expr.arguments;
        $sparklineSettingDiv.hide();
        if (formulaSparkline instanceof Sparklines.PieSparkline) {
            $("#pieSparklineSetting").show();
            $("#pieSparklineColorContainer").empty();
            getPieSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.AreaSparkline) {
            $("#areaSparklineSetting").show();
            getAreaSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.BoxPlotSparkline) {
            $("#boxplotSparklineSetting").show();
            getBoxPlotSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.BulletSparkline) {
            $("#bulletSparklineSetting").show();
            getBulletSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.CascadeSparkline) {
            $("#cascadeSparklineSetting").show();
            getCascadeSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.LineSparkline || formulaSparkline instanceof Sparklines.ColumnSparkline || formulaSparkline instanceof Sparklines.WinlossSparkline) {
            $("#compatibleSparklineSetting").show();
            if (expr.function.name) {
                sparklineName = expr.function.name;
            }
            getCompatibleSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.ScatterSparkline) {
            $("#scatterSparklineSetting").show();
            getScatterSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.HBarSparkline) {
            $("#hbarSparklineSetting").show();
            getHBarSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.VBarSparkline) {
            $("#vbarSparklineSetting").show();
            getVBarSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.ParetoSparkline) {
            $("#paretoSparklineSetting").show();
            getParetoSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.SpreadSparkline) {
            $("#spreadSparklineSetting").show();
            getSpreadSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.StackedSparkline) {
            $("#stackedSparklineSetting").show();
            getStackedSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.VariSparkline) {
            $("#variSparklineSetting").show();
            getVariSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.MonthSparkline) {
            $("#monthSparklineSetting").show();
            getMonthSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Sparklines.YearSparkline) {
            $("#yearSparklineSetting").show();
            getYearSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if (formulaSparkline instanceof Barcode.QRCode){
            $("#qrCodeSparklineSetting").show();
            getQRCodeSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.EAN8){
            $("#ean8SparklineSetting").show();
            getEAN8SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.EAN13){
            $("#ean13SparklineSetting").show();
            getEAN13SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.GS1_128){
            $("#gs1SparklineSetting").show();
            getGS1SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.Codabar){
            $("#codabarSparklineSetting").show();
            getCodabarSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.DataMatrix){
            $("#dataMatrixSparklineSetting").show();
            getDataMatrixSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.PDF417){
            $("#pdfSparklineSetting").show();
            getPDFSparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.Code39){
            $("#code39SparklineSetting").show();
            getCode39SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.Code49){
            $("#code49SparklineSetting").show();
            getCode49SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.Code93){
            $("#code93SparklineSetting").show();
            getCode93SparklineSetting(formulaArgs, row, col);
            return true;
        }
        else if(formulaSparkline instanceof Barcode.Code128){
            $("#code128SparklineSetting").show();
            getCode128SparklineSetting(formulaArgs, row, col);
            return true;
        }
    }
    return false;
}

function attachSparklineSettingEvents() {
    $("#setAreaSparkline").click(applyAreaSparklineSetting);
    $("#setBoxPlotSparkline").click(applyBoxPlotSparklineSetting);
    $("#setBulletSparkline").click(applyBulletSparklineSetting);
    $("#setCascadeSparkline").click(applyCascadeSparklineSetting);
    $("#setCompatibleSparkline").click(applyCompatibleSparklineSetting);
    $("#setScatterSparkline").click(applyScatterSparklineSetting);
    $("#setHbarSparkline").click(applyHbarSparklineSetting);
    $("#setVbarSparkline").click(applyVbarSparklineSetting);
    $("#setParetoSparkline").click(applyParetoSparklineSetting);
    $("#setSpreadSparkline").click(applySpreadSparklineSetting);
    $("#setStackedSparkline").click(applyStackedSparklineSetting);
    $("#setVariSparkline").click(applyVariSparklineSetting);
    $("#setPieSparkline").click(applyPieSparklineSetting);
    $("#setMonthSparkline").click(applyMonthSparklineSetting);
    $("#setYearSparkline").click(applyYearSparklineSetting);
    $("#setQRCodeSparkline").click(applyQRCodeSparklineSetting)
    $("#setEAN8Sparkline").click(applyEAN8SparklineSetting);
    $("#setEAN13Sparkline").click(applyEAN13SparklineSetting);
    $("#setGS1Sparkline").click(applyGS1SparklineSetting);
    $("#setCodabarSparkline").click(applyCodabarSparklineSetting);
    $("#setCode93Sparkline").click(applyCode93SparklineSetting);
    $("#setCode39Sparkline").click(applyCode39SparklineSetting);
    $("#setCode49Sparkline").click(applyCode49SparklineSetting);
    $("#setCode128Sparkline").click(applyCode128SparklineSetting);
    $("#setPDFSparkline").click(applyPDFSparklineSetting);
    $("#setDataMatrixSparkline").click(applyDataMatrixSparklineSetting);
}

function updateFormulaBar() {
    var sheet = spread.getActiveSheet();
    var formulaBar = $("#formulabox");
    if (formulaBar.length > 0) {
        var formula = sheet.getFormula(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex());
        if (formula) {
            formula = "=" + formula;
            formulaBar.text(formula);
        }
    }
}

function removeContinuousComma(parameter) {
    var len = parameter.length;
    while (len > 0 && parameter[len - 1] === ",") {
        len--;
    }
    return parameter.substr(0, len);
}

function formatFormula(paraArray) {
    var params = "";
    for (var i = 0; i < paraArray.length; i++) {
        var item = paraArray[i];
        if (item !== undefined && item !== null) {
            params += item + ",";
        }
        else {
            params += ",";
        }
    }
    params = removeContinuousComma(params);
    return params;
}

function getFormula(params) {
    var len = params.length;
    while (len > 0 && params[len - 1] === "") {
        len--;
    }
    var temp = "";
    for (var i = 0; i < len; i++) {
        temp += params[i];
        if (i !== len - 1) {
            temp += ",";
        }
    }
    return "=AREASPARKLINE(" + temp + ")";
}

function setFormulaSparkline(formula) {
    var sheet = spread.getActiveSheet();
    var row = sheet.getActiveRowIndex();
    var col = sheet.getActiveColumnIndex();
    if (formula) {
        sheet.setFormula(row, col, formula);
    }
}

function applyAreaSparklineSetting() {
    var points = getTextValue("areaSparklinePoints");
    var mini = getNumberValue("areaSparklineMinimumValue");
    var maxi = getNumberValue("areaSparklineMaximumValue");
    var line1 = getNumberValue("areaSparklineLine1");
    var line2 = getNumberValue("areaSparklineLine2");
    var colorPositive = "\"" + getBackgroundColor("areaSparklinePositiveColor") + "\"";
    var colorNegative = "\"" + getBackgroundColor("areaSparklineNegativeColor") + "\"";
    var paramArr = [points, mini, maxi, line1, line2, colorPositive, colorNegative];
    var formula = getFormula(paramArr);
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyBoxPlotSparklineSetting() {
    var pointsValue = getTextValue("boxplotSparklinePoints");
    var boxPlotClassValue = getDropDownValue("boxplotClassType");
    var showAverageValue = getCheckValue("boxplotSparklineShowAverage");
    var scaleStartValue = getTextValue("boxplotSparklineScaleStart");
    var scaleEndValue = getTextValue("boxplotSparklineScaleEnd");
    var acceptableStartValue = getTextValue("boxplotSparklineAcceptableStart");
    var acceptableEndValue = getTextValue("boxplotSparklineAcceptableEnd");
    var colorValue = getBackgroundColor("boxplotSparklineColorScheme");
    var styleValue = getDropDownValue("boxplotSparklineStyleType");
    var verticalValue = getCheckValue("boxplotSparklineVertical");

    var boxplotClassStr = boxPlotClassValue ? "\"" + boxPlotClassValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var paraPool = [
        pointsValue,
        boxplotClassStr,
        showAverageValue,
        scaleStartValue,
        scaleEndValue,
        acceptableStartValue,
        acceptableEndValue,
        colorStr,
        styleValue,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BOXPLOTSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyBulletSparklineSetting() {
    var measureValue = getTextValue("bulletSparklineMeasure");
    var targetValue = getTextValue("bulletSparklineTarget");
    var maxiValue = getTextValue("bulletSparklineMaxi");
    var goodValue = getTextValue("bulletSparklineGood");
    var badValue = getTextValue("bulletSparklineBad");
    var forecastValue = getTextValue("bulletSparklineForecast");
    var tickunitValue = getTextValue("bulletSparklineTickUnit");
    var colorSchemeValue = getBackgroundColor("bulletSparklineColorScheme");
    var verticalValue = getCheckValue("bulletSparklineVertical");

    var colorSchemeString = colorSchemeValue ? "\"" + colorSchemeValue + "\"" : null;
    var paraPool = [
        measureValue,
        targetValue,
        maxiValue,
        goodValue,
        badValue,
        forecastValue,
        tickunitValue,
        colorSchemeString,
        verticalValue
    ];

    var params = formatFormula(paraPool);
    var formula = "=BULLETSPARKLINE(" + params + ")";
    var sheet = spread.getActiveSheet();
    var sels = sheet.getSelections();
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCascadeSparklineSetting() {
    var pointsRangeValue = getTextValue("cascadeSparklinePointsRange");
    var pointIndexValue = getTextValue("cascadeSparklinePointIndex");
    var labelsRangeValue = getTextValue("cascadeSparklineLabelsRange");
    var minimumValue = getTextValue("cascadeSparklineMinimum");
    var maximumValue = getTextValue("cascadeSparklineMaximum");
    var colorPositiveValue = getBackgroundColor("cascadeSparklinePositiveColor");
    var colorNegativeValue = getBackgroundColor("cascadeSparklineNegativeColor");
    var verticalValue = getCheckValue("cascadeSparklineVertical");

    var colorPositiveStr = colorPositiveValue ? "\"" + colorPositiveValue + "\"" : null;
    var colorNegativeStr = colorNegativeValue ? "\"" + colorNegativeValue + "\"" : null;
    paraPool = [
        pointsRangeValue,
        pointIndexValue,
        labelsRangeValue,
        minimumValue,
        maximumValue,
        colorPositiveStr,
        colorNegativeStr,
        verticalValue
    ];

    var params = formatFormula(paraPool);
    var formula = "=CASCADESPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCompatibleSparklineSetting() {
    var data = getTextValue("compatibleSparklineData");
    var dataOrientation = getDropDownValue("dataOrientationType");
    var dateAxisData = getTextValue("compatibleSparklineDateAxisData");
    var dateAxisOrientation = getDropDownValue("dateAxisOrientationType");
    if (dateAxisOrientation === undefined) {
        dateAxisOrientation = "";
    }

    var sparklineSetting = {}, minAxisType, maxAxisType;
    sparklineSetting.displayEmptyCellsAs = getDropDownValue("emptyCellDisplayType");
    sparklineSetting.displayHidden = getCheckValue("showDataInHiddenRowOrColumn");
    sparklineSetting.showFirst = getCheckValue("compatibleSparklineShowFirst");
    sparklineSetting.showLast = getCheckValue("compatibleSparklineShowLast");
    sparklineSetting.showHigh = getCheckValue("compatibleSparklineShowHigh");
    sparklineSetting.showLow = getCheckValue("compatibleSparklineShowLow");
    sparklineSetting.showNegative = getCheckValue("compatibleSparklineShowNegative");
    sparklineSetting.showMarkers = getCheckValue("compatibleSparklineShowMarkers");
    minAxisType = getDropDownValue("minAxisType");
    sparklineSetting.minAxisType = Sparklines.SparklineAxisMinMax[minAxisType];
    sparklineSetting.manualMin = getTextValue("manualMin");
    maxAxisType = getDropDownValue("maxAxisType");
    sparklineSetting.maxAxisType = Sparklines.SparklineAxisMinMax[maxAxisType];
    sparklineSetting.manualMax = getTextValue("manualMax");
    sparklineSetting.rightToLeft = getCheckValue("rightToLeft");
    sparklineSetting.displayXAxis = getCheckValue("displayXAxis");

    sparklineSetting.negativeColor = getBackgroundColor("compatibleSparklineNegativeColor");
    sparklineSetting.markersColor = getBackgroundColor("compatibleSparklineMarkersColor");
    sparklineSetting.axisColor = getBackgroundColor("compatibleSparklineAxisColor");
    sparklineSetting.seriesColor = getBackgroundColor("compatibleSparklineSeriesColor");
    sparklineSetting.highMarkerColor = getBackgroundColor("compatibleSparklineHighMarkerColor");
    sparklineSetting.lowMarkerColor = getBackgroundColor("compatibleSparklineLowMarkerColor");
    sparklineSetting.firstMarkerColor = getBackgroundColor("compatibleSparklineFirstMarkerColor");
    sparklineSetting.lastMarkerColor = getBackgroundColor("compatibleSparklineLastMarkerColor");
    sparklineSetting.lineWeight = getTextValue("compatibleSparklineLastLineWeight");

    var settingArray = [];
    for (var item in sparklineSetting) {
        if (sparklineSetting[item] !== undefined && sparklineSetting[item] !== "") {
            settingArray.push(item + ":" + sparklineSetting[item]);
        }
    }
    var settingString = "";
    if (settingArray.length > 0) {
        settingString = "\"{" + settingArray.join(",") + "}\"";
    }

    var formula = "";
    if (settingString !== "") {
        formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
            "," + dateAxisData + "," + dateAxisOrientation + "," + settingString + ")";
    }
    else {
        if (dateAxisOrientation !== "") {
            formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
                "," + dateAxisData + "," + dateAxisOrientation + ")";
        }
        else {
            if (dateAxisData !== "") {
                formula = "=" + sparklineName + "(" + data + "," + dataOrientation +
                    "," + dateAxisData + ")";
            }
            else {
                formula = "=" + sparklineName + "(" + data + "," + dataOrientation + ")";
            }
        }
    }

    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyScatterSparklineSetting() {
    var paraPool = [];
    var inputList = ["scatterSparklinePoints1", "scatterSparklinePoints2", "scatterSparklineMinX", "scatterSparklineMaxX",
        "scatterSparklineMinY", "scatterSparklineMaxY", "scatterSparklineHLine", "scatterSparklineVLine",
        "scatterSparklineXMinZone", "scatterSparklineXMaxZone", "scatterSparklineYMinZone", "scatterSparklineYMaxZone"];
    for (var i = 0; i < inputList.length; i++) {
        var textValue = getTextValue(inputList[i]);
        paraPool.push(textValue);
    }
    var tags = getCheckValue("scatterSparklineTags");
    var drawSymbol = getCheckValue("scatterSparklineDrawSymbol");
    var drawLines = getCheckValue("scatterSparklineDrawLines");
    var color1 = getBackgroundColor("scatterSparklineColor1");
    var color2 = getBackgroundColor("scatterSparklineColor2");
    var dashLine = getCheckValue("scatterSparklineDashLine");

    color1 = color1 ? "\"" + color1 + "\"" : null;
    color2 = color2 ? "\"" + color2 + "\"" : null;

    paraPool.push(tags);
    paraPool.push(drawSymbol);
    paraPool.push(drawLines);
    paraPool.push(color1);
    paraPool.push(color2);
    paraPool.push(dashLine);
    var params = formatFormula(paraPool);
    var formula = "=SCATTERSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();

}

function applyHbarSparklineSetting() {
    var paraPool = [];
    var value = getTextValue("hbarSparklineValue");
    var colorScheme = getBackgroundColor("hbarSparklineColorScheme");

    colorScheme = "\"" + colorScheme + "\"";
    paraPool.push(value);
    paraPool.push(colorScheme);
    var params = formatFormula(paraPool);
    var formula = "=HBARSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyVbarSparklineSetting() {
    var paraPool = [];
    var value = getTextValue("vbarSparklineValue");
    var colorScheme = getBackgroundColor("vbarSparklineColorScheme");

    colorScheme = "\"" + colorScheme + "\"";
    paraPool.push(value);
    paraPool.push(colorScheme);
    var params = formatFormula(paraPool);
    var formula = "=VBARSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyParetoSparklineSetting() {
    var pointsRangeValue = getTextValue("paretoSparklinePoints");
    var pointIndexValue = getTextValue("paretoSparklinePointIndex");
    var colorRangeValue = getTextValue("paretoSparklineColorRange");
    var targetValue = getTextValue("paretoSparklineTarget");
    var target2Value = getTextValue("paretoSparklineTarget2");
    var highlightPositionValue = getTextValue("paretoSparklineHighlightPosition");
    var labelValue = getDropDownValue("paretoLabelType");
    var verticalValue = getCheckValue("paretoSparklineVertical");
    var paraPool = [
        pointsRangeValue,
        pointIndexValue,
        colorRangeValue,
        targetValue,
        target2Value,
        highlightPositionValue,
        labelValue,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=PARETOSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applySpreadSparklineSetting() {
    var pointsValue = getTextValue("spreadSparklinePoints");
    var showAverageValue = getCheckValue("spreadSparklineShowAverage");
    var scaleStartValue = getTextValue("spreadSparklineScaleStart");
    var scaleEndValue = getTextValue("spreadSparklineScaleEnd");
    var styleValue = getDropDownValue("spreadSparklineStyleType");
    var colorSchemeValue = getBackgroundColor("spreadSparklineColorScheme");
    var verticalValue = getCheckValue("spreadSparklineVertical");

    var colorSchemeString = colorSchemeValue ? "\"" + colorSchemeValue + "\"" : null;
    var paraPool = [
        pointsValue,
        showAverageValue,
        scaleStartValue,
        scaleEndValue,
        styleValue,
        colorSchemeString,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=SPREADSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyStackedSparklineSetting() {
    var pointsValue = getTextValue("stackedSparklinePoints");
    var colorRangeValue = getTextValue("stackedSparklineColorRange");
    var labelRangeValue = getTextValue("stackedSparklineLabelRange");
    var maximumValue = getNumberValue("stackedSparklineMaximum");
    var targetRedValue = getNumberValue("stackedSparklineTargetRed");
    var targetGreenValue = getNumberValue("stackedSparklineTargetGreen");
    var targetBlueValue = getNumberValue("stackedSparklineTargetBlue");
    var targetYellowValue = getNumberValue("stackedSparklineTargetYellow");
    var colorValue = getBackgroundColor("stackedSparklineColor");
    var highlightPositionValue = getNumberValue("stackedSparklineHighlightPosition");
    var verticalValue = getCheckValue("stackedSparklineVertical");
    var textOrientationValue = getDropDownValue("stackedSparklineTextOrientation");
    var textSizeValue = getNumberValue("stackedSparklineTextSize");

    var colorString = colorValue ? "\"" + colorValue + "\"" : null;
    var paraPool = [
        pointsValue,
        colorRangeValue,
        labelRangeValue,
        maximumValue,
        targetRedValue,
        targetGreenValue,
        targetBlueValue,
        targetYellowValue,
        colorString,
        highlightPositionValue,
        verticalValue,
        textOrientationValue,
        textSizeValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=STACKEDSPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyVariSparklineSetting() {
    var varianceValue = getTextValue("variSparklineVariance");
    var referenceValue = getTextValue("variSparklineReference");
    var miniValue = getTextValue("variSparklineMini");
    var maxiValue = getTextValue("variSparklineMaxi");
    var markValue = getTextValue("variSparklineMark");
    var tickunitValue = getTextValue("variSparklineTickUnit");
    var colorPositiveValue = getBackgroundColor("variSparklineColorPositive");
    var colorNegativeValue = getBackgroundColor("variSparklineColorNegative");
    var legendValue = getCheckValue("variSparklineLegend");
    var verticalValue = getCheckValue("variSparklineVertical");

    var colorPositiveStr = colorPositiveValue ? "\"" + colorPositiveValue + "\"" : null;
    var colorNegativeStr = colorNegativeValue ? "\"" + colorNegativeValue + "\"" : null;
    var paraPool = [
        varianceValue,
        referenceValue,
        miniValue,
        maxiValue,
        markValue,
        tickunitValue,
        legendValue,
        colorPositiveStr,
        colorNegativeStr,
        verticalValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=VARISPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyMonthSparklineSetting() {
    var dataRangeStr = getTextValue("monthSparklineData");
    var year = getTextValue("monthSparklineYear");
    var month = getTextValue("monthSparklineMonth");
    var emptyColor = getBackgroundColor("monthSparklineEmptyColor");
    var startColor = getBackgroundColor("monthSparklineStartColor");
    var middleColor = getBackgroundColor("monthSparklineMiddleColor");
    var endColor = getBackgroundColor("monthSparklineEndColor");
    var colorRangeStr = getTextValue("monthSparklineColorRange");
    var formulaStr;
    if (!colorRangeStr) {
        formulaStr = "=" + "MONTHSPARKLINE" + "(" + year + "," + month + "," + dataRangeStr + "," + parseSparklineColorOptions(emptyColor) + "," + parseSparklineColorOptions(startColor) + "," + parseSparklineColorOptions(middleColor) + "," + parseSparklineColorOptions(endColor) + ")";
    } else {
        formulaStr = "=" + "MONTHSPARKLINE" + "(" + year + "," + month + "," + dataRangeStr + "," + colorRangeStr + ")";
    }
    setFormulaSparkline(formulaStr);
    updateFormulaBar();
}

function applyYearSparklineSetting() {
    var dataRangeStr = getTextValue("yearSparklineData");
    var year = getTextValue("yearSparklineYear");
    var emptyColor = getBackgroundColor("yearSparklineEmptyColor");
    var startColor = getBackgroundColor("yearSparklineStartColor");
    var middleColor = getBackgroundColor("yearSparklineMiddleColor");
    var endColor = getBackgroundColor("yearSparklineEndColor");
    var colorRangeStr = getTextValue("yearSparklineColorRange");
    var formulaStr;
    if (!colorRangeStr) {
        formulaStr = "=" + "YEARSPARKLINE" + "(" + year + "," + dataRangeStr + "," + parseSparklineColorOptions(emptyColor) + "," + parseSparklineColorOptions(startColor) + "," + parseSparklineColorOptions(middleColor) + "," + parseSparklineColorOptions(endColor) + ")";
    } else {
        formulaStr = "=" + "YEARSPARKLINE" + "(" + year + "," + dataRangeStr + "," + colorRangeStr + ")";
    }
    setFormulaSparkline(formulaStr);
    updateFormulaBar();
}

function applyPieSparklineSetting() {
    var paraPool = [];
    var range = getTextValue("pieSparklinePercentage");
    paraPool.push(range);

    $("#pieSparklineColorContainer").find(".color-view").each(function () {
        var color = "\"" + $(this).css("background-color") + "\"";
        paraPool.push(color);
    });

    var params = formatFormula(paraPool);
    var formula = "=PIESPARKLINE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyQRCodeSparklineSetting(){
    var dataValue = getTextValue("qrCodeSparklineData");
    var colorValue = getBackgroundColor("qrCodeSparklineColor");
    var backgroundColorValue = getBackgroundColor("qrCodeSparklineBackgroundColor");
    var levelValue = getDropDownValue("errorCorrectionLevelType");
    var modelValue = getDropDownValue("qrCodeSparklineModel");
    var versionValue = getDropDownValue("qrCodeSparklineVersion");
    var maskValue = getDropDownValue("qrCodeSparklineMask");
    var connectionValue = getCheckValue("checkboxQRCodeSparklineConnection");
    var connectionNoValue = getDropDownValue("qrCodeSparklineConnectionNo");
    var charCodeValue = getTextValue("qrCodeSparklineCharCode");
    var charsetValue = getDropDownValue("qrCodeCharsetType");
    var quietZoneLeftValue = getNumberValue("qrCodeSparklineLeft");
    var quietZoneRightValue = getNumberValue("qrCodeSparklineRight");
    var quietZoneTopValue = getNumberValue("qrCodeSparklineTop");
    var quietZoneBottomValue = getNumberValue("qrCodeSparklineBottom");

    var versionStr = versionValue === "auto" ? versionValue ? "\"" + versionValue + "\"" : null : versionValue;
    var maskStr = maskValue === "auto" ? maskValue ? "\"" + maskValue + "\"" : null : maskValue;
    var levelStr = levelValue ? "\"" + levelValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var charCodeStr = charCodeValue ? "{" + charCodeValue + "}" : null;
    var charsetStr = charsetValue ? "\"" + charsetValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        levelStr,
        modelValue,
        versionStr,
        maskStr,
        connectionValue,
        connectionNoValue,
        charCodeStr,
        charsetStr,
        quietZoneLeftValue,
        quietZoneRightValue,
        quietZoneTopValue,
        quietZoneBottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_QRCODE(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyEAN8SparklineSetting(){
    var dataValue =  getTextValue("ean8SparklineData");
    var colorValue = getBackgroundColor("ean8SparklineColor");
    var backgroundColorValue = getBackgroundColor("ean8SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxEAN8SparklineShowLabel");
    var labelPositionValue = getDropDownValue("ean8SparklineLabelPosition");
    var fontFamilyValue = getDropDownText("ean8SparklineFontFamily");
    var fontStyleValue = getDropDownValue("ean8SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("ean8SparklineFontWeight");
    var textDecorationValue = getDropDownText("ean8SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("ean8SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("ean8SparklineFontSize");
    var leftValue = getNumberValue("ean8SparklineLeft");
    var rightValue = getNumberValue("ean8SparklineRight");
    var topValue = getNumberValue("ean8SparklineTop");
    var bottomValue = getNumberValue("ean8SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;

    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_EAN8(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyEAN13SparklineSetting(){
    var dataValue =  getTextValue("ean13SparklineData");
    var colorValue = getBackgroundColor("ean13SparklineColor");
    var backgroundColorValue = getBackgroundColor("ean13SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxEAN13SparklineShowLabel");
    var labelPositionValue = getDropDownValue("ean13SparklineLabelPosition");
    var addOnValue = getTextValue("ean13SparklineAddOn");
    var addOnLabelPositionValue = getDropDownValue("ean13SparklineAddOnLabelPosition");
    var fontFamilyValue = getDropDownText("ean13SparklineFontFamily");
    var fontStyleValue = getDropDownValue("ean13SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("ean13SparklineFontWeight");
    var textDecorationValue = getDropDownText("ean13SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("ean13SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("ean13SparklineFontSize");
    var leftValue = getNumberValue("ean13SparklineLeft");
    var rightValue = getNumberValue("ean13SparklineRight");
    var topValue = getNumberValue("ean13SparklineTop");
    var bottomValue = getNumberValue("ean13SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var addOnStr = addOnValue ? "\"" + addOnValue + "\"" : null;
    var addOnLabelPositionStr = addOnLabelPositionValue ? "\"" + addOnLabelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        addOnStr,
        addOnLabelPositionStr,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_EAN13(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyGS1SparklineSetting(){
    var dataValue =  getTextValue("gs1SparklineData");
    var colorValue = getBackgroundColor("gs1SparklineColor");
    var backgroundColorValue = getBackgroundColor("gs1SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxGS1SparklineShowLabel");
    var labelPositionValue = getDropDownValue("gs1SparklineLabelPosition");
    var fontFamilyValue = getDropDownText("gs1SparklineFontFamily");
    var fontStyleValue = getDropDownValue("gs1SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("gs1SparklineFontWeight");
    var textDecorationValue = getDropDownText("gs1SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("gs1SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("gs1SparklineFontSize");
    var leftValue = getNumberValue("gs1SparklineLeft");
    var rightValue = getNumberValue("gs1SparklineRight");
    var topValue = getNumberValue("gs1SparklineTop");
    var bottomValue = getNumberValue("gs1SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_GS1_128(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCodabarSparklineSetting(){
    var dataValue =  getTextValue("codabarSparklineData");
    var colorValue =  getBackgroundColor("codabarSparklineColor");
    var backgroundColorValue = getBackgroundColor("codabarSparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxCodabarSparklineShowLabel");
    var labelPositionValue = getDropDownValue("codabarSparklineLabelPosition");
    var checkDigitValue = getCheckValue("checkboxCodabarSparklineCheckDigit");
    var nwRatioValue = getDropDownValue("codabarNWRatio");
    var fontFamilyValue = getDropDownText("codabarSparklineFontFamily");
    var fontStyleValue = getDropDownValue("codabarSparklineFontStyle" );
    var fontWeightValue = getDropDownValue("codabarSparklineFontWeight");
    var textDecorationValue = getDropDownText("codabarSparklineFontTextDecoration");
    var textAlignValue = getDropDownText("codabarSparklineFontTextAlign");
    var fontSizeValue = getDropDownText("codabarSparklineFontSize");
    var leftValue = getNumberValue("codabarSparklineLeft");
    var rightValue = getNumberValue("codabarSparklineRight");
    var topValue = getNumberValue("codabarSparklineTop");
    var bottomValue = getNumberValue("codabarSparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        checkDigitValue,
        nwRatioValue,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_CODABAR(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCode93SparklineSetting(){
    var dataValue =  getTextValue("code93SparklineData");
    var colorValue =  getBackgroundColor("code93SparklineColor");
    var backgroundColorValue = getBackgroundColor("code93SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxCode93SparklineShowLabel");
    var labelPositionValue = getDropDownValue("code93SparklineLabelPosition");
    var checkDigitValue = getCheckValue("checkboxCode93SparklineCheckDigit");
    var fullASCIIValue = getCheckValue("checkCode93SparklineFullASCII");
    var fontFamilyValue = getDropDownText("code93SparklineFontFamily");
    var fontStyleValue = getDropDownValue("code93SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("code93SparklineFontWeight");
    var textDecorationValue = getDropDownText("code93SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("code93SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("code93SparklineFontSize");
    var leftValue = getNumberValue("code93SparklineLeft");
    var rightValue = getNumberValue("code93SparklineRight");
    var topValue = getNumberValue("code93SparklineTop");
    var bottomValue = getNumberValue("code93SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        checkDigitValue,
        fullASCIIValue,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_CODE93(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCode39SparklineSetting(){
    var dataValue =  getTextValue("code39SparklineData");
    var colorValue =  getBackgroundColor("code39SparklineColor");
    var backgroundColorValue = getBackgroundColor("code39SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxCode39SparklineShowLabel");
    var labelPositionValue = getDropDownValue("code39SparklineLabelPosition");
    var checkDigitValue = getCheckValue("checkboxCode39SparklineCheckDigit");
    var fullASCIIValue = getCheckValue("checkCode39SparklineFullASCII");
    var charaterValue = getCheckValue("checkboxCode39SparklineCharacter");
    var nwRatioValue = getDropDownValue("code39SparklineNWRatio");
    var fontFamilyValue = getDropDownText("code39SparklineFontFamily");
    var fontStyleValue = getDropDownValue("code39SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("code39SparklineFontWeight");
    var textDecorationValue = getDropDownText("code39SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("code39SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("code39SparklineFontSize");
    var leftValue = getNumberValue("code39SparklineLeft");
    var rightValue = getNumberValue("code39SparklineRight");
    var topValue = getNumberValue("code39SparklineTop");
    var bottomValue = getNumberValue("code39SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        charaterValue,
        checkDigitValue,
        nwRatioValue,
        fullASCIIValue,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_CODE39(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCode49SparklineSetting(){
    var dataValue =  getTextValue("code49SparklineData");
    var colorValue =  getBackgroundColor("code49SparklineColor");
    var backgroundColorValue = getBackgroundColor("code49SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxCode49SparklineShowLabel");
    var labelPositionValue = getDropDownValue("code49SparklineLabelPosition");
    var groupValue = getCheckValue("checkboxCode49SparklineGrouping");
    var groupNoValue = getNumberValue("code49SparklineGroupNo");
    var fontFamilyValue = getDropDownText("code49SparklineFontFamily");
    var fontStyleValue = getDropDownValue("code49SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("code49SparklineFontWeight");
    var textDecorationValue = getDropDownText("code49SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("code49SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("code49SparklineFontSize");
    var leftValue = getNumberValue("code49SparklineLeft");
    var rightValue = getNumberValue("code49SparklineRight");
    var topValue = getNumberValue("code49SparklineTop");
    var bottomValue = getNumberValue("code49SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        groupValue,
        groupNoValue,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_CODE49(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyCode128SparklineSetting(){
    var dataValue =  getTextValue("code128SparklineData");
    var colorValue =  getBackgroundColor("code128SparklineColor");
    var backgroundColorValue = getBackgroundColor("code128SparklineBackgroundColor");
    var showLabelValue = getCheckValue("checkboxCode128SparklineShowLabel");
    var labelPositionValue = getDropDownValue("code128SparklineLabelPosition");
    var codesetValue = getDropDownValue("code128Codeset");
    var fontFamilyValue = getDropDownText("code128SparklineFontFamily");
    var fontStyleValue = getDropDownValue("code128SparklineFontStyle" );
    var fontWeightValue = getDropDownValue("code128SparklineFontWeight");
    var textDecorationValue = getDropDownText("code128SparklineFontTextDecoration");
    var textAlignValue = getDropDownText("code128SparklineFontTextAlign");
    var fontSizeValue = getDropDownText("code128SparklineFontSize");
    var leftValue = getNumberValue("code128SparklineLeft");
    var rightValue = getNumberValue("code128SparklineRight");
    var topValue = getNumberValue("code128SparklineTop");
    var bottomValue = getNumberValue("code128SparklineBottom");

    var fontFamilyStr = fontFamilyValue ? "\"" + fontFamilyValue + "\"" : null;
    var fontStyleStr = fontStyleValue ? "\"" + fontStyleValue + "\"" : null;
    var fontWeightStr = fontWeightValue ? "\"" + fontWeightValue + "\"" : null;
    var textDecorationStr = textDecorationValue ? "\"" + textDecorationValue + "\"" : null;
    var textAlignStr = textAlignValue ? "\"" + textAlignValue + "\"" : null;
    var codesetStr = codesetValue ? "\"" + codesetValue + "\"" : null;
    var labelPositionStr = labelPositionValue ? "\"" + labelPositionValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        Number(showLabelValue),
        labelPositionStr,
        codesetStr,
        fontFamilyStr,
        fontStyleStr,
        fontWeightStr,
        textDecorationStr,
        textAlignStr,
        fontSizeValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_CODE128(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyPDFSparklineSetting(){
    var dataValue =  getTextValue("pdfSparklineData");
    var colorValue =  getBackgroundColor("pdfSparklineColor");
    var backgroundColorValue = getBackgroundColor("pdfSparklineBackgroundColor");
    var errorCorrectionLevelValue = getDropDownValue("pdfSparklineLevel");
    var rowsValue = getDropDownValue("pdfSparklineRows");
    var columnsValue = getDropDownValue("pdfSparklineColumns");
    var compactValue = getCheckValue("checkboxPDFSparklineCompact");
    var leftValue = getNumberValue("pdfSparklineLeft");
    var rightValue = getNumberValue("pdfSparklineRight");
    var topValue = getNumberValue("pdfSparklineTop");
    var bottomValue = getNumberValue("pdfSparklineBottom");

    var errorCorrectionLevelStr = errorCorrectionLevelValue === "auto" ? (errorCorrectionLevelValue ? "\"" + errorCorrectionLevelValue + "\"" : null) : errorCorrectionLevelValue;
    var rowsStr = rowsValue === "auto" ? (rowsValue ? "\"" + rowsValue + "\"" : null) : rowsValue ;
    var columnsStr = columnsValue === "auto" ? (columnsValue ? "\"" + columnsValue + "\"" : null) : columnsValue ;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        errorCorrectionLevelStr,
        rowsStr,
        columnsStr,
        compactValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_PDF417(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

function applyDataMatrixSparklineSetting(){
    var dataValue = getTextValue("dataMatrixSparklineData");
    var colorValue = getBackgroundColor("dataMatrixSparklineColor");
    var backgroundColorValue = getBackgroundColor("dataMatrixSparklineBackgroundColor");
    var eccModeValue = getDropDownValue("dataMatrixSparklineEccMode");
    var ecc200SymbolSizeValue = getTextValue("dataMatrixSparklineSize");
    var ecc200EndcodingModeValue = getTextValue("dataMatrixSparklineEndcodingMode");
    var ecc00140SymboleValue = getTextValue("dataMatrixSparklineSymbole");
    var structureNumberValue = getDropDownValue("dataMatrixSparklineStructureNumber");
    var fileIdentifierValue = getNumberValue("dataMatrixSparklineFileIdentifier");
    var structureAppendValue = getCheckValue("checkboxPDFSparklineStructureAppend");
    var leftValue = getNumberValue("dataMatrixSparklineLeft");
    var rightValue = getNumberValue("dataMatrixSparklineRight");
    var topValue = getNumberValue("dataMatrixSparklineTop");
    var bottomValue = getNumberValue("dataMatrixSparklineBottom");

    var eccModeStr = eccModeValue ? "\"" + eccModeValue + "\"" : null;
    var ecc200SymbolSizeStr = ecc200SymbolSizeValue ? "\"" + ecc200SymbolSizeValue + "\"" : null;
    var ecc200EndcodingModeStr = ecc200EndcodingModeValue ? "\"" + ecc200EndcodingModeValue + "\"" : null;
    var ecc00140SymboleStr = ecc00140SymboleValue ? "\"" + ecc00140SymboleValue + "\"" : null;
    var colorStr = colorValue ? "\"" + colorValue + "\"" : null;
    var backgroundColorStr = backgroundColorValue ? "\"" + backgroundColorValue + "\"" : null;
    var paraPool = [
        dataValue,
        colorStr,
        backgroundColorStr,
        eccModeStr,
        ecc200SymbolSizeStr,
        ecc200EndcodingModeStr,
        ecc00140SymboleStr,
        structureAppendValue,
        structureNumberValue,
        fileIdentifierValue,
        leftValue,
        rightValue,
        topValue,
        bottomValue
    ];
    var params = formatFormula(paraPool);
    var formula = "=BC_DataMatrix(" + params + ")";
    setFormulaSparkline(formula);
    updateFormulaBar();
}

// Sparkline related items (end)

// Zoom related items
function processZoomSetting(value, title) {
    if (typeof value === 'number') {
        spread.getActiveSheet().zoom(value);
    }
    else {
        //console.log("processZoomSetting not process with ", value, title);
    }
}
// Zoom related items (end)

function getResource(key) {
    key = key.replace(/\./g, "_");

    return resourceMap[key];
}

function getResourceMap(src) {
    function isObject(item) {
        return typeof item === "object";
    }

    function addResourceMap(map, obj, keys) {
        if (isObject(obj)) {
            for (var p in obj) {
                var cur = obj[p];

                addResourceMap(map, cur, keys.concat(p));
            }
        } else {
            var key = keys.join("_");
            map[key] = obj;
        }
    }

    addResourceMap(resourceMap, src, []);
}

function addShapesOnToolbar() {
    // var hiddenShapes = [
    //     'lineCallout1',
    //     'lineCallout1AccentBar',
    //     'lineCallout1NoBorder',
    //     'lineCallout1BorderandAccentBar',
    //     'actionButtonCustom', 'balloon'
    // ];

    var connectorShapes = ['noHeadStraight', 'endArrowHeadStraight', 'beginEndArrowHeadStraight', 'Elbow', 'endArrowHeadElbow', 'beginEndArrowHeadElbow']
    var basicShapes = [ 'parallelogram', 'trapezoid', 'diamond', 'octagon', 'isoscelesTriangle',
        'rightTriangle', 'oval', 'hexagon', 'cross', 'regularPentagon', 'can', 'cube', 'bevel', 'foldedCorner', 'smileyFace',
        'donut', 'noSymbol', 'blockArc', 'heart', 'lightningBolt', 'sun', 'moon', 'arc', 'doubleBracket', 'doubleBrace', 'plaque',
        'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'actionButtonHome', 'actionButtonHelp', 'actionButtonInformation',
        'actionButtonBackorPrevious', 'actionButtonForwardorNext', 'actionButtonBeginning', 'actionButtonEnd', 'actionButtonReturn',
        'actionButtonDocument', 'actionButtonSound', 'actionButtonMovie', 'diagonalStripe', 'pie', 'nonIsoscelesTrapezoid', 'decagon',
        'heptagon', 'dodecagon', 'star6Point', 'star7Point', 'star10Point', 'star12Point', 'frame', 'halfFrame', 'tear', 'chord', 'corner',
        'cornerTabs', 'squareTabs', 'plaqueTabs', 'gear6', 'gear9', 'funnel', 'pieWedge', 'cloud', 'chartX', 'chartStar', 'chartPlus', 'lineInverse'];
    var blockArrows = ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow',
        'leftRightUpArrow', 'bentArrow', 'uTurnArrow', 'leftUpArrow', 'bentUpArrow', 'curvedRightArrow',
        'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'stripedRightArrow', 'notchedRightArrow',
        'pentagon', 'chevron', 'rightArrowCallout', 'leftArrowCallout', 'upArrowCallout', 'downArrowCallout',
        'leftRightArrowCallout', 'upDownArrowCallout', 'quadArrowCallout', 'circularArrow', 'leftCircularArrow',
        'leftRightCircularArrow', 'swooshArrow'];
    var flowchart = ['flowchartProcess',
        'flowchartAlternateProcess', 'flowchartDecision', 'flowchartData', 'flowchartPredefinedProcess', 'flowchartInternalStorage',
        'flowchartDocument', 'flowchartMultidocument', 'flowchartTerminator', 'flowchartPreparation', 'flowchartManualInput',
        'flowchartManualOperation', 'flowchartConnector', 'flowchartOffpageConnector', 'flowchartCard', 'flowchartPunchedTape',
        'flowchartSummingJunction', 'flowchartOr', 'flowchartCollate', 'flowchartSort', 'flowchartExtract', 'flowchartMerge',
        'flowchartStoredData','flowchartDelay', 'flowchartSequentialAccessStorage', 'flowchartMagneticDisk', 'flowchartDirectAccessStorage',
        'flowchartDisplay', 'flowchartOfflineStorage'];
    var callOuts = ['rectangularCallout', 'roundedRectangularCallout', 'ovalCallout', 'cloudCallout', 'lineCallout2','lineCallout3',
        'lineCallout4', 'lineCallout2AccentBar', 'lineCallout3AccentBar', 'lineCallout4AccentBar', 'lineCallout2NoBorder',
        'lineCallout3NoBorder', 'lineCallout4NoBorder', 'lineCallout2BorderandAccentBar', 'lineCallout3BorderandAccentBar', 'lineCallout4BorderandAccentBar'];
    var rectangles = ['roundedRectangle', 'rectangle', 'round1Rectangle', 'round2SameRectangle', 'round2DiagRectangle', 'snipRoundRectangle',
        'snip1Rectangle', 'snip2SameRectangle', 'snip2DiagRectangle'];
    var equation = ['mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual'];
    var starsAndBanners = ['explosion1', 'explosion2', 'shape4pointStar', 'shape5pointStar', 'shape8pointStar', 'shape16pointStar', 'shape24pointStar',
        'shape32pointStar', 'upRibbon', 'downRibbon', 'curvedUpRibbon', 'curvedDownRibbon', 'leftRightRibbon', 'verticalScroll', 'horizontalScroll', 'wave',
        'doubleWave']

    var idMaps = [
        {id: 'connectorShapeTypeContainer', shapes: connectorShapes},
        {id: 'shapeRectanglesContainer', shapes: rectangles},
        {id: 'shapeBasicsContainer', shapes: basicShapes},
        {id: 'shapeBlockArrowsContainer', shapes: blockArrows},
        {id: 'shapeEquationsContainer', shapes: equation},
        {id: 'shapeFlowchartContainer',shapes: flowchart},
        {id: 'shapeStarsAndBannersContainer',shapes: starsAndBanners},
        {id: 'shapeCalloutsContainer',shapes: callOuts}
    ];

    idMaps.forEach(function(shapeIdMap) {
        var shapeHtmlStr = '';
        shapeIdMap.shapes.forEach(function(shapeName) {
            shapeHtmlStr += '<button type="button" class="btn-toolbar localize-tooltip" '
            + 'id="' + shapeName + '" '
            + 'title="' + shapeName + '">'
            + '<span class="shape-icon shape-' + shapeName + '"></span>'
            + '</button>';
        });

        $('#' + shapeIdMap.id).html(shapeHtmlStr);
    });
}

$(document).ready(function () {

    addShapesOnToolbar();

    function localizeUI() {
        function getLocalizeString(text) {
            var matchs = text.match(/(?:(@[\w\d\.]*@))/g);

            if (matchs) {
                matchs.forEach(function (item) {
                    var s = getResource(item.replace(/[@]/g, ""));
                    text = text.replace(item, s);
                });
            }

            return text;
        }

        $(".localize").each(function () {
            var text = $(this).text();

            $(this).text(getLocalizeString(text));
        });

        $(".localize-tooltip").each(function () {
            var text = $(this).prop("title");

            $(this).prop("title", getLocalizeString(text));
        });

        $(".localize-value").each(function () {
            var text = $(this).attr("value");

            $(this).attr("value", getLocalizeString(text));
        });
    }

    getResourceMap(uiResource);

    localizeUI();

    spread = new spreadNS.Workbook($("#ss")[0], {tabStripRatio: 0.88});
    excelIO = new GC.Spread.Excel.IO();
    getThemeColor();
    initSpread();
    initStatusBar();

    //Change default allowCellOverflow the same with Excel.
    spread.sheets.forEach(function (sheet) {
        sheet.options.allowCellOverflow = true;
    });

    //window resize adjust
    $(".insp-container").draggable();
    checkMediaSize();
    screenAdoption();
    var resizeTimeout = null;
    $(window).bind("resize", function () {
        if (resizeTimeout === null) {
            resizeTimeout = setTimeout(function () {
                screenAdoption();
                clearTimeout(resizeTimeout);
                resizeTimeout = null;
            }, 100);
        }
    });

    addMenu();
    doPrepareWork();

    $("ul.dropdown-menu>li>a").click(function () {
        var value = $(this).text(),
            $divhost = $(this).parents("div.btn-group"),
            groupName = $divhost.data("name"),
            sheet = spread.getActiveSheet();

        $divhost.find("button:first").text(value);

        switch (groupName) {
            case "fontname":
                setStyleFont(sheet, "font-family", false, [value], value);
                break;

            case "fontsize":
                setStyleFont(sheet, "font-size", false, [value], value);
                break;
        }
    });

    var toolbarHeight = $("#toolbar").height(),
        formulaboxDefaultHeight = $("#formulabox").outerHeight(true),
        verticalSplitterOriginalTop = formulaboxDefaultHeight - $("#verticalSplitter").height();
    $("#verticalSplitter").draggable({
        axis: "y",              // vertical only
        containment: "#inner-content-container",  // limit in specified range
        scroll: false,          // not allow container scroll
        zIndex: 100,            // set to move on top
        stop: function (event, ui) {
            var $this = $(this),
                top = $this.offset().top,
                offset = top - toolbarHeight - verticalSplitterOriginalTop;

            // limit min size
            if (offset < 0) {
                offset = 0;
            }
            // adjust size of related items
            $("#formulabox").css({height: formulaboxDefaultHeight + offset});
            var height = $("div.insp-container").height() - $("#formulabox").outerHeight(true);
            $("#controlPanel").height(height);
            $("#ss").height(height);
            spread.refresh();
            // reset
            $(this).css({top: 0});
        }
    });

    attachEvents();

    $("#download").on("click", function (e) {
        e.preventDefault();
        return false;
    });

    spread.focus();

    syncSheetPropertyValues();
    syncSpreadPropertyValues();

    onCellSelected();

    updatePositionBox(spread.getActiveSheet());

    //fix bug 220484
    if (isIE) {
        $("#formulabox").css('padding', 0);
    }


    window.richEditor.init({
        element: document.getElementById('richEditor'),
        defaultParagraphSeparator: defaultParagraphSeparator,
        styleWithCSS: false,
        onChange:function () {
            document.getElementById('richTextResult').innerText = JSON.stringify(getRichText());
        }
    });
});

function initStatusBar() {
    var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(document.getElementById('statusBar'));
    statusBar.bind(spread);
}

function getHitTest(pageX, pageY, sheet) {
    var offset = $("#ss").offset(),
        x = pageX - offset.left,
        y = pageY - offset.top;
    return sheet.hitTest(x, y);
}

// import / export related items
function processExportAction($dropdown, action) {
    switch (action) {
        case "exportJson":
            exportToJSON();
            break;
        case "exportExcel":
            exportToExcel();
            break;
        default:
            break;
    }
    hideExportActionDropDown();
}

function importFile(file) {
    var fileName = file.name;
    var index = fileName.lastIndexOf('.');
    var fileExt = fileName.substr(index + 1).toLowerCase();
    if (fileExt === 'json' || fileExt === 'ssjson') {
        importSpreadFromJSON(file);
    } else if (fileExt === 'xlsx') {
        importSpreadFromExcel(file);
    } else {
        alert(getResource("messages.invalidImportFile"));
    }
}

function importSpreadFromExcel(file, options) {
    function processPasswordDialog() {
        importSpreadFromExcel(file, {password: getTextValue("txtPassword")});
        setTextValue("txtPassword", "");
    }

    var PASSWORD_DIALOG_WIDTH = 300;
    excelIO.open(file, function (json) {
        importJson(json);
    }, function (e) {
        if (e.errorCode === 0 || e.errorCode === 1) {
            alert(getResource("messages.invalidImportFile"));
        } else if (e.errorCode === 2) {
            $("#passwordError").hide();
            showModal(uiResource.passwordDialog.title, PASSWORD_DIALOG_WIDTH, $("#passwordDialog").children(), processPasswordDialog);
        } else if (e.errorCode === 3) {
            $("#passwordError").show();
            showModal(uiResource.passwordDialog.title, PASSWORD_DIALOG_WIDTH, $("#passwordDialog").children(), processPasswordDialog);
        }
    }, options);
}

function importSpreadFromJSON(file) {
    function importSuccessCallback(responseText) {
        var spreadJson = JSON.parse(responseText);
        importJson(spreadJson);
    }

    var reader = new FileReader();
    reader.onload = function () {
        importSuccessCallback(this.result);
    };
    reader.readAsText(file);
    return true;
}

function importJson(spreadJson) {
    function updateActiveCells() {
        for (var i = 0; i < spread.getSheetCount(); i++) {
            var sheet = spread.getSheet(i);
            columnIndex = sheet.getActiveColumnIndex(),
                rowIndex = sheet.getActiveRowIndex();
            if (columnIndex !== undefined && rowIndex !== undefined) {
                spread.getSheet(i).setActiveCell(rowIndex, columnIndex);
            } else {
                spread.getSheet(i).setActiveCell(0, 0);
            }
        }
    }

    if (spreadJson.version && spreadJson.sheets) {
        spread.unbindAll();
        spread.fromJSON(spreadJson);
        attachSpreadEvents(true);
        updateActiveCells();
        spread.focus();
        fbx.workbook(spread);
        onCellSelected();
        syncSpreadPropertyValues();
        syncSheetPropertyValues();
    } else {
        alert(getResource("messages.invalidImportFile"));
    }
}

function getFileName() {
    function to2DigitsString(num) {
        return ("0" + num).substr(-2);
    }

    var date = new Date();
    return [
        "export",
        date.getFullYear(), to2DigitsString(date.getMonth() + 1), to2DigitsString(date.getDate()),
        to2DigitsString(date.getHours()), to2DigitsString(date.getMinutes()), to2DigitsString(date.getSeconds())
    ].join("");
}

function exportToJSON() {
    var json = spread.toJSON({includeBindingSource: true}),
        text = JSON.stringify(json);
    var fileName = getFileName();
    if (isSafari) {
        showModal(uiResource.toolBar.downloadTitle, DOWNLOAD_DIALOG_WIDTH, $("#downloadDialog").children(), function () {
            $("#downloadDialog").hide();
        });
        var link = $("#download");
        link[0].href = "data:text/plain;" + text;
    } else {
        saveAs(new Blob([text], {type: "text/plain;charset=utf-8"}), fileName + ".json");
    }
}

function exportToExcel() {
    var fileName = getFileName();
    var json = spread.toJSON({includeBindingSource: true});
    excelIO.save(json, function (blob) {
        if (isSafari) {
            var reader = new FileReader();
            reader.onloadend = function () {
                showModal(uiResource.toolBar.downloadTitle, DOWNLOAD_DIALOG_WIDTH, $("#downloadDialog").children(), function () {
                    $("#downloadDialog").hide();
                });
                var link = $("#download");
                link[0].href = reader.result;
            };
            reader.readAsDataURL(blob);
        } else {
            saveAs(blob, fileName + ".xlsx");
        }
    }, function (e) {
        alert(e);
    });
}

// import / export related items (end)

// format related items
function processFormatSetting(name, title) {
    switch (name) {
        case "nullValue":
            name = null;
        case "0.00":
        case "$#,##0.00":
        case "_($* #,##0.00_);_($* (#,##0.00);_($* '-'??_);_(@_)":
        case "m/d/yyyy":
        case "dddd, mmmm dd, yyyy":
        case "h:mm:ss AM/PM":
        case "0%":
        case "# ?/?":
        case "0.00E+00":
        case "@":
            setFormatter(name);
            break;

        default:
            //console.log("processFormatSetting not process with ", name, title);
            break;
    }
}

function setFormatter(value) {
    var sheet = spread.getActiveSheet();
    execInSelections(sheet, "formatter", function (sheet, row, column) {
        var style = sheet.getStyle(row, column);
        if (!style) {
            style = new spreadNS.Style();
        }
        style.formatter = value;
        sheet.setStyle(row, column, style);
    });
}

function execInSelections(sheet, styleProperty, func) {
    sheet.suspendPaint();
    var selections = sheet.getSelections();
    for (var k = 0; k < selections.length; k++) {
        var selection = selections[k];
        var col = selection.col, row = selection.row,
            rowCount = selection.rowCount, colCount = selection.colCount;
        if ((col === -1 || row === -1) && styleProperty) {
            var style, r, c;
            // whole sheet was selected, need set row / column' style one by one
            if (col === -1 && row === -1) {
                for (r = 0; r < rowCount; r++) {
                    if ((style = sheet.getStyle(r, -1)) && style[styleProperty] !== undefined) {
                        func(sheet, r, -1);
                    }
                }
                for (c = 0; c < colCount; c++) {
                    if ((style = sheet.getStyle(-1, c)) && style[styleProperty] !== undefined) {
                        func(sheet, -1, c);
                    }
                }
            }
            // Get actual range for whole rows / columns / sheet selection
            if (col === -1) {
                col = 0;
            }
            if (row === -1) {
                row = 0;
            }
            // set to each cell with style that in the adjusted selection range
            for (var i = 0; i < rowCount; i++) {
                r = row + i;
                for (var j = 0; j < colCount; j++) {
                    c = col + j;
                    if ((style = sheet.getStyle(r, c)) && style[styleProperty] !== undefined) {
                        func(sheet, r, c);
                    }
                }
            }
        }
        if (selection.col == -1 && selection.row == -1) {
            func(sheet, -1, -1);
        }
        else if (selection.row == -1) {
            for (var i = 0; i < selection.colCount; i++) {
                func(sheet, -1, selection.col + i);
            }
        }
        else if (selection.col == -1) {
            for (var i = 0; i < selection.rowCount; i++) {
                func(sheet, selection.row + i, -1);
            }
        }
        else {
            for (var i = 0; i < selection.rowCount; i++) {
                for (var j = 0; j < selection.colCount; j++) {
                    func(sheet, selection.row + i, selection.col + j);
                }
            }
        }
    }
    sheet.resumePaint();
}

function convertRichText2HTML(richTextObj, $container) {
    var textDecorationType = GC.Spread.Sheets.TextDecorationType;
    var vertAlign = GC.Spread.Sheets.VertAlign;
    var texts = richTextObj.richText;
    var placeholder = '%placeholder%';

    var _innerElement = function (_htmlStr, eleName, style) {
        style = style ? ' ' + style : '';
        var eleHtml = '<' + eleName + style + '>' + placeholder + '<' + eleName + '/>'
        return _htmlStr.replace(placeholder, eleHtml);
    }

    texts.forEach(function (text) {
        var ele = document.createElement('span');
        var $ele = $(ele)[0];
        var eleStyle = text.style;
        $ele.style.color = eleStyle.color || eleStyle.foreColor || $ele.style.color;
        $ele.style.font = eleStyle.font || $ele.style.font;
        var fontValues = eleStyle.font.split(' ');
        var fontSize = 0;
        fontValues.some(function(_fValue) {
            if(_fValue.indexOf('px') >= 0) {
                fontSize = parseFloat(_fValue.substring(0, _fValue.indexOf('px')));
                return true;
            } else {
                return false;
            }
        });

        var htmlStr = placeholder;
        if (eleStyle.textDecoration === textDecorationType.underline) {
            htmlStr = _innerElement(htmlStr, 'u');
        }
        if (eleStyle.textDecoration === textDecorationType.lineThrough) {
            htmlStr = _innerElement(htmlStr, 'strike');
        }
        if (eleStyle.textDecoration === 3) {
            htmlStr = _innerElement(htmlStr, 'u');
            htmlStr = _innerElement(htmlStr, 'strike');
        }
        if (eleStyle.vertAlign === vertAlign.subscript) {
            var originFontsize = fontSize / 0.75;
            htmlStr = _innerElement(htmlStr, 'span', 'style="font-size:' + originFontsize + 'px"');
            htmlStr = _innerElement(htmlStr, 'sub');
        } else if (eleStyle.vertAlign === vertAlign.superscript) {
            var originFontsize = fontSize / 0.75;
            htmlStr = _innerElement(htmlStr, 'span', 'style="font-size:' + originFontsize + 'px"');
            htmlStr = _innerElement(htmlStr, 'sup');
        }
        htmlStr = htmlStr.replace(placeholder, text.text);
        $ele.innerHTML = htmlStr;
        $container.append($ele);
    });
}


// format related items (end)

// dialog related items
function showModal(title, width, content, callback) {
    var sheet = spread.getActiveSheet(),
        row = sheet.getActiveRowIndex(),
        col = sheet.getActiveColumnIndex();
    if(content && content.prevObject && content.prevObject.selector === '#richtextdialog') {
        var container = $(".rich-editor-content")[0];
        var $container = $(container)
        $container.text('');
        var richTextObj = sheet.getValue(row, col, GC.Spread.Sheets.SheetArea.viewport, GC.Spread.Sheets.ValueType.richText);
        var ele, $ele;
        if(richTextObj && richTextObj.text) {
            convertRichText2HTML(richTextObj, $container);
        } else if (richTextObj) {
            ele = document.createElement('span');
            $ele = $(ele)[0];
            $ele.innerText = richTextObj;
            $container.append($ele);
        }
    }

    var $dialog = $("#modalTemplate"),
        $body = $(".modal-body", $dialog);

    $(".modal-title", $dialog).text(title);
    $dialog.data("content-parent", content.parent());
    $body.append(content);

    // remove old and add new event handler since this modal is common used (reused)
    $("#dialogConfirm").off("click");
    $("#dialogConfirm").on("click", function () {
        var result = callback();

        // return an object with  { canceled: true } to tell not close the modal, otherwise close the modal
        if (!(result && result.canceled)) {
            $("#modalTemplate").modal("hide");
        }
    });

    if (!$dialog.data("event-attached")) {
        $dialog.on("hidden.bs.modal", function () {
            var $originalParent = $(this).data("content-parent");
            if ($originalParent) {
                $originalParent.append($(".modal-body", this).children());
            }
        });
        $dialog.data("event-attached", true);
    }

    // set width of the dialog
    $(".modal-dialog", $dialog).css({width: width});

    $dialog.modal("show");
}

// dialog related items (end)

// clear related items
function processClearAction($dropdown, action) {
    switch (action) {
        case "clearAll":
            doClear(255, true);   // Laze mark all types with 255 (0xFF)
            break;
        case "clearFormat":
            doClear(spreadNS.StorageType.style, true);
            break;
        default:
            break;
    }
    hideClearActionDropDown();
}

function clearSpansInSelection(sheet, selection) {
    if (sheet && selection) {
        var ranges = [],
            row = selection.row, col = selection.col,
            rowCount = selection.rowCount, colCount = selection.colCount;

        sheet.getSpans().forEach(function (range) {
            if (range.intersect(row, col, rowCount, colCount)) {
                ranges.push(range);
            }
        });
        ranges.forEach(function (range) {
            sheet.removeSpan(range.row, range.col);
        });
    }
}

function doClear(types, clearSpans) {
    var sheet = spread.getActiveSheet(),
        selections = sheet.getSelections();

    selections.forEach(function (selection) {
        sheet.clear(selection.row, selection.col, selection.rowCount, selection.colCount, spreadNS.SheetArea.viewport, types);
        if (clearSpans) {
            clearSpansInSelection(sheet, selection);
        }
    });
}

// clear related items (end)

// positionbox related items
function getSelectedRangeString(sheet, range) {
    var selectionInfo = "",
        rowCount = range.rowCount,
        columnCount = range.colCount,
        startRow = range.row + 1,
        startColumn = range.col + 1;

    if (rowCount == 1 && columnCount == 1) {
        selectionInfo = getCellPositionString(sheet, startRow, startColumn);
    }
    else {
        if (rowCount < 0 && columnCount > 0) {
            selectionInfo = columnCount + "C";
        }
        else if (columnCount < 0 && rowCount > 0) {
            selectionInfo = rowCount + "R";
        }
        else if (rowCount < 0 && columnCount < 0) {
            selectionInfo = sheet.getRowCount() + "R x " + sheet.getColumnCount() + "C";
        }
        else {
            selectionInfo = rowCount + "R x " + columnCount + "C";
        }
    }
    return selectionInfo;
}

function getCellPositionString(sheet, row, column) {
    if (row < 1 || column < 1) {
        return null;
    }
    else {
        var letters = "";
        switch (spread.options.referenceStyle) {
            case spreadNS.ReferenceStyle.a1: // 0
                while (column > 0) {
                    var num = column % 26;
                    if (num === 0) {
                        letters = "Z" + letters;
                        column--;
                    }
                    else {
                        letters = String.fromCharCode('A'.charCodeAt(0) + num - 1) + letters;
                    }
                    column = parseInt((column / 26).toString());
                }
                letters += row.toString();
                break;
            case spreadNS.ReferenceStyle.r1c1: // 1
                letters = "R" + row.toString() + "C" + column.toString();
                break;
            default:
                break;
        }
        return letters;
    }
}

// positionbox related items (end)

// theme color related items
function setThemeColorToSheet(sheet) {
    sheet.suspendPaint();

    sheet.getCell(2, 3).text("Background 1").themeFont("Body");
    sheet.getCell(2, 4).text("Text 1").themeFont("Body");
    sheet.getCell(2, 5).text("Background 2").themeFont("Body");
    sheet.getCell(2, 6).text("Text 2").themeFont("Body");
    sheet.getCell(2, 7).text("Accent 1").themeFont("Body");
    sheet.getCell(2, 8).text("Accent 2").themeFont("Body");
    sheet.getCell(2, 9).text("Accent 3").themeFont("Body");
    sheet.getCell(2, 10).text("Accent 4").themeFont("Body");
    sheet.getCell(2, 11).text("Accent 5").themeFont("Body");
    sheet.getCell(2, 12).text("Accent 6").themeFont("Body");

    sheet.getCell(4, 1).value("100").themeFont("Body");

    sheet.getCell(4, 3).backColor("Background 1");
    sheet.getCell(4, 4).backColor("Text 1");
    sheet.getCell(4, 5).backColor("Background 2");
    sheet.getCell(4, 6).backColor("Text 2");
    sheet.getCell(4, 7).backColor("Accent 1");
    sheet.getCell(4, 8).backColor("Accent 2");
    sheet.getCell(4, 9).backColor("Accent 3");
    sheet.getCell(4, 10).backColor("Accent 4");
    sheet.getCell(4, 11).backColor("Accent 5");
    sheet.getCell(4, 12).backColor("Accent 6");

    sheet.getCell(5, 1).value("80").themeFont("Body");

    sheet.getCell(5, 3).backColor("Background 1 80");
    sheet.getCell(5, 4).backColor("Text 1 80");
    sheet.getCell(5, 5).backColor("Background 2 80");
    sheet.getCell(5, 6).backColor("Text 2 80");
    sheet.getCell(5, 7).backColor("Accent 1 80");
    sheet.getCell(5, 8).backColor("Accent 2 80");
    sheet.getCell(5, 9).backColor("Accent 3 80");
    sheet.getCell(5, 10).backColor("Accent 4 80");
    sheet.getCell(5, 11).backColor("Accent 5 80");
    sheet.getCell(5, 12).backColor("Accent 6 80");

    sheet.getCell(6, 1).value("60").themeFont("Body");

    sheet.getCell(6, 3).backColor("Background 1 60");
    sheet.getCell(6, 4).backColor("Text 1 60");
    sheet.getCell(6, 5).backColor("Background 2 60");
    sheet.getCell(6, 6).backColor("Text 2 60");
    sheet.getCell(6, 7).backColor("Accent 1 60");
    sheet.getCell(6, 8).backColor("Accent 2 60");
    sheet.getCell(6, 9).backColor("Accent 3 60");
    sheet.getCell(6, 10).backColor("Accent 4 60");
    sheet.getCell(6, 11).backColor("Accent 5 60");
    sheet.getCell(6, 12).backColor("Accent 6 60");

    sheet.getCell(7, 1).value("40").themeFont("Body");

    sheet.getCell(7, 3).backColor("Background 1 40");
    sheet.getCell(7, 4).backColor("Text 1 40");
    sheet.getCell(7, 5).backColor("Background 2 40");
    sheet.getCell(7, 6).backColor("Text 2 40");
    sheet.getCell(7, 7).backColor("Accent 1 40");
    sheet.getCell(7, 8).backColor("Accent 2 40");
    sheet.getCell(7, 9).backColor("Accent 3 40");
    sheet.getCell(7, 10).backColor("Accent 4 40");
    sheet.getCell(7, 11).backColor("Accent 5 40");
    sheet.getCell(7, 12).backColor("Accent 6 40");

    sheet.getCell(8, 1).value("-25").themeFont("Body");

    sheet.getCell(8, 3).backColor("Background 1 -25");
    sheet.getCell(8, 4).backColor("Text 1 -25");
    sheet.getCell(8, 5).backColor("Background 2 -25");
    sheet.getCell(8, 6).backColor("Text 2 -25");
    sheet.getCell(8, 7).backColor("Accent 1 -25");
    sheet.getCell(8, 8).backColor("Accent 2 -25");
    sheet.getCell(8, 9).backColor("Accent 3 -25");
    sheet.getCell(8, 10).backColor("Accent 4 -25");
    sheet.getCell(8, 11).backColor("Accent 5 -25");
    sheet.getCell(8, 12).backColor("Accent 6 -25");

    sheet.getCell(9, 1).value("-50").themeFont("Body");

    sheet.getCell(9, 3).backColor("Background 1 -50");
    sheet.getCell(9, 4).backColor("Text 1 -50");
    sheet.getCell(9, 5).backColor("Background 2 -50");
    sheet.getCell(9, 6).backColor("Text 2 -50");
    sheet.getCell(9, 7).backColor("Accent 1 -50");
    sheet.getCell(9, 8).backColor("Accent 2 -50");
    sheet.getCell(9, 9).backColor("Accent 3 -50");
    sheet.getCell(9, 10).backColor("Accent 4 -50");
    sheet.getCell(9, 11).backColor("Accent 5 -50");
    sheet.getCell(9, 12).backColor("Accent 6 -50");
    sheet.resumePaint();
}

function getColorName(sheet, row, col) {
    var colName = sheet.getCell(2, col).text();
    var rowName = sheet.getCell(row, 1).text();
    return colName + " " + rowName;
}

function getThemeColor() {
    var sheet = spread.getActiveSheet();
    setThemeColorToSheet(sheet);                                            // Set current theme color to sheet

    var $colorUl = $("#default-theme-color");
    var $themeColorLi, cellBackColor;
    for (var col = 3; col < 13; col++) {
        var row = 4;
        cellBackColor = sheet.getActualStyle(row, col).backColor;
        $themeColorLi = $("<li class=\"color-cell seed-color-column\"></li>");
        $themeColorLi.css("background-color", cellBackColor).attr("data-name", sheet.getCell(2, col).text()).appendTo($colorUl);
        for (row = 5; row < 10; row++) {
            cellBackColor = sheet.getActualStyle(row, col).backColor;
            $themeColorLi = $("<li class=\"color-cell\"></li>");
            $themeColorLi.css("background-color", cellBackColor).attr("data-name", getColorName(sheet, row, col)).appendTo($colorUl);
        }
    }

    sheet.clear(2, 1, 8, 12, spreadNS.SheetArea.viewport, 255);      // Clear sheet theme color
}

// theme color related items (end)

// slicer related items
function processAddSlicer() {
    addTableColumns();                          // get table header data from table, and add them to slicer dialog

    var SLICER_DIALOG_WIDTH = 230;              // slicer dialog width
    showModal(uiResource.slicerDialog.insertSlicer, SLICER_DIALOG_WIDTH, $("#insertslicerdialog").children(), addSlicerEvent);
}

function addTableColumns() {
    var table = _activeTable;
    if (!table) {
        return;
    }
    var $slicerContainer = $("#slicer-container");
    $slicerContainer.empty();
    for (var col = 0; col < table.range().colCount; col++) {
        var columnName = table.getColumnName(col);
        var $slicerDiv = $(
            "<div>"
            + "<div class='insp-row'>"
            + "<div>"
            + "<div class='insp-checkbox insp-inline-row'>"
            + "<div class='button insp-inline-row-item'></div>"
            + "<div class='text insp-inline-row-item localize'>" + columnName + "</div>"
            + "</div>"
            + "</div>"
            + "</div>"
            + "</div>");
        $slicerDiv.appendTo($slicerContainer);
    }
    $("#slicer-container .insp-checkbox").click(checkedChanged);
}

function getSlicerName(sheet, columnName) {
    var autoID = 1;
    var newName = columnName;
    while (sheet.slicers.get(newName)) {
        newName = columnName + '_' + autoID;
        autoID++;
    }
    return newName;
}

function addSlicerEvent() {
    var table = _activeTable;
    if (!table) {
        return;
    }
    var checkedColumnIndexArray = [];
    $("#slicer-container div.button").each(function (index) {
        if ($(this).hasClass("checked")) {
            checkedColumnIndexArray.push(index);
        }
    });
    var sheet = spread.getActiveSheet();
    var posX = 100, posY = 200;
    spread.suspendPaint();
    for (var i = 0; i < checkedColumnIndexArray.length; i++) {
        var columnName = table.getColumnName(checkedColumnIndexArray[i]);
        var slicerName = getSlicerName(sheet, columnName);
        var slicer = sheet.slicers.add(slicerName, table.name(), columnName);
        slicer.position(new spreadNS.Point(posX, posY));
        posX = posX + 30;
        posY = posY + 30;
    }
    spread.resumePaint();
    slicer.isSelected(true);
    initSlicerTab();
}

function bindSlicerEvents(sheet, slicer, propertyName) {
    if (!slicer) {
        return;
    }
    if (propertyName === "isSelected") {
        if (slicer.isSelected()) {
            if (sheet.options.protectionOptions.allowEditObjects || !(sheet.options.isProtected && slicer.isLocked())) {
                setActiveTab("slicer");
                initSlicerTab();
            }
        }
        else {
            // setActiveTab("cell");

            // The events' execution sequence is different between V10 and V9.
            // In V9, EnterCell event will execute after SlicerChanged event. But in V10, SlicerChanged event will execute after EnterCell event.
            // So, when I move focus from table slicer to table cell, table tab will not be active.
            // In this situation, code above should be removed to make table be active.
        }
    }
    else {
        changeSlicerInfo(slicer, propertyName);
    }
}

function initSlicerTab() {
    var sheet = spread.getActiveSheet();
    var selectedSlicers = getSelectedSlicers(sheet);
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return;
    }
    if (selectedSlicers.length > 1) {
        getMultiSlicerSetting(selectedSlicers);
        setTextDisabled("slicerName", true);
    }
    else if (selectedSlicers.length === 1) {
        getSingleSlicerSetting(selectedSlicers[0]);
        setTextDisabled("slicerName", false);
    }
}

function getSingleSlicerSetting(slicer) {
    if (!slicer) {
        return;
    }
    setTextValue("slicerName", slicer.name());
    setTextValue("slicerCaptionName", slicer.captionName());
    setDropDownValue("slicerItemSorting", slicer.sortState());
    setCheckValue("displaySlicerHeader", slicer.showHeader());
    setNumberValue("slicerColumnNumber", slicer.columnCount());
    setNumberValue("slicerButtonWidth", getSlicerItemWidth(slicer.columnCount(), slicer.width()));
    setNumberValue("slicerButtonHeight", slicer.itemHeight());
    if (slicer.dynamicMove()) {
        if (slicer.dynamicSize()) {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
        }
        else {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
        }
    }
    else {
        setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
    }
    setCheckValue("lockSlicer", slicer.isLocked());
    selectedCurrentSlicerStyle(slicer);
}

function getMultiSlicerSetting(selectedSlicers) {
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return;
    }
    var slicer = selectedSlicers[0];
    var isDisplayHeader = false,
        isSameSortState = true,
        isSameCaptionName = true,
        isSameColumnCount = true,
        isSameItemHeight = true,
        isSameItemWidth = true,
        isSameLocked = true,
        isSameDynamicMove = true,
        isSameDynamicSize = true;

    var sortState = slicer.sortState(),
        captionName = slicer.captionName(),
        columnCount = slicer.columnCount(),
        itemHeight = slicer.itemHeight(),
        itemWidth = getSlicerItemWidth(columnCount, slicer.width()),
        dynamicMove = slicer.dynamicMove(),
        dynamicSize = slicer.dynamicSize();

    for (var item in selectedSlicers) {
        var slicer = selectedSlicers[item];
        isDisplayHeader = isDisplayHeader || slicer.showHeader();
        isSameLocked = isSameLocked && slicer.isLocked();
        if (slicer.sortState() !== sortState) {
            isSameSortState = false;
        }
        if (slicer.captionName() !== captionName) {
            isSameCaptionName = false;
        }
        if (slicer.columnCount() !== columnCount) {
            isSameColumnCount = false;
        }
        if (slicer.itemHeight() !== itemHeight) {
            isSameItemHeight = false;
        }
        if (getSlicerItemWidth(slicer.columnCount(), slicer.width()) !== itemWidth) {
            isSameItemWidth = false;
        }
        if (slicer.dynamicMove() !== dynamicMove) {
            isSameDynamicMove = false;
        }
        if (slicer.dynamicSize() !== dynamicSize) {
            isSameDynamicSize = false;
        }
        selectedCurrentSlicerStyle(slicer);
    }
    setTextValue("slicerName", "");
    if (isSameCaptionName) {
        setTextValue("slicerCaptionName", captionName);
    }
    else {
        setTextValue("slicerCaptionName", "");
    }
    if (isSameSortState) {
        setDropDownValue("slicerItemSorting", sortState);
    }
    else {
        setDropDownValue("slicerItemSorting", "");
    }
    setCheckValue("displaySlicerHeader", isDisplayHeader);
    if (isSameDynamicMove && isSameDynamicSize && dynamicMove) {
        if (dynamicSize) {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-size");
        }
        else {
            setRadioItemChecked("slicerMoveAndSize", "slicer-move-nosize");
        }
    }
    else {
        setRadioItemChecked("slicerMoveAndSize", "slicer-nomove-size");
    }
    if (isSameColumnCount) {
        setNumberValue("slicerColumnNumber", columnCount);
    }
    else {
        setNumberValue("slicerColumnNumber", "");
    }
    if (isSameItemHeight) {
        setNumberValue("slicerButtonHeight", Math.round(itemHeight));
    }
    else {
        setNumberValue("slicerButtonHeight", "");
    }
    if (isSameItemWidth) {
        setNumberValue("slicerButtonWidth", itemWidth);
    }
    else {
        setNumberValue("slicerButtonWidth", "");
    }
    setCheckValue("lockSlicer", isSameLocked);
}

function changeSlicerInfo(slicer, propertyName) {
    if (!slicer) {
        return;
    }
    switch (propertyName) {
        case "width":
            setNumberValue("slicerButtonWidth", getSlicerItemWidth(slicer.columnCount(), slicer.width()));
            break;
    }
}

function setSlicerSetting(property, value) {
    var sheet = spread.getActiveSheet();
    var selectedSlicers = getSelectedSlicers(sheet);
    if (!selectedSlicers || selectedSlicers.length === 0) {
        return;
    }
    for (var item in selectedSlicers) {
        setSlicerProperty(selectedSlicers[item], property, value);
    }
}

function setSlicerProperty(slicer, property, value) {
    switch (property) {
        case "name":
            var sheet = spread.getActiveSheet();
            var slicerPreName = slicer.name();
            if (!value) {
                alert(getResource("messages.invalidSlicerName"));
                setTextValue("slicerName", slicerPreName);
            }
            else if (value && value !== slicerPreName) {
                if (sheet.floatingObjects.get(value)) {
                    alert(getResource("messages.duplicatedSlicerName"));
                    setTextValue("slicerName", slicerPreName);
                }
                else {
                    slicer.name(value);
                }
            }
            break;
        case "captionName":
            slicer.captionName(value);
            break;
        case "sortState":
            slicer.sortState(value);
            break;
        case "showHeader":
            slicer.showHeader(value);
            break;
        case "columnCount":
            slicer.columnCount(value);
            break;
        case "itemHeight":
            slicer.itemHeight(value);
            break;
        case "itemWidth":
            slicer.width(getSlicerWidthFromItem(slicer.columnCount(), value));
            break;
        case "moveSize":
            if (value === "slicer-move-size") {
                slicer.dynamicMove(true);
                slicer.dynamicSize(true);
            }
            if (value === "slicer-move-nosize") {
                slicer.dynamicMove(true);
                slicer.dynamicSize(false);
            }
            if (value === "slicer-nomove-size") {
                slicer.dynamicMove(false);
                slicer.dynamicSize(false);
            }
            break;
        case "lock":
            slicer.isLocked(value);
            break;
        case "style":
            slicer.style(value);
            break;
        default:
            //console.log("Slicer doesn't have property:", property);
            break;
    }
}

function setTextDisabled(name, isDisabled) {
    var $item = $("div.insp-text[data-name='" + name + "']");
    var $input = $item.find("input");
    if (isDisabled) {
        $item.addClass("disabled");
        $input.attr("disabled", true);
    }
    else {
        $item.removeClass("disabled");
        $input.attr("disabled", false);
    }
}

function setRadioItemChecked(groupName, itemName) {
    var $radioGroup = $("div.insp-checkbox[data-name='" + groupName + "']");
    var $radioItems = $("div.radiobutton[data-name='" + itemName + "']");

    $radioGroup.find(".radiobutton").removeClass("checked");
    $radioItems.addClass("checked");
}

function getSlicerItemWidth(count, slicerWidth) {
    if (count <= 0) {
        count = 1; //Column count will be converted to 1 if it is set to 0 or negative number.
    }
    var SLICER_PADDING = 6;
    var SLICER_ITEM_SPACE = 2;
    var itemWidth = Math.round((slicerWidth - SLICER_PADDING * 2 - (count - 1) * SLICER_ITEM_SPACE) / count);
    if (itemWidth < 0) {
        return 0;
    }
    else {
        return itemWidth;
    }
}

function getSlicerWidthFromItem(count, itemWidth) {
    if (count <= 0) {
        count = 1; //Column count will be converted to 1 if it is set to 0 or negative number.
    }
    var SLICER_PADDING = 6;
    var SLICER_ITEM_SPACE = 2;
    return Math.round(itemWidth * count + (count - 1) * SLICER_ITEM_SPACE + SLICER_PADDING * 2);
}

function getSelectedSlicers(sheet) {
    if (!sheet) {
        return null;
    }
    var slicers = sheet.slicers.all();
    if (!slicers || slicers.length === 0) {
        return null;
    }
    var selectedSlicers = [];
    for (var item in slicers) {
        if (slicers[item].isSelected()) {
            selectedSlicers.push(slicers[item]);
        }
    }
    return selectedSlicers;
}

function processSlicerItemSorting(sortValue) {
    switch (sortValue) {
        case 0:
        case 1:
        case 2:
            setSlicerSetting("sortState", sortValue);
            break;

        default:
            //console.log("processSlicerItemSorting not process with ", name);
            return;
    }
}

function selectedCurrentSlicerStyle(slicer) {
    var slicerStyle = slicer.style(),
        styleName = slicerStyle && slicerStyle.name();
    $("#slicerStyles .slicer-format-item").removeClass("slicer-format-item-selected");
    styleName = styleName.split("SlicerStyle")[1];
    if (styleName) {
        $("#slicerStyles .slicer-format-item div[data-name='" + styleName.toLowerCase() + "']").parent().addClass("slicer-format-item-selected");
    }
}

function changeSlicerStyle() {
    spread.suspendPaint();

    var styleName = $(">div", this).data("name");
    setSlicerSetting("style", spreadNS.Slicers.SlicerStyles[styleName]());
    $("#slicerStyles .slicer-format-item").removeClass("slicer-format-item-selected");
    $(this).addClass("slicer-format-item-selected");

    spread.resumePaint();
}

// slicer related items (end)

// spread theme related items
function processChangeSpreadTheme(value) {
    $("link[title='spread-theme']").attr("href", value);

    setTimeout(
        function () {
            spread.refresh();
        }, 300);
}

// spread theme related items (end)

//cell label related item
function setLabelOptions(sheet, value, option) {
    var selections = sheet.getSelections(),
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();
    sheet.suspendPaint();
    for (var n = 0; n < selections.length; n++) {
        var sel = getActualCellRange(sheet, selections[n], rowCount, columnCount);
        for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
            for (var c = sel.col; c < sel.col + sel.colCount; c++) {
                var style = sheet.getStyle(r, c);
                if (!style) {
                    style = new spreadNS.Style();
                }
                if (!style.labelOptions) {
                    style.labelOptions = {};
                }
                if (option === "foreColor") {
                    style.labelOptions.foreColor = value;
                } else if (option === "margin") {
                    style.labelOptions.margin = value;
                } else if (option === "visibility") {
                    style.labelOptions.visibility = GC.Spread.Sheets.LabelVisibility[value];
                } else if (option === "alignment") {
                    style.labelOptions.alignment = GC.Spread.Sheets.LabelAlignment[value];
                }
                sheet.setStyle(r, c, style);
            }
        }
    }
    sheet.resumePaint();
}

function setWatermark(sheet, value) {
    var selections = sheet.getSelections(),
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();
    sheet.suspendPaint();
    for (var n = 0; n < selections.length; n++) {
        var sel = getActualCellRange(sheet, selections[n], rowCount, columnCount);
        for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
            for (var c = sel.col; c < sel.col + sel.colCount; c++) {
                var style = sheet.getStyle(r, c);
                if (!style) {
                    style = new spreadNS.Style();
                }
                style.watermark = value;
                sheet.setStyle(r, c, style);
            }
        }
    }
    sheet.resumePaint();
}

function setCellPadding(sheet, value) {
    var selections = sheet.getSelections(),
        rowCount = sheet.getRowCount(),
        columnCount = sheet.getColumnCount();
    sheet.suspendPaint();
    for (var n = 0; n < selections.length; n++) {
        var sel = getActualCellRange(sheet, selections[n], rowCount, columnCount);
        for (var r = sel.row; r < sel.row + sel.rowCount; r++) {
            for (var c = sel.col; c < sel.col + sel.colCount; c++) {
                var style = sheet.getStyle(r, c);
                if (!style) {
                    style = new spreadNS.Style();
                }
                style.cellPadding = value;
                sheet.setStyle(r, c, style);
            }
        }
    }
    sheet.resumePaint();
}

//cell label related item (end)




//chart sample (begin)

function createComboChart(formula,chartName,type0,type1) {
    var sheet = spread.getActiveSheet();
    var chart = sheet.charts.add( chartName, type0, 0, 100, 400, 300, formula);
    var seriesItem = chart.series().get(0);
    seriesItem.chartType = type1;
    chart.series().set(0,seriesItem);
    return chart;
}
var dataLabelPosition = GC.Spread.Sheets.Charts.DataLabelPosition;
var chartGroupItemObj = {
    ColumnGroup:
        [
            {desc:uiResource.chartDataLabels.center,key:dataLabelPosition.center},
            {desc:uiResource.chartDataLabels.insideEnd,key:dataLabelPosition.insideEnd},
            {desc:uiResource.chartDataLabels.outsideEnd,key:dataLabelPosition.outsideEnd}
        ],
    LineGroup:
        [
            {desc:uiResource.chartDataLabels.center,key:dataLabelPosition.center},
            {desc:uiResource.chartDataLabels.above,key:dataLabelPosition.above},
            {desc:uiResource.chartDataLabels.below,key:dataLabelPosition.below}
        ],
    PieGroup:
        [
            {desc:uiResource.chartDataLabels.center,key:dataLabelPosition.center},
            {desc:uiResource.chartDataLabels.insideEnd,key:dataLabelPosition.insideEnd},
            {desc:uiResource.chartDataLabels.bestFit,key:dataLabelPosition.bestFit},
            {desc:uiResource.chartDataLabels.outsideEnd,key:dataLabelPosition.outsideEnd}
        ],
    BarGroup:
        [
            {desc:uiResource.chartDataLabels.center,key:dataLabelPosition.center},
            {desc:uiResource.chartDataLabels.insideEnd,key:dataLabelPosition.insideEnd},
            {desc:uiResource.chartDataLabels.outsideEnd,key:dataLabelPosition.outsideEnd}
        ],
    AreaGroup: [
    ],
    ScatterGroup:
        [
            {desc:uiResource.chartDataLabels.center,key:dataLabelPosition.center},
            {desc:uiResource.chartDataLabels.above,key:dataLabelPosition.above},
            {desc:uiResource.chartDataLabels.below,key:dataLabelPosition.below}
        ],
    StockGroup:[
    ],
    ComboGroup: {}
};

var chartTypeDict = {
    0: {
        chartType: "combo",
            chartGroup: "ComboGroup"
    },
    1: {
        chartType: "xyScatter",
            chartGroup: "ScatterGroup"
    },
    2: {
        chartType: "radar",
            chartGroup: "RadarGroup"
    },
    3: {
        chartType: "doughnut",
            chartGroup: "PieGroup"
    },
    8: {
        chartType: "area",
            chartGroup: "AreaGroup"
    },
    9: {
        chartType: "line",
            chartGroup: "LineGroup"
    },
    10: {
        chartType: "pie",
            chartGroup: "PieGroup"
    },
    11: {
        chartType: "bubble",
            chartGroup: "ScatterGroup"
    },
    12: {
        chartType: "columnClustered",
            chartGroup: "ColumnGroup"
    },
    13: {
        chartType: "columnStacked",
            chartGroup: "ColumnGroup"
    },
    14: {
        chartType: "columnStacked100",
            chartGroup: "ColumnGroup"
    },
    18: {
        chartType: "barClustered",
            chartGroup: "BarGroup"
    },
    19: {
        chartType: "barStacked",
            chartGroup: "BarGroup"
    },
    20: {
        chartType: "barStacked100",
            chartGroup: "BarGroup"
    },
    24: {
        chartType: "lineStacked",
            chartGroup: "LineGroup"
    },
    25: {
        chartType: "lineStacked100",
            chartGroup: "LineGroup"
    },
    26: {
        chartType: "lineMarkers",
            chartGroup: "LineGroup"
    },
    27: {
        chartType: "lineMarkersStacked",
            chartGroup: "LineGroup"
    },
    28: {
        chartType: "lineMarkersStacked100",
            chartGroup: "LineGroup"
    },
    33: {
        chartType: "xyScatterSmooth",
            chartGroup: "ScatterGroup"
    },
    34: {
        chartType: "xyScatterSmoothNoMarkers",
            chartGroup: "ScatterGroup"
    },
    35: {
        chartType: "xyScatterLines",
            chartGroup: "ScatterGroup"
    },
    36: {
        chartType: "xyScatterLinesNoMarkers",
            chartGroup: "ScatterGroup"
    },
    37: {
        chartType: "areaStacked",
            chartGroup: "AreaGroup"
    },
    38: {
        chartType: "areaStacked100",
            chartGroup: "AreaGroup"
    },
    42: {
        chartType: "radarMarkers",
            chartGroup: "RadarGroup"
    },
    43: {
        chartType: "radarFilled",
            chartGroup: "RadarGroup"
    },
    49: {
        chartType: "stockHLC",
            chartGroup: "StockGroup"
    },
    50: {
        chartType: "stockOHLC",
            chartGroup: "StockGroup"
    },
    51: {
        chartType: "stockVHLC",
            chartGroup: "StockGroup"
    },
    52: {
        chartType: "stockVOHLC",
            chartGroup: "StockGroup"
    },
    57: {
        chartType: "sunburst",
        chartGroup: "TreeGroup"
    },
    58: {
        chartType: "treemap",
        chartGroup: "TreeGroup"
    }
}
function getChartGroupString (typeValue) {
    var chartTypeInfo = chartTypeDict[typeValue];
    if (chartTypeInfo && chartTypeInfo.chartGroup) {
        return chartTypeInfo.chartGroup;
    }
}
function getChartTypeString (typeValue) {
    var chartTypeInfo = chartTypeDict[typeValue];
    if (chartTypeInfo && chartTypeInfo.chartType) {
        return chartTypeInfo.chartType;
    }
}

function getActiveChart() {
    var sheet = spread.getActiveSheet();
    var activeChart = null;
    sheet.charts.all().forEach(function (chart) {
        if (chart.isSelected()) {
            activeChart = chart;
        }
    });
    return activeChart;
}

function getColorByThemeColor(themeColor) {
    var sheet = spread.getActiveSheet();
    var theme = sheet.currentTheme();
    return theme.getColor(themeColor);
}

function createSeriesListMenu(host, nameArray){
    for(var i=0;i<nameArray.length;i++){
        var $text = $("<div></div>").addClass('text localize');
        $text.attr('data-value',i);
        $text.html(nameArray[i]);

        var $menuItem = $("<div></div>").addClass('menu-item');
        $menuItem.on('click', itemSelected);
        $menuItem.append($("<div></div>").addClass('image fa fa-check'));
        $menuItem.append($text);
        $menuItem.append($("<div></div>").addClass('shortcut'));
        host.append($menuItem);
    }
}

function getSeriesNameArrayWithChart(chart) {
    var nameArray = [];
    var seriesArray = chart.series().get();
    for (var i = 0; i < seriesArray.length; i++) {
        var series = seriesArray[i];
        var sheet = spread.getActiveSheet();
        if (series.name) {
            var name = '';
            var range = spreadNS.CalcEngine.formulaToRange(sheet, series.name);
            if(range === undefined || range === null) {
                name = series.name
            }else{
                var cell = sheet.getCell(range.row, range.col);
                name = cell.value();
            }
            nameArray.push(name);
        }
    }
    return nameArray;
}

function attachChartItemEvents() {

    $("#setChartArea").click(applyChartAreaSetting);
    $("#setChartTitle").click(applyChartTitle);
    $("#setChartSeries").click(applyChartSeries);
    $("#setChartLegend").click(applyChartLegendSetting);
    $("#setChartDataLabels").click(applyChartDataLabelsSetting);
    $("#setChartAxes").click(applyChartAxesSetting);
    $("#setDataPoints").click(applyDataPointSetting);
}

function showChartPanel(chart) {
    if (chart && chart.isSelected()) {
        setActiveTab("chartEx");
        updateChartOption(chart);
    }
}

function updateChartOption(chart) {
    updateChartAreaSetting(chart);
    updateChartTitleSetting(chart);
    updateChartSeriesSetting(chart, 0);
    updateChartLegendSetting(chart);
    updateChartDataLabelsSetting(chart);
    updateChartAxesSetting(chart);
    updateChartAnimationSetting(chart);
    updateDataPointSettinig(chart);
}

function getTransparency(name){
    var chart = getActiveChart();
    var shapes = getActiveShapes();
    var axesType = getDropDownValue("chartAxieType");
    var transparency, axesTY;
    if (chart && axesType >= 0) {
        switch(axesType){
            case 0:
                axesTY = chart.axes().primaryCategory;
                break;
            case 1:
                axesTY = chart.axes().primaryValue;
                break;
            case 2:
                axesTY = chart.axes().secondaryCategory;
                break;
            case 3:
                axesTY = chart.axes().secondaryValue;
                break;
        }
    }
    switch(name){
        case 'chartTitleColor':
            transparency = chart.title().transparency;
            break;
        case 'chartSeriesColor':
            var seriesIndex = getDropDownValue("chartSeriesIndexValue");
            var seriesItem = chart.series().get(seriesIndex);
            transparency = seriesItem.backColorTransparency;
            break;
        case 'chartSeriesLineColor':
            var seriesIndex = getDropDownValue("chartSeriesIndexValue");
            var seriesItem = chart.series().get(seriesIndex);
            transparency = seriesItem.border.transparency;
            break;
        case 'chartAreaBackColor':
            transparency = chart.chartArea().backColorTransparency;
            break;
        case 'chartAreaColor':
            transparency = chart.chartArea().transparency;
            break;
        case 'legendBackColor':
            transparency = chart.legend().backColorTransparency;
            break;
        case 'legendBorderColor':
            transparency = chart.legend().borderStyle.transparency;
            break;
        case 'chartAixsColor':
            transparency = axesTY.style.transparency;
            break;
        case 'chartAixsTitleColor':
            transparency = axesTY.title.transparency ? axesTY.title.transparency : 0;
            break;
        case 'chartAixsLineColor':
            transparency = axesTY.lineStyle.transparency;
            break;
        case 'chartAixsMajorGridlineColor':
            transparency = axesTY.majorGridLine.transparency ? axesTY.majorGridLine.transparency : 0;
            break;
        case 'chartAixsMinorGridlineColor':
            transparency = axesTY.minorGridLine.transparency ? axesTY.minorGridLine.transparency : 0;
            break;
        case 'dataPointColor':
            var currentPointIndex = getDropDownValue("chartDataPointsValue");
            transparency = chart.series().dataPoints().get(currentPointIndex).transparency;
            break;
        case "shapeColor":
        case "shapeBackgroundColor":
        case "shapeBorderColor":
            transparency = getShapeTransparency(shapes, name)
            break;
        default:
            transparency = 0;
    }

    $('#colorpickerTransparency').val(transparency);
}

function getShapeTransparency(shapes, transparencyName) {
    var transparency;

    var _getTransparency = function (_shapes) {
        _shapes.some(function (_shape) {
            var _shapeType = getShapeType(_shape);

            if (_shapeType === 'shapeGroup') {
                _getTransparency(_shape.all());
            } else {
                var shapeStyle = _shape.style();
                switch (transparencyName) {
                    case "shapeColor":
                        if(shapeStyle.textEffect) {
                            transparency = shapeStyle.textEffect.transparency;
                        }
                        break;
                    case "shapeBackgroundColor":
                        if(shapeStyle.fill) {
                            transparency = shapeStyle.fill.transparency;
                        }
                        break;
                    case "shapeBorderColor":
                        transparency = shapeStyle.line.transparency;
                        break;
                }
            }
        });
    }

    _getTransparency(shapes);

    return transparency;
}

function updateChartAreaSetting(chart) {
    if (chart) {
        var chartArea = chart.chartArea();
        setColorValue("chartAreaBackColor", getRGBAColor(getColorByThemeColor(chartArea.backColor), 1 - chartArea.backColorTransparency));
        setColorValue("chartAreaColor", getRGBAColor(getColorByThemeColor(chartArea.color), 1 - chartArea.transparency));
        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAreaFontFamily']"), chartArea.fontFamily);
        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAreaFontSize']"), parseInt(chartArea.fontSize));
    }
}
function applyChartAreaSetting() {
    var chart = getActiveChart();
    if(chart){
        var fontSize = parseInt(getDropDownText("chartAreaFontSize"));
        var fontFamily = getDropDownText("chartAreaFontFamily")
        var backColor = getBackgroundColor("chartAreaBackColor");
        var color = getBackgroundColor("chartAreaColor");
        var chartArea = chart.chartArea();
        chartArea.transparency = getColorTransparency("chartAreaColor");
        chartArea.backColorTransparency = getColorTransparency("chartAreaBackColor");
        chartArea.fontSize = fontSize;
        chartArea.backColor =  backColor ;
        chartArea.color = color;
        chartArea.fontFamily = fontFamily;
        chart.chartArea(chartArea);
    }
}

function updateChartTitleSetting(chart) {
    if(chart){
        var title = chart.title();
        if(title) {
            setTextValue('chartTitletext', title.text);
            setColorValue("chartTitleColor", getRGBAColor(title.color, 1 - title.transparency));
            setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartTitleFontFamily']"), title.fontFamily);
            setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartTitleFontSize']"), parseInt(title.fontSize));
        }
    }
}
function applyChartTitle() {
    var chart = getActiveChart();
    if(chart){
        var fontSize = parseInt(getDropDownText('chartTitleFontSize'));
        var fontFamily = getDropDownText("chartTitleFontFamily")
        var text = getTextValue('chartTitletext');
        var color = getColorByThemeColor(getBackgroundColor("chartTitleColor"));
        var title = chart.title();
        title.transparency = getColorTransparency("chartTitleColor");
        title.text = text;
        title.color  = color ;
        title.fontFamily = fontFamily;
        title.fontSize = fontSize;
        chart.title(title);
    }
}

function changeSeriesIndex(seriesIndex){
    var chart = getActiveChart();
    updateChartSeriesSetting(chart,seriesIndex);
}
function updateChartSeriesSetting(chart,seriesIndex) {
    var chartGroupString = getChartGroupString(chart.chartType());
    if(chartGroupString === "StockGroup" || chartGroupString === 'PieGroup' || chartGroupString === 'TreeGroup') {
        $("#chartSeriesGroup").hide();
        return;
    }
    $("#chartSeriesGroup").show();
    var nameArray = getSeriesNameArrayWithChart(chart);
    var $host = $('#chartSeriesIndexContner');
    $host.html('');
    createSeriesListMenu($host,nameArray);
    setDropDownValue("chartSeriesIndexValue", seriesIndex);
    var seriesItem = chart.series().get(seriesIndex);
    var axisGroup = seriesItem.axisGroup.toString();
    var lineWidth = seriesItem.border.width;
    if(chartGroupString === "ScatterGroup"){
        $('#chartSeriesLineWidth').hide();
        if(chart.chartType() === 11){
            $('#chartSeriesColor').show();
            $('#chartSeriesLineColor').hide();
        } else {
            $('#chartSeriesLineColor').show();
            $('#chartSeriesColor').hide();
        }

    } else {
        $('#chartSeriesColor').show();
        $('#chartSeriesLineColor').show();
        $('#chartSeriesLineWidth').show();
    }
    var lineColor = seriesItem.border.color;
    if(chartGroupString === "ScatterGroup" && lineColor === undefined){
        lineColor = "Accent " + (seriesIndex % 6 + 1);
    }
    var lineColorByTheme = getColorByThemeColor(lineColor);
    var backColor  = getColorByThemeColor(seriesItem.backColor);
    setDropDownValue("chartSeriesGroupValue", axisGroup);
    setColorValue("chartSeriesColor", getRGBAColor(backColor, 1 - seriesItem.backColorTransparency));
    setTextValue('chartSeriesLineWidth',lineWidth);
    setColorValue("chartSeriesLineColor", getRGBAColor(lineColorByTheme, 1 - seriesItem.border.transparency));
}

function applyChartSeries() {
    var chart = getActiveChart();
    if(chart){
        var seriesIndex = getDropDownValue("chartSeriesIndexValue");
        var axisGroup = getDropDownValue("chartSeriesGroupValue");
        var seriesItem = chart.series().get(seriesIndex);
        var backColor = getBackgroundColor('chartSeriesColor');
        var linwWidth = getTextValue('chartSeriesLineWidth');
        var lineColor = getBackgroundColor('chartSeriesLineColor');
        seriesItem.backColor = backColor;
        seriesItem.axisGroup  = axisGroup;
        seriesItem.border.width = parseInt(linwWidth);
        seriesItem.border.color = lineColor;
        seriesItem.border.transparency = getColorTransparency("chartSeriesLineColor");
        seriesItem.backColorTransparency = getColorTransparency("chartSeriesColor");chartSeriesColor
        chart.series().set(seriesIndex, seriesItem);
        updateChartAxesSetting(chart);
    }
}

//

function updateChartLegendSetting(chart) {
    var chartGroupString = getChartGroupString(chart.chartType());
    // there is no legend for stock chart, need to control whether to show legend group in panel.
    if (chartGroupString === "StockGroup" || chartGroupString === "TreeGroup") {
        $('#chartLegendGroup').hide();
        return;
    }
    $('#chartLegendGroup').show();
    var legend = chart.legend();
    setCheckValue("showChartLegend", legend.visible);
    var position = legend.position.toString();
    setDropDownValue("chartLegendPosition", position);
    setColorValue("legendBackColor", getRGBAColor(legend.backColor, 1 - legend.backColorTransparency));
    if (legend.borderStyle && legend.borderStyle && legend.borderStyle.transparency) {
        setColorValue("legendBorderColor", getRGBAColor(legend.borderStyle.color, 1 - legend.borderStyle.transparency));
    }
    if (legend.borderStyle && legend.borderStyle.width) {
        setNumberValue("legendBorderWidth", legend.borderStyle.width);
    }
}

function applyChartLegendSetting() {
    var chart = getActiveChart();
    var chartGroupString = getChartGroupString(chart.chartType());
    if(chart && chartGroupString !== "StockGroup"){
        var legend = chart.legend();
        var isShowLegend = getCheckValue("showChartLegend");
        legend.visible = isShowLegend;
        var currentPosition = getDropDownValue("chartLegendPosition");
        var legendBackColor = getBackgroundColor('legendBackColor');
        var legendBorderColor = getBackgroundColor("legendBorderColor")
        legend.position = currentPosition;
        legend.backColor = legendBackColor;
        legend.backColorTransparency = getColorTransparency("legendBackColor");
        legend.borderStyle = legend.borderStyle || {};
        legend.borderStyle.color = legendBorderColor;
        legend.borderStyle.transparency = getColorTransparency("legendBorderColor");
        legend.borderStyle.width = getNumberValue("legendBorderWidth");
        chart.legend(legend);
    }
}

function getStrIndex(str,cha,num){
    var x=str.indexOf(cha);
    for(var i=0;i<num;i++){
        x=str.indexOf(cha,x+1);
    }
    return x;
}

function getChartDataLabelsDescAndKey(chart) {
    var chartGroupString = getChartGroupString(chart.chartType());
    var chartTypeString = getChartTypeString(chart.chartType());
    var dataLabelsDescArray = [];
    var dataLabelsKeyArray = [];
    if(chartTypeString === 'doughnut'){
        dataLabelsDescArray = [];
        dataLabelsKeyArray = [];
    }else if(chartGroupItemObj[chartGroupString]){
        var array = chartGroupItemObj[chartGroupString];
        for(var i=0;i<array.length;i++){
            dataLabelsDescArray.push(array[i].desc);
            dataLabelsKeyArray.push(array[i].key);
        }
    }
    return {desc:dataLabelsDescArray,key:dataLabelsKeyArray};
}

function judjeDataLabelsIsShow(isShowObj){
    var isShow;
    var chart = getActiveChart();
    if(isShowObj !== undefined && isShowObj !== null){
        var itemString = isShowObj.item;
        switch (itemString){
            case "showDataLabelsValue":
                showValue = isShowObj.isShow;
                break;
            case "showDataLabelsSeriesName":
                showSeriesName = isShowObj.isShow;
                break;
            case "showDataLabelsCategoryName":
                showCategoryName = isShowObj.isShow;
                break;
            default:
                isShow = false;
                break;
        }
    }
    isShow = showCategoryName || showValue|| showSeriesName;
    return isShow;
}
function updateDataLabelsPositionDropDown(isShow){
    var chart = getActiveChart();
    if(chart){
        var obj = getChartDataLabelsDescAndKey(chart);
        var dataLabelsKeyArray = obj.key;
        var dataLabelsDescArray = obj.desc;
        var dataLabels = chart.dataLabels();
        if(isShow){
            var position = dataLabels.position;
            //get dropDownIndex
            var index = 0;
            for(var i=0;i<dataLabelsKeyArray.length;i++){
                if(position === dataLabelsKeyArray[i]){
                    index = i;
                    break;
                }
            }
            $('#dataLabelsColorCon').show();
            //create dropDownList
            if(dataLabelsDescArray.length>0){
                $('#chartDataLabelPositionDropDown').show();
                var $host = $('#chartDataLabelList');
                $host.html('');
                createSeriesListMenu($host,dataLabelsDescArray);
                setDropDownValue("chartDataLabelPosition", index);
            }else{
                //hide dropDown
                $('#chartDataLabelPositionDropDown').hide();
            }
        }else{
            //hide
            $('#chartDataLabelPositionDropDown').hide();
            $('#dataLabelsColorCon').hide();
        }
    }
}
function updateChartDataLabelsSetting(chart) {
    var chartGroupString = getChartGroupString(chart.chartType());
    if(chartGroupString === "StockGroup" ||  chartGroupString === "TreeGroup"){
        // there is no data labels for stock chart, hide this dom in panel.
        $("#chartDataLabelsGroup").hide();
        return;
    }
    $("#chartDataLabelsGroup").show();
    var dataLabels = chart.dataLabels();
    showValue = dataLabels.showValue;
    showSeriesName = dataLabels.showSeriesName;
    showCategoryName = dataLabels.showCategoryName;

    var isShow = judjeDataLabelsIsShow();
    updateDataLabelsPositionDropDown(isShow);
    setCheckValue("showDataLabelsValue",dataLabels.showValue);
    setCheckValue("showDataLabelsSeriesName",dataLabels.showSeriesName);
    setCheckValue("showDataLabelsCategoryName",dataLabels.showCategoryName);
    setColorValue("dataLabelsColor",getColorByThemeColor(dataLabels.color||"rgb(255,255,255)"));

}

function applyChartDataLabelsSetting() {
    var chart = getActiveChart();
    if(chart){
        var dataLabels = chart.dataLabels();
        var dataLabelPositionIndex = getDropDownValue("chartDataLabelPosition");
        if(dataLabelPositionIndex !== null && dataLabelPositionIndex !== undefined) {
            var dataLabelsKeyArray = getChartDataLabelsDescAndKey(chart).key;
            var position = dataLabelsKeyArray[dataLabelPositionIndex];
            dataLabels.position = position;
        }
        var showValue = getCheckValue("showDataLabelsValue");
        var showSeriesName = getCheckValue("showDataLabelsSeriesName");
        var showCategoryName = getCheckValue("showDataLabelsCategoryName");
        var dataLabelsColor = getBackgroundColor("dataLabelsColor");
        dataLabels.color = dataLabelsColor;
        dataLabels.showValue = showValue;
        dataLabels.showSeriesName = showSeriesName;
        dataLabels.showCategoryName = showCategoryName;
        chart.dataLabels(dataLabels);
    }
}

function changeAxieTypeIndex(nameValue) {
    var chart = getActiveChart();
    if (chart) {
    var axes = chart.axes();
    switch(nameValue){
        case 0:
            axesTY = axes.primaryCategory;
            break;
        case 1:
            axesTY = axes.primaryValue;
            break;
        case 2:
            axesTY = axes.secondaryCategory;
            break;
        case 3:
            axesTY = axes.secondaryValue;
            break;
    }
    var chartType = chart.chartType();
    if(chartType !== 10 && chartType !== 3){
        var text = axesTY.title.text;
        var aixsLineWidth = axesTY.lineStyle.width;
        var aixsMajorUnit = axesTY.majorUnit || 'Auto';
        var aixsMinorUnit = axesTY.minorUnit || 'Auto';
        var aixsMajorGridlineWidth = axesTY.majorGridLine.width;
        var aixsMinorGridlineWidth = axesTY.minorGridLine.width;

        var aixsFontFamily = axesTY.style.fontFamily;
        var aixsTitleFontFamily = axesTY.title.fontFamily || '';
        var aixsTitleFontSize = axesTY.title.fontSize || '';
        var aixsFontSize = axesTY.style.fontSize;

        var showMajorGridline = axesTY.majorGridLine.visible;
        var showMinorGridline = axesTY.minorGridLine.visible;
        var showAxis = axesTY.visible;

        var aixsTitleColor = axesTY.title.color || '#999999';
        var aixsColor = axesTY.style.color || '#999999';
        var aixsLineColor = axesTY.lineStyle.color || '#999999';
        var aixsMajorGridlineColor = axesTY.majorGridLine.color || '#999999';
        var aixsMinorGridlineColor = axesTY.minorGridLine.color || '#999999';

        var aixsTickLabelPosition = axesTY.tickLabelPosition.toString();
        var aixsMajorTickPosition = axesTY.majorTickPosition.toString();
        var aixsMinorTickPosition = axesTY.minorTickPosition.toString();

        setTextValue('chartAixsTitletext', text);
        setTextValue("chartAixsLineWidth", aixsLineWidth);
        setTextValue("chartAixsMajorUnit", aixsMajorUnit);
        setTextValue("chartAixsMinorUnit", aixsMinorUnit);
        setTextValue("chartAixsMajorGridlineWidth", aixsMajorGridlineWidth);
        setTextValue("chartAixsMinorMinorGridlineWidth", aixsMinorGridlineWidth);


        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAxesFontFamily']"), aixsFontFamily);
        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAxesFontSize']"), aixsFontSize);

        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAxesTitleFontFamily']"), aixsTitleFontFamily);
        setDropDownText($("#chartExTab div.insp-dropdown-list[data-name='chartAxesTitleFontSize']"), aixsTitleFontSize);

        setCheckValue("showMajorGridline", showMajorGridline);
        setCheckValue("showMinorGridline", showMinorGridline);
        setCheckValue("showAxis", showAxis);

        setColorValue("chartAixsTitleColor", getColorByThemeColor(aixsTitleColor));
        setColorValue("chartAixsColor", getColorByThemeColor(aixsColor));
        setColorValue("chartAixsLineColor", getColorByThemeColor(aixsLineColor));
        setColorValue("chartAixsMajorGridlineColor", getColorByThemeColor(aixsMajorGridlineColor));
        setColorValue("chartAixsMinorGridlineColor", getColorByThemeColor(aixsMinorGridlineColor));

        setDropDownValue("chartTickLabelPosition", aixsTickLabelPosition);
        setDropDownValue("chartMajorTickPosition", aixsMajorTickPosition);
        setDropDownValue("chartMinorTickPosition", aixsMinorTickPosition);
    }
}


}
function updateChartAxesSetting(chart) {
    var chartName = chart.name().toLowerCase();
    setDropDownValue("chartAxieType", 1);
    var chartGroupString = getChartGroupString(chart.chartType());
    if(chartGroupString === 'PieGroup' || chartGroupString === 'TreeGroup'){
        $('#chartAxesGroup').hide();
    }else{
        $('#chartAxesGroup').show();
        var secondaryAxis = $("#chartAxisTypeList .secondary-axis");
        var secondaryValue = $("#chartAxisTypeList .secondary-value");
        var primaryCategory = $("#chartAxisTypeList .primary-category");
        if(chartGroupString === 'RadarGroup' || chartName.indexOf('radar') != -1) {
            secondaryAxis.hide();
            primaryCategory.hide();
            if(Object.keys(chart.axes()).indexOf('secondaryValue') != -1){
                secondaryValue.show();
            }else{
                secondaryValue.hide();
            }
        } else {
            secondaryAxis.show();
            secondaryValue.show();
            primaryCategory.show();
        }
        changeAxieTypeIndex(1);
    }
}
function applyChartAxesSetting() {
    var chart = getActiveChart();
    var spreadCH = GC.Spread.Sheets.Charts;
    if(chart){
        var axes = chart.axes();
        var axesType = getDropDownValue("chartAxieType");
        var text = getTextValue("chartAixsTitletext");
        var showMajorGridline = getCheckValue("showMajorGridline");
        var showMinorGridline = getCheckValue("showMinorGridline");
        var showAxis = getCheckValue("showAxis");
        var aixsTitleColor = getBackgroundColor("chartAixsTitleColor");
        var aixsTitleFontFamily = getDropDownText("chartAxesTitleFontFamily");
        var aixsTitleFontSize = getDropDownText("chartAxesTitleFontSize");
        var aixsColor = getBackgroundColor("chartAixsColor");
        var aixsFontFamily = getDropDownText("chartAxesFontFamily");
        var aixsFontSize = getDropDownText("chartAxesFontSize");
        var aixsLineColor = getBackgroundColor("chartAixsLineColor");
        var aixsLineWidth = parseInt(getTextValue("chartAixsLineWidth"));
        var aixsMajorUnit = parseInt(getTextValue("chartAixsMajorUnit"));
        var aixsMinorUnit = parseInt(getTextValue("chartAixsMinorUnit"));
        var aixsMajorGridlineWidth = parseInt(getTextValue("chartAixsMajorGridlineWidth"));
        var aixsMajorGridlineColor = getBackgroundColor("chartAixsMajorGridlineColor");
        var aixsMinorGridlineWidth = parseInt(getTextValue("chartAixsMinorMinorGridlineWidth"));
        var aixsMinorGridlineColor = getBackgroundColor("chartAixsMinorGridlineColor");
        var aixsTickLabelPosition;
        switch (getDropDownValue("chartTickLabelPosition")){
            case 3:
                aixsTickLabelPosition = spreadCH.TickLabelPosition.none;
                break;
            case 2:
                aixsTickLabelPosition = spreadCH.TickLabelPosition.nextToAxis;
                break;
        }
        var aixsMajorTickPosition;
        var aixsMinorTickPosition;
        switch(getDropDownValue("chartMajorTickPosition")){
            case 0:
                aixsMajorTickPosition = spreadCH.TickMark.cross;
                break;
            case 1:
                aixsMajorTickPosition = spreadCH.TickMark.inside;
                break;
            case 2:
                aixsMajorTickPosition = spreadCH.TickMark.none;
                break;
            case 3:
                aixsMajorTickPosition = spreadCH.TickMark.outside;
                break;
        }

        switch(getDropDownValue("chartMinorTickPosition")){
            case 0:
                aixsMinorTickPosition = spreadCH.TickMark.cross;
                break;
            case 1:
                aixsMinorTickPosition = spreadCH.TickMark.inside;
                break;
            case 2:
                aixsMinorTickPosition = spreadCH.TickMark.none;
                break;
            case 3:
                aixsMinorTickPosition = spreadCH.TickMark.outside;
                break;
        }

        var axesTY;
        switch(axesType){
            case 0:
                axesTY = axes.primaryCategory;
                break;
            case 1:
                axesTY = axes.primaryValue;
                break;
            case 2:
                axesTY = axes.secondaryCategory;
                break;
            case 3:
                axesTY = axes.secondaryValue;
                break;

        }
        axesTY.style.color = aixsColor;
        axesTY.style.transparency = getColorTransparency("chartAixsColor");
        axesTY.style.fontFamily =  aixsFontFamily;
        axesTY.style.fontSize =  aixsFontSize;
        axesTY.title.text = text;
        if(axesTY.title.text){
            axesTY.title.color = aixsTitleColor;
            axesTY.title.transparency =  1 - (getStrIndex(aixsTitleColor,',',2) === -1 ? 1 : aixsTitleColor.slice(getStrIndex(aixsTitleColor,',',2)+1,-1));
            axesTY.title.fontFamily =  aixsTitleFontFamily;
        }
        if(aixsTitleFontSize){
            axesTY.title.fontSize =  aixsTitleFontSize;
        }
        axesTY.majorGridLine.visible = showMajorGridline;
        axesTY.minorGridLine.visible = showMinorGridline;
        axesTY.minorGridLine.visible = showMinorGridline;
        axesTY.lineStyle.color = aixsLineColor;
        axesTY.lineStyle.width = aixsLineWidth;
        axesTY.lineStyle.transparency = getColorTransparency("chartAixsLineColor");
        axesTY.majorTickPosition = aixsMajorTickPosition;
        axesTY.minorTickPosition = aixsMinorTickPosition;
        axesTY.visible = showAxis;
        axesTY.majorUnit = aixsMajorUnit;
        axesTY.minorUnit = aixsMinorUnit;
        if(axesTY.majorGridLine.visible){
            axesTY.majorGridLine.width = aixsMajorGridlineWidth;
            axesTY.majorGridLine.color = aixsMajorGridlineColor;
            axesTY.majorGridLine.transparency = 1 - (getStrIndex(aixsMajorGridlineColor,',',2) === -1 ? 1 : aixsMajorGridlineColor.slice(getStrIndex(aixsMajorGridlineColor,',',2)+1,-1));
        }
        if(axesTY.minorGridLine.visible){
            axesTY.minorGridLine.width = aixsMinorGridlineWidth;
            axesTY.minorGridLine.color = aixsMinorGridlineColor;
            axesTY.minorGridLine.transparency = 1 - (getStrIndex(aixsMinorGridlineColor,',',2) === -1 ? 1 : aixsMinorGridlineColor.slice(getStrIndex(aixsMinorGridlineColor,',',2)+1,-1));
        }
        axesTY.tickLabelPosition = aixsTickLabelPosition;

        chart.axes(axes);

        changeAxieTypeIndex(axesType);
    }
}

function updateChartAnimationSetting(chart) {
    var chartGroupString = getChartGroupString(chart.chartType());
    var animationChartGroups = ["ColumnGroup", "BarGroup", "LineGroup", "PieGroup"]

    if (animationChartGroups.indexOf(chartGroupString) >= 0) {
        $("#chartOptionsGroup").show();
        setCheckValue("useChartAnimation", chart.useAnimation());
    } else {
        $("#chartOptionsGroup").hide();
    }
}

function applyChartAnimationSetting(useChartAnimation) {
    var chart = chart || getActiveChart();
    if (chart) {
        chart.useAnimation(useChartAnimation);
    }
}

function changeDataPointIndex(index) {
    var chart = getActiveChart();
    updateDataPointSettinig(chart, index);
}

function updateDataPointSettinig(chart, currentPointIndex) {
    var chartGroupString = getChartGroupString(chart.chartType());
    currentPointIndex = currentPointIndex || 0;

    if (chartGroupString === "TreeGroup") {
        $('#chartDatapointsGroup').show();
        var dataPoints = chart.series().dataPoints();
        var dataPointIndex = 0;
        var dataPointNames = [];
        while(dataPoints.get(dataPointIndex)) {
            dataPointNames.push("Data Point "+dataPointIndex);
            dataPointIndex++;
        }
        var $host = $("#chartDataPointsContainer");
        $host.html("");
        createSeriesListMenu($host, dataPointNames);
        setDropDownValue("chartDataPointsValue", currentPointIndex);
        var currentPoint = dataPoints.get(currentPointIndex);
        if(currentPoint) {
            setColorValue("dataPointColor", getRGBAColor(currentPoint.fillColor, 1 - currentPoint.transparency) );
            setNumberValue("dataPointTransparency", currentPoint.transparency);
        }
    } else {
        $('#chartDatapointsGroup').hide();
        return;
    }
}

function applyDataPointSetting() {
    var chart = getActiveChart();
    var dataPoints = chart.series().dataPoints();

    var currentPointIndex = getDropDownValue("chartDataPointsValue");
    var currentDataPoint = dataPoints.get(currentPointIndex);
    if(currentDataPoint) {
        currentDataPoint.fillColor = getBackgroundColor("dataPointColor");
        currentDataPoint.transparency = getColorTransparency("dataPointColor");
        dataPoints.set(currentPointIndex, currentDataPoint);
    }
}

function changeModelIndex(currentPointIndex){
    if(currentPointIndex === 1){
        $("#versionList .no-common").hide();
    }else{
        $("#versionList .no-common").show();
    }
}

function getActiveShapes() {
    var sheet = spread.getActiveSheet();
    var activeShapes = [];
    sheet.shapes.all().forEach(function (shape) {
        if (shape.isSelected()) {
            activeShapes.push(shape);
        }
    });
    return activeShapes;
}

function setShapeGroup(type, sheet) {
    var shapes = getActiveShapes();

    if(type === "group") {
        var shapes = getActiveShapes();
        var groupShape = sheet.shapes.group(shapes);
        groupShape.isSelected(true);
    } else {
        var childrens = shapes[0].all();
        sheet.shapes.ungroup(shapes[0]);
        childrens.forEach(function(children) {
            children.isSelected(true);
        });
    }
}

function attachShapeEvents() {
    $("#setShape").click(applyShapeSetting);
}

function showShapePanel(shape, needSet) {
    var shapes = getActiveShapes();
    if (shapes && shapes.length > 0) {
        setActiveTab("shapeEx");
        if(!needSet){
            updateShapeSetting(shapes);
        }
    }
}

function getShapeType(shape) {
    var result = 'shape';
    if(shape instanceof GC.Spread.Sheets.Shapes.GroupShape) {
        result = 'shapeGroup';
    }
    if(shape instanceof GC.Spread.Sheets.Shapes.ConnectorShape) {
        result = 'connector';
    }

    return result;
}

function getShapeBorderTypeString(type) {
    var result = '';
    for(typeString in spreadNS.Shapes.PresetLineDashStyle) {
        if(spreadNS.Shapes.PresetLineDashStyle[typeString] === type){
            result = typeString;
            break;
        }
    }
    return result;
}

function getShapeArrowString(value) {
    var result = 'none';
    for(key in GC.Spread.Sheets.Shapes.ArrowheadStyle) {
        if(GC.Spread.Sheets.Shapes.ArrowheadStyle[key] === value) {
            result = key;
        }
    }
    return result;
}

/**
 * updateShapeSetting
 * @param {*} shape
 */
function updateShapeSetting(shapes) {
    var groupCount = 0,
        shapeCount = 0,
        connectorCount = 0;

    var _setConnector = function(shape, shapeStyle) {
        setDropDownValue("shapeType", shape.type());
        setDropDownValue("beginArrowWidth", shapeStyle.line.beginArrowheadWidth);
        setDropDownValue("beginArrowHeight", shapeStyle.line.beginArrowheadLength);
        setDropDownValue("endArrowWidth", shapeStyle.line.endArrowheadWidth);
        setDropDownValue("endArrowHeight", shapeStyle.line.endArrowheadLength);
        processArrowStyleSetting('beginArrowStyle', getShapeArrowString(shapeStyle.line.beginArrowheadStyle));
        processArrowStyleSetting('endArrowStyle', getShapeArrowString(shapeStyle.line.endArrowheadStyle));
    }

    var _setNormalShape = function(shape, shapeStyle) {
        var arr = shapeStyle.textEffect.font.split("px ");
        var size = arr[0];
        setTextValue("shapeText", shape.text());
        setColorValue("shapeBackgroundColor", getRGBAColor(shapeStyle.fill.color, 1 - shapeStyle.fill.transparency));
        setColorValue("shapeColor", getRGBAColor(shapeStyle.textEffect.color, 1 - shapeStyle.textEffect.transparency));
        setNumberValue("baseShapeWidth", shape.width());
        setNumberValue("baseShapeHeight", shape.height());
        setDropDownText($("#shapeExTab div.insp-dropdown-list[data-name='shapeFontSize']"), parseInt(size));
        setDropDownText($("#shapeExTab div.insp-dropdown-list[data-name='shapeFontFamily']"), arr[1]);
        $("#shape_setting_text_valign .insp-radio-button-group span.btn").removeClass('active');
        $("#shape_setting_text_halign .insp-radio-button-group span.btn").removeClass('active');

        // setting aligenments
        var _activeAlignBtn = function (alignType, alignValue) {
            var queryString = '#shape_setting_text_' + alignType + 'align .insp-radio-button-group span.btn[data-name="' + alignValue + '"]';
            $(queryString).addClass('active');
        }
        var alignMap = {
            vAlign: ['top', 'center', 'bottom'],
            hAlign: ['left', 'center', 'right']
        };
        _activeAlignBtn('v', alignMap.vAlign[shapeStyle.textFrame.vAlign]);
        _activeAlignBtn('h', alignMap.hAlign[shapeStyle.textFrame.hAlign]);
    }

    var _setCommonAttrs = function(shape, shapeStyle) {
        processShapeBorderLineSetting(getShapeBorderTypeString(shapeStyle.line.lineStyle));
        setColorValue("shapeBorderColor", getRGBAColor(shapeStyle.line.color, 1- shapeStyle.line.transparency));
        setTextValue("shapeBorderWidth", shapeStyle.line.width);
        setDropDownValue("shapeCapType", shapeStyle.line.capType);
        setDropDownValue("shapeJoinType", shapeStyle.line.joinType);
        setTextValue("shapeName", shape.name());
        setCheckValue("allowShapeMove", shape.allowMove());
        setCheckValue("allowShapeResize", shape.allowResize());
        setCheckValue("shapeCanPrint", shape.canPrint());
        setCheckValue("shapeIsVisible", shape.isVisible());
        setCheckValue("shapeDynamicMove", shape.dynamicMove());
        setCheckValue("shapeDynamicSize", shape.dynamicSize());
        setCheckValue("shapeIsLocked", shape.isLocked());
        setCheckValue("shpaeIsSelected", shape.isSelected());
        setTextValue("shapeRotate", Number((shape.rotate&&shape.rotate())||0));
    }

    var _setVisiableElements =  function(_groupCount, _shapeCount, _connectorCount) {
        var shapeElements = ['shape_setting_text', 'shape_setting_color','shape_setting_bgcolor', 'shape_setting_font_size',
            'shape_setting_font_family', 'shape_setting_text_valign', 'shape_setting_text_halign',
            'shape_setting_width', 'shape_setting_height'];
        var connectorElements = ['shape_connector_begin_arrow_style', 'shape_connector_begin_arrow_width', 'shape_connector_begin_arrow_height',
            'shape_connector_end_arrow_style', 'shape_connector_end_arrow_width', 'shape_connector_end_arrow_height', 'shape_connector_type'];

        if((_shapeCount + _connectorCount) > 1) {
            hiddenElements(['shape_name', 'shape_setting_text']);
        } else {
            showElements(['shape_name', 'shape_setting_text'])
        }
        if(_connectorCount > 0) {
            showElements(connectorElements);
        } else {
            hiddenElements(connectorElements);
        }
        if(_shapeCount > 0) {
            showElements(shapeElements);
        } else {
            hiddenElements(shapeElements);
        }
        if(shapes.length === 1) {
            hiddenElements(['shape_group_btn']);
            if(_groupCount >= 1) {
                showElements(['shape_group_container', 'shape_ungroup_btn']);
                hiddenElements(['shape_name', 'shape_setting_text'])
            } else {
                hiddenElements(['shape_group_container', 'shape_ungroup_btn']);
            }
        } else if(shapes.length > 1) {
            showElements(['shape_group_container', 'shape_group_btn']);
            hiddenElements(['shape_ungroup_btn']);
        }
    }

    var _digShapes = function(_shapes) {
        _shapes.forEach(function(shape){
            var shapeType = getShapeType(shape);

            if(shape && shapeType === "shapeGroup"){
                groupCount ++;
                _digShapes(shape.all());
            } else {
                var shapeStyle = shape.style();
                _setCommonAttrs(shape, shapeStyle);

                if (shapeType === "shape") {
                    shapeCount ++;
                    _setNormalShape(shape, shapeStyle);
                } else {
                    connectorCount ++;
                    _setConnector(shape, shapeStyle);
                }
            }
        });
    }

    _digShapes(shapes);
    _setVisiableElements(groupCount, shapeCount, connectorCount);
}

function hiddenElements(ids) {
    ids.forEach(function(id) {
        $('#' + id).hide();
    });
}

function showElements(ids) {
    ids.forEach(function(id) {
        $('#' + id).show();
    });
}

function applyShapeSetting() {
    var width = getNumberValue("baseShapeWidth");
    var height = getNumberValue("baseShapeHeight");
    var borderValueString = $('#shape-border-line-type').data('value')
    var borderStyle = spreadNS.Shapes.PresetLineDashStyle[borderValueString];
    var borderColor = getBackgroundColor("shapeBorderColor");
    var borderWidth = getNumberValue("shapeBorderWidth");
    var bgColor = getBackgroundColor("shapeBackgroundColor");
    var rotate = Number(getTextValue("shapeRotate"));
    var text = getTextValue("shapeText");
    var shapeColor = getBackgroundColor("shapeColor")
    var fontSize = getDropDownText("shapeFontSize");
    var fontFamily = getDropDownText("shapeFontFamily");
    var font = fontSize + "px " + fontFamily;
    var isSelected = getCheckValue("shpaeIsSelected");
    var allowMove = getCheckValue("allowShapeMove");
    var allowResize = getCheckValue("allowShapeResize");
    var canPrint = getCheckValue("shapeCanPrint");
    var isVisible = getCheckValue("shapeIsVisible");
    var dynamicMove = getCheckValue("shapeDynamicMove");
    var dynamicSize = getCheckValue("shapeDynamicSize");
    var isLocked = getCheckValue("shapeIsLocked");
    var beginArrowWidth = getDropDownValue("beginArrowWidth");
    var beginArrowLength = getDropDownValue("beginArrowHeight");
    var endArrowWidth = getDropDownValue("endArrowWidth");
    var endArrowLength = getDropDownValue("endArrowHeight");
    var endArrowStyle = getArrowStyleType($('#end-arrow-style-type')[0].className);
    var beginArrowStyle = getArrowStyleType($('#begin-arrow-style-type')[0].className);
    var capType = getDropDownValue("shapeCapType");
    var joinType = getDropDownValue("shapeJoinType");
    var _getConnector = function(_shape, _shapeStyle) {
        _shapeStyle.line.beginArrowheadStyle = beginArrowStyle;
        _shapeStyle.line.beginArrowheadLength = beginArrowLength;
        _shapeStyle.line.beginArrowheadWidth = beginArrowWidth;
        _shapeStyle.line.endArrowheadStyle = endArrowStyle;
        _shapeStyle.line.endArrowheadLength = endArrowLength;
        _shapeStyle.line.endArrowheadWidth = endArrowWidth;
        return _shapeStyle;
    }

    var _getShapeStyle = function(_shape, _shapeStyle, deep) {
        _shapeStyle.fill.color = bgColor;
        _shapeStyle.fill.transparency = getColorTransparency("shapeBackgroundColor");
        _shapeStyle.textEffect.color = shapeColor;
        _shapeStyle.textEffect.font = font;
        _shapeStyle.textEffect.transparency = getColorTransparency("shapeColor");
        _shape.text(text);

        if(deep === 0) {
            _shape.height(height);
            _shape.width(width);
        }
        return _shapeStyle;
    }

    var _getCommonStyle = function(_shape, _shapeStyle, deep) {
        _shapeStyle.line.capType = capType;
        _shapeStyle.line.joinType = joinType;
        _shapeStyle.line.lineStyle = borderStyle;
        _shapeStyle.line.color = borderColor;
        _shapeStyle.line.width = borderWidth;
        if(deep === 0) {
            _shape.rotate&&_shape.rotate(rotate);
        }
        _shapeStyle.line.transparency =  getColorTransparency("shapeBorderColor");
        return _shapeStyle;
    }

    var _applyBaseSettings = function(_shape) {
        _shape.isSelected(isSelected);
        _shape.allowMove(allowMove);
        _shape.allowResize(allowResize);
        _shape.canPrint(canPrint);
        _shape.dynamicMove(dynamicMove);
        _shape.dynamicSize(dynamicSize);
        _shape.isVisible(isVisible);
        _shape.isLocked(isLocked);
    }

    var _applayShapeSettingToItem = function (_shapes, deep) {
        _shapes.forEach(function(item) {
            var itemType = getShapeType(item);
            if(itemType === 'shapeGroup') {
                // deep may use for judge whether use style in group shape , using deep++ will make eroor in below case:
                // shapes is like this [GroupShape, shape],  in this case , deep will be > 0 for the shapes[1]
                _applayShapeSettingToItem(item.all(), deep + 1);
                item.height(height);
                item.width(width);
                _applyBaseSettings(item);
                item.rotate(rotate);
            } else {
                var shapeStyle = item.style();
                shapeStyle = _getCommonStyle(item, shapeStyle, deep);
                if(itemType === 'connector') {
                    shapeStyle = _getConnector(item, shapeStyle);
                }
                if(itemType === "shape"){
                    shapeStyle = _getShapeStyle(item, shapeStyle, deep);
                }
                if(_shapes.length === 1 && deep === 0) {
                    _applyBaseSettings(item);
                }
                item.style(shapeStyle)
            }
        });
    }
    var sheet = spread.getActiveSheet();
    sheet.suspendPaint();

    _applayShapeSettingToItem(getActiveShapes(), 0);

    sheet.resumePaint();

}

function changeCapTypeIndex(value){
    setDropDownValue("shapeCapType", value);
}

function changeJoinTypeIndex(value){
    setDropDownValue("shapeJoinType", value);
}

function changeShapeFontSize(value){
    setDropDownText($("#shapeExTab div.insp-dropdown-list[data-name='shapeFontSize']"), value);
}

function changeShapeFontFamily(value){
    setDropDownText($("#shapeExTab div.insp-dropdown-list[data-name='shapeFontFamily']"), value);
}

function getColorTransparency(colorRoot){
    var color = getBackgroundColor(colorRoot);
    return 1 - (getStrIndex(color,',',2) === -1 ? 1 : color.slice(getStrIndex(color,',',2) + 1, -1));
}

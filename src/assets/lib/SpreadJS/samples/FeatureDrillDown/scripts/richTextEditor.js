(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports) :
        typeof define === 'function' && define.amd ? define(['exports'], factory) :
            (factory((global.richEditor = {})));
}(this, (function (exports) {
    'use strict';

    var _extends = Object.assign || function (target) {
        for (var i = 1; i < arguments.length; i++) {
            var source = arguments[i];
            for (var key in source) {
                if (Object.prototype.hasOwnProperty.call(source, key)) {
                    target[key] = source[key];
                }
            }
        }
        return target;
    };

    function stopBubble(e) {
        if (e && e.stopPropagation) {
            e.stopPropagation();
        } else {
            window.event.cancelBubble = true;
        }
    }

    function colorRGB2Hex(color) {
        var rgb = color.split(',');
        var r = parseInt(rgb[0].split('(')[1]);
        var g = parseInt(rgb[1]);
        var b = parseInt(rgb[2].split(')')[0]);

        var hex = "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
        return hex;
    }

    var defaultParagraphSeparatorString = 'defaultParagraphSeparator';
    var formatBlock = 'formatBlock';
    var addEventListener = function addEventListener(parent, type, listener) {
        return parent.addEventListener(type, listener);
    };
    var appendChild = function appendChild(parent, child) {
        return parent.appendChild(child);
    };
    var createElement = function createElement(tag) {
        return document.createElement(tag);
    };
    var queryCommandState = function queryCommandState(command) {
        return document.queryCommandState(command);
    };
    var queryCommandValue = function queryCommandValue(command) {
        return document.queryCommandValue(command);
    };

    function exec(command) {
        var value = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
        return document.execCommand(command, false, value);
    };

    var defaultActions = {
        fontFamily: {
            icon: '<span id="fontFamilyValue" class="font-family-value">Calibri</span>' +
            '<span class="drop-down-arrow fa fa-caret-down fa-lg fa-pull-right"></span>',
            title: 'Bold',
            type: 'drop-down',
            specialStyle: {
                width: '120px'
            },
            dropDownListId: 'textFontFamilyList',
            queryValue: function () {
                var value = queryCommandValue('fontName').replace(/"/g, '');
                document.getElementById('fontFamilyValue').innerText = value;
            },
            result: function result(e) {
                e.currentTarget.style.display = 'none';
                stopBubble(e);
                var value =  e.target.innerText;
                document.getElementById('fontFamilyValue').innerText = value;
                return e.target.nodeName.toUpperCase() === 'LI' ? exec('fontName', value) : false;
            }
        },
        fontSize: {
            icon: '<span id="fontSizeValue">14</span>' +
            '<span class="drop-down-arrow fa fa-caret-down fa-lg fa-pull-right"></span>',
            title: 'Bold',
            type: 'drop-down',
            specialStyle: {
                width: '40px'
            },
            dropDownListId: 'textFontSizeList',
            queryValue: function () {
                var fontSizeMap = [0, 10, 13, 16, 18, 24, 32, 48];
                var value = queryCommandValue('fontSize');
                document.getElementById('fontSizeValue').innerText = fontSizeMap[value] ? fontSizeMap[value] : document.getElementById('fontSizeValue').innerText;
            },
            result: function result(e) {
                e.currentTarget.style.display = 'none';
                stopBubble(e);
                var value = e.target.value;
                var fontSizeMap = [0, 10, 13, 16, 18, 24, 32, 48];
                if (fontSizeMap[value]) {
                    document.getElementById('fontSizeValue').innerText = fontSizeMap[value];
                }
                return e.target.nodeName.toUpperCase() === 'LI' ? exec('fontSize', parseInt(value)) : false;
            }
        },
        bold: {
            icon: '<b style="font-weight: bold;">B</b>',
            title: 'Bold',
            state: function state() {
                return queryCommandState('bold');
            },
            result: function result() {
                return exec('bold');
            }
        },
        italic: {
            icon: '<i>I</i>',
            title: 'Italic',
            state: function state() {
                return queryCommandState('italic');
            },
            result: function result() {
                return exec('italic');
            }
        },
        underline: {
            icon: '<u>U</u>',
            title: 'Underline',
            state: function state() {
                return queryCommandState('underline');
            },
            result: function result() {
                return exec('underline');
            }
        },
        strikethrough: {
            icon: '<strike>S</strike>',
            title: 'Strike-through',
            state: function state() {
                return queryCommandState('strikeThrough');
            },
            result: function result() {
                return exec('strikeThrough');
            }
        },
        colorPicker: {
            icon: '<span id="foreColorValue" class="color_picker_result">&nbsp;A&nbsp;</span>' +
            '<span class="drop-down-arrow fa fa-caret-down fa-lg fa-pull-right"></span>',
            title: 'colorPicker',
            type: 'drop-down',
            specialStyle: {
                width: '40px'
            },
            dropDownListId: 'colorPicker',
            queryValue: function () {
                var value = queryCommandValue('foreColor');
                document.getElementById('foreColorValue').style.borderBottomColor = value;
            },
            result: function result(e) {
                e.currentTarget.style.display = 'none';
                stopBubble(e);
                return e.target.nodeName.toUpperCase() === 'LI' ? exec('foreColor', colorRGB2Hex(e.target.style.backgroundColor)) : false;
            }
        },
        superScript: {
            icon: 'X<sup>2</sup>',
            title: 'SuperScript',
            state: function state() {
                return queryCommandState('superscript');
            },
            result: function result() {
                return exec('superscript');
            }
        },
        subScript: {
            icon: 'X<sub>2</sub>',
            title: 'SubScript',
            state: function state() {
                return queryCommandState('subscript');
            },
            result: function result() {
                return exec('subscript');
            }
        }
    };

    var defaultClasses = {
        actionbar: 'rich-editor-actionbar',
        button: 'rich-editor-button',
        content: 'rich-editor-content',
        selected: 'rich-editor-button-selected'
    };

    var init = function init(settings) {
        var actions = [];
        if(settings.actions){
            settings.actions.map(function (action) {
                if (typeof action === 'string') {
                    action = defaultActions[action];
                } else if (defaultActions[action.name]) {
                    action = _extends({}, defaultActions[action.name], action);
                }
                actions.push(action);
            })
        } else {
            Object.keys(defaultActions).map(function (action) {
                actions.push(defaultActions[action]);
            });
        }
        var classes = _extends({}, defaultClasses, settings.classes);

        var defaultParagraphSeparator = settings[defaultParagraphSeparatorString] || 'div';

        var actionbar = createElement('div');
        actionbar.className = classes.actionbar;
        appendChild(settings.element, actionbar);

        var content = settings.element.content = createElement('div');
        content.contentEditable = true;
        content.className = classes.content;
        content.oninput = function (_ref) {
            var firstChild = _ref.target.firstChild;
            if (firstChild && firstChild.nodeType === 3) exec(formatBlock, '<' + defaultParagraphSeparator + '>'); else if (content.innerHTML === '<br>') content.innerHTML = '';
            if(settings.onChange){
                settings.onChange();
            }
        };
        content.onkeydown = function (event) {
            if (event.key === 'Tab') {
                event.preventDefault();
            } else if (event.key === 'Enter' && queryCommandValue(formatBlock) === 'blockquote') {
                setTimeout(function () {
                    return exec(formatBlock, '<' + defaultParagraphSeparator + '>');
                }, 0);
            }
        };
        appendChild(settings.element, content);

        actions.forEach(function (action) {
            var button = createElement('button');
            button.className = classes.button;
            button.innerHTML = action.icon;
            button.title = action.title;
            if (action.specialStyle) {
                for (var styleProp in action.specialStyle) {
                    if (action.specialStyle.hasOwnProperty(styleProp)) {
                        button.style[styleProp] = action.specialStyle[styleProp];
                    }
                }
            }
            button.setAttribute('type', 'button');
            button.onclick = function () {
                var lists = document.getElementsByClassName('list');
                var dropDownList = document.getElementById(action.dropDownListId);
                for (var i = 0; i < lists.length; i++) {
                    if (lists[i] !== dropDownList) {
                        lists[i].style.display = 'none';
                    }
                }
                if (action.type === 'drop-down') {
                    if (button.contains(dropDownList)) {
                        dropDownList.style.display === 'none' ? (dropDownList.style.display = 'block') : (dropDownList.style.display = 'none');
                    } else {
                        button.appendChild(dropDownList);
                        var hostOffsetHeight = button.offsetHeight;
                        dropDownList.style.top = hostOffsetHeight + 'px';
                        dropDownList.style.zIndex = 1000;
                        dropDownList.style.display = 'block';
                        dropDownList.onclick = function (e) {
                            return action.result(e) && content.focus();
                        }
                    }
                } else {
                    action.result() && content.focus();
                    if(action.title === 'Bold') {
                        var bElements = document.querySelectorAll('#richTextContainer b');
                        bElements.forEach(function(ele){
                            ele.style.fontWeight = 'bold';
                        });
                    } else if(action.title === 'Italic') {
                        var iElements = document.querySelectorAll('#richTextContainer i');
                        iElements.forEach(function(ele) {
                            ele.style.fontStyle = 'italic';
                        })
                    }
                }
            };

            if (action.state) {
                var handler = function handler() {
                    return button.classList[action.state() ? 'add' : 'remove'](classes.selected);
                };
                addEventListener(content, 'keyup', handler);
                addEventListener(content, 'mouseup', handler);
                addEventListener(button, 'click', handler);
            }

            if (action.queryValue) {
                var handler = action.queryValue;
                if (action.dropDownListId) {
                    var dropDownList = document.getElementById(action.dropDownListId);
                    addEventListener(dropDownList, 'click', handler);
                }
                addEventListener(content, 'keyup', handler);
                addEventListener(content, 'mouseup', handler);
                addEventListener(button, 'click', handler);
            }

            appendChild(actionbar, button);
        });

        if (settings.styleWithCSS) exec('styleWithCSS');
        exec(defaultParagraphSeparatorString, defaultParagraphSeparator);

        return settings.element;
    };

    var richEditor = {exec: exec, init: init};

    exports.exec = exec;
    exports.init = init;
    exports['default'] = richEditor;

    Object.defineProperty(exports, '__esModule', {value: true});

})));

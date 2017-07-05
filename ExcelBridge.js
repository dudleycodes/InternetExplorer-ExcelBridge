//JSLint Settings
/*jslint indent: 4, maxerr: 200, passfail: false, white: false, browser: true, devel: true, windows: true,
 rhino: false, debug: true, evil: true, continue: false, forin: false, sub: false, css: false, todo: false, on: false,
 fragment: false, es5: false, vars: false, undef: false, nomen: true, node: false, plusplus: true, bitwise: true,
 regexp: false, newcap: false, unparam: false, sloppy: true, eqeq: true, stupid: false */
//

//JSHint Settings
/*jshint -W018 */ // Msg: Confusing case of '!'.                Reason: easiest-to-read boolean typecasting
/*jshint -W088 */ // Msg: Creating global 'for' variable.       Reason: clashes w/ JSLint
/*global CollectGarbage: false */
//

/**
 * @author James Dudley <James@Dudley.codes>
 * @file Provide a real-time link between Internet Explorer and Microsoft Excel via ActiveX enabling basic document
 * manipulation and data retrieval using JavaScript. Purposely written to have no external dependencies.
 *
 * @example var ExcelLink = new excelBridgeFromFile("\\server1\docs\abc.xlsx");
 * ExcelLink.setSelected('Sheet2');
 * var data = ExcelLink.getRow(1, 0, 'Sheet3');
 * ExcelLink.setCell('B4', 'Data has been read successfully.', 'Sheet1');
 * ExcelLink.terminate();
 *
 * @copyright Unlicense - For more information, please refer to <http://unlicense.org/>
 * This is free and unencumbered software released into the public domain.
 * Anyone is free to copy, modify, publish, use, compile, sell, or distribute this software, either in source code
 * form or as a compiled binary, for any purpose, commercial or non-commercial, and by any means.
 *
 * In jurisdictions that recognize copyright laws, the author or authors of this software dedicate any and all 
 * copyright interest in the software to the public domain. We make this dedication for the benefit of the public
 * at large and to the detriment of our heirs and successors. We intend this dedication to be an overt act of 
 * relinquishment in perpetuity of all present and future rights to this software under copyright law.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
 * THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

/**
 * @class
 * @constructor
 *
 * @classdesc <p>Provides a real-time link between Internet Explorer and Microsoft Excel using ActiveX.</p>
 *
 * <h4>Requirements</h4>
 * <blockquote><ul>
 *      <li>Internet Explorer 8 or higher (ECMAScript 3rd Edition equivalent)</li>
 *      <li>Microsoft Office 2007 or higher</li>
 * </ul></blockquote>
 *
 * <h4>Object Instantiation</h4>
 * <blockquote><p>To initiate a new instance use one of the following included factory functions.  
 * In the event of  a critical error these factories will return null instead of an object. Creating
 * a new instance directly (w/o using a factory) is not recommended.<br /><br /></p>
 * <p>
 *      <ul>
 *          <li>excelBridgeFromFile(strUNCPath)</li>
 *          <li>excelBridgeFromNew()</li>
 *          <li>excelBridgeFromUserOpen()</li>
 *      </ul>
 * </p></blockquote>
 *
 * <h4>Excel Application Busy</h4>
 * <blockquote><p>When Excel is busy its activeX connection will become unresponsive and direct
 * calls will halt all javascript execution.  For the most part 'Excel application busy' will 
 * occur predictably such as when the user is actively manipulating a document or when a modular 
 * dialogue box is displayed. Occasionally 'application busy' will not be so predictable such as 
 * during an auto-save or VBA  script is modifying the document.  The function isReady() can be used
 * to determine if Excel is available for read/writes.<br/ ><br/ ></p>
 *
 * <p>To smooth basic retrieval operations around 'application busy' this object uses caching. All
 * caching is stored in containers prefaced with 'cache' which should not be accessed or manipulated
 * from outside of the object instance.</p></blockquote>
 *
 * <h4>Function Design</h4>
 * <blockquote><p>Functions default to using a 'selected' workbook/sheet rather than the 'active' as 
 *  the 'active' workbook/sheet depends on user selection and can be unpredictable.<br /><br /></p>
 *
 * <p>Common Function Returns:<br /><br /></p>
 * <p><table>
 * <tr>
 *      <td>null</td>
 *      <td>&nbsp;</td>
 *      <td>Indicates the function cannot be preformed at this time.  Occurs when when the Excel 
 *      Application is busy and caching is not available.</td>
 * </tr>
 * <tr>
 *      <td>varaint</td>
 *      <td>&nbsp;</td>
 *      <td>Function was successfully executed and is returning an appropriate value.</td>
 * </tr>
 * <tr>
 *      <td>false</td>
 *      <td>&nbsp;</td>
 *      <td>If not expecting a boolean return false indicates the function was successfully executed
 *      but cannot be  preformed.  Example: the worksheet to be modified does not exist.</td>
 * </tr>
 * </table><br /></p>
 *
 * <p>Functions that take parameters of 'worksheetName' and 'workbookName' do so in that order. 
 * While stark in contrast to Excel's document model tree the vast majority of use-cases involves
 * only one workbook and as such is the more optional parameter.</p></blockquote>
 *
 * <h4>Error Control</h4>
 * <blockquote><p>Internet Explorer 8's jscript implementation has issues catching extended ECMA  
 * standard errors. So rather than throwing errors to bubble messages up through the stack this 
 * object uses an event log. If a function returns null (be it from application busy or otherwise) 
 * or fails the member  variable 'logMessage' will contain a human-friendly string communicating the
 * issue(s) that occurred immediately after a function call.  In the even no issues occurred 
 * logMessage will contain an empty string.  Functions that manipulate this variable should not be 
 * accessed outside of the object instance.<br/ ><br/ ></p>
 *
 * <p>Functions with the verbs 'convertTo', 'format', and 'is' will never affect log messaging - all
 * others will.</p></blockquote>
 *
 * <h4>Development Notes</h4>
 * <blockquote><p>This class file is designed to be validated using JSLint - all setting flags are
 * embedded.</p></blockquote>
 *
 * @summary <tt>It is not recommended to create an instance directly - use an appropriate provided
 * factory function!</tt>
 * @param {string} locus - Determines the Excel Document, if any, to load. By default treats the 
 * passed value as a 'Universal Naming Convention' file path. Pass a null or empty value to create a 
 * new document or the string '@userOpenFile' to prompt the user to open a document from their 
 * workstation.
 * @example var ExcelLink = new ExcelBridge("\\server\dir\sub\test.xlsx");@s
 */
function ExcelBridge(locus) {
    /**
    * Cache store for simple values.
    * @private
    * @type object
    */
    this.cache = {};


    /**
     * Object that sanitizes and validates a reference to a cell.
     * @param {string|array|object} seed - Seed is the cell reference and can take multiple formats:<br />
     * <ul><li>String: "BAA163"</li>
     * <li>Array: [15, 312]<li>
     * <li>Array: ['BAA', 312]</li>
     * <li>Object returned from cellReference: cellReference("BAA163")</li></ul>
     * @returns {object}
     * @example var columnIdx = ExcelLink.convertToColumn("YBB");
     */
    this.cellReference = function (seed) {
        var blankObj = { "col": null, "colIndex": null, "errorMsg": '', "isValid": false, "name": null, "row": null },
            convertAlphaToIndex,
            i = 0,
            returnObj = { "col": null, "colIndex": null, "errorMsg": '', "isValid": false, "name": null, "row": null },
            temp;

        /**
         * Returns an integer representing a column index from a column name (e.g. 'AU' becomes '47').
         * @private
         */
        convertAlphaToIndex = function (strLetter) {
            var returnValue = 0;
            strLetter = strLetter.toUpperCase();
            if (/^[A-Z]+$/.test(strLetter)) {
                returnValue = 0;
                for (i = 0; i < strLetter.length; ++i) {
                    returnValue *= 26;
                    returnValue += strLetter.charCodeAt(i) - ("A".charCodeAt() - 1);
                }
            }
            return returnValue;
        };

        if (seed === undefined || seed === null) {
            returnObj.errorMsg = "`" + String(seed) + "` is not a valid cell.";
        } else if (typeof seed === 'string') {
            //= Cell name
            seed = seed.replace(/^\s+|\s+$/g, '').toUpperCase();
            if (/^[A-Z]+[0-9]+$/.test(String(seed))) {
                returnObj.name = String(seed).toUpperCase();
                returnObj.row = String(seed.match(/\d+$/));
                returnObj.col = String(seed).replace(returnObj.row, '');
                returnObj.row = parseInt(returnObj.row, 10);
                temp = convertAlphaToIndex(returnObj.col);
                if (temp) {
                    returnObj.colIndex = temp;
                } else {
                    returnObj = blankObj;
                    returnObj.errorMsg = "'" + returnObj.col + "' is not a valid column.";
                }
            } else {
                returnObj.errorMsg = "'" + seed + "' is not a valid cell.";
            }
        } else if (Object.prototype.toString.call(seed) === '[object Array]') {
            //= Coordinate set [col, row]
            if (seed.length !== 2) {
                returnObj.errorMsg = 'Cell coordinates expects 2 values [col, row] not ' + seed.length;
            } else {
                //first element
                if (isNaN(seed[0])) {
                    seed[0] = String(seed[0]).toUpperCase().replace(/^\s+|\s+$/g, '');
                    temp = convertAlphaToIndex(seed[0]);
                    if (temp) {
                        returnObj.col = seed[0];
                        returnObj.colIndex = temp;
                    } else {
                        returnObj.errorMsg = "'" + seed[0] + "' is not a valid column.";
                    }
                    temp = null;
                } else if (parseInt(seed[0], 10) === parseFloat(seed[0])) {
                    temp = (function (colIndex) {
                        var col = '',
                            dividend = colIndex,
                            modulo = 0;
                        while (dividend > 0) {
                            modulo = (dividend - 1) % 26;
                            col = String.fromCharCode(65 + modulo) + col;
                            dividend = Math.floor((dividend - modulo) / 26);
                        }
                        return col.toUpperCase();
                    }(seed[0]));
                    returnObj.col = temp;
                    returnObj.colIndex = parseInt(seed[0], 10);
                } else {
                    returnObj.errorMsg = "'" + seed[0] + "' is not a valid column.";
                }

                if (!returnObj.errorMsg) {
                    //second element
                    if (!isNaN(seed[1]) && parseInt(seed[1], 10) === parseFloat(seed[1])) {
                        returnObj.row = parseInt(seed[1], 10);
                    } else {
                        returnObj.errorMsg = "'" + seed[1] + "' is not a valid row.";
                    }
                }

                if (!returnObj.errorMsg) {
                    returnObj.name = String(returnObj.col + returnObj.row);
                }
            }
        } else if (typeof seed === 'object' && typeof seed.hasOwnProperty === 'function') {     //is object
            if (seed.hasOwnProperty('row') && (seed.hasOwnProperty('col') || seed.hasOwnProperty('colIndex'))) {
                returnObj = this.cellReference([(seed.col || seed.colIndex), seed.row]);
            } else {
                returnObj.errorMsg = "Invalid object - must have properties 'col' or 'colIndex' as well as 'row'.";
            }
        } else {
            returnObj.errorMsg = "'" + String(seed) + "' is not a valid cell.";
        }

        //Value boundary checks
        if (!returnObj.errorMsg) {
            if (returnObj.colIndex < 1) {
                returnObj.errorMsg = "Column '" + String(returnObj.colIndex) +
                    "' is too low - must be one or higher.";
            } else if (returnObj.colIndex > this.settingColMaxCount) {
                returnObj.errorMsg = "Column '" + String(returnObj.colIndex) + "' is too high - must be "
                    + this.formatCommaSeperatedNumber(this.settingColMaxCount) + " or lower";
            } else if (returnObj.row < 1) {
                returnObj.errorMsg = "Row '" + String(returnObj.row)
                    + "' is too low - must be one or higher.";
            } else if (returnObj.row > this.settingRowMaxCount) {
                returnObj.errorMsg = "Row '" + String(returnObj.row) + "' is too high - must be "
                    + this.formatCommaSeperatedNumber(this.settingRowMaxCount) + " or lower";
            }
        }

        if (returnObj.errorMsg) {
            temp = returnObj.errorMsg;
            returnObj = null;
            returnObj = blankObj;
            returnObj.errorMsg = temp;
        }
        returnObj.isValid = !returnObj.errorMsg;

        return returnObj;
    };


    /**
     * Object that sanitizes and validates a reference to a color.
     * @param {int|string|object} seed1
     * @param {int} seed2
     * @param {int} seed3
     * @returns {object}
     * @private
     */
    this.colorReference = function (seed1, seed2, seed3) {
        var i = 0,
            regEx = '',
            returnObj = {
                "blue": null,
                "errorMsg": '',
                "green": null,
                "hex": null,
                "red": null,
                "isValid": false
            },
            temp;

        if (seed1 === undefined && seed2 === undefined && seed3 === undefined) {
            seed1 = '#000000';
        } else if (typeof seed1 === 'string') {
            //Check if color name was provided
            seed1 = seed1.toLowerCase().replace(/^\s+|\s+$/g, '');  //IE8 has no trim()
            if ((seed1.length !== 4 && seed1.length !== 7) || seed1.substring(0, 1) !== '#') {
                seed2 = undefined;
                seed3 = undefined;
                temp = null;
                for (i = 1; i < this.settingColorNames.length; ++i) {
                    if (this.settingColorNames[i] === seed1) {
                        temp = this.settingColorCodes[i];
                        break;
                    }
                }
                if (temp !== null) {
                    seed1 = temp;
                    temp = null;
                } else {
                    returnObj.errorMsg = "'" + seed1 + "' is not a recognized color name.";
                }
            }
        } else if (typeof seed1 === 'object' && Object.prototype.toString.call(seed1) !== '[object Array]') {
            temp = seed1;
            if (temp.hex !== undefined && temp.hex !== null) {
                seed1 = temp.hex;
                seed2 = undefined;
                seed3 = undefined;
            } else {
                seed1 = (temp.r || temp.red || undefined);
                seed2 = (temp.g || temp.green || undefined);
                seed3 = (temp.b || temp.blue || undefined);
                if (seed1 === undefined && seed2 === undefined && seed3 === undefined) {
                    returnObj.errorMsg = 'Recieved object did not contain any color values.';
                }
            }
        } else if (typeof seed1 === 'boolean' || Object.prototype.toString.call(seed1) === '[object Array]') {
            returnObj.errorMsg = "seed1 is invalid.";
        } else if (typeof seed2 === 'boolean' || Object.prototype.toString.call(seed2) === '[object Array]') {
            returnObj.errorMsg = "seed2 is invalid.";
        } else if (typeof seed3 === 'boolean' || Object.prototype.toString.call(seed3) === '[object Array]') {
            returnObj.errorMsg = "seed3 is invalid.";
        } else if (!isNaN(seed1) && seed1 > 0 && seed1 <= 56 && seed2 === undefined && seed3 === undefined) {
            seed1 = this.settingColorCodes[parseInt(seed1, 10)];        //Excel color code lookup
        }

        if (returnObj.errorMsg === '') {
            temp = null;
            if (typeof seed1 === 'string' && seed1.substring(0, 1) === '#') {
                if (seed1.length === 4) {       //Expand shorthand form (#03F -=> #0033FF)
                    regEx = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
                    seed1 = seed1.replace(regEx, function (m, r, g, b) {
                        m = null;   //prevent JSLint from squawking about an unused variable
                        return '#' + r + r + g + g + b + b;
                    });
                }
                temp = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(seed1);
            }

            if (temp !== null) {
                returnObj.blue = parseInt(temp[3], 16);
                returnObj.green = parseInt(temp[2], 16);
                returnObj.hex = temp[0];
                returnObj.red = parseInt(temp[1], 16);
            } else {
                seed1 = (parseInt(seed1, 10) || 0);
                if (seed1 < 0 || seed1 > 255) { returnObj.errorMsg = "'" + seed1 + "' is not a valid red value. "; }
                seed2 = (parseInt(seed2, 10) || 0);
                if (seed2 < 0 || seed2 > 255) { returnObj.errorMsg += "'" + seed2 + "' is not a valid green value. "; }
                seed3 = (parseInt(seed3, 10) || 0);
                if (seed3 < 0 || seed3 > 255) { returnObj.errorMsg += "'" + seed3 + "' is not a valid blue value. "; }

                if (returnObj.errorMsg === '') {
                    returnObj.blue = seed3;
                    returnObj.green = seed2;
                    returnObj.hex = "#" + ((1 << 24) + (seed1 << 16) + (seed2 << 8) + seed3).toString(16).slice(1);
                    returnObj.red = seed1;
                }
            }
        }

        returnObj.b = returnObj.blue;
        returnObj.g = returnObj.green;
        if (typeof returnObj.hex === 'string') { returnObj.hex = returnObj.hex.toUpperCase(); }
        returnObj.isValid = !returnObj.errorMsg;
        returnObj.r = returnObj.red;

        return returnObj;
    };

    /**
     * Object that sanitizes and validates a reference to a column.
     * @param {string|array|object} seed - Seed is the column reference and can take multiple formats.  It can also take
     * optional row values:<br />
     * <ul><li>String: "BAA"</li>
     * <li>Integer: 15<li>
     * <li>Array: ['BAA', 312]</li>
     * <li>String: "BAA15"</li>
     * <li>Object returned from cellReference: cellReference("BAA163")</li></ul>
     * @returns {object}
     * @private
     * @example var columnIdx = ExcelLink.colReference("YBB25");
     */
    this.colReference = function (seed) {
        var fakeRow = null,
            returnObj = {
                "col": null,
                "colIndex": null,
                "colReference": true,
                "errorMsg": '',
                "isValid": false,
                "row": false
            };

        if (typeof seed === 'string') {
            seed = String(seed).toUpperCase().replace(/^\s+|\s+$/g, '');
        }

        if (seed === undefined || seed === null) {
            returnObj.errorMsg = "`" + String(seed) + "` is not a valid column.";
        } else if ((!isNaN(seed) && parseInt(seed, 10) === parseFloat(seed)) || /^[A-Z]+$/.test(seed)) {
            fakeRow = 1;
            seed = this.cellReference([seed, fakeRow]);
        } else if (typeof seed === 'object' && Object.prototype.toString.call(seed) !== '[object Array]') {
            if (seed.hasOwnProperty('col') || seed.hasOwnProperty('colIndex')) {
                if (seed.hasOwnProperty('row') && seed.row !== false) {
                    seed = this.cellReference(seed);
                } else {
                    fakeRow = 1;
                    seed = this.cellReference([(seed.col || seed.colIndex), fakeRow]);
                }
            } else {
                returnObj.errorMsg = "Invalid colReference object - must have property 'col' and/or 'colIndex'.";
            }
        } else {
            seed = this.cellReference(seed);
        }

        if (!returnObj.errorMsg) {
            if (!seed.isValid) {
                returnObj.errorMsg = seed.errorMsg;
            } else {
                returnObj.col = seed.col;
                returnObj.colIndex = seed.colIndex;
                returnObj.isValid = true;
                if (fakeRow === null) {
                    returnObj.row = seed.row;
                }
            }
        }

        return returnObj;
    };


    this.logger = function (bolDebugMode) {
        this.debugMode = !!bolDebugMode;
        this.reset();
    };

    this.logger.prototype.isEmpty = function () {
        return !(this.msg);
    };


    this.logger.prototype.push = function (msg) {
        if (msg !== null && typeof msg === 'object' && msg.number !== undefined && msg.description !== undefined) {
            msg = (this.debugMode) ? String(msg.number + ": " + msg.description) : 'An unexpected error occured.';
        } else {
            msg = String(msg).replace(/^\s+|\s+$/g, '');    //No .trim() in IE8 :(
        }

        if (msg) {
            this.msg = (!this.msg) ? msg : this.msg + " " + msg;
            this.msg += ".";
        }

        return true;
    };


    this.logger.prototype.reset = function () {
        this.setName('default');
        this.msg = '';
    };


    this.logger.prototype.setName = function (name) {
        name = String(name).replace(/^\s+|\s+$/g, '');

        switch (name.toLowerCase()) {
        case 'workbook':
        case 'badworkbook':
            this.name = 'ExcelErrorBadWorkbook';
            break;
        case 'worksheet':
        case 'badworksheet':
            this.name = 'ExcelErrorBadWorksheet';
            break;
        case 'busy':
            this.name = 'ExcelErrorBusy';
            break;
        case 'terminated':
            this.name = 'ExcelErrorTerminated';
            break;
        default:
            this.name = 'ExcelError';
        }
    };


    this.logger.prototype.throwError = function (unshiftMsg, bolDisableClear) {
        if (!this.isEmpty()) {
            if (unshiftMsg) { this.unshift(unshiftMsg); }
            var ErrObj = function (name, description) {
                this.description = description;
                this.ExcelBridgeError = true;
                this.message = description;
                this.name = name;
                this.number = -1;
                if (typeof this.toString !== 'function') {
                    this.toString = function () { return this.name + ': ' + this.description; };
                }
            };
            ErrObj.prototype = Error.prototype;

            if (!bolDisableClear) {
                this.reset();
            }
            this.setName(this.name);
            throw new ErrObj(this.name, this.msg);
        }
    };


    this.logger.prototype.toString = function () {
        var returnValue = '';
        if (this.msg) {
            returnValue = (this.name) ? this.name + ': ' + this.msg : this.msg;
        }

        return returnValue;
    };


    this.logger.prototype.unshift = function (msg) {
        if (msg !== null && typeof msg === 'object' && msg.number !== undefined && msg.description !== undefined) {
            msg = (this.debugMode) ? String(msg.number + ": " + msg.description) : 'An unexpected error occured.';
        } else {
            msg = String(msg).replace(/^\s+|\s+$/g, '');    //No .trim() in IE8 :(
        }

        if (msg) {
            this.msg = (this.msg) ? msg + ". " + this.msg : msg + ".";
        }

        return true;
    };


    /**
     * Used to toggle ExcelBridge in and out of 'debug mode'.  If debug mode is enabled a warning alert will appear
     * when a class instance is created.  When debug mode is off warnings and errors will be suppressed from the 
     * end-user.
     * @type boolean
     * @default false
     */
    this.debugMode = false;

    /**
     * Container object for ActiveX link to Microsoft Excel
     * @private
     * @type object
     * @default null
     */
    this.ExcelApp = null;


    /**
     * Used to store "last-known" hashes for the document tree
     * @default {}
     * @example this.hashTreeCurrent['book1']['sheet2'] = -363664769;
     * @private
     * @type object
     */
    this.hashTreeCurrent = {};


    /**
     * Used to store hashes for the document tree when it was last synced.
     * @default {}
     * @example this.hashTreeCurrent['book1']['sheet2'] = 436899445;
     * @private
     * @type object
     */
    this.hashTreeSynced = {};


    /**
     * Used when ExcelBridge shouldn't actively connect to ActiveXObject "Excel.Application".
     * Primarily for debugging and UnitTests.  Invoked by using a locus of "@noconnect" during class
     * instantiation.
     * @private
     * @type boolean
     */
    this.passiveMode = false;


    /**
     * Object that sanitizes and validates a reference to a range.
     * param {string|array|object} [seed1] - Either a range name "a5:t4" or a cell reference
     * param {string|array|object} [seed2=null] - A cell reference
     * @returns {object}
     * @private
     * @example var rangeRef = ExcelLink.rangeReference('a5:t5');
     * var rangeRef = ExcelLink.rangeReference('a5', 't4');
     * var rangeRef = ExcelLink.rangeReference(['a', 5], ['t', 4]);
     * var rangeRef = ExcelLink.rangeReference([1, 5], [20, 4]);
     * var rangeRef = ExcelLink.rangeReference(cellRef1, cellRef2);
     */
    this.rangeReference = function (seed1, seed2) {
        var errorMsg = '',
            leftCell = { "errorMsg" : "No data provided.", "isValid": false },
            returnObj = {
                "colSpan": 0,
                "errorMsg": "",
                "isValid": false,
                "leftCell": { "isValid": false },
                "name": null,
                "rightCell": { "isValid": false },
                "rowSpan": 0
            },
            rightCell = { "errorMsg": "No data provided.", "isValid": false },
            temp;

        if (seed1 !== undefined && seed2 !== undefined) {
            leftCell = this.cellReference(seed1);
            rightCell = this.cellReference(seed2);
            if (!leftCell.isValid) {
                errorMsg = "seed1: " + leftCell.errorMsg + " ";
            }
            if (!rightCell.isValid) {
                errorMsg = ("seed2: " + rightCell.errorMsg).replace(/^\s+|\s+$/g, '');
            }
        } else if (seed1 !== undefined && (seed2 === undefined || seed2 === null || seed2 === '')) {
            temp = String(seed1).split(":");
            if (typeof seed1 === 'string' && temp.length === 2) {
                leftCell = this.cellReference(temp[0]);
                rightCell = this.cellReference(temp[1]);
            } else if (seed1 && seed1.leftCell !== undefined && seed1.rightCell !== undefined) {
                leftCell = this.cellReference(seed1.leftCell);
                rightCell = this.cellReference(seed1.rightCell);
            }
            if (!leftCell.isValid || !rightCell.isValid) {
                errorMsg = "'" + String(seed1) + "' is not a valid range of cells.";
            }
        } else {
            errorMsg = "'" + String(seed1) + String(seed2) + "' is not a valid range of cells.";
        }

        if (!errorMsg) {    // make sure the left cell really is on the left
            temp = null;
            if (leftCell.row > rightCell.row) {
                temp = leftCell.row;
                leftCell.row = rightCell.row;
                rightCell.row = temp;
            }

            if (leftCell.colIndex > rightCell.colIndex) {
                temp = leftCell.colIndex;
                leftCell.colIndex = rightCell.colIndex;
                rightCell.colIndex = temp;
            }

            if (temp !== null) {
                //Rebuild so .name is correct
                leftCell = this.cellReference([leftCell.colIndex, leftCell.row]);
                rightCell = this.cellReference([rightCell.colIndex, rightCell.row]);
                temp = null;
            }

            //calculate colSpan and rowSpan
            returnObj.colSpan = Math.abs(rightCell.colIndex - leftCell.colIndex) + 1;
            returnObj.rowSpan = Math.abs(rightCell.row - leftCell.row) + 1;
        }

        returnObj.errorMsg = errorMsg;
        returnObj.isValid = !errorMsg;
        returnObj.leftCell = leftCell;
        returnObj.rightCell = rightCell;
        if (returnObj.leftCell.isValid && returnObj.rightCell.isValid) {
            returnObj.name = String(leftCell.name + ':' + rightCell.name);
        }

        return returnObj;
    };


    /**
     * Object that sanitizes and validates a reference to a row.
     * @param {string|array|object} seed - Seed is the row reference and can take multiple formats.  It can also take
     * optional col values:<br />
     * <ul><li>Integer: 15<li>
     * <li>Array: ['BAA', 312]</li>
     * <li>String: "BAA15"</li>
     * <li>Object returned from cellReference: cellReference("BAA163")</li></ul>
     * @returns {object}
     * @private
     * @example var rowIdx = ExcelLink.colReference("YBB25").row;
     */
    this.rowReference = function (seed) {
        var fakeCol = null,
            returnObj = {
                "col": false,
                "colIndex": false,
                "errorMsg": '',
                "isValid": false,
                "row": null,
                "rowReference": true
            };

        if (typeof seed === 'string') {
            seed = String(seed).toUpperCase().replace(/^\s+|\s+$/g, '');
        }

        if (seed === undefined || seed === null) {
            returnObj.errorMsg = "`" + String(seed) + "` is not a valid row.";
        } else if (!isNaN(seed) && parseInt(seed, 10) === parseFloat(seed)) {
            fakeCol = 1;
            seed = this.cellReference([fakeCol, seed]);
        } else if (typeof seed === 'object' && Object.prototype.toString.call(seed) !== '[object Array]') {
            if (seed.hasOwnProperty('row')) {
                if ((seed.hasOwnProperty('col') && seed.col !== false)
                        || (seed.hasOwnProperty('colIndex') && seed.colIndex !== false)) {
                    seed = this.cellReference(seed);
                } else {
                    fakeCol = 1;
                    seed = this.cellReference([fakeCol, seed.row]);
                }
            } else {
                returnObj.errorMsg = "Invalid rowReference object - must have property 'row'.";
            }
        } else {
            seed = this.cellReference(seed);
        }

        if (!returnObj.errorMsg) {
            if (!seed.isValid) {
                returnObj.errorMsg = seed.errorMsg;
            } else {
                returnObj.row = seed.row;
                returnObj.isValid = true;
                if (fakeCol === null) {
                    returnObj.col = seed.col;
                    returnObj.colIndex = seed.colIndex;
                }
            }
        }
        return returnObj;
    };


    /**
     * String containing the selected cell color.
     * @private
     * @default #FFFFFF
     * @type string
     */
    this.selectedCellColor = "#FFFFFF";


    /**
     * String containing the selected font color.
     * @private
     * @default #000000
     * @type string
     */
    this.selectedFontColor = "#000000";


    /**
     * String containing the selected workbook (if any).
     * @private
     * @default false
     * @type false|string
     */
    this.selectedWorkbookName = '';


    /**
     * String containing the selected worksheet (if any).
     * @private
     * @default false
     * @type false|string
     */
    this.selectedWorksheetName = '';


    /**
     * Determines if ExcelBridge should automatically terminate the Excel instance when the user
     * closes the Application Window
     * @public
     * @default true
     * @type boolean
     */
    this.settingAutoTerminate = true;


    /**
     * Array containing valid cell border index names.
     * @private
     * @type array
     */
    this.settingBorderIndexes = ['xldiagonaldown', 'xldiagonalup', 'xledgebottom', 'xledgeleft',
        'xledgeright', 'xledgetop', 'xlinsidehorizontal', 'xlinsidevertical'];


    /**
     * Object containing valid cell border weights.  The index being the constant and the value
     * being the recognized name. Duplicate entry is to allow for compatibility with Microsoft's
     * XLBorderWeight constant values.
     * @link http://msdn.microsoft.com/en-us/library/office/aa221100%28v=office.11%29.aspx
     * @private
     * @type object
     */
    this.settingBorderWeights = {
        "1": "xlhairline",
        "2": "xlthin",
        "3": "xlmedium",        //id not recognized by Excel
        "4": "xlthick",
        "-4138": "xlmedium"
    };


    /**
     * Specifies the maximum number of characters Excel supports in a single cell.
     * @private
     * @default 32767 (Excel 2007+)
     * @type integer
     */
    this.settingCellMaxChacters = 32767;


    /**
     * Specifies the maximum number of columns Excel supports in a worksheet.
     * @private
     * @default 16384 (Excel 2007+)
     * @type integer
     */
    this.settingColMaxCount = 16384;


    /**
     * Array containing the hexadecimal values of Excel's 56 predefined colors.  The index
     * correlates with the 'Excel Color Code' whereas the value contains the hexadecimal.
     * @private
     * @type array
     */
    this.settingColorCodes = [undefined, '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF',
            '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000', '#000080', '#808000', '#800080',
            '#008080', '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066',
            '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080',
            '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF',
            '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900',
            '#FF6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300',
            '#993366', '#333399', '#333333'];


    /**
     * Array containing the recognized Excel color names.  The index correlates with the 'Excel
     * Color Code' whereas the value contains the color name.  Not all colors have defined names.
     * @private
     * @type array
     */
    this.settingColorNames = [undefined, 'black', 'white', 'red', 'bright green', 'blue', 'yellow',
            'pink', 'turqoise', 'dark red', 'green', 'dark blue', 'dark yellow', 'violet', 'teal',
            'gray-25%', 'gray-50%'];
    this.settingColorNames[32] = undefined;
    this.settingColorNames.push('sky blue', 'light turqoise', 'light green', 'light yellow',
            'pale blue', 'rose', 'lavendar', 'tan', 'light blue', 'aqua', 'lime', 'gold',
            'light orange', 'orange', 'blue-gray', 'gray-40%', 'dark teal', 'sea green',
            'dark green', 'olive green', 'brown', 'plum', 'indigo', 'gray-80%');

    /**
     * Specifies the maximum width of a column (in characters)
     * @private
     * @default 255 (Excel 2007+)
     * @type integer
     */
    this.settingColMaxWidth = 255;


    /**
     * Array containing valid cell border index names.
     * @private
     * @type array
     */
    this.settingLineStyles = ['xlContinuous', 'xlDash', 'xlDashDot', 'xlDashDotDot', 'xlDot',
        'xlDouble', 'xlLineStyleNone', 'xlSlantDashDot'];


    /**
     * Specifies the maximum number of rows Excel supports in a worksheet.
     * @private
     * @default 10485776 (Excel 2007+)
     * @type integer
     */
    this.settingRowMaxCount = 10485776;


    /**
     * Specifies the maximum height of a row (in points)
     * @private
     * @default 409 (Excel 2007+)
     * @type integer
     */
    this.settingRowMaxHeight = 409;


    /**
     * Specifies the maximum length of a worksheet name
     * @private
     * @default 31 (Excel 2000+)
     * @type integer
     */
    this.settingWorksheetNameMaxLength = 31;


    /**
     * Specifies that maximum zoom percentage Excel supports
     * @private
     * @default 400 (Excel 2007+)
     * @type integer
     */
    this.settingZoomMax = 400;


    /**
     * Specifies that minimum zoom percentage Excel supports
     * @private
     * @default 10 (Excel 2007+)
     * @type integer
     */
    this.settingZoomMin = 10;


    /**
     * Contains the ID of the interval timer (if any).
     * @private
     * @default 0
     * @type integer
     */
    this.timerID = 0;


    /**
     * When timing interval has been enabled contains the current internal tick count
     * @private
     * @default 0
     * @type integer
     */
    this.timerTickCount = 0;


    /**
     * Determines the timing interval (in milliseconds), of any, of the internal tick-rate.  Set to
     * 0 to disable timer.
     * @public
     * @default 0
     * @type integer
     */
    this.timerTickRate = 0;


    // Process debug environment
    this.debugMode = !!this.debugMode;  //Enforce Boolean casting
    window.EUDDEV = (window.EUDDEV || {});
    window.EUDDEV.ExcelBridge = (window.EUDDEV.ExcelBridge || {});
    window.EUDDEV.ExcelBridge.intInstanceCount = (window.EUDDEV.ExcelBridge.intInstanceCount || 0);
    window.EUDDEV.ExcelBridge.intInstanceCount++;

    if (this.debugMode) {
        window.EUDDEV.ExcelBridge.intInstanceDebugCount = (window.EUDDEV.ExcelBridge.intInstanceDebugCount || 0);
        window.EUDDEV.ExcelBridge.intInstanceDebugCount++;

        //Unit Tests may initiate instance multiple times - only display warning once
        if (window.EUDDEV.ExcelBridge.bolDebugAlertSent !== true) {
            window.EUDDEV.ExcelBridge.bolDebugAlertSent = true;
            window.setTimeout(function () {
                alert("WARNING: ExcelBridge.js has been loaded with the 'Debug Mode' flag enabled."
                    + " If this is a production environment contact the support team!");
            }, 500);    //Don't want alert() halting executions - delay by 1/2 a second
        }
    }


    // Set up logging instance
    this.log = new this.logger(this.debugMode);

    //Clean up constructor arguments
    locus = String(locus || '').replace(/^\s+|\s+$/g, '');

    //Create ActiveX Link
    if (locus.length > 0 && locus.toLowerCase() === '@noconnect') {
        this.passiveMode = true;
        this.timerID = 0;
        this.timerTickRate = 0;
    } else {
        this.passiveMode = false;
        //Create ActiveX Link
        try {
            this.ExcelApp = new ActiveXObject("Excel.Application");
        } catch (err) {
            this.ExcelApp = null;
            if (!document.documentMode) {
                //IE 8+ has document.documentMode
                if (this.debugMode) {
                    throw new Error('Failed to create ActiveXObject("Excel.Application"). Browser does not appear to '
                        + 'be IE8+ but ' + window.navigator.appName + ' ' + window.navigator.appVersion + "'.");
                } else {
                    throw new Error("Microsoft Internet Explorer 8 or higher is required!  You appear to be using '"
                        + window.navigator.appName + "'.");
                }
            } else if ((Object.getOwnPropertyDescriptor && Object.getOwnPropertyDescriptior(window, 'ActiveXObject'))
                    || (window.hasOwnProperty('ActiveXObject'))) {
                alert("If you are accessing the network remotely you must configure IE's"
                        + " ActiveX settings.\nSettings > Internet Options > Security (tab) and "
                        + "click 'Custom Level'.\nUnder 'ActiveX controls and plug-ins set "
                        + "'Initialize and script ActiveX controls not marked as safe...' to "
                        + "enabled.\n\n");
                if (this.debugMode) {
                    throw err;
                } else {
                    throw new Error("Microsoft Office must be installed!");
                }
            } else {
                if (this.debugMode) {
                    throw new Error('window.ActiveXObject is undefined!');
                }
                throw new Error('ActiveX must be enabled inside of Internet Explorer!');
            }
        }

        if (typeof VBArray !== 'function') {
            throw new Error('Internet Explorer is required - ExcelBridge requires VBArray() which is not available in '
                + 'Windows Store applications!');
        }
    }

    //Create or load workbook
    var initResult;
    if (!this.passiveMode && this.ExcelApp !== null) {
        if (locus.length === 0) {       //New file
            this.ExcelApp.Workbooks.Add();
            initResult = true;
        } else {
            if (locus.toLowerCase() === '@userOpenFile') {      //Prompt user to select file
                this.setVisible(true);      //Make the the 'open' dialogue flash in window's taskbar
                initResult = !!(this.ExcelApp.FindFile);
            } else {
                if (locus.substring(0, 1) !== '\\' && !locus.match('^[a-zA-z]:')) {     //using URL not Windows UNC
                    if (locus.substring(0, 1) === '/') {
                        locus = location.protocol + '//' + location.host + locus;
                    }
                }
                initResult = !!(this.ExcelApp.Workbooks.Open(locus));
            }

            if (!this.ExcelApp.ActiveSheet) {   //no file was opened
                this.terminateInstance();
            }
        }
    }

    //Build Internals
    if (this.passiveMode === true) {
        this.cache.activeWorkbookName = this.selectedWorkbookName;
        this.cache.activeWorksheetName = this.selectedWorksheetName;
    } else {
        try {
            this.setSelected(this.getActiveWorksheet(true), this.getActiveWorkbook(true));
            this.setVisible(false);
        } catch (err2) {
            this.log.unshift('Critical ExcelBridge Initialization error.');
            this.terminateInstance();
            this.log.throwError();
        }

        //this.queueCacheUpdate();
        //this.updateInternals();
    }
}


ExcelBridge.prototype.commandWrap = function (executionCode, worksheetName, workbookName) {
    this.log.reset();

    var cmdString = 'this.ExcelApp.',
        returnValue = false;

    do {
        if ((workbookName !== undefined && workbookName !== null)
                || (worksheetName !== undefined && worksheetName !== null)) {

            workbookName = (workbookName !== undefined && workbookName !== null && workbookName !== true)
                ? String(workbookName) : this.getSelectedWorkbook();

            if (worksheetName === null || worksheetName === undefined) {
                if (workbookName.toLowerCase() === this.getSelectedWorkbook().toLowerCase()) {
                    worksheetName = this.getSelectedWorksheet();
                } else {
                    this.log.push('A worksheet must be specified when providing the workbook!');
                    this.log.name('worksheet');
                    break;
                }
            }

            cmdString += 'Workbooks("' + workbookName + '").Sheets("' + worksheetName + '").';
        }

        try {
            returnValue = eval(cmdString + executionCode);
        } catch (err) {
            returnValue = this.ping(worksheetName, workbookName);
            if (returnValue) {
                returnValue = false;
                this.log.push(err);
            }
            break;
        }
    } while (returnValue === undefined);

    return returnValue;
};


ExcelBridge.prototype.commandWrapChain = function (executionCode, worksheetName, workbookName) {
    this.commandWrap(executionCode, worksheetName, workbookName);
    this.log.throwError();
    return this;
};


/**
 * Format a number by adding commas - making its display more user-friendly.
 * @param {numeric} [number] - Number to be formatted.
 * @returns {string|false} - Returns false if an invalid number was received
 * @private
 * @example var strUserNumber = ExcelLink.formatCommaSeperatedNumber(436329347.2);  //returns "436,329,347.2"
 */
ExcelBridge.prototype.formatCommaSeperatedNumber = function (number) {
    var parts = [],
        returnValue = false;

    if (!isNaN(parseFloat(number)) && isFinite(number)) {
        parts = number.toString().split(".");
        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        returnValue = parts.join(".");
    }

    return returnValue;
};


/**
 * Returns the last known active workbook
 * @param {boolean} [bolDisableCache=false] - Set to true to disable pulling values from cache.
 * @returns {string|false|null}
 * @example var workbookName = ExcelLink.getActiveWorkbook(true);
 */
ExcelBridge.prototype.getActiveWorkbook = function (bolDisableCache) {

    var returnValue = this.commandWrap('ActiveWorkbook.Name;');

    if (this.log.isEmpty()) {
        this.cache.activeWorkbookName = returnValue;
    } else if (!bolDisableCache && !this.isClosed() && !this.isTerminated()) {
        this.log.reset();
        returnValue = this.cache.activeWorkbookName;
    } else {
        returnValue = (this.isClosed() || this.isTerminated()) ? false : null;
        this.log.unshift('Unable to determine active workbook.');
    }

    return returnValue;
};


/**
 * Returns the last known active worksheet
 * @param {boolean} [bolDisableCache=false] - Set to true to disable pulling values from cache.
 * @returns {string|false|null}
 * var worksheetName = ExcelLink.getActiveWorksheet(true);
 */
ExcelBridge.prototype.getActiveWorksheet = function (bolDisableCache) {

    var returnValue = this.commandWrap('ActiveSheet.Name;');

    if (this.log.isEmpty()) {
        this.cache.activeWorkbookName = returnValue;
    } else if (!bolDisableCache && !this.isClosed() && !this.isTerminated()) {
        this.log.reset();
        returnValue = this.cache.activeWorkbookName;
    } else {
        returnValue = (this.isClosed() || this.isTerminated()) ? false : null;
        this.log.unshift('Unable to determine active workbook.');
    }

    return returnValue;
};


/**
 * Returns the value of a specified cell.
 *
 * @param {string|array} cellRef - TO BE SPECIFIED
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the worksheet to read
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 *
 * @example var cellContents = ExcelLink.getCell([4, 11], 'Sheet1', 'test1.xlsx');  //Gets contents of col 4, row 11
 * var cellContents = ExcelLink.getCell('D11', 'Sheet1', 'test1.xlsx');     //Gets contents of col 4, row 11
 *
 * @returns {string|false|null}
 */
ExcelBridge.prototype.getCell = function (cellRef, worksheetName, workbookName) {
    this.log.reset();

    var returnValue = false;
    cellRef = this.cellReference(cellRef);

    if (!cellRef.isValid) {
        returnValue = false;
        this.log.push(cellRef.errorMsg);
    } else {
        worksheetName = (worksheetName === undefined || worksheetName === null) ? this.getSelectedWorksheet()
            : String(worksheetName);
        workbookName = (workbookName === undefined || workbookName === null) ? this.getSelectedWorkbook()
            : String(workbookName);

        returnValue = this.commandWrap('Range("' + cellRef.name + '").Text', worksheetName, workbookName);
        if (this.log.isEmpty()) {
            returnValue = (returnValue === null || returnValue === undefined)
                ? '' : String(returnValue).replace(/^\s+|\s+$/g, '');
        }
    }

    if (!returnValue && returnValue !== '') {
        this.log.unshift('Unable to get cell contents.');
    }
    return returnValue;
};


/**
 * Returns an array containing a column's cell values.
 *
 * @param {integer|string|array} [rowRef] - Specified the column to be read.  If a row is provided 
 * getCol() will start on that row rather than row '1'.
 * <ul><li>Integer - specifies the column be read.  E.g. "42".</li>
 * <li>String - specifies the column to be read. E.g. "F"</li>
 * <li>String - cell specifies the column to be read and the row to start from.  E.g. "b35"</li>
 * <li>Array - cell specifies the column to be read and the row to start from.  E.g. [5, 6] or ['e', 6]</li></ul>
 * @param {integer} [length=dynamic] The number of cells to extract.  If length is omitted getCol() extracts cells to
 * the end of the row.  If length is negative getCol() uses it as an index from colReference.  If length is negative and
 * abs(length) exceeds the number of cells to the left of colRef getCol() starts with row 1.  If length is set to 0
 * getCol() will return an empty array.
 * @param {string} [worksheetName=this.getSelectedWorksheet()] - The name of the worksheet to read from.  If left 
 * unspecified the current selected worksheet will be used.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook to read from.  If left
 * unspecified the current selected workbook will be used.
 *
 * @example var colArray = ExcelLink.getCol([4, 2], 8 'Sheet3', 'book1.xlsx');
 *
 * @returns {array|false|null} getRow() will return an array of values.  If Excel is currently inaccessible getCol()
 * will return `null`.  If bad parameters are passed (such as an invalid cell or invalid worksheet) or if Excel has been
 * terminated getCol() will return `false`.
 */
ExcelBridge.prototype.getCol = function (colRef, length, worksheetName, workbookName) {
    this.log.reset();

    var depth = 0,
        i = 0,
        temp = null,
        leftCell,
        returnValue = false,
        rowCount = 0;

    do {
        //Get number of rows.  Will also verify worksheetName/workbookName are valid and excel is ready.
        rowCount = this.getRowCount(worksheetName, workbookName);
        if (!this.log.isEmpty()) {
            returnValue = rowCount;     //Make sure `null` (busy signal) can bubble up through stack
            break;
        }

        //Process colRef
        colRef = this.colReference(colRef);
        if (!colRef.isValid) {
            this.log.push(colRef.errorMsg);
            break;
        }

        //Process length
        if (length === undefined) { length = null; }
        if (length !== null) {
            temp = parseInt(length, 10);
            if (isNaN(temp) || temp !== parseFloat(length)) {
                this.log.push("Specified length ' " + String(length) + "' is not a valid integer.");
                break;
            } else {
                length = temp;
                if (length === 0) {
                    returnValue = [];
                    break;
                }
            }
        }

        //Build "left cell"
        leftCell = this.cellReference([colRef.col, (colRef.row || 1)]);

        //Determine row "depth" from "left cell"
        if (length === null) {
            depth = (rowCount - leftCell.row + 1);
        } else if (length < 0) {
            i = leftCell.row;
            temp = (leftCell.row - Math.abs(length) + 1);
            if (temp < 1) { temp = 1; }
            leftCell = this.cellReference([leftCell.colIndex, temp]);
            depth = Math.abs(length);
            if (depth > i) {
                depth = (i - leftCell.row + 1);
            }
        } else {
            depth = length;
        }

        if ((leftCell.row + depth) > this.settingRowMaxCount) {
            this.log.push('Can not retreive ' + this.formatCommaSeperatedNumber(depth) + ' rows starting with row '
                + leftCell.row + '. Excel only supports ' + this.formatCommaSeperatedNumber(this.settingRowMaxCount)
                + ' total rows.');
            break;
        }

        //Cannot use range().Text on multiple cells - must use brute force
        temp = [];
        for (i = 0; i < depth; ++i) {
            returnValue = this.getCell([leftCell.colIndex, leftCell.row + i], worksheetName, workbookName);
            if (returnValue === false || returnValue === null) {
                temp = false;
                break;
            } else if (length || (returnValue !== '' || temp.length)) {
                //If no length specified we don't want empty cell values being added to front of array
                temp.push(returnValue);
            }
        }

        if (temp !== false) {
            if (length === null) {
                while (temp[temp.length - 1] === '') {
                    i = temp.pop();
                }
            }
            returnValue = temp;
        }
    } while (temp === undefined);

    if (!this.log.isEmpty()) {
        if (this.debugMode) {
            temp = 'Failed on .getCol(' + JSON.stringify(colRef) + ', ' + JSON.stringify(length) + ', '
                + JSON.stringify(worksheetName) + ', ' + JSON.stringify(workbookName) + ')';
        } else {
            temp = 'Unable to retrieve column from ' + String(colRef.name);
        }
        this.log.unshift(temp);
    }

    return returnValue;
};


/**
 * Returns the number of populated columns contained in a worksheet.
 *
 * @param {string} [worksheetName=this.getSelectedWorksheet()] - The name of the worksheet to check.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 *
 * @example var colCount = ExcelLink.getColCount('Sheet1', 'workbook2.xlsx');
 *
 * @returns {integer|false|null}
 */
ExcelBridge.prototype.getColCount = function (worksheetName, workbookName) {
    this.log.reset();

    worksheetName = (worksheetName === undefined || worksheetName === null)
        ? this.getSelectedWorksheet() : String(worksheetName).replace(/^\s+|\s+$/g, '');

    workbookName = (workbookName === undefined || workbookName === null)
        ? this.getSelectedWorkbook() : String(workbookName).replace(/^\s+|\s+$/g, '');

    var depth = 0,
        leftIndex = 0,
        returnValue;

    try {
        leftIndex = this.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).UsedRange.Column;
        depth = this.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).UsedRange.Columns.Count;
        returnValue = parseInt((leftIndex + depth - 1), 10);
        //If no colums are used we want to return 0 not 1
        if (returnValue === 1 && this.getHash(null, worksheetName, workbookName) === '') {
            returnValue = 0;
        }
    } catch (err) {
        returnValue = this.ping(worksheetName, workbookName);
        if (returnValue) {
            returnValue = false;
            this.log.push(err.message);
        }
        this.log.unshift("Unable to determine column count.");
    }

    return returnValue;
};


/**
 * Returns a hash representing a cell-range, a worksheet, a workbook, or all open workbooks. Returns
 * a string upon a successful hash.  If there were no values to be hashed the string will be empty 
 * (e.g. "").  Returns 'null' (busy) or 'false' (invalid) if a hash cannot be preformed.
 *
 * @param {string} [cellRange=null]
 * @param {string} [worksheetName=null]
 * @param {string} [workbookName=null]
 *
 * @example var hashValue = ExcelLink.getHash();            //all open workbooks
 * hashValue = ExcelLink.getHash('a5:t7');                  //cell range "A5" through "T7"
 * hashValue = ExcelLink.getHash('t7);                      //cell range A5 through the worksheet (bottom-right)
 * hashValue = ExcelLink.getHash('a5:t7', 'sheet1');        //cell range hash on worksheet 'sheet1' in selected workbook
 * hashValue = ExcelLink.getHash(a5:t7, 'sheet1', 'book1'); //cell range hash on 'sheet1' in workbook 'book1'
 * hashValue = ExcelLink.getHash('', 'sheet1');             //worksheet hash of sheet1 in selected workbook
 * hashValue = ExcelLink.getHash('', '', 'book1');          //workbook hash of 'book1'
 *
 * @returns {string|false|null}
 */
ExcelBridge.prototype.getHash = function (cellRange, worksheetName, workbookName) {
    //This function ain't purdy but it's fast
    this.log.reset();

    var funcGetRawValue,
        funcVBArrayToString,
        returnValue = '',
        temp,
        that = this;

    funcVBArrayToString = function (vbArr) {
        var returnValue = '';

        try {
            returnValue = String(new VBArray(vbArr).toArray());
            //If all cells are empty but were ever modified a string of commas could be returned - treat as empty.
            returnValue = (returnValue.length > 0 && !(/^[,]*$/.test(returnValue))) ? returnValue : '';
        } catch (err) {
            if (err.number === -2146823275) {       //-2146823275 === VBArray: argument is not a VBArray object."
                returnValue = (!vbArr && vbArr !== 0) ? '' : String(vbArr);
            } else {
                throw err;
            }
        }

        return returnValue;
    };

    funcGetRawValue = function (cellRange, worksheetName, workbookName) {
        var index = 0,
            returnValue = false,
            temp,
            temp2;

        try {
            if (cellRange !== null && cellRange !== undefined) {
                worksheetName = (worksheetName === null || worksheetName === undefined)
                    ? that.getSelectedWorksheet() : String(worksheetName);
                workbookName = (workbookName === null || workbookName === undefined) ? that.getSelectedWorkbook()
                    : String(workbookName);

                temp = that.rangeReference(cellRange);
                if (!temp.isValid) {
                    temp = that.rangeReference(cellRange, [that.getColCount(), that.getRowCount()]);
                }

                if (!temp.isValid) {
                    that.log.push(temp.errorMsg);
                    returnValue = false;
                } else {
                    temp = that.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).Range(temp.name).Value2;
                    returnValue = funcVBArrayToString(temp);
                }
            } else if (worksheetName !== null && worksheetName !== undefined) {
                //hash used-range inside given worksheet
                worksheetName = String(worksheetName);
                workbookName = (workbookName === null || workbookName === undefined) ? that.getSelectedWorkbook()
                    : String(workbookName);
                temp = that.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).UsedRange.Value2;
                returnValue = funcVBArrayToString(temp);
            } else if (workbookName !== null && workbookName !== undefined) {
                //hash all worksheets in given workbook
                workbookName = String(workbookName);
                temp = that.getWorksheets(workbookName);
                for (index = 0; index < (temp.length || 0); ++index) {
                    temp2 = funcGetRawValue(null, String(temp[index]), workbookName);
                    if (temp2 === false || temp2 === null) {
                        returnValue = temp2;
                        break;
                    } else {
                        returnValue += temp2 + String(temp[index]).toLowerCase();
                    }
                }
            } else {
                //hash all workbooks
                temp = that.getWorkbooks();
                for (index = 0; index < (temp.length || 0); ++index) {
                    temp2 = funcGetRawValue(null, null, temp[index]);
                    if (temp2 === false || temp2 === null) {
                        returnValue = temp2;
                        break;
                    } else {
                        returnValue += temp2 + String(temp[index]).toLowerCase();
                    }
                }
            }
        } catch (err) {
            returnValue = that.ping(worksheetName, workbookName);
            if (returnValue) {
                returnValue = false;
                that.log.push(err);
            }
        }

        return returnValue;
    };

    returnValue = funcGetRawValue(cellRange, worksheetName, workbookName);

    if (!this.log.isEmpty()) {
        temp = 'Failed to getHash() ';

        if (cellRange !== undefined && cellRange !== null) {
            temp += 'of range ' + String(cellRange.name || cellRange) + ' ';
        }

        if (worksheetName !== undefined && worksheetName !== null) {
            temp += 'on worksheet ' + String(worksheetName) + ' ';
        }

        if (workbookName !== undefined && workbookName !== null) {
            temp += 'in workbook ' + String(workbookName) + ' ';
        }
        this.log.unshift(temp);
    } else {
        //Create Hash using CRC
        returnValue = (function (str) {
            str = String(str);
            var hashValue = 0,
                i = 0,
                strlen = str.length;
            if (str.length > 0) {
                for (i = 0; i < strlen; ++i) {
                    hashValue = str.charCodeAt(i) + ((hashValue << 5) - hashValue);
                    hashValue = hashValue & hashValue;
                }
            }
            return (hashValue !== 0) ? String(hashValue) : '';
        }(returnValue));
    }

    return returnValue;
};


/**
 * Returns an array containing a row's cell values.
 *
 * @param {integer|string|array} [rowRef] - Specified the row to be read.  If a column or column index is provided 
 * getRow() will start on that column rather than column 'A'.
 * <ul><li>Integer - specifies the row be read.  E.g. "42".</li>
 * <li>String - cell specifies the row to be read and the column to start from.  E.g. "b35"</li>
 * <li>Array - cell specifies the row to be read and the column to start from.  E.g. [5, 6] or ['e', 6]</li></ul>
 * @param {integer} [length=dynamic] The number of cells to extract.  If length is omitted getRow() extracts cells to
 * the end of the row.  If length is negative getRow() uses it as an index from rowReference.  If length is negative and
 * abs(length) exceeds the number of cells to the left of rowRef getRow() starts with column 1.  If length is set to 0
 * getRow() will return an empty array.
 * @param {string} [worksheetName=this.getSelectedWorksheet()] - The name of the worksheet to read from.  If left 
 * unspecified the current selected worksheet will be used.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook to read from.  If left
 * unspecified the current selected workbook will be used.
 *
 * @example var rowArray = ExcelLink.getRow([4, 2], 8 'Sheet3', 'book1.xlsx');
 *
 * @returns {array|false|null} getRow() will return an array of values.  If Excel is currently inaccessible getRow()
 * will return `null`.  If bad parameters are passed (such as an invalid cell or invalid worksheet) or if Excel has been
 * terminated getRow() will return `false`.
 */
ExcelBridge.prototype.getRow = function (rowRef, length, worksheetName, workbookName) {
    this.log.reset();

    var colCount = 0,
        depth = 0,
        i = 0,
        temp = null,
        leftCell,
        returnValue = false;

    do {
        //Get number of rows.  Will also verify worksheetName/workbookName are valid and excel is ready.
        colCount = this.getRowCount(worksheetName, workbookName);
        if (!this.log.isEmpty()) {
            returnValue = colCount;     //Make sure `null` (busy signal) can bubble up through stack
            break;
        }

        //Process colRef
        rowRef = this.rowReference(rowRef);
        if (!rowRef.isValid) {
            this.log.push(rowRef.errorMsg);
            break;
        }

        //Process length
        if (length === undefined) { length = null; }
        if (length !== null) {
            temp = parseInt(length, 10);
            if (isNaN(temp) || temp !== parseFloat(length)) {
                this.log.push("Specified length ' " + String(length) + "' is not a valid integer.");
                break;
            } else {
                length = temp;
                if (length === 0) {
                    returnValue = [];
                    break;
                }
            }
        }

        //Build "left cell"
        leftCell = this.cellReference([(rowRef.colIndex || 1), rowRef.row]);

        //Determine row "depth" from "left cell"
        if (length === null) {
            depth = (colCount - leftCell.colIndex + 1);
        } else if (length < 0) {
            i = leftCell.colIndex;
            temp = (leftCell.colIndex - Math.abs(length) + 1);
            if (temp < 1) { temp = 1; }
            leftCell = this.cellReference([temp, leftCell.row]);
            depth = Math.abs(length);
            if (depth > i) {
                depth = (i - leftCell.colIndex + 1);
            }
        } else {
            depth = length;
        }

        if ((leftCell.colIndex + depth) > this.settingColMaxCount) {
            this.log.push('Can not retreive ' + this.formatCommaSeperatedNumber(depth)
                + ' columns starting with column ' + leftCell.col + '. Excel only supports '
                + this.formatCommaSeperatedNumber(this.settingColMaxCount) + ' total columns.');
            break;
        }

        //Cannot use range().Text on multiple cells - must use brute force
        temp = [];
        for (i = 0; i < depth; ++i) {
            returnValue = this.getCell([leftCell.colIndex + i, leftCell.row], worksheetName, workbookName);
            if (returnValue === false || returnValue === null) {
                temp = false;
                break;
            } else if (length || (returnValue !== '' || temp.length)) {
                //If no length specified we don't want empty cell values being added to front of array
                temp.push(returnValue);
            }
        }

        if (temp !== false) {
            if (length === null) {
                while (temp[temp.length - 1] === '') {
                    i = temp.pop();
                }
            }
            returnValue = temp;
        }
    } while (temp === undefined);

    if (!this.log.isEmpty()) {
        if (this.debugMode) {
            temp = 'Failed on .getRow(' + JSON.stringify(rowRef) + ', ' + JSON.stringify(length) + ', '
                + JSON.stringify(worksheetName) + ', ' + JSON.stringify(workbookName) + ')';
        } else {
            temp = 'Unable to retrieve row from ' + String(rowRef.name);
        }
        this.log.unshift(temp);
    }

    return returnValue;
};


/**
 * Calculates the number of populated rows contained in a worksheet
 * @param {string} [worksheetName=this.getSelectedWorksheet()] - The name of the worksheet to count.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 * @returns {integer|false|null}
 * @example var rowDepth = ExcelLink.getRowCount('Sheet3');
 */
ExcelBridge.prototype.getRowCount = function (worksheetName, workbookName) {
    this.log.reset();

    worksheetName = (worksheetName === undefined || worksheetName === null)
        ? this.getSelectedWorksheet() : String(worksheetName).replace(/^\s+|\s+$/g, '');

    workbookName = (workbookName === undefined || workbookName === null)
        ? this.getSelectedWorkbook() : String(workbookName).replace(/^\s+|\s+$/g, '');

    var depth = 0,
        topIndex = 0,
        returnValue;

    try {
        topIndex = this.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).UsedRange.Row;
        depth = this.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).UsedRange.Rows.Count;
        returnValue = parseInt((topIndex + depth - 1), 10);
        //If no colums are used we want to return 0 not 1
        if (returnValue === 1 && this.getHash(null, worksheetName, workbookName) === '') {
            returnValue = 0;
        }
    } catch (err) {
        returnValue = this.ping(worksheetName, workbookName);
        if (returnValue) {
            returnValue = false;
            this.log.push(err.message);
        }
        this.log.unshift("Unable to determine row count.");
    }

    return returnValue;
};


/**
 * Returns a string containing the selected workbook.  Returns false if Excel has been closed.
 * @returns {string|false}
 * @example var workbookName = ExcelLink.getSelectedWorkbook();
 */
ExcelBridge.prototype.getSelectedWorkbook = function () {
    return (this.isTerminated() || this.isClosed()) ? false : this.selectedWorkbookName;
};


/**
 * Returns a string containing the selected worksheet.  Returns false if Excel has been closed.
 * @returns {string|false}
 * @example var worksheetName = ExcelLink.getSelectedWorksheet();
 */
ExcelBridge.prototype.getSelectedWorksheet = function () {
    return (this.isClosed()) ? false : this.selectedWorksheetName;
};


/**
 * Gets list of all open workbooks
 * @param {boolean} [bolDisableCache=false] - Set to true to disable pulling data from cache when Excel is busy.
 * @returns {string[]|false|null}
 * @example var workbookArray = ExcelLink.getWorkbooks(true);
 */
ExcelBridge.prototype.getWorkbooks = function (bolDisableCache) {
    this.log.reset();

    bolDisableCache = !!bolDisableCache;

    var bolReturnFlag = true,
        workbookCount = 0,
        workbookIndex = 0,
        workbookName = "",
        workbookNameArray = [];

    try {
        workbookCount = this.ExcelApp.Workbooks.Count;
        for (workbookIndex = 1; workbookIndex <= workbookCount; ++workbookIndex) {
            workbookNameArray.push(this.ExcelApp.Workbooks(workbookIndex).Name);
        }
    } catch (err) {
        if (this.isClosed() || this.isTerminated()) {
            this.log.push("Unable to get workbooks. Excel application is no longer running.");
            bolReturnFlag = false;
        } else if (!bolDisableCache) {
            for (workbookName in this.hashTreeCurrent) {
                if (this.hashTreeCurrent.hasOwnProperty(workbookName)) {
                    //Edge Case: Excel can become unavailable resulting in a partial array
                    //      Additionally the array could contain members not yet in cache
                    if (workbookNameArray[workbookName] === undefined) {
                        workbookNameArray.push(workbookName);
                    }
                }
            }
        } else if (!this.isReady()) {
            this.log.push("Unable to get workbooks. Excel application is busy.");
            bolReturnFlag = null;
        } else {
            this.log.push("Unable to get workbooks.");
            this.log.push(err.message);
            bolReturnFlag = false;
        }
    }
    return (bolReturnFlag) ? workbookNameArray : bolReturnFlag;
};


/**
 * Get an array of the worksheets available in a workbook.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the workbook to select.
 * @param {string} [bolDisableCache=false] - Disables reading from cache.
 * @returns {string[]|false|null}
 * @example var worksheetArray = ExcelLink.getWorksheets('book1.xlsx');
 */
ExcelBridge.prototype.getWorksheets = function (workbookName, bolDisableCache) {
    this.log.reset();
    workbookName = String(workbookName || this.getSelectedWorkbook());
    bolDisableCache = !!bolDisableCache;

    var i = 0,
        pingCheck = null,
        sheetCount = 0,
        sheetName = "",
        sheetNameArray = [];

    pingCheck = this.ping(null, workbookName);
    if (pingCheck === false) {
        this.log.unshift("Unable to retrieve worksheets.");
        return false;
    }

    try {
        sheetCount = this.ExcelApp.Workbooks(workbookName).Worksheets.Count;
        for (i = 1; i <= sheetCount; ++i) {
            sheetName = this.ExcelApp.Workbooks(workbookName).Worksheets(i).Name;
            if (sheetName) {
                sheetNameArray.push(sheetName);
            }
        }
        return sheetNameArray;
    } catch (err) {
        if (!this.isReady()) {
            if (bolDisableCache) {
                return null;
            }
            for (sheetName in this.hashTreeCurrent[workbookName]) {
                if (this.hashTreeCurrent[workbookName].hasOwnProperty(sheetName)) {
                    if (this.hashTreeCurrent.hasOwnProperty(sheetName)) {
                        if (sheetNameArray[sheetName] === undefined) {
                            sheetNameArray.push(sheetName);
                        }
                    }
                }
            }
            return sheetNameArray;
        }
        this.log.unshift("Unable to retrieve worksheets.");
        this.log.unshift(err.message);
        return false;
    }
};


/**
 * Inserts a column into the worksheet, shifting all other columns to the right.
 * @returns {boolean | integer}
 * @example var bolResult = ExcelLink.insertColumn(2, 'Sheet1');
 */
ExcelBridge.prototype.insertCol = function (colRef, worksheetName, workbookName) {
    this.log.reset();


    var bolResult = false;

    colRef = this.colReference(colRef);
    if (!colRef.isValid) {
        this.log.push(colRef.errorMsg);
    } else {
        workbookName = (workbookName === undefined || workbookName === null) ? true : String(workbookName);
        bolResult = this.commandWrap('Columns(' + colRef.colIndex + ').Insert();', worksheetName, workbookName);
    }

    if (!this.log.isEmpty()) {
        this.log.unshift("Unable to insert column '" + colRef.col + "'.");
    }

    return (this.log.isEmpty()) ? colRef.colIndex : bolResult;
};


/**
 * Inserts a row into the worksheet, shifting all other rows down.
 * @returns {boolean | integer}
 * @example var bolResult = ExcelLink.insertRow(2, 'Sheet2');
 */
ExcelBridge.prototype.insertRow = function (rowRef, worksheetName, workbookName) {
    this.log.reset();

    var bolResult = false;

    rowRef = this.rowReference(rowRef);
    if (!rowRef.isValid) {
        this.log.push(rowRef.errorMsg);
    } else {
        workbookName = (workbookName === undefined || workbookName === null) ? true : String(workbookName);
        bolResult = this.commandWrap('Rows(' + rowRef.row + ').Insert();', worksheetName, workbookName);
    }

    if (!this.log.isEmpty()) {
        this.log.unshift("Unable to insert row '" + rowRef.row + "'.");
    }

    return (this.log.isEmpty()) ? rowRef.row : bolResult;
};


/**
 * Inserts a worksheet into a workbook.
 * @param {string} [newWorksheetName=null] newWorksheetName - The name of the worksheet to be created
 * @param {boolean} [bolNoAutoSelect=false]
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 * @returns {string|null|false}
 * @example var sheetName = ExcelLink.insertWorksheet();
 * var sheetName = ExcelLink.insertWorksheet('Status Sheet');
 */
ExcelBridge.prototype.insertWorksheet = function (newWorksheetName, bolNoAutoSelect, workbookName) {
    this.log.reset();

    var prevWorksheetHandle,
        returnValue = false,
        temp = null;

    bolNoAutoSelect = !!bolNoAutoSelect;
    workbookName = (workbookName === null || workbookName === undefined) ? this.getSelectedWorkbook()
        : String(workbookName);

    do {
        //newWorksheetName
        if (newWorksheetName === null || newWorksheetName === undefined) {
            newWorksheetName = false;
        } else {
            if (!this.validateWorksheetName(newWorksheetName)) { break; }
            newWorksheetName = String(newWorksheetName);
        }

        //insert new sheet and store name in returnValue
        try {
            temp = this.ExcelApp.Workbooks(workbookName).Worksheets.Count;
            prevWorksheetHandle = this.ExcelApp.Workbooks(workbookName).Worksheets(temp);
            returnValue = (this.ExcelApp.Workbooks(workbookName).Worksheets.Add(null, prevWorksheetHandle)).Name;
        } catch (err) {
            returnValue = this.ping(null, workbookName);
            if (returnValue) {
                returnValue = false;
                this.log.push(err);
            }
            break;
        }

        //Rename worksheet if newWorksheetName was provided
        if (newWorksheetName) {
            temp = this.renameWorksheet(returnValue, newWorksheetName, workbookName);
            if (temp) {
                returnValue = newWorksheetName;
            } else {
                prevWorksheetHandle = this.ExcelApp.Workbooks(workbookName).Sheets(returnValue).Delete;
                returnValue = temp;
                break;
            }
        }

        //Auto select new worksheet if applicable
        if (!bolNoAutoSelect) {
            this.setSelected(returnValue, workbookName);
        }
    } while (temp === undefined);

    if (!returnValue) {
        temp = (newWorksheetName) ? "'" + newWorksheetName + "' " : '';
        this.log.unshift("Unable to insert worksheet " + temp + "into '" + workbookName + "'.");
    }

    return returnValue;
};


/**
 * Determine if the Excel Application has been closed by the user.
 * @returns {Boolean}
 * @example var bolApplicationClosed = ExcelLink.isClosed();
 */
ExcelBridge.prototype.isClosed = function () {
    var bolReturnValue = false;

    //ActiveWindow; window closed: null -=> null
    try {
        if (this.ExcelApp.ActiveWindow === null) {
            bolReturnValue = true;
            if (this.settingAutoTerminate) {
                this.terminateInstance();
                this.log.reset();
            }
        }
    } catch (err) {
        //Either terminated or busy.  Not closed if just busy.
        bolReturnValue = this.isTerminated();
    }

    return bolReturnValue;
};


/**
 * Determines if all documents, a workbook, or a worksheet have been modified
 * @param {string} [worksheetkName=this.getSelectedWorksheet()] - Name of the worksheet to check.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 * @param {integer} [recursionSum=0] - INTERNAL USE ONLY - Used internally when building stacks.
 * @returns {Boolean|null}
 * @example var bolWorksheetModified = ExcelLink.isModified('Sheet2');
 */
ExcelBridge.prototype.isModified = function (worksheetName, workbookName, recursionSum) {
    recursionSum = (recursionSum) ? parseInt(++recursionSum, 10) : 0;

    var bolReturnValue = false,
        key = "",
        valueCurrent = 0,
        valuePrevious = 0;

    if (!this.timerID && recursionSum === 0) {
        this.queueCacheUpdate(worksheetName, workbookName);
        this.updateInternals();
    }

    if (worksheetName) {
        //Comparing single worksheet (no stack)
        worksheetName = String(worksheetName);
        workbookName = String(workbookName || this.getSelectedWorkbook());

        valueCurrent = this.hashTreeCurrent[workbookName][worksheetName];
        valuePrevious = this.hashTreeSynced[workbookName][worksheetName];

        if (valueCurrent === true) {
            bolReturnValue = true;
        } else if (valueCurrent !== valuePrevious) {
            bolReturnValue = (valueCurrent === null) ? null : true;
        }
    } else if (workbookName) {
        //Comparing single workbook (mid stack)
        workbookName = String(workbookName);
        for (key in this.hashTreeCurrent[workbookName]) {
            if (this.hashTreeCurrent[workbookName].hasOwnProperty(key)) {
                bolReturnValue = this.isModified(key, workbookName, recursionSum);
                if (bolReturnValue !== false) {
                    break;
                }
            }
        }
    } else {
        //Comparing all workbooks (high stack)
        for (key in this.hashTreeCurrent) {
            if (this.hashTreeCurrent.hasOwnProperty(key)) {
                bolReturnValue = this.isModified(null, key, recursionSum);
                if (bolReturnValue !== false) {
                    break;
                }
            }
        }
    }

    return bolReturnValue;
};


/**
 * Determine if the Excel Application is available for read and writes.
 * @returns {Boolean}
 * @example if (ExcelLink.isReady()) { ... }
 */
ExcelBridge.prototype.isReady = function () {
    var bolReturn = false;
    //possible better choice - http://msdn.microsoft.com/en-us/library/office/ff197917%28v=office.15%29.aspx
    if (!this.passiveMode) {
        try {
            bolReturn = !!(this.ExcelApp.ActiveWindow !== null);
        } catch (err) {
            bolReturn = false;
        }
    }

    return bolReturn;
};


/**
 * Determine if the Excel.exe instance has been terminated.
 * @returns {Boolean}
 * @example if (ExcelLink.isTerminated()) { ... }
 */
ExcelBridge.prototype.isTerminated = function () {
    var bolReturnValue = false,
        throwAwayVar;

    try {
        if (this.ExcelApp === null) {
            bolReturnValue = true;
        } else {
            throwAwayVar = this.ExcelApp.ActiveWindow;
        }
    } catch (err) {
        if (err.number === -2146827826) {
            bolReturnValue = true;
            this.ExcelApp = null;
        }
    }

    return bolReturnValue;
};


/**
 * Determine visibility of a worksheet, a workbook (leave worksheetName null), or the Excel 
 * Application (leave both worksheetName and workbookName null).
 * @param {string} [worksheetName=null] - Name of the worksheet to check.
 * @param {string} [workbookName=null] - Name of the workbook to check.
 * @returns {boolean|null}
 * @example var bolVisible = ExcelLink.isVisible('Sheet2');
 */
ExcelBridge.prototype.isVisible = function (worksheetName, workbookName) {
    var bolReturnValue = false;

    if (!this.isClosed() && !this.isReady()) {
        bolReturnValue = null;
    } else if (!this.isClosed()) {
        try {
            if (worksheetName) {
                bolReturnValue = !!(this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).Visible);
            } else if (workbookName) {
                bolReturnValue = !!(this.ExcelApp.Workbooks(workbookName).Visible);
            } else {
                bolReturnValue = !!(this.ExcelApp.Application.Visible);
            }
        } catch (err) {
            //Nothing to do here -> "is" verb functions do not modify logging.
            bolReturnValue = false;
        }
    }

    return bolReturnValue;
};


ExcelBridge.prototype.paste = function (worksheetName, workbookName) {
    this.log.reset();

    this.ExcelApp.Application.ScreenUpdating = false;
    this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).paste;
    this.ExcelApp.Application.ScreenUpdating = true;

    return this;
};


ExcelBridge.prototype.ping = function (worksheetName, workbookName) {
    this.log.reset();

    var bolReturnValue = false;

    do {
        //sanitize worksheetName
        if (worksheetName !== null && worksheetName !== undefined) {
            if ((typeof worksheetName !== 'string' || worksheetName === '') && typeof worksheetName !== 'number') {
                this.validateWorksheetName(worksheetName);
                break;
            }
            worksheetName = String(worksheetName);
        }

        //sanitize workbookName
        if (workbookName !== null && workbookName !== undefined) {
            if ((typeof workbookName !== 'string' || workbookName === '') && typeof workbookName !== 'number') {
                this.validateWindowsFilename(workbookName);
                break;
            }
            workbookName = String(workbookName);
        }

        //Check Excel Application
        bolReturnValue = this.isReady();
        if (!bolReturnValue) {
            if (this.isTerminated()) {
                this.log.push('Excel application is terminated.');
                this.log.setName('terminated');
            } else {
                this.log.push('Excel application is busy.');
                this.log.setName('busy');
                bolReturnValue = null;
            }
            break;
        }

        if (workbookName || worksheetName) {
            if (!worksheetName) {        //Check workbook only
                try {
                    bolReturnValue = this.ExcelApp.Workbooks(workbookName).Name;
                    bolReturnValue = true;
                } catch (err) {
                    bolReturnValue = false;
                    this.log.push(err);
                    this.log.setName('workbook');
                }
            } else {                    //Check worksheet (include workbook)
                if (!workbookName) { workbookName = this.getSelectedWorkbook(); }
                try {
                    bolReturnValue = this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).Name;
                    bolReturnValue = true;
                } catch (err2) {
                    bolReturnValue = false;
                    this.log.push(err2);
                    this.log.setName('worksheet');
                }
            }
        }
    } while (bolReturnValue === undefined);

    if (!bolReturnValue) {
        if (worksheetName) {
            this.log.unshift("Could not access worksheet '" + worksheetName + "' on workbook '" + workbookName + "'.");
        } else if (workbookName) {
            this.log.unshift("Could not access workbook '" + workbookName + "'.");
        } else {
            this.log.unshift('Could not access Excel.');
        }
    }

    return bolReturnValue;
};


/**
 * Removes a column from a worksheet, shifting all subsequent columns to the left
 * @returns {Boolean|null}
 * @example var bolResult = ExcelLink.removeColumn(2, 'sheet2');
 */
ExcelBridge.prototype.removeCol = function (colRef, worksheetName, workbookName) {
    this.log.reset();

    worksheetName = String(worksheetName || this.getSelectedWorksheet());

    var bolResult = false;

    colRef = this.colReference(colRef);
    if (!colRef.isValid) {
        this.log.push(colRef.errorMsg);
    } else {
        bolResult = this.commandWrap('Columns(' + colRef.colIndex + ').Delete();', worksheetName, workbookName);
    }

    if (!this.log.isEmpty()) {
        this.log.unshift("Unable to remove column '" + colRef.col + "'.");
    }

    return bolResult;
};


/**
 * Removes a row from a worksheet, shifting all subsequent rows up
 * @returns {Boolean|null}
 * @example var bolResult = ExcelLink.removeRow(13, 'Sheet1');
 */
ExcelBridge.prototype.removeRow = function (rowRef, worksheetName, workbookName) {
    this.log.reset();

    worksheetName = String(worksheetName || this.getSelectedWorksheet());
    workbookName = String(workbookName || this.getSelectedWorkbook());

    var bolResult = false;

    rowRef = this.rowReference(rowRef);
    if (!rowRef.isValid) {
        this.log.push(rowRef.errorMsg);
    } else {
        bolResult = this.commandWrap('Rows(' + rowRef.row + ').Delete()', worksheetName, workbookName);
    }

    if (!this.log.isEmpty()) {
        this.log.unshift("Unable to remove row '" + rowRef.row + "'.");
    }

    return bolResult;
};


/**
 * Removes a worksheet from a workbook
 * @param {string} worksheetName - The name of the worksheet to remove.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the parent workbook.
 * @returns {Boolean|null}
 * @example var bolResult = ExcelLink.removeWorksheet('Sheet3', 'book1.xlsx');
 */
ExcelBridge.prototype.removeWorksheet = function (worksheetName, workbookName) {
    this.log.reset();

    workbookName = (workbookName === null || workbookName === undefined) ? this.getSelectedWorkbook()
        : String(workbookName);

    var prevSetting = false,
        returnValue = this.ping(worksheetName, workbookName);

    if (returnValue) {
        try {
            prevSetting = !!(this.ExcelApp.DisplayAlerts);
            this.ExcelApp.DisplayAlerts = false;
            returnValue = !!(this.ExcelApp.Workbooks(workbookName).Sheets(worksheetName).Delete);
            if (workbookName.toUpperCase() === this.getSelectedWorkbook().toUpperCase()
                    && worksheetName.toUpperCase() === this.getSelectedWorksheet().toUpperCase()) {
                if (this.getActiveWorkbook().toUpperCase === workbookName.toUpperCase()) {
                    this.setSelected(this.getActiveWorksheet());
                } else {
                    this.setSelected(this.getWorksheets(workbookName, true)[0]);
                }
            }
        } catch (err) {
            if (err.number === -2146827284) {
                this.log.push('Cannot remove the last visable worksheet in a workbook.');
            } else {
                this.log.push(err);
            }
            returnValue = false;
        }

        if (prevSetting !== false) { this.ExcelApp.DisplayAlerts = prevSetting; }
    }

    if (!returnValue) {
        this.log.unshift("Unable to remove worksheet '" + worksheetName + "' from workbook '" + workbookName + "'");
    }

    return returnValue;
};


/**
 * Renames a worksheet
 * @param {string} worksheetName - The worksheet to be renamed
 * @param {string} newName - The new name for the worksheet
 * @param {string} [workbookName=getSelectedWorkbook()] - The workbook to which the worksheet belongs
 * @returns {boolean|null}
 * @example var bolResult = ExcelLink.renameWorksheet('Sheet2', 'Awesome Sheet', 'book1.xlsx');
 */
ExcelBridge.prototype.renameWorksheet = function (currentName, newName, workbookName) {
    this.log.reset();

    var bolChangeSelected = false,
        returnValue = false;

    do {
        //newName
        if (!this.validateWorksheetName(newName)) { break; }
        newName = String(newName);

        if (this.ping(newName, workbookName)) {
            this.log.push('Worksheet "' + newName + '" already exists in workbook "' + workbookName + '".');
            break;
        } else {
            this.log.reset();
        }

        //currentName
        returnValue = this.ping(currentName, workbookName);
        if (!returnValue) {
            break;
        }
        currentName = String(currentName);

        //bolChangeSelected
        if (this.getSelectedWorkbook(true).toUpperCase() === workbookName.toUpperCase()
                && this.getSelectedWorksheet(true).toUpperCase() === currentName.toUpperCase()) {
            bolChangeSelected = true;
        }

        //preform rename
        returnValue = this.commandWrap('Name = "' + newName + '";', currentName, workbookName);
        if (!returnValue) { break; }
        returnValue = true;

        //Change selected worksheet if needed
        if (bolChangeSelected) {
            this.selectedWorksheet = newName;
            this.selectedWorkbook = workbookName;
        }
    } while (bolChangeSelected === undefined);

    if (!this.log.isEmpty()) {
        this.log.push('Cannot change worksheet name!');
    }

    return returnValue;
};


//index presets: none, outline, inside
ExcelBridge.prototype.setBorder = function (rangeRef, indexes, weight, color, style) {
    this.log.reset();

    weight = String(weight || '').toLowerCase();
    weight = String(color || '').toLowerCase();
    weight = String(style || '').toLowerCase();

    var bordersArray = [],
        i,
        i2,
        temp;

    do {
        //validate rangeRef
        rangeRef = this.rangeReference(rangeRef);
        if (!rangeRef.isValid) {
            temp = this.cellReference(rangeRef);
            if (temp.isValid) {
                rangeRef = this.rangeReference(rangeRef = temp, temp);
            } else {
                this.log.push(rangeRef.errorMsg);
                break;
            }
        }

        //presets
        if (!indexes) {
            indexes = this.settingBorderIndexes;
            style = 'xlLineStyleNone';
        } else if (indexes === true) {
            indexes = ['xlEdgeBottom', 'xlEdgeLeft', 'xlEdgeRight', 'xlEdgeTop'];
        } else if (weight === '0') {
            weight = '';
            style = 'xlLineStyleNone';
        }

        //validate indexes
        if (indexes.prototype.toString.call(indexes) !== '[object Array]') {
            indexes = [String(indexes)];
        }
        for (i = 0; i < indexes.length; ++i) {
            indexes[i] = String(indexes[i]).toLowerCase();
            for (i2 = 0; i2 < this.settingsBorderIndexes.length; ++i2) {
                temp = this.settingBorderIndexes[i2].toLowerCase();
                if (indexes[i] === temp) {
                    indexes[i] = temp;
                    temp = true;
                    break;
                }
            }
            if (temp !== true && i === indexes.length - 1) {
                this.log.push("'" + indexes[i] + "' is not a valid border index.");
                break;
            }
        }
        if (!this.log.isEmpty()) { break; }

        //validate weight
        if (!weight) {
            weight = 'xlThin';
        } else {
            if (this.settingBorderWeights[weight]) {
                weight = this.settingBorderWeights[weight];
            } else {
                temp = String(weight).toLowerCase();
                for (i = 0; i < this.settingBorderWeights.length; ++i) {
                    if (temp === this.settingBorderWeights[i].toLowerCase()) {
                        temp = true;
                        weight = this.settingBorderWeights[i];
                        break;
                    }
                }
                if (temp !== true) {
                    this.log.push("'" + weight + "' is not a valid weight.");
                    break;
                }
            }
        }

        //validate color
        if (!color) { color = [0, 0, 0]; }
        color = this.colorReference(color);
        if (!color.isValid) {
            this.log.push(color.errorMsg);
            break;
        }

        //validate style
        if (!style) {
            style = 'xlContinuous';
        } else {
            if (this.settingLineStyles[style]) {
                style = this.settingLineStyles[style];
            } else {
                temp = String(style).toLowerCase();
                for (i = 0; i < this.settingLineStyles.length; ++i) {
                    if (temp === this.settingLineStyles[i].toLowerCase()) {
                        temp = true;
                        style = this.settingLineStyles[i];
                        break;
                    }
                }
                if (temp !== true) {
                    this.log.push("'" + style + "' is not a valid style.");
                    break;
                }
            }
        }

        for (i = 0; i < indexes.length; ++i) {

        }

        temp = false;
    } while (temp);

    this.log.throwError('Unable to set borders.');

    return this;
};


ExcelBridge.prototype.setCell = function (cellRef, value) {
    this.log.reset();

    cellRef = this.cellReference(cellRef);
    if (!cellRef.isValid) {
        this.log.push(cellRef.errorMsg);
    } else if (value === undefined) {
        this.log.push("Value `undefined` is not valid.");
    } else {
        value = (value === null) ? '' : String(value);
        this.commandWrap('Range("' + cellRef.name + '").Value2 = "' + value + '";', true);
    }

    if (!this.log.isEmpty()) {
        this.log.unshift("Could not set content of cell '" + cellRef.name + "'.");
        this.log.throwError();
    } else {
        return this;
    }
};


/**
 * Sets a worksheet's orientation
 * @param {boolean} [bolValue=true] - True for landscape; false for portrait
 * @param {string} [worksheetName=getSelectedWorksheet()] - Name of the worksheet to modify
 * @param {string} [workbookName=getSelectedWorkbook()] - Name of the workbook to modify
 * @returns {Boolean|null}
 * @example var bolResult = ExcelLink.setLandscape(false);
 */
ExcelBridge.prototype.setLandscape = function (bolValue, worksheetName, workbookName) {
    this.log.reset();

    bolValue = (bolValue !== false && bolValue !== 0) ? 2 : 1;
    worksheetName = String(worksheetName || this.getSelectedWorksheet());
    workbookName = String(workbookName || this.getSelectedWorkbook());

    this.commandWrap('PageSetup.Orientation = ' + bolValue, worksheetName, workbookName);

    if (!this.log.isEmpty()) {
        this.log.unshift("Unable to change page orientation of worksheet '" + worksheetName
            + "' in workbook '" + workbookName + "'");
    }
    this.log.throwError();
    return this;
};


/**
 * Sets the worksheet's top, right, bottom, and left margins.  Functionality mimics CSS's shorthand
 * properties for margin.
 * @param {integer} top - Value for the top margin
 * @param {integer} [right=null] - Value for right margin, leave null for shorthand
 * @param {integer} [bottom=null] - Value for bottom margin, leave null for shorthand
 * @param {integer} [left=null] - Value for left margin, leave null for shorthand
 * @param {string} [worksheetName=getSelectedWorksheet()] - Name of the worksheet to modify
 * @param {string} [workbookName=getSelectedWorkbook()] - Name of the workbook to modify
 * @returns {Boolean|null}
 */
ExcelBridge.prototype.setMargins = function (top, right, bottom, left, worksheetName, workbookName) {
    this.log.reset();

    top = parseInt(top, 10);
    right = (right || right === 0) ? parseInt(right, 10) : null;
    bottom = (bottom || bottom === 0) ? parseInt(bottom, 10) : null;
    left = (left || left === 0) ? parseInt(left, 10) : null;
    worksheetName = String(worksheetName || this.getSelectedWorksheet());
    workbookName = String(workbookName || this.getSelectedWorkbook());

    var bolReturnValue = false;

    if (isNaN(top)) {
        this.log.push("Margin property '" + top + "' is not a valid integer");
    } else if (right !== null && isNaN(right)) {
        this.log.push("Margin property '" + right + "' is not a valid integer");
    } else if (bottom !== null && isNaN(bottom)) {
        this.log.push("Margin property '" + bottom + "' is not a valid integer");
    } else if (left !== null && isNaN(left)) {
        this.log.push("Margin property '" + left + "' is not a valid integer");
    } else {
        bolReturnValue = this.ping(worksheetName, workbookName);

        if (bolReturnValue) {
            if (bottom === null && left === null) {
                if (right === null) {
                    bottom = left = right = top;
                } else {
                    bottom = top;
                    left = right;
                }
            } else if (left === null) {
                left = right;
            }

            try {
                this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).PageSetup.TopMargin = top;
                this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).PageSetup.RightMargin = right;
                this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).PageSetup.BottomMargin = bottom;
                this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).PageSetup.LeftMargin = left;
            } catch (err) {
                this.log.push(err.message);
                bolReturnValue = false;
            }
        }
    }

    if (!bolReturnValue) {
        this.log.unshift("Cannot change margin properties.");
    }

    return bolReturnValue;
};


ExcelBridge.prototype.setMerge = function (rangeRef, bolUnmerge) {
    this.log.reset();

    var temp;

    rangeRef = this.rangeReference(rangeRef);
    if (!rangeRef.isValid) {
        rangeRef = this.cellReference(rangeRef);
    }

    if (!rangeRef.isValid) {
        this.log.push(this.rangeReference(rangeRef).errorMsg);
    } else {
        temp = 'Range("' + rangeRef.name + '").' + ((bolUnmerge) ? 'UnMerge();' : 'Merge();');
        this.commandWrap(temp, true);
    }

    if (!this.log.isEmpty()) {
        if (bolUnmerge) {
            this.log.throwError('Unable to unmerge cells.');
        } else {
            this.log.throwError('Unable to merge cells.');
        }
    }

    this.log.throwError();
    return this;
};


/**
 *
 *
 */
ExcelBridge.prototype.setRange = function (rangeRef, valueArray, strFillValue) {
    //stacking ActiveX dictionaries does not produce a valid multi-dim VBArray
    this.log.reset();

    rangeRef = this.rangeReference(rangeRef);
    strFillValue = String(strFillValue || '');

    if (!rangeRef.isValid) {
        this.log.push(rangeRef.errorMsg);
    } else if (valueArray === undefined) {
        this.log.push("A defined value must be supplied for valueArray.");
    } else {
        try {
            var cellValue,
                i = 0,
                i2 = 0,
                rowRange,
                rowValue,
                VBArrayRow = new ActiveXObject('Scripting.Dictionary');

            this.ExcelApp.ScreenUpdating = false;

            for (i = 0; i < rangeRef.rowSpan; ++i) {
                if (Object.prototype.toString.call(valueArray) === '[object Array]') {
                    rowValue = (valueArray[i] || strFillValue);
                } else {
                    rowValue = String(valueArray || strFillValue);
                }
                for (i2 = 0; i2 < rangeRef.colSpan; ++i2) {
                    if (Object.prototype.toString.call(rowValue) === '[object Array]') {
                        cellValue = String(rowValue[i2] || strFillValue);
                    } else {
                        cellValue = String(rowValue || strFillValue);
                    }
                    VBArrayRow.add(i2 + 1, cellValue);
                }
                rowRange = this.rangeReference(
                    [rangeRef.leftCell.colIndex, rangeRef.leftCell.row + i],
                    [rangeRef.rightCell.colIndex, rangeRef.leftCell.row + i]
                );
                this.ExcelApp.Workbooks(this.getSelectedWorkbook()).Sheets(this.getSelectedWorksheet())
                    .Range(rowRange.name).Value2 = VBArrayRow.Items();
                VBArrayRow.RemoveAll();
            }
        } catch (err) {
            this.log.push(err);
        } finally {
            this.ExcelApp.ScreenUpdating = true;
        }
    }

    this.log.throwError("'Unable to set values for range '" + rangeRef.name + "'.");

    return this;
};


/**
 * Marks the instance, in a given workbook, and/or a given worksheet as synced.  When used with 
 * isModified() allows you to manually determine whether data has been 'modified' since it has
 * been 'synced'.
 * @param {string} worksheetName - The name of the worksheet to mark as synced.
 * @param {string} [workbookName=getSelectedWorkbook()] - The name of the workbook to mark as synced.
 * @returns {Boolean}
 * @example var bolResult = ExcelLink.setSynced('Sheet3', 'book1.xlsx');
 */
ExcelBridge.prototype.setSynced = function (worksheetName, workbookName) {
    if (!worksheetName && !workbookName) {
        this.hashTreeSynced = this.hashTreeCurrent;
    } else {
        workbookName = String(workbookName || this.getSelectedWorkbook());

        if (worksheetName) {
            worksheetName = String(worksheetName);
            if (!this.hashTreeCurrent[workbookName][workbookName]) {
                delete this.hashTreeSynced[workbookName][workbookName];
            } else {
                this.hashTreeSynced[workbookName][workbookName] = this.hashTreeCurrent[workbookName][workbookName];
            }
        } else {
            if (!this.hashTreeCurrent[workbookName][workbookName]) {
                delete this.hashTreeSynced[workbookName];
            } else {
                this.hashTreeSynced[workbookName] = this.hashTreeCurrent[workbookName];
            }
        }
    }
    return true;
};


/**
 * Sets the selected worksheet and workbook.
 * @param {string} worksheetName - The name of the worksheet to select.
 * @param {string} [workbookName=this.getSelectedWorkbook()] - The name of the workbook to select.
 * @example var bolResult = ExcelLink.setSelected('Sheet2');
 */
ExcelBridge.prototype.setSelected = function (worksheetName, workbookName) {
    this.log.reset();

    if (workbookName && (worksheetName === null || worksheetName === undefined)) {
        this.log.push('A worksheet must be specified when changing selected workbook.');
        this.log.setName('worksheet');
    } else if ((worksheetName === undefined || worksheetName === null)
            && (workbookName === undefined || workbookName === null)) {
        this.log.push('No workbook and/or worksheet was provided.');
        this.log.setName('workbook');
    } else {
        if (workbookName === null || workbookName === undefined) {
            workbookName = this.getSelectedWorkbook();
        }
        if (this.ping(worksheetName, workbookName)) {
            this.selectedWorkbookName = String(workbookName);
            this.selectedWorksheetName = String(worksheetName);
        } else {
            if (!workbookName) {
                this.log.unshift('Could not change selected worksheet.');
            } else {
                this.log.unshift('Could not change selected workbook/worksheet.');
            }
        }
    }

    this.log.throwError();
    return this;
};


/**
 * Toggle visibility on a worksheet, a workbook (leave worksheetName null), or the Excel Application
 * (leave both worksheetName and workbookName null).
 * @param {boolean} [bolValue=true] - Whether to make visible (true) or invisible (false).
 * @param {string} [worksheetName=null] - Worksheet to toggle (if any).
 * @param {string} [workbookName=null] - Workbook to toggle (if any).
 * @returns {Boolean|null}
 * @example var bolResult = ExcelLink.setVisible(false, 'Sheet3', 'book3.xlsx');
 */
ExcelBridge.prototype.setVisible = function (bolValue, worksheetName, workbookName) {
    this.log.reset();

    bolValue = (bolValue === undefined || bolValue === "" || bolValue === null) ? true : !!bolValue;

    var bolReturnValue = false,
        pingCheck,
        strTargetDescription = "";

    try {
        if (!worksheetName && !workbookName) {
            strTargetDescription = "Excel Application";
            this.ExcelApp.Application.Visible = bolValue;
            this.ExcelApp.ScreenUpdating = bolValue;
        } else if (worksheetName) {
            worksheetName = String(worksheetName);
            workbookName = String(workbookName || this.getSelectedWorkbook());
            strTargetDescription = "worksheet '" + worksheetName + "' in workbook '" + workbookName + "'";
            this.ExcelApp.Workbooks(workbookName).Worksheets(worksheetName).Visible = bolValue;
        } else {
            workbookName = String(workbookName);
            strTargetDescription = "workbook '" + workbookName + "'";
            this.ExcelApp.Workbooks(workbookName).Visible = bolValue;
        }
        bolReturnValue = true;
    } catch (err) {
        if (worksheetName) {
            pingCheck = this.ping(worksheetName, workbookName);
        } else if (workbookName) {
            pingCheck = this.ping(null, workbookName);
        } else {
            pingCheck = true;
            if (this.isClosed()) {
                this.log.push("Excel Application is not running.");
            } else if (!this.isReady()) {
                this.log.push("Excel Application is busy.");
                bolReturnValue = null;
            }
        }

        if (!pingCheck) {
            bolReturnValue = pingCheck;
        } else {
            if (workbookName && workbookName.toUpperCase() === this.getActiveWorkbook().toUpperCase()) {
                if (worksheetName && worksheetName.toUpperCase() === this.getActiveWorksheet().toUpperCase()) {
                    this.log.push("Cannot hide active worksheet.");
                } else {
                    this.log.push("Cannot hide active workbook.");
                }
            } else {
                this.log.push(err);
            }
        }

        if (bolValue) {
            this.log.unshift("Unable to make " + strTargetDescription + " visible.");
        } else {
            this.log.unshift("Unable to make " + strTargetDescription + " hidden.");
        }
    }

    return bolReturnValue;
};


/**
 * Terminate the bound Excel.exe instance
 * @returns {boolean}
 * @example var bolResult = ExcelLink.terminateInstance();
 */
ExcelBridge.prototype.terminateInstance = function () {
    this.log.reset();

    var returnValue = false,
        timeoutID = 0;

    if (this.isTerminated()) {
        this.log.push("Excel instance is already terminated.");
        returnValue = true;
    } else if (this.ExcelApp !== null) {
        try {
            this.ExcelApp.DisplayAlerts = false;
            this.ExcelApp.Application.Quit();
            this.ExcelApp = null;
            if (CollectGarbage !== undefined) {
                /*ignore jslint start*/
                timeoutID = window.setTimeout(function () { CollectGarbage(); }, 200);  // jshint ignore:line
                /*ignore jslint end*/
            }
            returnValue = true;
        } catch (err) {
            if (this.isReady()) {
                this.log.push("Excel is busy - try again later.");
            } else {
                this.log.push(err);
            }
        }
    }
    if (!returnValue) { this.log.unshift("Was unable to terminate Excel."); }

    return returnValue;
};


/**
 * Queues a document/workbook/worksheet for updating when internals are updated.  Leave both
 * parameters empty to queue all open excel documents
 * @param {string} [worksheetName=null] - The name of the workbook to queue.
 * @param {string} [workbookName=null] - The name of the workbook queue.
 * @returns {Boolean}
 * @example var bolResult = ExcelLink.queueCacheUpdate('Sheet4');
 */
ExcelBridge.prototype.queueCacheUpdate = function (worksheetName, workbookName) {
    var i = 0,
        i2 = 0,
        newValue,
        workbookArray = [],
        worksheetArray = [];

    if (worksheetName) {
        workbookArray = (workbookName) ? [workbookName] : [this.getSelectedWorkbook()];
        worksheetArray = [worksheetName];
    } else {
        workbookArray = (workbookName) ? [workbookName] : [this.getWorkbooks()];
    }

    for (i = 0; i < workbookArray.length; ++i) {
        if (!this.hashTreeCurrent[workbookArray[i]]) {
            this.hashTreeCurrent[workbookArray[i]] = {};
        }
        if (!worksheetArray.length) {
            worksheetArray = this.getWorksheets(workbookArray[i]);
        }
        for (i2 = 0; i2 < worksheetArray.length; ++i2) {
            newValue = (this.hashTreeCurrent[workbookArray[i]][worksheetArray[i2]] === true) ? true : null;
            this.hashTreeCurrent[workbookArray[i]][worksheetArray[i2]] = newValue;
        }
    }

    return true;
};


/**
 * Examines the instance's internal variables and updates when possible.  Should be either run
 * regularly with a timer or manually - not both!
 * @returns {boolean}
 * @example var bolResult = ExcelLink.updateInternals();
 */
ExcelBridge.prototype.updateInternals = function () {
    var i = 0,
        i2 = 0,
        newValue,
        prevValue,
        tempInterval = 0,
        tempTree = {},
        workbookArray = [],
        workbookCount = 0,
        workbookName = "",
        worksheetArray = [],
        worksheetCount = 0,
        worksheetName = "";

    if (this.isTerminated()) {
        if (this.TimerID) {
            this.currentCacheTree = {};
            clearInterval(this.TimerID);
            this.timerTickRate = 0;
        }
    } else {
        //Add worksheets to update queue
        if (this.timerTickRate) {
            //Being executed on a timer.
            if ((this.timerTickCount % 160) === 0) {
                //Rehash everything every 160 ticks
                this.queueCacheUpdate();
                this.timerTickCount = 0;
            } else {
                //If known changes rehash every 15 ticks, otherwise every 3
                tempInterval = (this.isModified(this.getActiveWorksheet(), this.getActiveWorkbook())) ? 15 : 3;
                if ((this.timerTickCount % 3) === 0) {
                    this.queueCacheUpdate(this.getActiveWorksheet(), this.getActiveWorkbook());
                }
            }
            this.timerTickCount++;
        } else {
            //Being executed manually without a timer.
            this.queueCacheUpdate();
        }

        //Rebuild cacheStoreCurrent to make sure created/deleted workbooks/worksheets are accounted for
        //Calculate missing checksums where needed
        workbookArray = this.getWorkbooks(true);
        if (workbookArray) {
            workbookCount = workbookArray.length;
            for (i = 0; i < workbookCount; ++i) {
                workbookName = workbookArray[i];
                tempTree[workbookName] = {};

                worksheetArray = this.getWorksheets(workbookName, true);
                if (worksheetArray) {
                    worksheetCount = worksheetArray.length;

                    for (i2 = 0; i2 < worksheetCount; ++i2) {
                        worksheetName = worksheetArray[i2];

                        if (this.hashTreeCurrent) {
                            prevValue = this.hashTreeCurrent[workbookName][worksheetName] || null;
                        }

                        newValue = prevValue || this.getWorksheetHash(worksheetName, workbookName);

                        tempTree[workbookName][worksheetName] = (newValue || prevValue);
                    }
                }
            }

            this.hashTreeCurrent = tempTree;
        }

        if (workbookArray) {
            //Remove any invalid references in selected workbook/worksheet
            if (this.hashTreeCurrent[this.getSelectedWorkbook()] === undefined) {
                this.selectedWorkbookName = null;
                this.selectedWorksheetName = null;
            } else if (this.hashTreeCurrent[this.getSelectedWorkbook()][this.getSelectedWorksheet()] === undefined) {
                this.selectedWorksheetName = null;
            }
        }
    }

    return true;
};


/**
 * Determine if a filename is valid windows filename.
 * @private
 *
 * @param {string} filename - The name of the filename to validate.
 *
 * @example var bolIsValid = this.validateWorksheetName('qwerty.txt');
 *
 * @returns {boolean}
 */
ExcelBridge.prototype.validateWindowsFilename = function (filename) {
    this.log.reset();

    var i = 0,
        prohibitedChars = ['<', '>', ':', '"', '/', String.fromCharCode(92), '|', '?', '*'],
        strlen = 0;

    if (typeof filename === 'boolean') {                                        //Verify isn't a boolean
        this.log.push('Windows filename cannot be a boolean value.');
    } else if (!filename && filename !== 0) {                                   //Verify isn't empty
        this.log.push('Windows filename cannot be empty.');
    } else if (typeof filename === 'object') {
        this.log.push('Windows filename cannot be an object reference.');       //Verify isn't an object
    } else {
        filename = String(filename).replace(/^\s+|\s+$/g, '');
        strlen = filename.length;

        if (strlen < 1) {
            this.log.push('Windows filename must contain at least one character.');
        } else if (strlen > 260) {
            this.log.push("Windows filename '" + filename + "' is too long - must be 260 characters or less.");
        } else {
            i = prohibitedChars.length;
            while (i--) {
                if (filename.indexOf(prohibitedChars[i]) !== -1) {
                    this.log.push("Windows filename '" + filename + "' is invalid, '" + prohibitedChars[i]
                        + "' is a prohibited character.");
                    break;
                }
            }
        }
    }

    return (this.log.isEmpty()) ? true : false;
};


/**
 * Determine if a worksheet name is valid.
 * @private
 *
 * @param {string} worksheetName - The name of the worksheet to validate.
 *
 * @example var bolIsValid = this.validateWorksheetName('qwerty');
 *
 * @returns {boolean}
 */
ExcelBridge.prototype.validateWorksheetName = function (worksheetName) {
    this.log.reset();

    var i = 0,
        prohibitedChars = ['/', String.fromCharCode(92), '*', '?', '[', ']'],
        strlen = 0;

    if (typeof worksheetName === 'boolean') {                               //Verify isn't a boolean
        this.log.push('Worksheet name cannot be a boolean value');
    } else if (!worksheetName && worksheetName !== 0) {                     //Verify isn't empty
        this.log.push('Worksheet name cannot be empty.');
    } else if (typeof worksheetName === 'object') {
        this.log.push('Worksheet name cannot be an object reference.');     //Verify isn't an object
    } else {
        worksheetName = String(worksheetName);
        strlen = worksheetName.length;

        if (strlen < 1) {
            this.log.push('Worksheet name must contain at least one character.');
        } else if (strlen > this.settingWorksheetNameMaxLength) {
            this.log.push("Worksheet name '" + worksheetName + "' is too long - must be "
                + this.settingWorksheetNameMaxLength + " characters or less.");
        } else {
            i = prohibitedChars.length;
            while (i--) {
                if (worksheetName.indexOf(prohibitedChars[i]) !== -1) {
                    this.log.push("Worksheet name '" + worksheetName + "' is invalid, '" + prohibitedChars[i]
                        + "' is a prohibited character.");
                    break;
                }
            }
        }
    }

    if (!this.log.isEmpty()) { this.log.setName('badworksheet'); }
    return (this.log.isEmpty()) ? true : false;
};


/////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////// FACTORIES ///////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////

var excelBridgeFactory = function (param) {
    var objTemp;

    try {
        objTemp = new ExcelBridge(param);

        if (objTemp.timerTickRate > 0) {
            objTemp.timerID = setInterval(function () { objTemp.updateInternals(); }, objTemp.timerTickRate);
        }

        return objTemp;
    } catch (err) {
        if (objTemp) {
            objTemp = null;
        }
        throw err;
    }
};

var excelBridgeFromFile = function (strUNCPath) {     // jshint ignore:line
    return excelBridgeFactory(strUNCPath);
};

var excelBridgeFromUserOpen = function () {           // jshint ignore:line
    return excelBridgeFactory('@userOpenFile');
};

var excelBridgeFromNew = function () {                // jshint ignore:line
    return excelBridgeFactory();
};
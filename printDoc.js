/* 
 * printDoc.js
 *
 * Copyright (c) 2021 Christopher Crawford
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

/*
 * Manually set your default printer to 
 * "Adobe PDF" or "Microsoft Print to PDF"
 * BEFORE using this script.
 *
 */

var GLOBALSHELL = new ActiveXObject("WScript.Shell");

function send(s)  { GLOBALSHELL.SendKeys(s); }
function sleep(n) { WScript.Sleep(n);        }

// https://stackoverflow.com/a/6630849
function Array.prototype.pop() {
  var len = this.length - 1;
  var value = this[len];
  this.length = len;
  delete this[len];
  return value;
}

// var p is the full path to the doc

// file paths in this script need escaped backslashes
// paths provided at the CLI do not

//var p = 'C:\\Users\\u\\Desktop\\test.docx';
var p = WScript.Arguments.Item(0);

var pArray = p.split("\\");
var f = pArray.pop();
var path = pArray.join("\\");
path = path + "\\";

var fArray = f.split(".");
var fBase = fArray[0];
var fExt = fArray[1];

GLOBALSHELL.Run('cmd /c start "" "' + p + '"', 9);
sleep(2000);
send("%f");
sleep(1000);
send("p");
send("p");
send(path + fBase + ".pdf");
send("%s");
sleep(2000);
send("%{F4}");
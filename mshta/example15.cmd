@set @E=1; /*
@echo Off
SETLOCAL
set DAYS=%1
if "%DAYS%" equ "" set DAYS=362

for /F "tokens=*" %%_ in ('cscript.exe //NoLogo /E:jscript "%~dpnx0" %DAYS% 1 ^| more') do set VALUE=%%_
set NEW_DATE=%VALUE%
REM altertive collect result as exit status
REM set /A NEW_DATE=%ErrorLevel%

echo NEW_DATE=%NEW_DATE%
exit /B




*/
var arguments = WScript.Arguments;

var o = new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1);
// https://stackoverflow.com/questions/563406/how-to-add-days-to-date
Date.prototype.addDays = function(x) {
  var d = new Date(this.valueOf());
  d.setDate(d.getDate() + x);
  return d
};
// https://stackoverflow.com/questions/23593052/format-javascript-date-as-yyyy-mm-dd
var n;
if (arguments.length > 0){
  n = parseInt(arguments(0));
} else {
  n = 1;
}
n = (arguments.length > 0) ? parseInt(arguments(0)) :1
var currentDate = new Date();

// NOTE toLocaleString() is available but is Locale specific
// o.Write(d.addDays(n).toLocaleString());

function pad(num){
  var text = '0' + num.toString();
  return text.substring(text.length - 2)
}

function fmt(y){
  return [y.getFullYear(),pad(y.getMonth()+1),pad(y.getDate())].join('.')
}

var result = fmt(currentDate.addDays(n));
o.Write( result + '\n');
// altertive return result as exit status
// Wscript.exit(result);
// https://stackoverflow.com/questions/23593052/format-javascript-date-as-yyyy-mm-dd

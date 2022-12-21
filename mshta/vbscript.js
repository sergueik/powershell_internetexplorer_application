// origin: http://with-love-from-siberia.blogspot.com/2009/12/msgbox-inputbox-in-jscript.html
// https://github.com/ildar-shaimordanov/jsxt
// advise from the original developer:
// To launch this code with success on 64-bit systems you need to run the 32-bit version of the scripting host located at "%windir%\SysWOW64".

// see also:
// https://cwestblog.com/2012/03/10/jscript-using-inputbox-and-msgbox/

// see also: 
// https://www.codeproject.com/Tips/1196971/InputBox-or-MessageBox-with-JavaScript
var vb = {};


vb.Function = function(func)
{
    return function()
    {
        return vb.Function.eval.call(this, func, arguments);
    };
};


vb.Function.eval = function(func)
{
    var args = Array.prototype.slice.call(arguments[1]);
    for (var i = 0; i < args.length; i++) {
        if ( typeof args[i] != 'string' ) {
            continue;
        }
        args[i] = '"' + args[i].replace(/"/g, '" + Chr(34) + "') + '"';
    }

    var vbe;
    vbe = new ActiveXObject('ScriptControl');
    vbe.Language = 'VBScript';

    return vbe.eval(func + '(' + args.join(', ') + ')');
};


/**
 *
 * InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
 *
 */
var InputBox = vb.Function('InputBox');


/**
 *
 * MsgBox(prompt[, buttons][, title][, helpfile, context])
 *
 */
var MsgBox = vb.Function('MsgBox');



var title, res, msg;

title = 'VBScript Emulating';

// The resulting string
res = InputBox('Enter a string', title);

// The message to be displayed
if ( res ) {
    msg = 'You have entered: "' + res + '".';
} else {
    msg = 'Nothing has been entered.';
}

// Displaying of the message
MsgBox(msg, 0, title);
function get_text_in_file(filepath) {
    var ProgID = "Scripting.FileSystemObject";
    var FS = new ActiveXObject(ProgID);
    return FS.OpenTextFile(filepath, 1).ReadAll(); // 1 = for reading
}

function require(filepath) {
    eval(get_text_in_file(filepath));
    return true;
}

// base.js sets up the goog.global namespace
var BASE_DIR = "closure-library-master\\closure\\goog\\";
eval(get_text_in_file(BASE_DIR + "base.js"));
require(BASE_DIR + "deps.js");
goog.global.CLOSURE_IMPORT_SCRIPT = require;
goog.basePath = BASE_DIR

goog.require("goog.crypt.base64");
var result = goog.crypt.base64.encodeString("TEST");
WScript.echo(result)

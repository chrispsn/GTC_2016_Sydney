function get_text_in_file(filepath) {
    var ProgID = "Scripting.FileSystemObject";
    var FS = new ActiveXObject(ProgID);
    return FS.OpenTextFile(filepath, 1).ReadAll(); // 1 = for reading
}

eval(get_text_in_file("handlebars-v4.0.5.js", 1));

var source = "{{person}} has a {{dog}}"
var template = Handlebars.compile(source)
var data = {person: "Bob", dog: "terrier"}
var output = template(data)

WScript.echo(output)

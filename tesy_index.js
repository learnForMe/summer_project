'use strict';

var PythonShell = require('python-shell');
var pyshell = new PythonShell('luck.py',{mode:'text '});
var readline = require('readline');
// sends a message to the Python script via stdin
var ttys = require('ttys');
var rl = readline.createInterface({
    input: ttys.stdin,
    output: ttys.stdout
});



/*
var rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});
*/
var name = rl.question ("This student is NEW -> ", (name) => {

//console.log("Student", name);

rl.close();

pyshell.send(name);

pyshell.on('message', function (message) {
  // received a message sent from the Python script (a simple "print" statement)
  console.log(message);
});

// end the input stream and allow the process to exit
pyshell.end(function (err) {
  if (err) throw err;
  console.log('Student added as', name );
});

});


/*
var options={
    mode: 'text',
    pythonpath: '/Users/garytsai/Desktop/rfid-reader-http/summer_project/luck.py',
    pythonOptios:['-u'],
    scriptPath:'/Users/garytsai/Desktop/rfid-reader-http/summer_project/',
    args: ["values"]
};


var stdin = process.openStdin();
stdin.setEncoding('utf8');
stdin.on('data', function(input){

    console.log(input);
})




PythonShell.run('luck.py', options, function (err, results) {
  if (err) throw err;
  // results is an array consisting of messages collected during execution
  console.log('results: %j', results);
});

*/

//var PythonShell = require('python-shell');
var execSync = require('child_process').execSync;
code = execSync('ls -la');

console.log('return code' + code);
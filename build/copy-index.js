var fs = require('fs-extra');
var root = __dirname + '/../';

fs.copySync(root + 'login.html', root + 'dist/login.html');

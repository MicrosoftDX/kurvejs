// node umdify.js module.js umd.js ModuleName
var fs = require('fs');

var regexModuleName = /MODULE_NAME/g;
var regexModule = /MODULE_SOURCE/g;
var fnModule = process.argv[2];
var fnTemplate = process.argv[3];
var stModuleName = process.argv[4];

fs.writeFileSync(fnModule,
    fs.readFileSync(fnTemplate, 'utf8')
        .replace(regexModuleName, stModuleName)
        .replace(regexModule, 
            fs.readFileSync(fnModule, 'utf8')),
    'utf8');
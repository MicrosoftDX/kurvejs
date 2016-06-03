// node typify.js typings.d.ts ModuleName
var fs = require('fs');

var fnTypings = process.argv[2];
var stModuleName = process.argv[3];

fs.writeFileSync(fnTypings,
    fs.readFileSync(fnTypings, 'utf8')
        .concat(`\nexport = ${stModuleName}\n`),
    'utf8');
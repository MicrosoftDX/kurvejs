// node typify.js typings.d.ts ModuleName
var fs = require('fs');

var fnTypings = process.argv[2];
var stModuleName = process.argv[3];

var stOriginal = fs.readFileSync(fnTypings, 'utf8')

fs.writeFileSync(fnTypings, stOriginal.concat(`\nexport = ${stModuleName};\n`), 'utf8');
fs.writeFileSync(fnTypings.replace(".d.ts", "-global.d.ts"), stOriginal, 'utf8');

var http = require('http');
var https = require('https');
var crypto = require('crypto');
var kurve = require('../../../dist/kurve');

var port = process.env.port || 1337;

http.createServer(function (req, res) {   
    var identity=new kurve.Identity("de50e8bc-7b32-4722-9705-3195d1ac942a", "http://localhost:1337", {
        appSecret:"9QLw+Xf5zK6cGP5I8xzRm1JqX1dWY1zFzVdfnWpaD4k=",
        mode:kurve.Mode.Node,
    });
    
    identity.handleNodeCallback(req, res, https, crypto, persistCallback, retrieveCallback).then(result => {
        if (result) {
           var graph=new kurve.Graph(identity, {https: https});
           graph.me.GetUser().then(function (result){
            
            res.writeHead(200, { 'Content-Type': 'text/plain' });
            res.end(result.displayName + '\n');
               
           });
        }        
    });

}).listen(port);

var memcache={};
function persistCallback(key,value,expiry){
    memcache[key]={
        value:value,
        expiry:expiry
    }
}

function retrieveCallback(key){
    return memcache[key]?memcache[key].value:null;    
}
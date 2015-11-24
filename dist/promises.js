// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var Kurve;
(function (Kurve) {
    // Based on https://www.npmjs.com/package/typescript-promises
    var Deferred = (function () {
        function Deferred() {
            this.doneCallbacks = [];
            this.failCallbacks = [];
            this.progressCallbacks = [];
            this._promise = new Promise(this);
            this._state = 'pending';
        }
        Object.defineProperty(Deferred.prototype, "promise", {
            get: function () {
                return this._promise;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Deferred.prototype, "state", {
            get: function () {
                return this._state;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Deferred.prototype, "rejected", {
            get: function () {
                return this.state === 'rejected';
            },
            set: function (rejected) {
                this._state = rejected ? 'rejected' : 'pending';
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Deferred.prototype, "resolved", {
            get: function () {
                return this.state === 'resolved';
            },
            set: function (resolved) {
                this._state = resolved ? 'resolved' : 'pending';
            },
            enumerable: true,
            configurable: true
        });
        Deferred.prototype.resolve = function () {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i - 0] = arguments[_i];
            }
            args.unshift(this);
            return this.resolveWith.apply(this, args);
        };
        Deferred.prototype.resolveWith = function (context) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            this._result = args;
            this.doneCallbacks.forEach(function (callback) {
                callback.apply(context, args);
            });
            this.doneCallbacks = [];
            this.resolved = true;
            return this;
        };
        Deferred.prototype.reject = function () {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i - 0] = arguments[_i];
            }
            args.unshift(this);
            return this.rejectWith.apply(this, args);
        };
        Deferred.prototype.rejectWith = function (context) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            this.failCallbacks.forEach(function (callback) {
                callback.apply(context, args);
            });
            this.failCallbacks = [];
            this.rejected = true;
            return this;
        };
        Deferred.prototype.progress = function () {
            var _this = this;
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            var d = new Deferred();
            if (this.resolved || this.rejected) {
                callbacks.forEach(function (callback) {
                    callback.apply(_this._notifyContext, _this._notifyArgs);
                });
                return d;
            }
            callbacks.forEach(function (callback) {
                _this.progressCallbacks.push(_this.wrap(d, callback, d.notify));
            });
            this.checkStatus();
            return d;
        };
        Deferred.prototype.notify = function () {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i - 0] = arguments[_i];
            }
            args.unshift(this);
            return this.notifyWith.apply(this, args);
        };
        Deferred.prototype.notifyWith = function (context) {
            var args = [];
            for (var _i = 1; _i < arguments.length; _i++) {
                args[_i - 1] = arguments[_i];
            }
            if (this.resolved || this.rejected) {
                return this;
            }
            this._notifyContext = context;
            this._notifyArgs = args;
            this.progressCallbacks.forEach(function (callback) {
                callback.apply(context, args);
            });
            return this;
        };
        Deferred.prototype.checkStatus = function () {
            if (this.resolved) {
                this.resolve.apply(this, this._result);
            }
            else if (this.rejected) {
                this.reject.apply(this, this._result);
            }
        };
        Deferred.prototype.then = function (doneFilter, failFilter, progressFilter) {
            var d = new Deferred();
            this.progressCallbacks.push(this.wrap(d, progressFilter, d.progress));
            this.doneCallbacks.push(this.wrap(d, doneFilter, d.resolve));
            this.failCallbacks.push(this.wrap(d, failFilter, d.reject));
            this.checkStatus();
            return this;
        };
        Deferred.prototype.wrap = function (d, f, method) {
            return function () {
                var args = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    args[_i - 0] = arguments[_i];
                }
                var result = f.apply(f, args);
                if (result && result instanceof Promise) {
                    result.then(function () { d.resolve(); }, function () { d.reject(); });
                }
                else {
                    method.apply(d, [result]);
                }
            };
        };
        Deferred.prototype.done = function () {
            var _this = this;
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            var d = new Deferred();
            callbacks.forEach(function (callback) {
                _this.doneCallbacks.push(_this.wrap(d, callback, d.resolve));
            });
            this.checkStatus();
            return d;
        };
        Deferred.prototype.fail = function () {
            var _this = this;
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            var d = new Deferred();
            callbacks.forEach(function (callback) {
                _this.failCallbacks.push(_this.wrap(d, callback, d.reject));
            });
            this.checkStatus();
            return d;
        };
        Deferred.prototype.always = function () {
            var _this = this;
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            var d = new Deferred();
            callbacks.forEach(function (callback) {
                _this.doneCallbacks.push(_this.wrap(d, callback, d.resolve));
                _this.failCallbacks.push(_this.wrap(d, callback, d.reject));
            });
            this.checkStatus();
            return d;
        };
        return Deferred;
    })();
    Kurve.Deferred = Deferred;
    var Promise = (function () {
        function Promise(deferred) {
            this.deferred = deferred;
        }
        Promise.prototype.then = function (doneFilter, failFilter, progressFilter) {
            return this.deferred.then(doneFilter, failFilter, progressFilter).promise;
        };
        Promise.prototype.done = function () {
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            return this.deferred.done.apply(this.deferred, callbacks).promise;
        };
        Promise.prototype.fail = function () {
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            return this.deferred.fail.apply(this.deferred, callbacks).promise;
        };
        Promise.prototype.always = function () {
            var callbacks = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                callbacks[_i - 0] = arguments[_i];
            }
            return this.deferred.always.apply(this.deferred, callbacks).promise;
        };
        Object.defineProperty(Promise.prototype, "resolved", {
            get: function () {
                return this.deferred.resolved;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Promise.prototype, "rejected", {
            get: function () {
                return this.deferred.rejected;
            },
            enumerable: true,
            configurable: true
        });
        return Promise;
    })();
    Kurve.Promise = Promise;
    Kurve.when = function () {
        var args = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            args[_i - 0] = arguments[_i];
        }
        if (args.length === 1) {
            var arg = args[0];
            if (arg instanceof Deferred) {
                return arg.promise;
            }
            if (arg instanceof Promise) {
                return arg;
            }
        }
        var done = new Deferred();
        if (args.length === 1) {
            done.resolve(args[0]);
            return done.promise;
        }
        var pending = args.length;
        var results = [];
        var onDone = function () {
            var resultArgs = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                resultArgs[_i - 0] = arguments[_i];
            }
            results.push(resultArgs);
            if (--pending === 0) {
                done.resolve.apply(done, results);
            }
        };
        var onFail = function () {
            done.reject();
        };
        args.forEach(function (a) {
            a.then(onDone, onFail);
        });
        return done.promise;
    };
})(Kurve || (Kurve = {}));
//*********************************************************   
//   
//Kurve js, https://github.com/matvelloso/kurve-js
//  
//Copyright (c) Microsoft Corporation  
//All rights reserved.   
//  
// MIT License:  
// Permission is hereby granted, free of charge, to any person obtaining  
// a copy of this software and associated documentation files (the  
// ""Software""), to deal in the Software without restriction, including  
// without limitation the rights to use, copy, modify, merge, publish,  
// distribute, sublicense, and/or sell copies of the Software, and to  
// permit persons to whom the Software is furnished to do so, subject to  
// the following conditions:  
// The above copyright notice and this permission notice shall be  
// included in all copies or substantial portions of the Software.  
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,  
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND  
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE  
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION  
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
//   
//*********************************************************   

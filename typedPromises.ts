///<reference path="promises.ts" />

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
module Kurve { 
    export class TypedDeferred<T> {
        private _promise: TypedPromise<T>;
        private _wrappedDeferred: Deferred;

        constructor() {
            this._promise = new TypedPromise<T>(this);
            this._wrappedDeferred = new Deferred();
        }

        get promise(): TypedPromise<T> {
            return this._promise;
        }

        get state(): string {
            return this._wrappedDeferred.state;
        }

        get rejected(): boolean {
            return this._wrappedDeferred.rejected;
        }

        set rejected(rejected: boolean) {
            this._wrappedDeferred.rejected = rejected;
        }

        get resolved(): boolean {
            return this._wrappedDeferred.resolved;
        }

        set resolved(resolved: boolean) {
            this._wrappedDeferred.resolved = resolved;
        }

        resolve(...args: any[]): TypedDeferred<T> {
            args.unshift(this);
            return this.resolveWith.apply(this, args);
        }

        resolveWith(context: any, ...args: any[]): TypedDeferred<T> {
            this._wrappedDeferred.resolveWith(context, ...args);
            return this;
        }

        reject(...args: any[]): TypedDeferred<T> {
            args.unshift(this);
            return this.rejectWith.apply(this, args);
        }

        rejectWith(context: any, ...args: any[]): TypedDeferred<T> {
            this._wrappedDeferred.rejectWith(context, ...args);
            return this;
        }

        progress(...callbacks: Function[]): TypedDeferred<T> {
            var d = new TypedDeferred<T>();
            d._wrappedDeferred = this._wrappedDeferred.progress(...callbacks);
            return d;
        }

        notify(...args: any[]): TypedDeferred<T> {
            args.unshift(this);
            return this.notifyWith.apply(this, args);
        }

        notifyWith(context: any, ...args: any[]): TypedDeferred<T> {
            this._wrappedDeferred.notifyWith(context, ...args);
            return this;
        }


        then(doneFilter: Function, failFilter?: Function, progressFilter?: Function): TypedDeferred<T> {
            this._wrappedDeferred.then(doneFilter, failFilter, progressFilter);
            return this;
        }

        done(...callbacks: Function[]): TypedDeferred<T> {
            var d = new TypedDeferred<T>();
            d._wrappedDeferred = this._wrappedDeferred.done(...callbacks);
            return d;
        }

        fail(...callbacks: Function[]): TypedDeferred<T> {
            var d = new TypedDeferred<T>();
            d._wrappedDeferred = this._wrappedDeferred.fail(...callbacks);
            return d;
        }

        always(...callbacks: Function[]): TypedDeferred<T> {
            var d = new TypedDeferred<T>();
            d._wrappedDeferred = this._wrappedDeferred.always(...callbacks);
            return d;
        }
    }


    export class TypedPromise<T>  {
        constructor (protected deferred: TypedDeferred<T>) {
            
        }

        then(doneFilter: (T) => void, failFilter?: Function, progressFilter?: Function): TypedPromise<T> {
            return this.deferred.then(doneFilter, failFilter, progressFilter).promise as TypedPromise<T>;
        }

        done(...callbacks: ((T) => void)[]): TypedPromise<T> {
            return (<TypedDeferred<T>>this.deferred.done.apply(this.deferred, callbacks)).promise as TypedPromise<T>;
        }

        fail(...callbacks: Function[]): TypedPromise<T> {
            return (<TypedDeferred<T>>this.deferred.fail.apply(this.deferred, callbacks)).promise as TypedPromise<T>
        }

        always(...callbacks: Function[]): TypedPromise<T> {
            return (<TypedDeferred<T>>this.deferred.always.apply(this.deferred, callbacks)).promise as TypedPromise<T>
        }

        get resolved(): boolean {
            return this.deferred.resolved;
        }

        get rejected(): boolean {
            return this.deferred.rejected;
        }

    }
}


//*********************************************************   
//   
//Kurve js, https://github.com/microsoftdx/kurvejs
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

module kurve {
    
// Adapted from the original source: https://github.com/DirtyHairy/typescript-deferred



  export class Error {
        public status: number;
        public statusText: string;
        public text: string;
        public other: any;
    }
    
    function DispatchDeferred(closure: () => void) {
        setTimeout(closure, 0);
    }

     enum PromiseState { Pending, ResolutionInProgress, Resolved, Rejected }

    export interface PromiseCallback<T> {
        (error: Error, result?: T): void;
    }

    class Client {
        constructor(
            private _dispatcher: (closure: () => void) => void,
            private _successCB: any,
            private _errorCB: any
        ) {
            this.result = new Deferred<any, any>(_dispatcher);
        }

        resolve(value: any, defer: boolean): void {
            if (typeof (this._successCB) !== 'function') {
                this.result.resolve(value);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._successCB, value));
            } else {
                this._dispatchCallback(this._successCB, value);
            }
        }

        reject(error: any, defer: boolean): void {
            if (typeof (this._errorCB) !== 'function') {
                this.result.reject(error);
                return;
            }

            if (defer) {
                this._dispatcher(() => this._dispatchCallback(this._errorCB, error));
            } else {
                this._dispatchCallback(this._errorCB, error);
            }
        }

        private _dispatchCallback(callback: (arg: any) => any, arg: any): void {
            var result: any,
                then: any,
                type: string;

            try {
                result = callback(arg);
                this.result.resolve(result);
            } catch (err) {
                this.result.reject(err);
                return;
            }
        }

        result: Deferred<any, any>;
    }

    export class Deferred<T, E>  {
        private _dispatcher: (closure: () => void)=> void;

        constructor();
        constructor(dispatcher: (closure: () => void) => void);
        constructor(dispatcher?: (closure: () => void) => void) {
            if (dispatcher)
                this._dispatcher = dispatcher;
            else
                this._dispatcher = DispatchDeferred;
            this.promise = new Promise<T, E>(this);
        }

        private DispatchDeferred(closure: () => void) {
            setTimeout(closure, 0);
        }

        then(successCB: any, errorCB: any): any {
            if (typeof (successCB) !== 'function' && typeof (errorCB) !== 'function') {
                return this.promise;
            }

            var client = new Client(this._dispatcher, successCB, errorCB);

            switch (this._state) {
                case PromiseState.Pending:
                case PromiseState.ResolutionInProgress:
                    this._stack.push(client);
                    break;

                case PromiseState.Resolved:
                    client.resolve(this._value, true);
                    break;

                case PromiseState.Rejected:
                    client.reject(this._error, true);
                    break;
            }

            return client.result.promise;
        }

        resolve(value?: T): Deferred<T, E>;

        resolve(value?: Promise<T, E>): Deferred<T, E>;

        resolve(value?: any): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._resolve(value);
        }

        private _resolve(value: any): Deferred<T, E> {
            var type = typeof (value),
                then: any,
                pending = true;

            try {
                if (value !== null &&
                    (type === 'object' || type === 'function') &&
                    typeof (then = value.then) === 'function') {
                    if (value === this.promise) {
                        throw new TypeError('recursive resolution');
                    }

                    this._state = PromiseState.ResolutionInProgress;
                    then.call(value,
                        (result: any): void => {
                            if (pending) {
                                pending = false;
                                this._resolve(result);
                            }
                        },
                        (error: any): void => {
                            if (pending) {
                                pending = false;
                                this._reject(error);
                            }
                        }
                    );
                } else {
                    this._state = PromiseState.ResolutionInProgress;

                    this._dispatcher(() => {
                        this._state = PromiseState.Resolved;
                        this._value = value;

                        var i: number,
                            stackSize = this._stack.length;

                        for (i = 0; i < stackSize; i++) {
                            this._stack[i].resolve(value, false);
                        }

                        this._stack.splice(0, stackSize);
                    });
                }
            } catch (err) {
                if (pending) {
                    this._reject(err);
                }
            }

            return this;
        }

        reject(error?: E): Deferred<T, E> {
            if (this._state !== PromiseState.Pending) {
                return this;
            }

            return this._reject(error);
        }

        private _reject(error?: any): Deferred<T, E> {
            this._state = PromiseState.ResolutionInProgress;

            this._dispatcher(() => {
                this._state = PromiseState.Rejected;
                this._error = error;

                var stackSize = this._stack.length,
                    i = 0;

                for (i = 0; i < stackSize; i++) {
                    this._stack[i].reject(error, false);
                }

                this._stack.splice(0, stackSize);
            });

            return this;
        }

        promise: Promise<T, E>;

        private _stack: Array<Client> = [];
        private _state = PromiseState.Pending;
        private _value: T;
        private _error: any;
    }

    export class Promise<T, E>  {
        constructor(private _deferred: Deferred<T, E>) { }

        then<R>(
            successCallback?: (result: T) => R,
            errorCallback?: (error: E) => R
        );

        then(successCallback: any, errorCallback: any): any {
            return this._deferred.then(successCallback, errorCallback);
        }

        fail<R>(
            errorCallback?: (error: E) => R
        );

        fail(errorCallback: any): any {
            return this._deferred.then(undefined, errorCallback);
        }
    }

} //remove during bundling


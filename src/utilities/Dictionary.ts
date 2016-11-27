export interface IDictionary {
    add(key: string, value: any): void;
    remove(key: string): void;
    containsKey(key: string): boolean;
    keys(): string[];
    values(): any[];
}

export class Dictionary {

    private _keys: string[] = new Array<string>();
    private _values: any[] = new Array<any>();

    constructor(init: { key: string; value: any; }[]) {

        for (var x = 0; x < init.length; x++) {
            this[init[x].key] = init[x].value;
            this._keys.push(init[x].key);
            this._values.push(init[x].value);
        }
    }

    public add(key: string, value: any) {
        this[key] = value;
        this._keys.push(key);
        this._values.push(value);
    }

    public remove(key: string) {
        var index = this._keys.indexOf(key, 0);
        this._keys.splice(index, 1);
        this._values.splice(index, 1);

        delete this[key];
    }

    public keys(): string[] {
        return this._keys;
    }

    public values(): any[] {
        return this._values;
    }

    public containsKey(key: string) {
        if (typeof this[key] === "undefined") {
            return false;
        }

        return true;
    }

    public toLookup(): IDictionary {
        return this;
    }
}
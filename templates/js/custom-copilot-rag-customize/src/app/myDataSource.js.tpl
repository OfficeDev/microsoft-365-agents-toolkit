const path = require("path");
const fs = require("fs");

/**
 * A data source that searches through a local directory of files for a given query.
 */
class MyDataSource {
    /**
     * Creates a new instance of the MyDataSource instance.
     */
    constructor(name) {
        this.name = name;
    }

    /**
     * Initializes the data source.
     */
    init() {
        const filePath = path.join(__dirname, "../data");
        const files = fs.readdirSync(filePath);
        this._data = files.map(file => {
            const data = 
            {
                content:fs.readFileSync(path.join(filePath, file), "utf-8"),
                citation:file
            };
            return data;
        });
    }

    /**
     * Renders the data source as a string of text.
     */
    async renderData(context, memory, tokenizer, maxTokens) {
        const query = memory.getValue("temp.input");
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        for (let data of this._data) {
            if (data.content.includes(query)) {
                return { output: this.formatDocument(`${data.content}\n Citation title:${data.citation}`), length: data.content.length, tooLong: false };
            }
        }
        if (query.toLocaleLowerCase().includes("perksplus")) {
            return { output: this.formatDocument(`${this._data[0].content}\n Citation title:${this._data[0].citation}`), length: this._data[0].content.length, tooLong: false };
        } else if (query.toLocaleLowerCase().includes("company") || query.toLocaleLowerCase().includes("history")) {
            return { output: this.formatDocument(`${this._data[1].content}\n Citation title:${this._data[1].citation}`), length: this._data[1].content.length, tooLong: false };
        } else if (query.toLocaleLowerCase().includes("northwind") || query.toLocaleLowerCase().includes("health") || query.toLocaleLowerCase().includes("plan")) {
            return { output: this.formatDocument(`${this._data[2].content}\n Citation title:${this._data[2].citation}`), length: this._data[2].content.length, tooLong: false };
        }
        return { output: "", length: 0, tooLong: false };
    }

    /**
     * Formats the result string 
     */
    formatDocument(result) {
        return `<context>${result}</context>`;
    }
}

module.exports = {
  MyDataSource,
};
import * as JSZip from "jszip";
import { saveAs } from "file-saver";

interface IContents {
    [file: string]: string;
}

export class ExportPackage {
    private _contents: IContents = {};
    constructor(public packageName: string) {

    }

    public addOrUpdateFile(fileName: string, content: string) {
        this._contents[fileName] = content;
    }

    public get allFiles(): string[] {
        return Object.keys(this._contents);
    }

    public getFileContent(fileName: string): string {
        return this._contents[fileName];
    }

    public hasContent(fileName: string): boolean {
        return !!this._contents[fileName];
    }

    private get hasSingleFile(): boolean {
        return this.allFiles.length == 1;
    }

    public async download(): Promise<void> {
        if (this.hasSingleFile) {
            const fileName = this.allFiles[0];
            const fileContent = this._contents[fileName];
            const blob = new Blob([fileContent], { type: "octet/steam" });
            saveAs(blob, fileName);
        }
        else {
            const zip = new JSZip();
            this.allFiles.forEach(f => {
                zip.file(f, this._contents[f]);
            });
            await zip.generateAsync({ type: "blob" })
                .then((content) => {
                    saveAs(content, `${this.packageName}.zip`);
                });
        }
    }
}
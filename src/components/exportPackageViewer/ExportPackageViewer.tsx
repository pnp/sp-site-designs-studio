import * as React from "react";
import { ExportPackage } from "../../helpers/ExportPackage";
import styles from "./ExportPackageViewer.module.scss";
import { useState, useEffect } from "react";
import CodeEditor from "@monaco-editor/react";
import { Icon, ActionButton } from "office-ui-fabric-react";

export interface IExportPackageViewerProps {
    exportPackage: ExportPackage;
}

export const ExportPackageViewer = (props: IExportPackageViewerProps) => {

    const [currentFile, setCurrentFile] = useState(props.exportPackage.allFiles && props.exportPackage.allFiles.length && props.exportPackage.allFiles[0]);

    useEffect(() => {
        const allFiles = props.exportPackage.allFiles;
        if (allFiles && allFiles.length > 0) {
            setCurrentFile(allFiles[0]);
        }
    }, [props.exportPackage]);

    const viewFileContent = (file: string) => {
        setCurrentFile(file);
    };

    const getContentLanguage = (fileName: string) => {
        if (!fileName) {
            return null;
        }

        const fileNameParts = fileName.split(".");
        const extension = fileNameParts.length > 1 ? fileNameParts[fileNameParts.length - 1].toLowerCase() : null;

        switch (extension) {
            case "json":
                return "json";
            case "ps1":
                return "powershell";
            default:
                return "";
        }
    };

    const currentContent = props.exportPackage.hasContent(currentFile)
        ? props.exportPackage.getFileContent(currentFile)
        : "";
    return <div className={styles.ExportPackageViewer}>
        <div className={styles.row}>
            <div className={styles.column}>
                <h3><Icon styles={{
                    root: {
                        fontSize: 24,
                        verticalAlign: "text-bottom",
                        marginRight: 10
                    }
                }} iconName="Package" />{(props.exportPackage.packageName) || ""}</h3>
            </div></div>
        <div className={styles.row}>
            <div className={styles.column2}>
                <ul className={styles.filesList}>
                    {props.exportPackage.allFiles.map(f => <li>
                        <ActionButton iconProps={{ iconName: "TextDocument" }}
                            styles={{
                                root: { fontWeight: currentFile == f ? "bold" : "normal" }
                            }}
                            checked={currentFile == f} text={f} onClick={() => viewFileContent(f)} />
                    </li>)}
                </ul>
            </div>
            <div className={styles.column10}>
                <CodeEditor
                    height="65vh"
                    language={getContentLanguage(currentFile)}
                    options={{
                        readOnly: true,
                        folding: true,
                        renderIndentGuides: true,
                        minimap: {
                            enabled: false
                        }
                    }}
                    value={currentContent}
                />
            </div>
        </div>
    </div>;
};
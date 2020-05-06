import * as React from "react";
import { ActionType } from "../../../app/IApplicationAction";
import { useAppContext } from "../../../app/App";
import { IApplicationState } from "../../../app/ApplicationState";
import schema from "../../../schema/schema";
export const Debugger = (props: any) => {
    const [appContext, action] = useAppContext<IApplicationState, ActionType>();

    return <>
        <h1>Debugger</h1>
        <h2>App Context</h2>
        <textarea>
            {JSON.stringify({ ...appContext, serviceScope: null, componentContext: null })}
        </textarea>

        <h2>Schema</h2>
        <textarea>
            {JSON.stringify(schema)}
        </textarea>
    </>;
};
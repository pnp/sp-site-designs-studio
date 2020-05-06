import * as React from "react";
import { useReducer, useContext, Dispatch } from "react";

export interface IAction<TActionType> {
    type: TActionType;
}

export interface IAppReactContext<TAppState, TActionType> {
    applicationState: TAppState;
    dispatch: React.Dispatch<TActionType>;
}


let appReactContext: React.Context<any> = null;
const getAppReactContext: <TAppState, TActionType>() => React.Context<IAppReactContext<TAppState, TActionType>> =
    <TAppState, TActionType>() => {
        if (!appReactContext) {
            appReactContext = React.createContext({} as IAppReactContext<TAppState, TActionType>);
        }
        return appReactContext as React.Context<IAppReactContext<TAppState, TActionType>>;
    };

export interface IBaseAppState<TActionType> { }

export interface IAppProps<TAppState extends IBaseAppState<TAction>, TAction> {
    children?: any;
    applicationState: TAppState;
    reducers: (appState: TAppState, action: any) => TAppState;
}


export const useAppContext: <TAppState, TActionType>() => [TAppState, (actionType: TActionType, actionArgs: any) => void] = <TAppState, TActionType>() => {
    const ctx = getAppReactContext<TAppState, { type: TActionType }>();
    const appContext = useContext(ctx);
    return [appContext.applicationState, (actionType, actionArgs) => appContext.dispatch({ type: actionType, ...actionArgs } as any)];
};

export function App<TAppState, TActionType>(props: IAppProps<TAppState, TActionType>) {

    const [applicationState, dispatch] = useReducer(props.reducers, props.applicationState);
    const ctx = getAppReactContext<TAppState, TActionType>();
    return <ctx.Provider value={{ applicationState, dispatch }}>
        {props.children}
    </ctx.Provider>;
}
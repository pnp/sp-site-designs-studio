import { useRef, useEffect } from "react";

export const usePrevious = <T extends {}>(value: T): T | undefined => {
    const ref = useRef<T>();
    useEffect(() => {
        ref.current = value;
    });
    return ref.current;
};

export const useCompare = (val: any) => {
    const prevVal = usePrevious(val);
    return prevVal !== val;
};

export const useTraceUpdate = (componentName: string, props: any) => {
    if (DEBUG) {
        console.debug(`Rendering ${componentName}`);
        const prev = useRef(props);
        useEffect(() => {
            const changedProps = Object.entries(props).reduce((ps, [k, v]) => {
                if (prev.current[k] !== v) {
                    ps[k] = [prev.current[k], v];
                }
                return ps;
            }, {});
            if (Object.keys(changedProps).length > 0) {
                console.debug(`[${componentName}]::[Changed props]: `, changedProps);
            }
            prev.current = props;
        });
    }
};


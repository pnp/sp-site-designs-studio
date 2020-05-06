import { IPrincipal } from "../models/IPrincipal";

const GROUP_PRINCIPAL_NAME_PREFIX = "c:0o.c|federateddirectoryclaimprovider|";
const USER_PRINCIPAL_NAME_PREFIX = "i:0#.f|membership|";

export function getPrincipalTypeFromName(principalName: string): "group" | "user" {
    if (principalName.indexOf(GROUP_PRINCIPAL_NAME_PREFIX) == 0) {
        return "group";
    } else if (principalName.indexOf(USER_PRINCIPAL_NAME_PREFIX) == 0) {
        return "user";
    } else {
        console.warn("Cannot resolve principal type from specified principal name");
        return null;
    }
}

export function getUserLoginNameFromPrincipalName(principalName: string): string {
    if (getPrincipalTypeFromName(principalName) == "user") {
        return principalName.replace(USER_PRINCIPAL_NAME_PREFIX, "");
    }
    console.warn("Specified principal name is not a valid user principal name");
    return null;
}

export function getPrincipalAlias(principal: IPrincipal): string {
    switch (principal.type) {
        case "group":
            // Only for ensured groups
            return principal.principalName;
        case "user":
            return principal.principalName.replace(USER_PRINCIPAL_NAME_PREFIX, "").split('@')[0];
        default:
            return null;
    }
}

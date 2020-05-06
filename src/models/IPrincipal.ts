export interface IPrincipal {
    id: string;
    displayName: string;
    principalName: string;
    type: "user" | "group";
}
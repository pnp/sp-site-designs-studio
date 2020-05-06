export function toJSON(object: any): string {
    if (!object) {
        return '';
    }
    return JSON.stringify(object, null, 4);
}
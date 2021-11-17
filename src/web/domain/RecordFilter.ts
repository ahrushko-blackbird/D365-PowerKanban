export interface RecordFilter {
    selected?: boolean;
    logicalName?: string;
    displayName?: string;
    operator?: "equals" | "contains";
}
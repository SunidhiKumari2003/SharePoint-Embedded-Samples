export interface IColumn {
    description: string;
    displayName: string;
    id: string;
    indexed: boolean;
    isDeletable: boolean;
    name: string;
    type: string;
    text: JSON;
    number: JSON;
    datetime: Date;
}
declare interface IStrings {
  PropertyPaneDescription: string;
  PropertyPaneAdvancedDescription: string;
  BasicGroupName: string;
  AdvancedGroupName: string;
  QueryFieldLabel: string;
  FieldsTitleLabel: string;
  FieldsFieldLabel: string;
  FieldsTemplateLabel: string;
  FieldsMaxResults: string;
  FieldsSorting: string;
  Fieldsfiltering: string;
  QueryInfoDescription: string;
  FieldsExternalLabel: string;
  FieldsExternalTempLabel: string;
}

declare module 'mystrings' {
  const strings: IStrings;
  export = strings;
}

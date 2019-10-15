declare interface IHelloWorldWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TextToEcho: string;
}

declare module "HelloWorldWebPartStrings" {
  const strings: IHelloWorldWebPartStrings;
  export = strings;
}

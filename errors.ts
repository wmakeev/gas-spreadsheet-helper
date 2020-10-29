namespace SpreadsheetHelper.Errors {
  export class SpreadsheetHelperError extends Error {
    constructor(
      message: string,
      public sheetName?: string,
      public sheetRow?: number,
      public sheetColumn?: number
    ) {
      super(message)
      this.name = this.constructor.name
    }
  }
}

export class IronsheetError extends Error {
  constructor(
    message: string,
    readonly code: string
  ) {
    super(message);
    this.name = new.target.name;
  }
}

export class ZipError extends IronsheetError {
  constructor(message: string) {
    super(message, "ZIP_ERROR");
  }
}

export class PackageError extends IronsheetError {
  constructor(message: string) {
    super(message, "PACKAGE_ERROR");
  }
}

export class WorkbookError extends IronsheetError {
  constructor(message: string) {
    super(message, "WORKBOOK_ERROR");
  }
}

export class WorksheetError extends IronsheetError {
  constructor(message: string) {
    super(message, "WORKSHEET_ERROR");
  }
}

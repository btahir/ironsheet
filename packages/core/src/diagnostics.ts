export type DiagnosticSeverity = "info" | "warning" | "error";

export type Diagnostic = {
  severity: DiagnosticSeverity;
  code: string;
  message: string;
  part?: string;
};

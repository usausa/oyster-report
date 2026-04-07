namespace OysterReport.Generator.Models;

internal sealed class ReportDiagnostic
{
    public ReportDiagnosticSeverity Severity { get; init; } // Severity level

    public string Code { get; init; } = string.Empty; // Diagnostic code

    public string Message { get; init; } = string.Empty; // Diagnostic message for the user

    public string? SheetName { get; init; } // Associated sheet name (null if not sheet-specific)

    public string? CellAddress { get; init; } // Associated cell address (null if not cell-specific)
}

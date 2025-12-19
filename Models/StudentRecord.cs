// Models/StudentRecord.cs

namespace Playwrighter.Models;

public class StudentRecord
{
    public string StudentNo { get; set; } = string.Empty;
    public string Decision { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Programme { get; set; } = string.Empty;
    public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;
    public string ErrorMessage { get; set; } = string.Empty;
}

public enum ProcessingStatus
{
    Pending,
    Processing,
    Success,
    Failed,
    Skipped
}

public enum DecisionType
{
    Accept,
    Reject,
    Unknown
}

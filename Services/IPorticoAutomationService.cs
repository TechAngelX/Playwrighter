// Services/IPorticoAutomationService.cs

using Playwrighter.Models;

namespace Playwrighter.Services;

public interface IPorticoAutomationService
{
    event EventHandler<string>? StatusUpdated;
    event EventHandler<StudentRecord>? StudentProcessed;
    
    Task InitialiseAsync(AppConfig config);
    Task<bool> LoginAsync();
    Task<bool> NavigateToUclSelectAsync();
    Task ProcessStudentAcceptAsync(StudentRecord student);
    Task ProcessStudentRejectAsync(StudentRecord student);
    Task CloseAsync();
    bool IsInitialised { get; }
}

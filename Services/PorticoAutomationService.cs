// Services/PorticoAutomationService.cs

using Microsoft.Playwright;
using Playwrighter.Models;

namespace Playwrighter.Services;

public class PorticoAutomationService : IPorticoAutomationService
{
    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private IBrowserContext? _context;
    private IPage? _page;
    private AppConfig? _config;

    public bool DebugMode { get; set; } = false;

    private readonly Dictionary<string, string> _shortToLongProgCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "AIBH", "TMSARTSINT03" },
        { "AISD", "TMSARTSINT02" },
        { "AIDE", "TMSCOMSSAD18" },
        { "ISEC", "TMSCOMSINF01" },
        { "CF",   "TMSCOMSCFI01" },
        { "FRM",  "TMSCOMSFRM01" },
        { "FT",   "TMSFINSTEC01" },
        { "EDT",  "TMSCOMSEDT01" },
        { "ML",   "TMSCOMSMCL01" },
        { "DSML", "TMSDATSMLE01" },
        { "CSML", "TMSCOMSSML01" },
        { "RAI",  "TMSROBAARI01" },
        { "SEIOT","TMSCOMSEIT01" },
        { "DDI",  "TMSCOMSDDI19" },
        { "CS",   "TMSCOMSING01" },
        { "SSE",  "TMSCOMSSSE01" },
        { "CGVI", "TMSCOMSCGV01" }
    };

    public event EventHandler<string>? StatusUpdated;
    public event EventHandler<StudentRecord>? StudentProcessed;

    public bool IsInitialised => _page != null;

    public async Task InitialiseAsync(AppConfig config)
    {
        _config = config;
        LogStatus("Initialising Playwright...");
        _playwright = await Playwright.CreateAsync();

        var userDataDir = config.EdgeUserDataDir;
        if (string.IsNullOrEmpty(userDataDir))
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            userDataDir = Path.Combine(appDataPath, "Playwrighter", "EdgeProfile");
            Directory.CreateDirectory(userDataDir);
        }

        var contextOptions = new BrowserTypeLaunchPersistentContextOptions
        {
            Headless = config.HeadlessMode,
            Channel = "msedge",
            SlowMo = config.ActionDelayMs,
            AcceptDownloads = true,
            ViewportSize = null,
            Args = new[] { "--start-maximized" }
        };

        _context = await _playwright.Chromium.LaunchPersistentContextAsync(userDataDir, contextOptions);
        _page = _context.Pages.FirstOrDefault() ?? await _context.NewPageAsync();
        LogStatus("Browser initialised.");
    }

    public async Task<bool> LoginAsync()
    {
        if (_page == null || _config == null) throw new InvalidOperationException("Service not initialised.");
        LogStatus($"Navigating to Portico: {_config.PorticoUrl}");
        await _page.GotoAsync(_config.PorticoUrl);
        
        try {
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 5000 });
            LogStatus("Session valid. Already logged in.");
            return true;
        } catch { LogStatus("Session check: Login required."); }

        var staffLoginButton = _page.GetByRole(AriaRole.Button, new() { Name = "Staff and Students Login" });
        if (await staffLoginButton.IsVisibleAsync()) await staffLoginButton.ClickAsync();

        LogStatus("Waiting for manual SSO/MFA authentication...");
        try {
            await _page.WaitForSelectorAsync("text=My Portico", new PageWaitForSelectorOptions { Timeout = 240000 });
            LogStatus("Successfully logged in to Portico.");
            return true;
        } catch { return false; }
    }

    public async Task<bool> NavigateToUclSelectAsync()
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        
        LogStatus("Navigating to UCLSelect...");
        var uclSelectLink = _page.Locator("text=UCLSelect").First;
        if (await uclSelectLink.IsVisibleAsync()) 
        {
            await uclSelectLink.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }

        LogStatus("Clicking Search tab...");
        var searchTab = _page.Locator("a").Filter(new() { HasText = "Search" }).First;
        if (await searchTab.IsVisibleAsync()) 
        {
            await searchTab.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }
        
        LogStatus("Ready to search.");
        return true;
    }

    public async Task ProcessStudentAcceptAsync(StudentRecord student)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);
        
        try
        {
            LogStatus($"Processing OFFER for: {student.StudentNo} (Prog: '{student.Programme}')");
            await SearchForStudentAsync(student.StudentNo);
            await ClickStudentLinkAsync(student); 
            await NavigateToActionsTabAsync();
            await RecommendOfferAsync();
            
            student.Status = ProcessingStatus.Success;
            LogStatus($"SUCCESS: Offer processed for {student.StudentNo}");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"FAILED {student.StudentNo}: {ex.Message}");
        }
        StudentProcessed?.Invoke(this, student);
    }
    
    public async Task ProcessStudentRejectAsync(StudentRecord student)
    {
        if (_page == null) throw new InvalidOperationException("Service not initialised.");
        student.Status = ProcessingStatus.Processing;
        StudentProcessed?.Invoke(this, student);
        
        try
        {
            LogStatus($"Processing REJECT for: {student.StudentNo} (Prog: '{student.Programme}')");
            await SearchForStudentAsync(student.StudentNo);
            await ClickStudentLinkAsync(student); 
            await NavigateToActionsTabAsync();
            await RecommendRejectAsync();
            
            student.Status = ProcessingStatus.Success;
            LogStatus($"SUCCESS: Rejection processed for {student.StudentNo}");
        }
        catch (Exception ex)
        {
            student.Status = ProcessingStatus.Failed;
            student.ErrorMessage = ex.Message;
            LogStatus($"FAILED {student.StudentNo}: {ex.Message}");
        }
        StudentProcessed?.Invoke(this, student);
    }

    private async Task SearchForStudentAsync(string studentNo)
    {
        if (_page == null) return;
        
        LogStatus($"Searching: {studentNo}");
        
        var radioLabel = _page.Locator("text=Student Number").First;
        if (await radioLabel.IsVisibleAsync()) await radioLabel.ClickAsync();

        ILocator? searchInput = null;
        var textboxes = _page.GetByRole(AriaRole.Textbox);
        if (await textboxes.CountAsync() > 0) searchInput = textboxes.First;
        if (searchInput == null || !await searchInput.IsVisibleAsync())
             searchInput = _page.Locator("input[type='text']").First;

        if (searchInput != null && await searchInput.IsVisibleAsync()) 
        {
            await searchInput.ClickAsync();
            await searchInput.ClearAsync();
            await searchInput.FillAsync(studentNo);
        } 
        else 
        {
            throw new Exception("Could not find search input field.");
        }
        
        var searchBtn = _page.Locator("input[value='Search']").First;
        if (!await searchBtn.IsVisibleAsync()) 
        {
            searchBtn = _page.Locator("button").Filter(new() { HasText = "Search" }).First;
        }
        
        await searchBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(500);
    }

    private async Task ClickStudentLinkAsync(StudentRecord student)
    {
        if (_page == null) return;
        
        LogStatus($"Looking for student {student.StudentNo} with Prog '{student.Programme}'");

        try {
            await _page.WaitForSelectorAsync("table", new PageWaitForSelectorOptions { Timeout = 10000 });
            await Task.Delay(500); // Let table fully render
        } catch {
            throw new Exception("Search results table did not appear.");
        }

        string inputProg = student.Programme?.Trim() ?? "";
        string searchCode = "";

        if (!string.IsNullOrEmpty(inputProg))
        {
            if (_shortToLongProgCodes.TryGetValue(inputProg, out var longCode))
            {
                searchCode = longCode;
                LogStatus($"Mapped '{inputProg}' -> '{searchCode}'");
            }
            else
            {
                searchCode = inputProg;
            }
        }
        else
        {
            throw new Exception("Programme column is empty.");
        }

        // New approach: Find all links and check their parent row for the prog code
        var allLinks = await _page.Locator("table tbody tr td a").AllAsync();
        LogStatus($"Found {allLinks.Count} links in result rows");
        
        ILocator? targetLink = null;
        
        for (int i = 0; i < allLinks.Count; i++)
        {
            var link = allLinks[i];
            
            // Get the parent row of this link using XPath
            var parentRow = link.Locator("xpath=ancestor::tr[1]");
            var rowHtml = await parentRow.InnerHTMLAsync();
            var rowText = await parentRow.InnerTextAsync();
            
            // Check if this row has BOTH the student number AND the prog code
            bool hasStudentNo = rowText.Contains(student.StudentNo);
            bool hasProgCode = rowText.Contains(searchCode);
            
            var linkText = await link.InnerTextAsync();
            var href = await link.GetAttributeAsync("href") ?? "";
            
            LogStatus($"Link {i}: '{linkText.Trim()}' | StudentNo={hasStudentNo} | ProgCode={hasProgCode}");
            
            if (hasStudentNo && hasProgCode)
            {
                LogStatus($">>> Found matching link! href ends with: ...{href.Substring(Math.Max(0, href.Length - 50))}");
                targetLink = link;
                break;
            }
        }
        
        if (targetLink == null)
        {
            throw new Exception($"Could not find link in row with StudentNo='{student.StudentNo}' AND ProgCode='{searchCode}'");
        }
        
        LogStatus("Clicking the matched link...");
        await targetLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
    }

    private async Task NavigateToActionsTabAsync()
    {
        if (_page == null) return;
        
        LogStatus("Clicking Actions tab...");
        var actionsTab = _page.Locator("text=Actions").First;
        await actionsTab.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(300);
    }

    private async Task RecommendOfferAsync()
    {
        if (_page == null) return;
        
        LogStatus("Clicking 'Recommend Offer or Reject'...");
        var recommendLink = _page.Locator("text=Recommend Offer or Reject").First;
        await recommendLink.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(300);

        LogStatus("Selecting 'Offer recommendation'...");
        var offerRadio = _page.Locator("text=Offer recommendation").First;
        await offerRadio.ClickAsync();
        await Task.Delay(200);

        LogStatus("Clicking Process...");
        var processBtn = _page.Locator("input[value='Process']").First;
        await processBtn.ClickAsync();
        await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        await Task.Delay(500);
        
        LogStatus("Offer recommendation processed.");
    }
    
  private async Task RecommendRejectAsync()
  {
      if (_page == null) return;
      
      LogStatus("Clicking 'Recommend Offer or Reject'...");
      var recommendLink = _page.Locator("a").Filter(new() { HasText = "Recommend Offer or Reject" }).First;
      if (!await recommendLink.IsVisibleAsync())
      {
          recommendLink = _page.Locator("text=Recommend Offer or Reject").First;
      }
      await recommendLink.ClickAsync();
      await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
      await Task.Delay(1000);
  
      LogStatus("Selecting 'Reject' radio button...");
      var radioButtons = await _page.Locator("input[type='radio']").AllAsync();
      LogStatus($"Found {radioButtons.Count} radio buttons");
      
      if (radioButtons.Count >= 2)
      {
          await radioButtons[1].ClickAsync();
          LogStatus("Clicked Reject radio button (index 1)");
      }
      else
      {
          var rejectRadio = _page.GetByLabel("Reject");
          await rejectRadio.ClickAsync();
          LogStatus("Clicked Reject radio button (by label)");
      }
      
      await Task.Delay(2000);
      
      LogStatus("Looking for dropdown elements...");
      var allSelects = await _page.Locator("select").AllAsync();
      LogStatus($"Total <select> elements found: {allSelects.Count}");
      
      if (allSelects.Count < 2)
      {
          throw new Exception($"CRITICAL: Need at least 2 dropdowns, found: {allSelects.Count}");
      }
      
      var jsCode = @"
          () => {
              const selects = document.querySelectorAll('select');
              if (selects.length < 2) return 'ERROR: Need at least 2 dropdowns';
              
              const select = selects[1];
              let debugInfo = 'Reason 1 dropdown:\n';
              
              for (let i = 0; i < select.options.length; i++) {
                  debugInfo += '  ' + i + ': ' + select.options[i].text + '\n';
              }
              
              let foundOption = null;
              
              for (let i = 0; i < select.options.length; i++) {
                  const text = select.options[i].text;
                  if ((text.startsWith('8.') || text.startsWith('8 ')) && 
                      text.toLowerCase().includes('not competitive')) {
                      foundOption = i;
                      break;
                  }
              }
              
              if (foundOption === null) {
                  for (let i = 0; i < select.options.length; i++) {
                      const text = select.options[i].text.toLowerCase();
                      if (text.includes('not competitive') && text.includes('oversubscribed')) {
                          foundOption = i;
                          break;
                      }
                  }
              }
              
              if (foundOption !== null) {
                  select.selectedIndex = foundOption;
                  select.value = select.options[foundOption].value;
                  select.dispatchEvent(new Event('change', { bubbles: true }));
                  
                  return 'SUCCESS: Selected option ' + foundOption + ': ' + select.options[foundOption].text;
              }
              
              return 'FAILED: Could not find option 8\n' + debugInfo;
          }
      ";
      
      var jsResult = await _page.EvaluateAsync<string>(jsCode);
      
      LogStatus("=== DROPDOWN SELECTION RESULT ===");
      LogStatus(jsResult);
      LogStatus("================================");
      
      if (!jsResult.Contains("SUCCESS"))
      {
          throw new Exception("CRITICAL: Failed to select option 8!\n" + jsResult);
      }
      
      await Task.Delay(1000);
  
      if (DebugMode)
      {
          LogStatus("⏸️⏸️⏸️ DEBUG MODE: PAUSED ⏸️⏸️⏸️");
          LogStatus("⏸️ Check the browser NOW:");
          LogStatus("⏸️   1. Is 'Reject' selected?");
          LogStatus("⏸️   2. Does 'Reason 1' show option 8?");
          LogStatus("⏸️ Process button will NOT be clicked");
          return;
      }
  
      LogStatus("Clicking Process button...");
      var processBtn = _page.Locator("input[value='Process']").First;
      if (!await processBtn.IsVisibleAsync())
      {
          processBtn = _page.Locator("button").Filter(new() { HasText = "Process" }).First;
      }
      await processBtn.ClickAsync();
      await _page.WaitForLoadStateAsync(LoadState.NetworkIdle);
      await Task.Delay(500);
      
      LogStatus("Rejection processed.");
  }
    public async Task CloseAsync()
    {
        LogStatus("Closing browser...");
        if (_context != null) 
        { 
            await _context.CloseAsync(); 
            _context = null; 
        }
        _playwright?.Dispose(); 
        _playwright = null;
        _page = null;
    }

    private void LogStatus(string message) => StatusUpdated?.Invoke(this, message);
}

using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using Playwrighter.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Playwrighter.Views;

public partial class ProcessingWindow : Window
{
    private ObservableCollection<ProcessingStudentViewModel> _students = new();
    private int _totalStudents;
    private int _processedCount;

    // Explicitly define the controls to satisfy the compiler
    private ProgressBar ProgressBar = null!;
    private TextBlock ProgressText = null!;
    private ItemsControl StudentsList = null!;
    private ScrollViewer StudentsScrollViewer = null!;
    private TextBox StatusLog = null!;
    private TextBlock FooterStatus = null!;
    private TextBlock SubtitleText = null!;
    private Button CancelButton = null!;
    private Button CloseButton = null!;
    
    public ProcessingWindow()
    {
        InitializeComponent();
        StudentsList.ItemsSource = _students;
        
        // Attach event handlers
        CancelButton.Click += (s, e) => OnCancelRequested();
        CloseButton.Click += (s, e) => Close();
    }
    
    private void InitializeComponent()
    {
        // Load the XAML
        AvaloniaXamlLoader.Load(this);

        // Manually find the controls by their x:Name
        ProgressBar = this.FindControl<ProgressBar>("ProgressBar")!;
        ProgressText = this.FindControl<TextBlock>("ProgressText")!;
        StudentsList = this.FindControl<ItemsControl>("StudentsList")!;
        StudentsScrollViewer = this.FindControl<ScrollViewer>("StudentsScrollViewer")!;
        StatusLog = this.FindControl<TextBox>("StatusLog")!;
        FooterStatus = this.FindControl<TextBlock>("FooterStatus")!;
        SubtitleText = this.FindControl<TextBlock>("SubtitleText")!;
        CancelButton = this.FindControl<Button>("CancelButton")!;
        CloseButton = this.FindControl<Button>("CloseButton")!;
    }
    
    public event EventHandler? CancelRequested;
    
    public void Initialize(List<StudentRecord> students)
    {
        _totalStudents = students.Count;
        _processedCount = 0;
        
        _students.Clear();
        
        foreach (var student in students)
        {
            var nameParts = (student.Name ?? "").Split(new[] { ' ' }, 2);
            var fName = nameParts.Length > 0 ? nameParts[0] : "";
            var sName = nameParts.Length > 1 ? nameParts[1] : "";

            _students.Add(new ProcessingStudentViewModel
            {
                StudentNo = student.StudentNo,
                Forename = fName,
                Surname = sName,
                Decision = student.Decision,
                StatusIcon = "⏳",
                StatusText = "Pending",
                StatusColor = "#A0AEC0"
            });
        }
        
        UpdateProgress();
    }
    
    public void UpdateStudentStatus(string studentNo, ProcessingStatus status, string? errorMessage = null)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var student = _students.FirstOrDefault(s => s.StudentNo == studentNo);
            if (student == null) return;
            
            switch (status)
            {
                case ProcessingStatus.Processing:
                    student.StatusIcon = "⚙";
                    student.StatusText = "Processing";
                    student.StatusColor = "#4299E1";
                    break;
                
                case ProcessingStatus.Success:
                    student.StatusIcon = "✓";
                    student.StatusText = "Done";
                    student.StatusColor = "#48BB78";
                    _processedCount++;
                    break;
                
                case ProcessingStatus.Failed:
                    student.StatusIcon = "✗";
                    student.StatusText = "Failed";
                    student.StatusColor = "#E53E3E";
                    _processedCount++;
                    break;
            }
            
            UpdateProgress();
            ScrollToCurrentStudent(studentNo);
        });
    }
    
    public void LogMessage(string message)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            StatusLog.Text += $"[{timestamp}] {message}\n";
            StatusLog.CaretIndex = StatusLog.Text?.Length ?? 0;
        });
    }
    
    public void UpdateFooterStatus(string status)
    {
        Dispatcher.UIThread.Post(() =>
        {
            FooterStatus.Text = status;
        });
    }
    
    public void ProcessingComplete()
    {
        Dispatcher.UIThread.Post(() =>
        {
            var successCount = _students.Count(s => s.StatusText == "Done");
            var failedCount = _students.Count(s => s.StatusText == "Failed");
            
            SubtitleText.Text = $"Complete: {successCount} successful, {failedCount} failed";
            FooterStatus.Text = "Processing complete!";
            
            CancelButton.IsVisible = false;
            CloseButton.IsVisible = true;
        });
    }
    
    private void UpdateProgress()
    {
        var percentage = _totalStudents > 0 ? (_processedCount * 100.0 / _totalStudents) : 0;
        ProgressBar.Value = percentage;
        ProgressText.Text = $"{_processedCount} / {_totalStudents} ({percentage:F0}%)";
    }
    
    private void ScrollToCurrentStudent(string studentNo)
    {
        var index = _students.ToList().FindIndex(s => s.StudentNo == studentNo);
        if (index >= 0)
        {
            var itemHeight = 60;
            var scrollPosition = index * itemHeight;
            StudentsScrollViewer.Offset = new Avalonia.Vector(0, scrollPosition);
        }
    }
    
    private void OnCancelRequested()
    {
        CancelRequested?.Invoke(this, EventArgs.Empty);
    }
}

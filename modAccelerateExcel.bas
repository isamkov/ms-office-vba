Attribute VB_Name = "modAccelerateExcel"
'-------------------------------------------------------------------------
' Source: risksir.com
' Description: Acceleration of vba calculations by disabling slow services
'-------------------------------------------------------------------------


'To accelerate, use "Call AccelerateExcel" in the beginning of your function
 Public Sub AccelerateExcel()
 
  'Don't update screen
  Application.ScreenUpdating = False
 
  'Disable automatic calculation
  Application.Calculation = xlCalculationManual
 
  'Don't show page breaks
  If Workbooks.Count Then
      ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
  End If
 
  'Disable events
  Application.EnableEvents = False
  
  'Don't show status bar
  Application.DisplayStatusBar = False
 
  'Disable alerts
  Application.DisplayAlerts = False
 
 End Sub
 
'Don't forget to enable disabled services by including "Call disAccelerateExcel" at the end of your function
Public Sub disAccelerateExcel()
 
  'Turn on screen update
  Application.ScreenUpdating = True
 
  'Turn on automatic calculation
  Application.Calculation = xlCalculationAutomatic
  
  'Turn on page breaks
  If Workbooks.Count Then
      ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True
  End If
 
  'Turn on events
  Application.EnableEvents = True
 
  'Turn on status bar
  Application.DisplayStatusBar = True
 
  'Turn on alerts
  Application.DisplayAlerts = True
 
End Sub

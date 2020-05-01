
Imports System
Imports System.Diagnostics

Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI

Imports Rstyx.LoggingConsole
Imports Rstyx.Excel.ActionsNET.UI

''' <summary> Integrates LoggingConsole and provides LoggingConsole related public actions, accessible by VBA. </summary>
Public Module LoggingConsole
    
    Dim Logger          As Logger
    Dim LoggerName      As String = "Actions.NET"
    'Dim LogViewer       As WpfHostForm
    Dim LogViewerDock   As CustomTaskPane
    
    #Region "Public Actions"
        
        ''' <summary> Writes a error message to LoggingConsole. </summary>
         ''' <param name="Message"> The message to log. </param>
        <ExcelFunction(Description:="Schreibt eine Fehler-Nachricht in die Protokoll-Konsole")>
        Public Sub LoggingConsoleLogError(Message As String)
            Logger.logError(Message)
        End Sub
        
        ''' <summary> Writes a warning message to LoggingConsole. </summary>
         ''' <param name="Message"> The message to log. </param>
        <ExcelFunction(Description:="Schreibt eine Warnungs-Nachricht in die Protokoll-Konsole")>
        Public Sub LoggingConsoleLogWarning(Message As String)
            Logger.logWarning(Message)
        End Sub
        
        ''' <summary> Writes a info message to LoggingConsole. </summary>
         ''' <param name="Message"> The message to log. </param>
        <ExcelFunction(Description:="Schreibt eine Info-Nachricht in die Protokoll-Konsole")>
        Public Sub LoggingConsoleLogInfo(Message As String)
            Logger.logInfo(Message)
        End Sub
        
        ''' <summary> Writes a debug message to LoggingConsole. </summary>
         ''' <param name="Message"> The message to log. </param>
        <ExcelFunction(Description:="Schreibt eine Debug-Nachricht in die Protokoll-Konsole")>
        Public Sub LoggingConsoleLogDebug(Message As String)
            Logger.logDebug(Message)
        End Sub
        
        ''' <summary> Shows LoggingConsole. </summary>
        <ExcelFunction(Description:="Zeigt die Protokoll-Konsole an")>
        Public Sub LoggingConsoleShow()
            LogBox.Instance.showFloatingConsoleView(suppressErrorOnFail:=False)
        End Sub
        
        ''' <summary> Hides LoggingConsole if shown. </summary>
        <ExcelFunction(Description:="Versteckt die Protokoll-Konsole")>
        Public Sub LoggingConsoleHide()
            LogBox.Instance.hideFloatingConsoleView()
        End Sub
        
    #End Region
    
    #Region "Initialization and Integration of LoggingConsole"
        
        ''' <summary> Configures <see cref="Rstyx.LoggingConsole.LogBox"/> to use Excel dock panel (Custom Task Pane) as floating window. </summary>
         ''' <remarks> This will be called at AddIn startup ... </remarks>
        Friend Sub LoggingConsoleInit()
        
            'LogBox.Instance.DisplayName = LogBox.Instance.DisplayName & " (Excel)"
            
            ' Office 365 (x64) hides the custom task pane whenever a workbook is open.
            ' So, use LoggingConsole's built-in window.
            'LogBox.Instance.ShowFloatingConsoleViewAction = AddressOf ShowExcelConsole
            'LogBox.Instance.HideFloatingConsoleViewAction = AddressOf HideExcelConsole
        
            Logger = LogBox.getLogger(LoggerName)
        End Sub
        
        '' <summary> Shows the <see cref="Rstyx.LoggingConsole.ConsoleView"/> in a window as Excel's child. </summary>
        'Private Sub showExcelConsole2()
            'If ((LogViewer Is Nothing) OrElse (LogViewer.IsDisposed)) Then
            '    
            '    LogViewer = New WpfHostForm()
            '    'LogViewer.Size = My.Settings.LoggingConsoleSize
            '    'LogViewer.Location = My.Settings.LoggingConsoleLocation
            '    LogViewer.Text = LogBox.Instance.DisplayName & " (Excel)"
            '    LogViewer.WpfHost.Child = LogBox.Instance.Console.ConsoleView
            'End If
            '
            'If (Not LogViewer.Visible) Then
            '    Dim Hwnd As IntPtr = Process.GetCurrentProcess().MainWindowHandle
            '    
            '    If (Hwnd = IntPtr.Zero) Then
            '        LogViewer.Show()
            '    Else
            '        LogViewer.Show(New Win32Window(Hwnd))
            '    End If
            'End If
            
        'End Sub
        
        ''' <summary> Shows the <see cref="Rstyx.LoggingConsole.ConsoleView"/> in a Excel dock window (Custom Task Pane). </summary>
        Private Sub ShowExcelConsole()
            
            If (IsLogViewerDockAlive()) Then
                LogViewerDock.Visible = True
            Else
                Try
                    LogViewerDock = CustomTaskPaneFactory.CreateCustomTaskPane(GetType(WpfHostUserControl), LogBox.Instance.DisplayName)
                    
                    Dim ucLogViewer As WpfHostUserControl = CType(LogViewerDock.ContentControl, WpfHostUserControl)
                    ucLogViewer.WpfHost.Child = LogBox.Instance.Console.ConsoleView
                    
                    LogViewerDock.Visible = True
                    
                    Dim RecentDockPosition As MsoCTPDockPosition
                    If ([Enum].TryParse(Of MsoCTPDockPosition)(My.Settings.LoggingConsoleDockPosition, RecentDockPosition)) Then
                        LogViewerDock.DockPosition = RecentDockPosition
                    Else
                        LogViewerDock.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom
                    End If
                    Try
                        Select Case LogViewerDock.DockPosition
                            Case MsoCTPDockPosition.msoCTPDockPositionBottom, MsoCTPDockPosition.msoCTPDockPositionTop
                                LogViewerDock.Height = My.Settings.LoggingConsoleSize.Height
                            Case MsoCTPDockPosition.msoCTPDockPositionLeft, MsoCTPDockPosition.msoCTPDockPositionRight
                                LogViewerDock.Width  = My.Settings.LoggingConsoleSize.Width
                            Case MsoCTPDockPosition.msoCTPDockPositionFloating
                                LogViewerDock.Height = My.Settings.LoggingConsoleSize.Height
                                LogViewerDock.Width  = My.Settings.LoggingConsoleSize.Width
                        End Select
                    Finally
                        'AddHandler LogViewerDock.DockPositionStateChange, AddressOf ctp_DockPositionStateChange
                        AddHandler LogViewerDock.VisibleStateChange, AddressOf CTP_VisibleStateChange
                    End Try
                Catch ex As System.Exception
                    System.Diagnostics.Debug.Print(ex.ToString())
                End Try 
            End If
        End Sub
        
        ''' <summary> Hides the Excel docking window, which is showing LoggingConsole. </summary>
        Private Sub HideExcelConsole()
            If IsLogViewerDockAlive() Then LogViewerDock.Visible = False
        End Sub
        
        ''' <summary> Tells if the Excel docking window, which is showing LoggingConsole, is alive (hence not disposed). </summary>
        Private Function IsLogViewerDockAlive() As Boolean
            Dim RetValue As Boolean = False
            
            If (LogViewerDock IsNot Nothing) Then
                Try
                    Dim dummy As Boolean = LogViewerDock.Visible
                    RetValue = True
                Catch ex As Exception
                    ' Dock has been disposed.
                End Try
            End If
            
            Return RetValue
        End Function
        
        ''' <summary> Handles VisibleStateChange event. </summary>
        ''' <param name="ctp"> The CustomTaskPane. </param>
        ''' <remarks> Saves dock panel settings if it's just closed. </remarks>
        Private Sub CTP_VisibleStateChange(ctp As CustomTaskPane)
            If (Not ctp.Visible) Then
                My.Settings.LoggingConsoleDockPosition = ctp.DockPosition
                My.Settings.LoggingConsoleSize = New System.Drawing.Size(ctp.Width, ctp.Height)
                My.Settings.Save()
            End If
        End Sub
        
        Private Sub CTP_DockPositionStateChange(ctp As CustomTaskPane)
            ' Size can't be changed while event handler is active!
            If (ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating) Then
                Try
                    ctp.Height = 500
                    ctp.Width  = 500
                Catch ex As System.Exception
                    System.Diagnostics.Debug.Print(ex.ToString())
                End Try 
            End If
        End Sub
        
    #End Region
    
End Module

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:

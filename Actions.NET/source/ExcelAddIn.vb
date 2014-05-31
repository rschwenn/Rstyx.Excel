
Imports ExcelDna.Integration


Class ExcelAddIn
    Implements IExcelAddIn
    
    
    Public Sub  Start() Implements IExcelAddIn.AutoOpen
        'XL = ExcelDna.Integration.ExcelDnaUtil.Application
        LoggingConsoleInit()
        LoggingConsoleLogDebug("Loaded Excel-XLL-AddIn:  " & System.Reflection.Assembly.GetExecutingAssembly().FullName)
    End sub
    
    Public Sub  Close() Implements IExcelAddIn.AutoClose
        'Fires when addin is removed from the addins list
        'but not when excel closes - this is to
        'avoid issues caused by the Excel option to cancel
        ' out of the close     'after the event has fired.
        'msgbox ("Bye bye, from Actions.NET")
    End sub

End Class

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:

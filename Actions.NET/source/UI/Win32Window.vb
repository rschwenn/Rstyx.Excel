
Imports System

Namespace UI
    
    ''' <summary> A pure implementation of <see cref="System.Windows.Forms.IWin32Window"/> </summary>
     ''' <remarks>
     ''' Usage: Create a new instance passing a window handle (i.e. got by <c>Process.GetCurrentProcess().MainWindowHandle</c>) to the constructor.
     ''' </remarks>
    Public Class Win32Window
        Implements System.Windows.Forms.IWin32Window
        
        ''' <summary> Creates a new instance of <c>Win32Window</c>, that stores the <paramref name="handle"/>. </summary>
         ''' <param name="handle"> Valid Window handle. </param>
         ''' <remarks>             The handle may be retrieved by <c>Process.GetCurrentProcess().MainWindowHandle</c>. </remarks>
        Public Sub New(handle As System.IntPtr)
            _Handle = handle
        End Sub
        
        Private _Handle As IntPtr
        
        ''' <summary> Delivers the window handle. </summary>
         ''' <returns> The stored window handle. </returns>
        Public ReadOnly Property Handle As IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _Handle
            End Get
        End Property
        
    End Class
    
End Namespace

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:

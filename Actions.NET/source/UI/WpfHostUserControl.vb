
Imports System.Runtime.InteropServices

Namespace UI
    
    ''' <summary> A Forms UserControl that hosts WPF content. </summary>
     ''' <remarks> 
     ''' <para>
     ''' A <see cref="System.Windows.Forms.UserControl"/>, which is intended
     ''' to show a WPF UserControl in a host application's window.
     ''' </para>
     ''' <para>
     ''' The Forms UserControl contains one single <see cref="System.Windows.Forms.Integration.ElementHost"/> 
     ''' control that completely filles the whole Forms UserControl. It will match the WPF UserControl's size.
     ''' </para>
     ''' <para>
     ''' This class is intended for use by the ExcelDNA's <c>CustomTaskPaneFactory.CreateCustomTaskPane()</c> 
     ''' </para>
     ''' <para>
     ''' For ExcelDNA: Would need to be marked with [ComVisible(true)] if in a project that is marked as [assembly:ComVisible(false)] which is the default for VS projects. 
     ''' </para>
     ''' </remarks>
    <ComVisible(True)>
    Public Class WpfHostUserControl
        Inherits System.Windows.Forms.UserControl
        
        #Region "Private Fields"
            
            'Private _WpfHost    As New System.Windows.Forms.Integration.ElementHost
            Friend WithEvents _WpfHost  As New System.Windows.Forms.Integration.ElementHost
            
        #End Region
        
        #Region "Constuctors and Finalizers"
            
            Public Sub New()
                InitializeComponent()
            End Sub
            
            ''' <summary> Disposes resources used by the form. </summary>
             ''' <param name="disposing"> <see langword="true"/> if managed resources should be disposed; otherwise, false. </param>
            Protected Overrides Sub Dispose(ByVal disposing As Boolean)
                #If DEBUG Then
                    Dim msg As String = String.Format("Disposed {0} '{1}' (Hashcode {2})", Me.GetType().Name, Me.Text, Me.GetHashCode())
                    System.Diagnostics.Debug.WriteLine(msg)
                #End If
                If disposing Then
                    'If (components IsNot Nothing) Then
                    '    components.Dispose()
                    'End If
                    If ((_WpfHost IsNot Nothing) AndAlso (Not _WpfHost.IsDisposed)) Then
                        _WpfHost.Child = Nothing
                    End If
                End If
                MyBase.Dispose(disposing)
            End Sub
            
        #End Region
        
        #Region "Properties"
            
            ''' <summary> Returns the <see cref="System.Windows.Forms.Integration.ElementHost"/>. </summary>
            Public ReadOnly Property WpfHost() As System.Windows.Forms.Integration.ElementHost
                Get
                    Return _WpfHost
                End Get
            End Property
            
        #End Region
        
        #Region "System.Windows.Forms.Form (Designer) Members"
            
            '' <summary> Designer variable used to keep track of non-visual components. </summary>
            'Private components As System.ComponentModel.IContainer
            
            ''' <summary> Modified designer code. </summary>
            Private Sub InitializeComponent()
                'Me.SuspendLayout()
                '
                '_WpfHost
                '
                Me._WpfHost = New System.Windows.Forms.Integration.ElementHost()
                Me._WpfHost.AutoSize = True
                Me._WpfHost.Dock = System.Windows.Forms.DockStyle.Fill
                Me._WpfHost.Location = New System.Drawing.Point(0, 0)
                Me._WpfHost.Name = "_WpfHost"
                Me._WpfHost.TabIndex = 0
                Me._WpfHost.TabStop = False
                Me._WpfHost.Text = "_WpfHost"
                '
                'WpfHostUserControl
                '
                Me.AutoScaleDimensions = New System.Drawing.SizeF(8!, 16!)
                Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
                Me.AutoSize = False
                Me.Controls.Add(Me._WpfHost)
                Me.MinimumSize = New System.Drawing.Size(150, 100)
                'Me.ResumeLayout(false)
                'Me.PerformLayout()
            End Sub
            
        #End Region
        
        #Region "Event Handlers"
            
            'Private Sub WpfHostUserControl_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
                '
                'If ((Me._WpfHost IsNot Nothing) AndAlso (Not Me._WpfHost.IsDisposed)) Then
                '    My.Settings.LoggingConsoleLocation = Me.Location
                '    My.Settings.LoggingConsoleSize = Me.Size
                '    Me._WpfHost.Child = Nothing
                '    Try
                '        My.Settings.Save()
                '    Catch ex As System.Exception
                '        System.Diagnostics.Debug.Fail("WpfHostUserControl_FormClosing(): Save settings failed!")
                '    End Try
                'End If
            'End Sub
            
        #End Region
        
    End Class
    
End Namespace

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:

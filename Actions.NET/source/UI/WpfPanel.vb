
Imports System

Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI

'Imports Rstyx.Utilities
'Imports Rstyx.Excel.Utilities.ExceptionFactory
'Imports Rstyx.Excel.Utilities.My.Resources

Namespace UI
    
    ''' <summary> Panel types supported by <see cref="WpfPanel"/> </summary>
    Public Enum WpfPanelType As Integer
        
        ''' <summary> Excel Dock Panel (<see cref="ExcelDna.Integration.CustomUI.CustomTaskPane"/>). </summary>
        Dock = 0
        
        ''' <summary> <see cref="System.Windows.Forms.Form"/>, shown modal. </summary>
        Modal = 1
        
        ''' <summary> <see cref="System.Windows.Forms.Form"/>, shown non-modal. </summary>
        NonModal = 2
    End Enum
    
    
    ''' <summary> 
    ''' Kind of UIVisualizerService: <b>Shows a WPF UserControl</b> inside a Excel child window of choice.
    ''' </summary>
    ''' <remarks> 
     ''' <para>
     ''' <b>Features:</b>
     ''' <list type="bullet">
     ''' <item><description> Creates a Excel child window of choice (see <see cref="WpfPanel.PanelType"/>). </description></item>
     ''' <item><description> The given UserControl will fill the whole window. </description></item>
     ''' <item><description> Tied to the one UserControl instance. </description></item>
     ''' <item><description> For this UserControl its <c>UseLayoutRounding</c> property will be set to <c>True</c>. </description></item>
     ''' <item><description> View (UserControl) and Viewmodel will be wired up (Cinch.IViewAwareStatus and Cinch.IViewStatusAwareInjectionAware are supported). </description></item>
     ''' <item><description> Supports closing the window on ESC key press. </description></item>
     ''' <item><description> Supports window dock settings. </description></item>
     ''' </list>
     ''' </para>
     ''' <para>
     ''' If the UserControl's <b>data context</b> inherits <see cref="Rstyx.Utilities.UI.ViewModel.ViewModelBase"/>, then:
     ''' <list type="bullet">
     ''' <item><description> <c>ViewModel.CloseRequest</c> event is subscribed. When it is raised, the WpfPanel is closed. </description></item>
     ''' <item><description> <c>ViewModel.DisplayName</c> will become the panel title. </description></item>
     ''' </list>
     ''' </para>
     ''' <para>
     ''' The WpfPanel <b>watches</b> the following <b>events</b> in order to call it's own close() or hide() method:
     ''' <list type="bullet">
     ''' <item><description> <see cref="Rstyx.Utilities.UI.ViewModel.ViewModelBase"/>.CloseRequest </description></item>
     ''' <item><description> Bentley.Windowing.<see cref="Rstyx.Excel.ActionsNET.UI.WpfHostUserControl"/>.ContentCloseQuery </description></item>
     ''' <item><description> Bentley.Excel.WinForms.<see cref="Bentley.Excel.WinForms.Form"/>.FormClosing </description></item>
     ''' <item><description> <see cref="System.Windows.Controls.UserControl"/>.KeyUp (ESC key release) </description></item>
     ''' </list>
     ''' </para>
     ''' <para>
     ''' The <c>MinWidth</c>, <c>MinHeight</c>, <c>MaxWidth</c> and <c>MaxHeight</c> properties of the <see cref="WpfPanel.WpfUserControl"/>
     ''' will be respected when opening the Excel child window.
     ''' If these are not present,the "preferred" size values are used that has been calculated automatically by Excel.
     ''' </para>
    ''' </remarks>
    Public Class WpfPanel
        
        #Region "Private Fields"
            
            Private Shared Logger           As Rstyx.LoggingConsole.Logger = Rstyx.LoggingConsole.LogBox.getLogger("Rstyx.Excel.ActionsNET.UI.WpfPanel")
            
            Private _Form                   As WpfHostForm = Nothing
            Private _Caption                As String = String.Empty
            Private _UserControl            As Rstyx.Excel.ActionsNET.UI.WpfHostUserControl = Nothing
            Private _WpfUserControl         As System.Windows.Controls.UserControl
            
            Private _PanelType              As WpfPanelType = WpfPanelType.Dock
            Private _CloseOnEscape          As Boolean = False
            Private _HideOnUserClose        As Boolean = False
            'Private _CanDockVertically      As Boolean = True
            'Private _CanDockHorizontally    As Boolean = True
            'Private _CanDockInCenter        As Boolean = False
            Private _AutoSizeFactorWidth    As Double = 1.02
            Private _AutoSizeFactorHeight   As Double = 1.02
            Private _AutoSize               As Boolean = False
            Private _AutoSizeMode           As System.Windows.Forms.AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink
            
            Private DummySize               As System.Drawing.Size = New System.Drawing.Size(1, 1)
            Private SizeFactorWPF2Forms     As Double = 96 / 72 * 0.95
            Private FirstGetSize            As Boolean = True
            Private IsSizeInitialized       As Boolean = False
            Private IsPanelInitialized      As Boolean = False
            Private IsDockPanel             As Boolean = True
            Private MinSize                 As System.Drawing.Size
            Private MaxSize                 As System.Drawing.Size
            
        #End Region
        
        #Region "Constuctors and Finalizers"
            
            ''' <summary> Instantiates this class. </summary>
             ''' <param name="WpfUserControl"> The <see cref="System.Windows.Controls.UserControl"/> which should be displayed by this WpfPanel. </param>
             ''' <param name="ViewModel">      The ViewModel to use as DataContext for the WpfUserControl. If <see langword="null"/> it's ignored. </param>
             ''' <remarks>                     Initialization of a real Excel window/panel is delayed. </remarks>
             ''' <exception cref="T:System.ArgumentNullException"> <paramref name="WpfUserControl"/> is <see langword="null"/>. </exception>
            Public Sub New(WpfUserControl As System.Windows.Controls.UserControl,
                           ViewModel As Rstyx.Utilities.UI.ViewModel.ViewModelBase
                           )
                
                Me.New(WpfUserControl, String.Empty)
                
                ' ViewModel wiring.
                If (ViewModel IsNot Nothing) Then
                    If (TypeOf ViewModel Is Cinch.IViewStatusAwareInjectionAware) Then
                        Dim ViewAwareService As New Cinch.ViewAwareStatus()
                        ViewAwareService.InjectContext(WpfUserControl)
                        ViewModel.InitialiseViewAwareService(ViewAwareService)
                    End If
                    
                    WpfUserControl.DataContext = ViewModel
                End If
            End Sub
            
            ''' <summary> Instantiates this class. </summary>
             ''' <param name="WpfUserControl"> The <see cref="System.Windows.Controls.UserControl"/> which should be displayed by this WpfPanel. </param>
             ''' <param name="Caption">        Caption for window title bar (Will be overridden by <c>ViewModel.DisplayName</c> if available). </param>
             ''' <param name="AddIn">          The Excel AddIn which owns this WpfPanel. </param>
             ''' <remarks>                     Initialization of a real Excel panel is delayed. </remarks>
             ''' <exception cref="T:System.ArgumentNullException"> <paramref name="WpfUserControl"/> is <see langword="null"/>. </exception>
             ''' <exception cref="T:System.ArgumentNullException"> <paramref name="AddIn"/> is <see langword="null"/>. </exception>
            Public Sub New(WpfUserControl As System.Windows.Controls.UserControl,
                           byVal Caption As String
                           )
                ' Check arguments.
                If (WpfUserControl Is Nothing) Then Throw New System.ArgumentNullException("WpfUserControl")
                If (AddIn Is Nothing) Then Throw New System.ArgumentNullException("AddIn")
                
                ' Store input
                _AddIn = AddIn
                _Caption = Caption
                _WpfUserControl = WpfUserControl
                
                _WpfUserControl.UseLayoutRounding = True
                
                'AddHandler _WpfUserControl.Dispatcher.UnhandledException, AddressOf OnUnhandledException
            End Sub
            
            ' ''' <summary>
            ' ''' The static constructor registers a handler for the <see cref="System.Windows.Threading.Dispatcher.UnhandledException" /> event
            ' ''' of the current thread. This assumes that the current thread isn't Excel's main thread but the main thread of "DefaultDomain" AppDomain.
            ' ''' </summary>
            ' Shared Sub New()
            '     AddHandler System.Windows.Threading.Dispatcher.CurrentDispatcher.UnhandledException, AddressOf OnUnhandledException
            ' End Sub
            
            #If DEBUG Then
                ''' <summary> Finalizes the object. </summary>
                Protected Overrides Sub Finalize()
                    Dim msg As String = String.Format("Finalized {0} '{1}'  (Type {2}, Hashcode {3})", Me.GetType().Name, Me.Caption, Me.PanelType.ToString(), Me.GetHashCode())
                    System.Diagnostics.Debug.WriteLine(msg)
                End Sub
            #End If
            
        #End Region
        
        #Region "Properties"
            
            ''' <summary> Returns the <see cref="System.Windows.Controls.UserControl"/> which is displayed by this WpfPanel. </summary>
            Public ReadOnly Property WpfUserControl() As System.Windows.Controls.UserControl
                Get
                    Return _WpfUserControl
                End Get
            End Property
            
            ''' <summary> Returns the <see cref="Rstyx.Excel.ActionsNET.UI.WpfHostUserControl"/> created with and connected to this WpfPanel. </summary>
            Public ReadOnly Property UserControl() As Rstyx.Excel.ActionsNET.UI.WpfHostUserControl
                Get
                    Return _UserControl
                End Get
            End Property
            
            ''' <summary> Returns the <see cref="Bentley.Excel.WinForms.Form"/> created with and connected to this WpfPanel. </summary>
            Public ReadOnly Property Form() As Bentley.Excel.WinForms.Form
                Get
                    Return TryCast(_Form, Bentley.Excel.WinForms.Form)
                End Get
            End Property
            
            
            ''' <summary> Indicates the panel type to perform. Defaults to <see cref="WpfPanelType.Dock"/>. </summary>
            Public Property PanelType() As WpfPanelType
                Get
                    PanelType = _PanelType
                End Get
                Set(value As WpfPanelType)
                    If (not (value = _PanelType)) Then
                        _PanelType = value
                        
                        IsDockPanel = ((_PanelType = WpfPanelType.Dock) OrElse (_PanelType = WpfPanelType.DockToolBar))
                        
                        If (IsPanelInitialized) Then InitPanel()
                    End if
                End Set
            End Property
            
            ''' <summary> Sets or gets the panel's caption. </summary>
             ''' <remarks> When a panel is created this will be <b>overridden by <c>ViewModelBase.DisplayNameLong</c></b> if view model is available. </remarks>
            Public Property Caption() As String
                Get
                    Caption = _Caption
                End Get
                Set(value As String)
                    _Caption = value
                    if (_Form isNot Nothing) then _Form.Text = _Caption
                    if (_UserControl isNot Nothing) then _UserControl.Caption = _Caption
                End Set
            End Property
            
            ''' <summary> Indicates whether or not the Dock should be closed on ESC key. Defaults to <c>False</c>. </summary>
             ''' <remarks> Changes of this property hasn't any effect while the UserControl is shown (or hidden but not closed). </remarks>
            Public Property CloseOnEscape() As Boolean
                Get
                    CloseOnEscape = _CloseOnEscape
                End Get
                Set(value As Boolean)
                    if (Not (value = _CloseOnEscape)) then
                        _CloseOnEscape = value
                    end if
                End Set
            End Property
            
            ''' <summary>
            ''' Indicates whether or not the Panel should be hidden
            ''' instead to be closed and trashed, when the user closes the window. Defaults to <c>False</c>.
            ''' </summary>
             ''' <remarks> This is ignored if panel type is ToolSettings. </remarks>
            Public Property HideOnUserClose() As Boolean
                Get
                    HideOnUserClose = _HideOnUserClose
                End Get
                Set(value As Boolean)
                    _HideOnUserClose = value
                End Set
            End Property
            
            ''' <summary> Indicates whether or not the Panel can be docked vertically. Defaults to <c>True</c>. </summary>
             ''' <remarks> This is only applicabe for panel type <see cref="WpfPanelType.Dock"/>. </remarks>
            Public Property CanDockVertically() As Boolean
                Get
                    CanDockVertically = _CanDockVertically
                End Get
                Set(value As Boolean)
                    if (not (value = _CanDockVertically)) then
                        _CanDockVertically = value
                        if (_UserControl isNot Nothing) then _UserControl.CanDockVertically = _CanDockVertically
                    end if
                End Set
            End Property
            
            ''' <summary> Indicates whether or not the Panel can be docked horizontally. Defaults to <c>True</c>. </summary>
             ''' <remarks> This is only applicabe for panel type <see cref="WpfPanelType.Dock"/>. </remarks>
            Public Property CanDockHorizontally() As Boolean
                Get
                    CanDockHorizontally = _CanDockHorizontally
                End Get
                Set(value As Boolean)
                    if (not (value = _CanDockHorizontally)) then
                        _CanDockHorizontally = value
                        if (_UserControl isNot Nothing) then _UserControl.CanDockHorizontally = _CanDockHorizontally
                    end if
                End Set
            End Property
            
            ''' <summary> Indicates whether or not the Panel can be docked in center. Defaults to <c>False</c>. </summary>
             ''' <remarks> This is only applicabe for panel type <see cref="WpfPanelType.Dock"/>. </remarks>
            Public Property CanDockInCenter() As Boolean
                Get
                    CanDockInCenter = _CanDockInCenter
                End Get
                Set(value As Boolean)
                    if (not (value = _CanDockInCenter)) then
                        _CanDockInCenter = value
                        if (_UserControl isNot Nothing) then _UserControl.CanDockInCenter = _CanDockInCenter
                    end if
                End Set
            End Property
            
            ''' <summary> An adjustment factor for auto-calculated desired panel width. Defaults to 1.02. </summary>
             ''' <remarks> The desired panel width is calculated automatically, but isn't the best choice in every case. </remarks>
            Public Property AutoSizeFactorWidth() As Double
                Get
                    AutoSizeFactorWidth = _AutoSizeFactorWidth
                End Get
                Set(value As Double)
                    If (not (value = _AutoSizeFactorWidth) AndAlso (value > 0.1)) Then
                        _AutoSizeFactorWidth = value
                        If (IsPanelInitialized) Then setSizeRestrictions()
                    End If
                End Set
            End Property
            
            ''' <summary> An adjustment factor for auto-calculated desired panel height. Defaults to 1.02. </summary>
             ''' <remarks> The desired panel height is calculated automatically, but isn't the best choice in every case. </remarks>
            Public Property AutoSizeFactorHeight() As Double
                Get
                    AutoSizeFactorHeight = _AutoSizeFactorHeight
                End Get
                Set(value As Double)
                    If (not (value = _AutoSizeFactorHeight) AndAlso (value > 0.1)) Then
                        _AutoSizeFactorHeight = value
                        If (IsPanelInitialized) Then setSizeRestrictions()
                    End If
                End Set
            End Property
            
            ''' <summary> Determines whether or not the panel should be auto-sized. Defaults to <c>False</c>. </summary>
             ''' <remarks> This is only applicabe for Form panel types (<see cref="WpfPanelType.Modal"/>, <see cref="WpfPanelType.NonModal"/>, <see cref="WpfPanelType.ToolSettings"/>). </remarks>
            Public Property AutoSize() As Boolean
                Get
                    AutoSize = _AutoSize
                End Get
                Set(value As Boolean)
                    If (value XOR _AutoSize) Then
                        _AutoSize = value
                        'If (IsPanelInitialized()) Then setSizeRestrictions()
                    End If
                End Set
            End Property
            
            ''' <summary> Determines the way the panel should be auto-sized if <see cref="AutoSize"/> is <c>True</c>. Defaults to <c>GrowAndShrink</c>. </summary>
             ''' <remarks> This is only applicabe for Form panel types (<see cref="WpfPanelType.Modal"/>, <see cref="WpfPanelType.NonModal"/>, <see cref="WpfPanelType.ToolSettings"/>). </remarks>
            Public Property AutoSizeMode() As System.Windows.Forms.AutoSizeMode
                Get
                    AutoSizeMode = _AutoSizeMode
                End Get
                Set(value As System.Windows.Forms.AutoSizeMode)
                    If (Not (value = _AutoSizeMode)) Then
                        _AutoSizeMode = value
                        'If (IsPanelInitialized()) Then setSizeRestrictions()
                    End If
                End Set
            End Property
            
        #End Region
        
        #Region "Methods"
            
            ''' <summary> Shows the panel. It will be initialized if necessary. </summary>
             ''' <remarks> If a panel exists but doesn't matches the current panel type, it's trashed before. </remarks>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Public Sub Show()
                ' When calling this method without changing default _PanelType
                ' before, no panel is initialized yet...
                Try
                    Select Case _PanelType
                        Case WpfPanelType.Dock
                            If (_UserControl is Nothing) Then InitPanel()
                            _UserControl.Show()
                            
                        Case WpfPanelType.DockToolBar
                            If (_UserControl is Nothing) Then InitPanel()
                            _UserControl.Show()
                            
                        Case WpfPanelType.ToolSettings
                            If ((_Form is Nothing) OrElse _Form.IsDisposed OrElse (Not _Form.IsHandleCreated)) Then InitPanel()
                            
                        Case WpfPanelType.Modal 
                            If ((_Form is Nothing) OrElse _Form.IsDisposed OrElse (Not _Form.IsHandleCreated)) Then InitPanel()
                            If (Not _Form.Visible) Then _Form.ShowDialog(New Bentley.Excel.WinForms.ExcelWin32)
                            ' The modal dialog returns and continues here.
                            
                        Case WpfPanelType.NonModal 
                            If ((_Form is Nothing) OrElse _Form.IsDisposed OrElse (Not _Form.IsHandleCreated)) Then InitPanel()
                            If (Not _Form.Visible) Then _Form.Show(New Bentley.Excel.WinForms.ExcelWin32)
                            
                    End Select
                    
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorShow, _Caption), ex)
                End Try
            End Sub
            
            ''' <summary> Hides the panel. </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Public Sub Hide()
                Try
                    Select Case _PanelType
                        Case WpfPanelType.Dock, WpfPanelType.DockToolBar
                            If (_UserControl isNot Nothing) then _UserControl.Hide()
                            
                        Case WpfPanelType.ToolSettings
                            If (Not ((_Form is Nothing) OrElse _Form.IsDisposed)) Then Me.Close()
                            
                        Case WpfPanelType.Modal, WpfPanelType.NonModal
                            If (Not ((_Form is Nothing) OrElse _Form.IsDisposed)) Then _Form.Hide()
                    End Select
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorHide, _Caption), ex)
                End Try
            End Sub
            
            ''' <summary> Closes and trashes the appropriate panel, but not this WpfPanel object. </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Public Sub Close()
                Try
                    ' Notify the world ..
                    RaisePanelClosing()
                    
                    ' Dispose the appropriate panel, but not this WpfPanel object!
                    DisposeWindowComponents()
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorClose, _Caption), ex)
                End Try
            End Sub 
            
        #End Region
        
        #Region "Private Routines"
            
            ''' <summary> Initializes or reconfigures the appropriate Panel. Ensures that no other Panel stands in the way. </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Private Sub InitPanel()
                Try
                    DisposeWindowComponents()
                    initViewModelConnection()
                    
                    ' Create a new Form and place the given UserControl into it's WPF-Host
                    _Form = New WpfHostForm()
                    _Form.Text = _Caption
                    _Form.WpfHost.Child = _WpfUserControl
                    
                    ' Panel type specific preparations.
                    Select Case _PanelType
                        Case WpfPanelType.Dock
                            
                            ' Create new Dock based on prepared Form
                            _UserControl = WinManager.Dock(_Form, _Form, _Caption, Bentley.Windowing.DockLocation.Floating)
                            
                            ' Hide the automatically popped up _UserControl
                            _UserControl.Hide()
                            
                            ' Restrict size according to the WPF user control.
                            setSizeRestrictions()
                            '_Form.AutoSize = Me.AutoSize
                            '_Form.AutoSizeMode = Me.AutoSizeMode
                            
                            ' Apply settings
                            _UserControl.CanDockHorizontally = _CanDockHorizontally
                            _UserControl.CanDockVertically   = _CanDockVertically
                            _UserControl.CanDockInCenter     = _CanDockInCenter
                            
                            ' Add special behavior to the Dock
                            AddHandler _UserControl.ContentCloseQuery, AddressOf OnContentCloseQuery
                            
                            
                        Case WpfPanelType.DockToolBar
                            
                            ' Create new DockToolBar based on prepared Form
                            _UserControl = WinManager.DockToolBar(_Form, _Form, _Caption, Bentley.Windowing.DockLocation.Floating)
                            
                            ' Hide the automatically popped up _UserControl
                            _UserControl.Hide()
                            
                            ' Restrict size according to the WPF user control.
                            setSizeRestrictions()
                            '_Form.AutoSize = Me.AutoSize
                            '_Form.AutoSizeMode = Me.AutoSizeMode
                            
                            ' Add special behavior to the Dock
                            AddHandler _UserControl.ContentCloseQuery, AddressOf OnContentCloseQuery
                            
                            
                        Case WpfPanelType.ToolSettings
                            
                            ' Use Form directly for ToolSettings
                            _Form.AttachToToolSettings(addIn:=_AddIn)
                            
                            ' Size: auto resizable.
                            setSizeRestrictions()
                            _Form.AutoSize = Me.AutoSize
                            _Form.AutoSizeMode = Me.AutoSizeMode
                            
                            ' Isn't recognized:
                            'AddHandler _Form.FormClosing, AddressOf OnFormClosing
                            
                        Case WpfPanelType.Modal, WpfPanelType.NonModal
                            
                            ' Use Form directly as modal dialog
                            _Form.AttachAsTopLevelForm(addIn:=_AddIn, useExcelPositioning:=True, name:=_Caption)
                            
                            ' Restrict size according to the WPF user control.
                            setSizeRestrictions()
                            _Form.AutoSize = Me.AutoSize
                            _Form.AutoSizeMode = Me.AutoSizeMode
                            
                            ' Add special behavior to Form
                            AddHandler _Form.FormClosing, AddressOf OnFormClosing
                    End Select
                    
                    If (_CloseOnEscape) Then
                        'AddHandler _WpfUserControl.KeyUp, AddressOf OnUserControlESC
                    End If
                    
                    IsPanelInitialized = True
                    
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorInitPanel, _Caption), ex)
                End Try
            End Sub
            
            ''' <summary> Disposes the _Form and _UserControl and disconnects the _WpfUserControl from owner. </summary>
            Private Sub DisposeWindowComponents()
                
                suspendViewModelConnection()
                
                ' Dispose _WpfUserControl
                'RemoveHandler _WpfUserControl.KeyUp, AddressOf OnUserControlESC
                
                ' Dispose special unmanaged resources from _Form, which should always exist.
                If (_Form IsNot Nothing) then
                    RemoveHandler _Form.FormClosing, AddressOf OnFormClosing
                End If
                
                ' Dispose _UserControl
                If (_UserControl IsNot Nothing) then
                    RemoveHandler _UserControl.ContentCloseQuery, AddressOf OnContentCloseQuery
                    WinManager.RemoveContent(_UserControl)
                    _UserControl.Close()  ' Disposes _Form!
                    _UserControl = Nothing
                    _Form = Nothing
                    
                ElseIf (Not ((_Form is Nothing) OrElse _Form.IsDisposed)) then
                    _Form.DetachFromExcel()
                    _Form.Dispose()
                    _Form = Nothing
                End If
                
                IsPanelInitialized = False
            End Sub
            
            ''' <summary> Sets minimal and/or maximal size (width and/or height). </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Private Sub setSizeRestrictions()
                Try
                    ' Calculate the sizes only once for the WpfUserControl.
                    If (MinSize.IsEmpty) Then
                        MinSize = getMinWindowSize()
                        'If (Not IsDockPanel) Then MinSize.Height = MinSize.Height * 0.98
                    End If
                    If (MaxSize.IsEmpty) Then
                        MaxSize = getMaxWindowSize()
                    End If
                    
                    ' Apply the restrictions every time, because the Form has just been re-created.
                    If (_UserControl IsNot Nothing) then
                        
                        _UserControl.Zone.TopLevelControl.MinimumSize = MinSize
                        _UserControl.Zone.TopLevelControl.MaximumSize = MaxSize
                        
                        ' Force Layout to be updated because it may not fit anymore now...
                        DummySize.Width = CInt(IIf(DummySize.Width = 1, -1, 1))
                        _UserControl.Zone.TopLevelControl.Size += DummySize
                        
                    ElseIf (Not ((_Form is Nothing) OrElse _Form.IsDisposed)) then
                        
                        _Form.MinimumSize = MinSize
                        _Form.MaximumSize = MaxSize
                        
                    End If
                    
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorSetSizeRestrictions, _Caption), ex)
                End Try
            End Sub
            
            ''' <summary> Returns a size that should be a reasonable minimal size for the child window. </summary>
             ''' <remarks>
             ''' If _WpfUserControl has set a minimum size, this is used. 
             ''' Otherwise the automatically calculated PreferredSize of ElementHost is used, if the panel shouldn't grow and shrink automatically.
             ''' </remarks>
            Private Function getMinWindowSize() As System.Drawing.Size
                Dim MinSize  As System.Drawing.Size = System.Drawing.Size.Empty
                
                If ((Not ((_Form is Nothing) OrElse _Form.IsDisposed)) AndAlso (_Form.WpfHost IsNot Nothing) AndAlso (_WpfUserControl IsNot Nothing)) then
                    
                    Dim AutoGrowAndShrink As Boolean = ((Not IsDockPanel) AndAlso Me.AutoSize AndAlso (Me.AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink))
                    
                    If (_WpfUserControl.MinWidth > 0)  Then
                        MinSize.Width = CInt(_WpfUserControl.MinWidth * SizeFactorWPF2Forms)
                    ElseIf ((_Form.WpfHost.PreferredSize.Width > 0) AndAlso (_Form.WpfHost.PreferredSize.Width < Integer.MaxValue)) Then
                        If (Not AutoGrowAndShrink) Then
                            MinSize.Width = CInt(_Form.WpfHost.PreferredSize.Width * Me.AutoSizeFactorWidth)
                        End If
                    End If
                    
                    If (_WpfUserControl.MinHeight > 0)  Then
                        MinSize.Height = CInt(_WpfUserControl.MinHeight  * SizeFactorWPF2Forms)
                    ElseIf ((_Form.WpfHost.PreferredSize.Height > 0) AndAlso (_Form.WpfHost.PreferredSize.Height < Integer.MaxValue)) Then
                        If (Not AutoGrowAndShrink) Then
                            MinSize.Height = CInt(_Form.WpfHost.PreferredSize.Height * Me.AutoSizeFactorHeight)
                        End If
                    End If
                End If
                
                Return MinSize
            End Function
            
            ''' <summary> Returns the maximum size that is set for _WpfUserControl. If it's not set then Size.Empty is returned. </summary>
            Private Function getMaxWindowSize() As System.Drawing.Size
                Dim MaxSize  As System.Drawing.Size = System.Drawing.Size.Empty
                
                If ((Not ((_Form is Nothing) OrElse _Form.IsDisposed)) AndAlso (_Form.WpfHost IsNot Nothing) AndAlso (_WpfUserControl IsNot Nothing)) then
                    
                    If ((_WpfUserControl.MaxWidth > 0) AndAlso (_WpfUserControl.MaxWidth < Double.PositiveInfinity))  Then
                        MaxSize.Width = CInt(_WpfUserControl.MaxWidth  * SizeFactorWPF2Forms)
                    End If
                    
                    If ((_WpfUserControl.MaxHeight > 0) AndAlso (_WpfUserControl.MaxHeight < Double.PositiveInfinity))  Then
                        MaxSize.Height = CInt(_WpfUserControl.MaxHeight  * SizeFactorWPF2Forms)
                    End If
                    
                    ' If one size component is set, the other one has to be set too, otherwise now it would be really interpreted as zero:
                    If (Not MaxSize.IsEmpty) Then
                        If(MaxSize.Height = 0) Then
                            MaxSize.Height = Integer.MaxValue
                        Else
                            MaxSize.Width = Integer.MaxValue
                        End If
                    End If
                End If
                
                Return MaxSize
            End Function
            
            ''' <summary> Calculates the size of a WPF UI element in actual physical pixels. </summary>
             ''' <param name="Element"> The element of interest. </param>
             ''' <returns> Size in pixels. </returns>
             ''' <remarks> see http://stackoverflow.com/questions/3286175/how-do-i-convert-a-wpf-size-to-physical-pixels </remarks>
		    Private Function getElementPixelSize(Element As System.Windows.UIElement) As System.Windows.Size
                
		    	Dim transformToDevice   As System.Windows.Media.Matrix
		    	Dim WpfSource           As System.Windows.PresentationSource = System.Windows.PresentationSource.FromVisual(Element)
                
		    	If (WpfSource IsNot Nothing) Then
		    		transformToDevice = WpfSource.CompositionTarget.TransformToDevice()
		    	Else
		    		Using WindowHandle As New System.Windows.Interop.HwndSource(New System.Windows.Interop.HwndSourceParameters())
		    			transformToDevice = WindowHandle.CompositionTarget.TransformToDevice()
		    		End Using
		    	End If
		    	
		    	If (Not Element.IsMeasureValid) Then
		    		Element.Measure(New System.Windows.Size(Double.PositiveInfinity, Double.PositiveInfinity))
		    	End If
		    	
		    	Return CType(transformToDevice.Transform(CType(Element.DesiredSize, System.Windows.Vector)), System.Windows.Size)
		    End Function
		    
            ''' <summary> Subscribes to the view model's CloseRequest event, and takes it's DisplayName as dialog caption. </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Private Sub initViewModelConnection()
                Try
                    Dim ViewModel As Rstyx.Utilities.UI.ViewModel.ViewModelBase = TryCast(_WpfUserControl.DataContext, Rstyx.Utilities.UI.ViewModel.ViewModelBase)
                    If (ViewModel IsNot Nothing) Then
                        AddHandler ViewModel.CloseRequest, AddressOf OnViewModelCloseRequest
                        
                        Me.Caption = ViewModel.DisplayNameLong
                    End If
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorInitViewModelConnection, _Caption), ex)
                End Try
            End Sub
		    
            ''' <summary> Unsubscribes the view model's CloseRequest event. </summary>
             ''' <exception cref="RemarkException"> Wraps any exception. </exception>
            Private Sub suspendViewModelConnection()
                Try
                    Dim ViewModel As Rstyx.Utilities.UI.ViewModel.ViewModelBase = TryCast(_WpfUserControl.DataContext, Rstyx.Utilities.UI.ViewModel.ViewModelBase)
                    If (ViewModel IsNot Nothing) Then
                        RemoveHandler ViewModel.CloseRequest, AddressOf OnViewModelCloseRequest
                    End If
                Catch ex As System.Exception
                    Throw New RemarkException(StringUtils.sprintf(Resources.Message_WpfPanel_ErrorSuspendViewModelConnection, _Caption), ex)
                End Try
            End Sub
            
            '' <summary> Determines, if the Excel panel is currently initialized. </summary>
		    'Private Function IsPanelInitialized() As Boolean
             '   Dim RetValue As Boolean
             '   RetValue = ((_Form IsNot Nothing) AndAlso (Not _Form.IsDisposed))
             '   
             '   If (RetValue AndAlso ((_PanelType = WpfPanelType.Dock) OrElse (_PanelType = WpfPanelType.DockToolBar)))
             '           RetValue = (_UserControl IsNot Nothing)
             '   End If
             '   Return RetValue
		    'End Function
            
        #End Region
        
        #Region "Event Handlers"
            
            ' ''' <summary> Catches and logs an unhandled exception from the main thread of "DefaultDomain" AppDomain. </summary>
            ' Shared Sub OnUnhandledException(sender As object, e As System.Windows.Threading.DispatcherUnhandledExceptionEventArgs)
            '     Logger.logErrorMC(e.Exception, "WpfPanel: Unerwarteter Fehler in einem DotNet-AddIn.")
            '     e.Handled = True
            ' End Sub
            
            ''' <summary> Response to a <see cref="System.Windows.Forms.FormClosingEventHandler"/>: Hide or close the Form. </summary>
            Private Sub OnFormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
                Try
                    If ((sender IsNot Nothing) AndAlso (e IsNot Nothing) AndAlso (_Form IsNot Nothing) AndAlso sender.Equals(_Form)) then
                        If (_HideOnUserClose AndAlso (e.CloseReason = System.Windows.Forms.CloseReason.UserClosing)) then
                            e.Cancel = True
                            Me.Hide()
                        Else
                            e.Cancel = True
                            Me.Close()
                        End If
                    End If
                Catch ex As System.Exception
                    Logger.logError(ex, StringUtils.sprintf(Resources.Message_WpfPanel_ErrorOnFormClosing, _Caption))
                End Try
            End Sub
            
            ''' <summary> Response to a <see cref="Rstyx.Excel.ActionsNET.UI.WpfHostUserControl.ContentCloseQuery"/>: Hide or close the wincontent. </summary>
            Private Sub OnContentCloseQuery(sender As Object, e As Bentley.Windowing.ContentCloseEventArgs)
                Try
                    If ((sender IsNot Nothing) AndAlso (e IsNot Nothing) AndAlso (_UserControl IsNot Nothing) AndAlso sender.Equals(_UserControl)) then
                        If (_HideOnUserClose And e.UserAction) then
                            ' This works well, but we rather use the intended interface.
                            'e.CloseAction = Bentley.Windowing.ContentCloseAction.Hide
                            e.CloseAction = Bentley.Windowing.ContentCloseAction.Disabled
                            Me.Hide()
                        Else
                            e.CloseAction = Bentley.Windowing.ContentCloseAction.Disabled
                            Me.Close()
                        End If
                    End If
                Catch ex As System.Exception
                    Logger.logError(ex, StringUtils.sprintf(Resources.Message_WpfPanel_ErrorOnContentCloseQuery, _Caption))
                End Try
            End Sub
            
            ''' <summary> Response to ESC key </summary>
            Private Sub OnUserControlESC(sender As Object, e As System.Windows.Input.KeyEventArgs)
                Try
                    If (_CloseOnEscape) Then
                        If ((e IsNot Nothing) AndAlso (e.Key = System.Windows.Input.Key.Escape)) Then
                            If (_UserControl isNot Nothing) then
                                'If (_PanelType = WpfPanelType.Dock) then
                                    If (_UserControl.Zone.AutoHideState = Bentley.Windowing.Docking.AutoHideState.None) Then
                                        ' Floating or pinned Dock
                                        e.Handled = True
                                        If (_HideOnUserClose) then
                                            Me.Hide()
                                        Else
                                            Me.Close()
                                        End If
                                    ElseIf (_UserControl.Zone.AutoHideState = Bentley.Windowing.Docking.AutoHideState.UnpinnedExpanded) Then
                                        ' AutoExpanded Dock
                                        e.Handled = True
                                        _UserControl.Zone.Collapse()
                                    End If
                                'ElseIf (_PanelType = WpfPanelType.DockToolBar) then
                                    '?
                                'End If
                            ElseIf ((_PanelType = WpfPanelType.Modal) OrElse (_PanelType = WpfPanelType.NonModal)) Then
                                ' Modal Dialog
                                e.Handled = True
                                If (_HideOnUserClose) then
                                    Me.Hide()
                                Else
                                    Me.Close()
                                End If
                            End If
                        End If
                    End If
                Catch ex As System.Exception
                    Logger.logError(ex, StringUtils.sprintf(Resources.Message_WpfPanel_ErrorOnUserControlESC, _Caption))
                End Try
            End Sub
            
            ''' <summary> Response to a view model's close request. </summary>
            Private Sub OnViewModelCloseRequest(sender As Object, e As Cinch.CloseRequestEventArgs)
                Try
                    If (_HideOnUserClose) then
                        Me.Hide()
                    Else
                        Me.Close()
                    End If
                Catch ex As System.Exception
                    Logger.logError(ex, StringUtils.sprintf(Resources.Message_WpfPanel_ErrorOnViewModelCloseRequest, _Caption))
                End Try
            End Sub
            
        #End Region
        
        #Region "Events"
            
            Private ReadOnly PanelClosingWeakEvent As New Cinch.WeakEvent(Of EventHandler(Of EventArgs))
            
            ''' <summary> Raises when this WpfPanel is being to be closed. (Internaly managed in a weakly way). </summary>
             ''' <remarks> It isn't raised when <see cref="WpfPanel.PanelType"/> is  <see cref="WpfPanelType.ToolSettings"/>. </remarks>
            Public Custom Event PanelClosing As EventHandler(Of EventArgs)
                
                AddHandler(ByVal value As EventHandler(Of EventArgs))
                    PanelClosingWeakEvent.Add(value)
                End AddHandler
                
                RemoveHandler(ByVal value As EventHandler(Of EventArgs))
                    PanelClosingWeakEvent.Remove(value)
                End RemoveHandler
                
                RaiseEvent(ByVal sender As Object, ByVal e As EventArgs)
                    Try
                        PanelClosingWeakEvent.Raise(sender, e)
                    Catch ex As System.Exception
                    End Try
                End RaiseEvent
                
            End Event
            
            ''' <summary> Raises the PanelClosing event. </summary>
             ''' <remarks> This event indicates that this WpfPanel is being to be closed. </remarks>
            Private Sub RaisePanelClosing()
                RaiseEvent PanelClosing(Me, System.EventArgs.Empty)
            End Sub
            
        #End Region
        
    End Class
    
End Namespace

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:

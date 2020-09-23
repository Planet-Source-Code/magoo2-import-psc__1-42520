VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''CSEH: Skip
'*****************************************************************************************
'* Class       : CColorDialog
'* Description : Class for displaying the standard Color dialog box.
'* Notes       : Use this class to provide in your application the posibility of
'*               displaying the standard Color dialog box without using the file
'*               CommDlg32.ocx.
'*               For advanced use, please see the Source+ Library online documentation
'*               in the Source+ Explorer (the "Documentation" tab).
'*****************************************************************************************

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 17120 + vbObjectError
Public Enum EErrColorDialog
    eErrColorDialog_InvalidWindowHandle = E_ERR_BASE + 1
    eErrColorDialog_ComponentFailure
End Enum
Private Const S_ERR_InvalidWindowHandle = "Invalid window handle"
Private Const S_ERR_ComponentFailure = "CColorDialog component failure"

' Public class enums
Public Enum EColorFlags
    eColorFlag_FullOpen = 2
    eColorFlag_PreventFullOpen = 4
    eColorFlag_HelpButton = 8
End Enum

' Private class constants
Private Const eColorFlag_RGBInit = 1
Private Const eColorFlag_EnableHook = &H10

' Private class type definitions
Private Type ChooseColor
    lStructSize      As Long
    HWndOwner        As Long
    hInstance        As Long
    rgbResult        As Long
    lpCustColors     As Long
    Flags            As Long
    lCustData        As Long
    lpfnHook         As Long
    lpTemplateName   As String
End Type

' Private class API function declarations
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

' Private variables to hold property values
Private m_eFlags             As EColorFlags
Private m_lHWndOwner         As Long
Private m_lInitialColor      As Long
Private m_lSelectedColor     As Long

' Private variables for internal class use
Private m_bFirstTime             As Boolean
Private m_lArrayCustomColors(15) As Long   'array to store the 16 custom colors


'*****************************************************************************************
'* Sub         : Init
'* Notes       : Use this routine for basic object initialization.
'*****************************************************************************************
Public Sub Init(Optional InitialColor As Long = -1, Optional DialogTitle As String = "", Optional HWndOwner As Long = 0, Optional Flags As EColorFlags)
    On Error GoTo hComponentFailure
    
    m_lInitialColor = InitialColor
    g_sColorDialogTitle = DialogTitle
    
    If HWndOwner > 0 Then
        If IsWindow(HWndOwner) = 0 Then
            On Error GoTo 0
            Err.Raise eErrColorDialog_InvalidWindowHandle, App.EXEName & ".CColorDialog", S_ERR_InvalidWindowHandle
        End If
    End If
    
    m_lHWndOwner = HWndOwner
    
    m_eFlags = Flags

    Exit Sub

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Property    : HWndOwner
'* Notes       : Returns or sets a handle to the owner window for the standard Color
'*               dialog box.
'*****************************************************************************************
Public Property Let HWndOwner(ByVal Value As Long)
    On Error GoTo hComponentFailure
    
    If Value > 0 Then
        If IsWindow(Value) = 0 Then
            On Error GoTo 0
            Err.Raise eErrColorDialog_InvalidWindowHandle, App.EXEName & ".CColorDialog", S_ERR_InvalidWindowHandle
        End If
    End If
    
    m_lHWndOwner = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property

Public Property Get HWndOwner() As Long
    On Error GoTo hComponentFailure
    
    HWndOwner = m_lHWndOwner

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : Flags
'* Notes       : Returns or sets the flags specifying the options for the standard Color
'*               dialog box.
'*****************************************************************************************
Public Property Let Flags(ByVal Value As EColorFlags)
    On Error GoTo hComponentFailure
    
    m_eFlags = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property

Public Property Get Flags() As EColorFlags
    On Error GoTo hComponentFailure
    
    Flags = m_eFlags

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : InitialColor
'* Notes       : Returns or sets the standard Color dialog box initial color.
'*****************************************************************************************
Public Property Let InitialColor(ByVal Value As Long)
    On Error GoTo hComponentFailure
    
    m_lInitialColor = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property

Public Property Get InitialColor() As Long
    On Error GoTo hComponentFailure
    
    InitialColor = m_lInitialColor

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : SelectedColor
'* Notes       : Returns the selected color if the OK button was pressed, -1 if the
'*               standard Color dialog was canceled.
'*****************************************************************************************
Public Property Get SelectedColor()
    On Error GoTo hComponentFailure

    SelectedColor = m_lSelectedColor

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : DialogTitle
'* Notes       : Returns or sets the standard Color dialog title.
'*****************************************************************************************
Public Property Get DialogTitle() As String
    On Error GoTo hComponentFailure
    
    DialogTitle = g_sColorDialogTitle

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property

Public Property Let DialogTitle(ByVal Value As String)
    On Error GoTo hComponentFailure
    
    g_sColorDialogTitle = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : CenterDialog
'* Notes       : Returns or sets if the standard Color dialog box should appear centered
'*               on the screen or in its default position.
'*****************************************************************************************
Public Property Let CenterDialog(Value As Boolean)
    On Error GoTo hComponentFailure
    
    g_bCenterColorDialog = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property

Public Property Get CenterDialog() As Boolean
    On Error GoTo hComponentFailure
    
    CenterDialog = g_bCenterColorDialog

    Exit Property

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Function    : Show
'* Notes       : Displays the standard Color selection dialog box.
'*               Returns true if OK was selected, false otherwise.
'*****************************************************************************************
Public Function Show() As Boolean
    On Error GoTo hComponentFailure
    
    Dim CC              As ChooseColor
    Dim iColorMember    As Integer
    Dim lReturn         As Long
    Dim m_Flags         As Long
    
    CC.lStructSize = Len(CC)
    CC.rgbResult = m_lInitialColor
    
    If Not m_bFirstTime Then
        
        For iColorMember = 0 To 15
            m_lArrayCustomColors(iColorMember) = &HFFFFFF
        Next
        
        m_bFirstTime = True
    
    End If
    
    CC.lpCustColors = VarPtr(m_lArrayCustomColors(0))
    
    If m_lInitialColor <> -1 Then m_eFlags = m_eFlags + eColorFlag_RGBInit
    If g_bCenterColorDialog Or (Len(g_sColorDialogTitle) > 0) Then m_eFlags = m_eFlags + eColorFlag_EnableHook
        
    CC.Flags = m_eFlags
    
    If m_eFlags And eColorFlag_EnableHook = eColorFlag_EnableHook Then
        CC.lpfnHook = FarProc(AddressOf ColorDialogCallback)
    End If
    
    lReturn = ChooseColor(CC)
    
    If lReturn = 0 Then
        ' The user canceled
        m_lSelectedColor = -1
        Show = False
    Else
        ' A valid choice was made
        m_lSelectedColor = CC.rgbResult
        Show = True
    End If

    Exit Function

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Function    : FarProc
'* Notes       : Returns the adddress of the Color dialog callback proc.
'*               This function is necessary because AddressOff can be used only as a
'*               parameter part, i.e. it will be used in something like
'*               m_Address = FarProc(AddressOff ProcName)
'*****************************************************************************************
Private Function FarProc(pfn As Long) As Long
    On Error GoTo hComponentFailure
    
    FarProc = pfn
    
    Exit Function
    
hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Function


'*****************************************************************************************
'* Sub         : Class_Initialize
'* Notes       : Class data space initialization.
'*****************************************************************************************
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure
    
    m_lSelectedColor = -1
    g_bCenterColorDialog = True
    g_sColorDialogTitle = "Color"
    Exit Sub

hComponentFailure:
    Err.Raise eErrColorDialog_ComponentFailure, App.EXEName & ".CColorDialog", S_ERR_ComponentFailure
End Sub

Attribute VB_Name = "MColorDialog"
''CSEH: Skip
'*****************************************************************************************
'* Module      : MColorDialog
'* Description : Module that provides the callback function for the Color dialog box.
'* Notes       : The callback procedure receives messages or notifications intended for
'*               the default dialog box procedure of the Color dialog, so you can subclass
'*               the standard controls of the common dialog box. In this implementation
'*               we can center the standard dialog on the screen or we can set another
'*               title. If you want to perform other changes modify the following function.
'*****************************************************************************************

Option Explicit

' Private type definitions
Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type

' Private constants
Private Const WM_INITDIALOG = &H110

' Private API function declarations
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' Public variable used for communication between a CColorDialog object and the callback routine
' implemented in this module
Public g_bCenterColorDialog  As Boolean
Public g_sColorDialogTitle   As String


'*****************************************************************************************
'* Function    : ColorDialogCallback
'* Notes       : Callback (global) routine for the CColorDialog class. It is used to
'*               center the dialog box and to set the caption text.
'*****************************************************************************************
Public Function ColorDialogCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next
    
    Dim rcHeight     As Long
    Dim rcWidth      As Long
    Dim rc           As RECT
    Dim rcDesk       As RECT
    
    Select Case uMsg
        
        Case WM_INITDIALOG
            ' Set the new title
            If Len(Trim$(g_sColorDialogTitle)) > 0 Then SetWindowText hwnd, g_sColorDialogTitle
            
            ' Center the window
            If g_bCenterColorDialog Then
                Call GetWindowRect(GetDesktopWindow, rcDesk)
                Call GetWindowRect(hwnd, rc)
            
                rcHeight = rc.Bottom - rc.Top
                rcWidth = rc.Right - rc.Left
                rc.Left = Abs(((rcDesk.Right - rcDesk.Left) - rcWidth) / 2)
                rc.Top = Abs(((rcDesk.Bottom - rcDesk.Top) - rcHeight) / 2)
            
                MoveWindow hwnd, rc.Left, rc.Top, rcWidth, rcHeight, 1
            End If
        Case Else
            '
    
    End Select
    
    ColorDialogCallback = 0&
End Function

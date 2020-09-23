Attribute VB_Name = "modModifyTreeview"
'--------------------------------------------------------------
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------
'
Public Const GWL_STYLE As Long = (-16)
Public Const COLOR_WINDOW As Long = 5
Public Const COLOR_WINDOWTEXT As Long = 8

Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003

Public Const TVIF_STATE As Long = &H8

'treeview styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Public Const TVIS_BOLD  As Long = &H10

Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TV_ITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long





Public Function GetTVBackColour(tvTree As TreeView) As Long

   Dim clrref As Long
   Dim hwndTV As Long
   
   hwndTV = tvTree.hwnd   ' TV1.hwnd
   
  'try for the treeview backcolor
   clrref = SendMessage(hwndTV, TVM_GETBKCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'a treeview background is COLOR_WINDOW.
   If clrref = -1 Then
      clrref = GetSysColor(COLOR_WINDOW)  ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVBackColour = clrref
   
End Function


Public Function GetTVForeColour(tvTree As TreeView) As Long

   Dim clrref As Long
   Dim hwndTV As Long
   
   hwndTV = tvTree.hwnd   ' TV1.hwnd
   
  'try for the treeview text colour
   clrref = SendMessage(hwndTV, TVM_GETTEXTCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'treeview text is COLOR_WINDOWTEXT.
   If clrref = -1 Then
      clrref = GetSysColor(COLOR_WINDOWTEXT) ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVForeColour = clrref
   
End Function


Public Sub SetTVBackColour(tvTree As TreeView, clrref As Long)

   Dim hwndTV As Long
   Dim style As Long
   
   hwndTV = tvTree.hwnd   ' TV1.hwnd
   
  'Change the background
   Call SendMessage(hwndTV, TVM_SETBKCOLOR, 0, ByVal clrref)
   
  'reset the treeview style so the
  'tree lines appear properly
   style = GetWindowLong(tvTree.hwnd, GWL_STYLE)   ' (TV1.hwnd, GWL_STYLE)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
  
End Sub


Public Sub SetTVForeColour(tvTree As TreeView, clrref As Long)

   Dim hwndTV As Long
   Dim style As Long
   
   hwndTV = tvTree.hwnd   ' TV1.hwnd
   
  'Change the background
   Call SendMessage(hwndTV, TVM_SETTEXTCOLOR, 0, ByVal clrref)
   
  'reset the treeview style so the
  'tree lines appear properly
   style = GetWindowLong(tvTree.hwnd, GWL_STYLE)   ' (TV1.hwnd, GWL_STYLE)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
   
End Sub




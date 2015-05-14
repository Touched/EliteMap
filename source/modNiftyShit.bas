Attribute VB_Name = "modNiftyShit"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
  Dim lhWndTop As Long
  Dim lhWnd As Long
  Dim cx As Long
  Dim cy As Long
  Dim hIconLarge As Long
  Dim hIconSmall As Long
  
  If (bSetAsAppIcon) Then
    ' Find VB's hidden parent window:
    lhWnd = hwnd
    lhWndTop = lhWnd
    Do While Not (lhWnd = 0)
      lhWnd = GetWindow(lhWnd, GW_OWNER)
      If Not (lhWnd = 0) Then
        lhWndTop = lhWnd
      End If
    Loop
  End If
  
  cx = GetSystemMetrics(SM_CXICON)
  cy = GetSystemMetrics(SM_CYICON)
  hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
  If (bSetAsAppIcon) Then
    SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
  End If
  SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
  
  cx = GetSystemMetrics(SM_CXSMICON)
  cy = GetSystemMetrics(SM_CYSMICON)
  hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
  If (bSetAsAppIcon) Then
    SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
  End If
  SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
End Sub

Public Function INIRead(ByVal FileName As String, ByVal INIHeader As String, ByVal Variable As String) As String
  Dim AppPath As String
  AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
  INIRead = String(512, Chr(0))
  INIRead = Left$(INIRead, GetPrivateProfileString(INIHeader, ByVal Variable, "", INIRead, Len(INIRead), AppPath & FileName & ".ini"))
End Function

Public Sub INIWrite(ByVal FileName As String, ByVal INIHeader As String, ByVal Variable As String, ByVal TheValue As String)
  Dim AppPath As String
  Dim TempReturn As String
  AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
  TempReturn = WritePrivateProfileString(INIHeader, Variable, TheValue, AppPath & FileName & ".ini")
End Sub


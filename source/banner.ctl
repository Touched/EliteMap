VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4080
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3600
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private FileList(1024) As String
Private NumFiles As Integer
Private FileShown As Integer

Const m_def_BaseURL = "" '"http://helmetedrodent.kickassgamers.com"
Const m_def_ListFile = "files.txt"
Dim m_BaseURL As String
Dim m_ListFile As String

Private Sub Timer1_Timer()
  UserControl.AsyncRead m_BaseURL & FileList(FileShown), vbAsyncTypePicture, "picture"
  Debug.Print "Downloading: " & m_BaseURL & FileList(FileShown)
  FileShown = FileShown + 1
  If FileShown = NumFiles Then FileShown = 0
End Sub

Private Sub Timer2_Timer()
  UserControl.AsyncRead m_BaseURL & m_ListFile, vbAsyncTypeFile, "filelist"
  Debug.Print "Downloading: " & m_BaseURL & m_ListFile
  Timer2.Enabled = False
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
  Dim f As Integer
  f = FreeFile
  If AsyncProp.PropertyName = "filelist" Then
    Open AsyncProp.Value For Input As f
    Do
      Line Input #f, FileList(NumFiles)
      NumFiles = NumFiles + 1
      If NumFiles = 1024 Then Exit Do
    Loop Until EOF(f)
    Close f
    Kill AsyncProp.Value
    Timer1.Enabled = True
    Timer1_Timer
  End If
  If AsyncProp.PropertyName = "picture" Then
    Image1.Picture = AsyncProp.Value
  End If
End Sub

Private Sub UserControl_Initialize()
  Image1.Width = 400
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 6000
  UserControl.Height = 1215
End Sub

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
  Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
  Timer1.Interval() = New_Interval
  PropertyChanged "Interval"
End Property

Public Property Get BaseURL() As String
Attribute BaseURL.VB_Description = "The base URL for all downloads"
  BaseURL = m_BaseURL
End Property

Public Property Let BaseURL(ByVal New_BaseURL As String)
  m_BaseURL = New_BaseURL
  PropertyChanged "BaseURL"
End Property

Public Property Get ListFile() As String
Attribute ListFile.VB_Description = "The name of the text file that contains all the graphics file names"
  ListFile = m_ListFile
End Property

Public Property Let ListFile(ByVal New_ListFile As String)
  m_ListFile = New_ListFile
  PropertyChanged "ListFile"
End Property

Private Sub UserControl_InitProperties()
  m_BaseURL = m_def_BaseURL
  m_ListFile = m_def_ListFile
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Timer1.Interval = PropBag.ReadProperty("Interval", 5000)
  m_BaseURL = PropBag.ReadProperty("BaseURL", m_def_BaseURL)
  m_ListFile = PropBag.ReadProperty("ListFile", m_def_ListFile)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Interval", Timer1.Interval, 5000)
  Call PropBag.WriteProperty("BaseURL", m_BaseURL, m_def_BaseURL)
  Call PropBag.WriteProperty("ListFile", m_ListFile, m_def_ListFile)
End Sub


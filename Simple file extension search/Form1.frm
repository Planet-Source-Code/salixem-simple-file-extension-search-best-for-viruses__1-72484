VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "File extension search"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   795
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Text            =   "*.exe"
      Top             =   3660
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Text            =   "D:\"
      Top             =   3300
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   5595
   End
   Begin VB.CommandButton Cmd_Search 
      Caption         =   "Search"
      Height          =   1455
      Left            =   3600
      TabIndex        =   0
      Top             =   3300
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Extension to search:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   3660
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Directory to search:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   3300
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By SaLar Zeynali
'Salixem@Gmail.Com
'S4LiX3M@Yahoo.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function
Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)
Dim FileName As String
Dim DirName As String
Dim dirNames() As String
Dim nDir As Integer
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
If Right(path, 1) <> "\" Then path = path & "\"
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
If (DirName <> ".") And (DirName <> "..") Then
If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
dirNames(nDir) = DirName
DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD)
Loop
Cont = FindClose(hSearch)
End If
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
FileCount = FileCount + 1
List1.AddItem path & FileName
End If
Cont = FindNextFile(hSearch, WFD)
Wend
Cont = FindClose(hSearch)
End If
If nDir > 0 Then
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
Next i
End If
End Function
Sub Cmd_Search_Click()
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Integer, NumDirs As Integer
List1.Clear
SearchPath = Text1.Text
FindStr = Text2.Text
FileSize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs)
Text3.Text = "Files found: " & NumFiles & vbCrLf & "Subfolders searched: " & NumDirs + 1 & vbCrLf & "Size of files found: " & Format((FileSize / 1024), "#,###,###,##0") & " KB"
End Sub


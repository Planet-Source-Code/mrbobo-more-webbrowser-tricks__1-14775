VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "More Webbrowser Tricks"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Extras"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   5415
      Begin VB.CommandButton Command9 
         Caption         =   "Windows File Property Dialog"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Hide Cursor - Completely !"
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Windows Find Files Dialog"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Press Enter or Space or click like mad  to Unhide the Cursor"
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Downloader"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5415
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   300
         Left            =   4920
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   300
         Left            =   4920
         TabIndex        =   10
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Download File"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Uses Internet Explorers' downloader without any Dialog to get in the way."
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Destination (Local)"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "File to Download (URL) or browse to test with local file ."
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Favorites Tools"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "AddFaves Dlg"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Import Bookmarks"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Export Favorites"
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "WARNING : Recommend you backup your favorites before importing Bookmarks. Links folders will be emptied of shortcuts for example."
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   4935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Remember to reference SHDOCVW.DLL in your project.
'This is for access to Internet Explorers Favorites
Dim objShellHelper As New SHDocVw.ShellUIHelper

'This is for downloading files
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Command1_Click()
'Add a link to favorites with IEs' dialog
objShellHelper.AddFavorite "http://www.planetsourcecode.com/", "Planet Source Code"
End Sub
Private Sub Command2_Click()
'Download your selected file to your selected destination
Dim llRetVal As Long
llRetVal = URLDownloadToFile(0, Text1.Text, Text2.Text, 0, 0)
End Sub

Private Sub Command3_Click()
'Standard CommonDialog stuff for you to choose
'a bookmark file for conversion
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Import Netscape Bookmarks to Favorites"
    .CancelError = True
    .Filter = "Bookmark files (*.htm)|*.htm"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
'Import a Bookmark file and convert to IEs' favorites
objShellHelper.ImportExportFavorites True, sfile
woops: Exit Sub

End Sub

Private Sub Command4_Click()
'Standard CommonDialog stuff for you to choose
'a path to save a Bookmark file to.
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Export Favorites as Netscape Bookmark File"
    .CancelError = True
    .Filter = "Bookmark files (*.htm)|*.htm"
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
'Convert the favorites folder to an htm file in Netscapes bookmark format
objShellHelper.ImportExportFavorites False, sfile
woops: Exit Sub

End Sub

Private Sub Command5_Click()
'Standard CommonDialog stuff for you to choose
'a path to download to
On Error GoTo woops
With CommonDialog1
    .DialogTitle = "Download file to ...."
    .CancelError = True
    .Filter = "Downloaded files (*.*)|*.*"
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    Text2.Text = .FileName
End With
woops: Exit Sub

End Sub

Private Sub Command6_Click()
'Standard CommonDialog stuff for you to choose
'a file for testing
On Error GoTo woops
With CommonDialog1
    .DialogTitle = "Test Downloading with local file"
    .CancelError = True
    .Filter = "Downloaded files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    Text1.Text = .FileName
End With
woops: Exit Sub

End Sub

Private Sub Command7_Click()
'Find file Dialog
ShellExecute 0, "find", "c:\", "", "", 5
End Sub

Private Sub Command8_Click()
If hidden = False Then
    GetWindowRect Me.Hwnd, R 'Get the window size/location
    ClipCursor R 'Restrict cursor movement to our window
                 'otherwise it will become unhidden when not
                 'over the form.
    ShowCursor 0 'Hide the cursor when over our window
    SetCursorPos 1, 1 'just to be mean - move the cursor
    hidden = True     'away from the button
Else
    ShowCursor 1 'unhide the cursor
    ClipCursor ByVal 0& 'allow the cursor free movement again
    hidden = False
End If
End Sub

Private Sub Command9_Click()
'Call on API declares in the Module
'Chose notepad - what else
GetPropDlg Me, "c:\windows\notepad.exe"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Better release the cursor if something goes wrong
ShowCursor 1 'unhide the cursor
ClipCursor ByVal 0& 'allow the cursor free movement again

End Sub

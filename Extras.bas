Attribute VB_Name = "Extras"
'Find Dialog API
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Hiding the Cursor API
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public R As RECT
Public hidden As Boolean
'Properties API
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    Hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (Prop As SHELLEXECUTEINFO) As Long
Public Sub GetPropDlg(frm As Form, mfile As String)
'This code is available all over the place
'I just minimised it
Dim Prop As SHELLEXECUTEINFO
Dim R As Long
With Prop
    .cbSize = Len(Prop)
    .fMask = &HC
    .Hwnd = frm.Hwnd
    .lpVerb = "properties"
    .lpFile = mfile
End With
R = ShellExecuteEx(Prop)
End Sub


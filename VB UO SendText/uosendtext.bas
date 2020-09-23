Attribute VB_Name = "Module1"
' Module for sending text to Ultima Online's Window & Declares here
' Code by Jason (jason@filex.org)

Option Explicit

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Public Const WM_CHAR = &H102

Public Sub SendUOText(text$)
    Dim k As Integer
    Dim handle As Long
    Dim anerror As Boolean
    Dim lretval As Long
    Dim title$
    Dim uotitle$
    text$ = text$ & Chr$(13) 'adds Chr$(13) {which is enter} to actually send text

        handle = 0
        handle = FindWindow("Ultima Online", CLng(0)) 'normal UO:2D
        If handle = 0 Then
            handle = FindWindow("Ultima Online Third Dawn", CLng(0)) 'UO:Third Dawn
        End If
        'this part of the code can also be used to determine window name (character and shard)
        title$ = String(GetWindowTextLength(handle) + 1, Chr$(0))
        lretval = GetWindowText(handle, title$, Len(title$))
        If handle <> 0 Then
            For k = 1 To Len(text$)
               lretval = PostMessage(handle, WM_CHAR, Asc(Mid$(text$, k, 1)), 0) 'actually sends text to UO
            Next
            anerror = False 'no error, home free
            AppActivate (title$) 'brings it forward once text sent to screen
        Else
            anerror = True 'there was an error, message for error below
        End If
    If anerror = True Then Err = MsgBox("Something went wrong," + vbCr + "either UO is not running or another error ocurred" + vbCr + "Please run UO first. ", vbOKOnly, "Oops!")
End Sub





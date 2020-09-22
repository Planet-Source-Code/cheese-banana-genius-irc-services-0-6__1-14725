Attribute VB_Name = "KnK"
'  ___________________________________________
'/                                                                                      \
'\___________________________________________/
'  Y                        KnKFounders.bas #1                         Y
'  |              These codes were written by: PooK              |
' /  Visit his site at: http://knk.tierranet.com/PooK        /
'|               This .bas was compiled by: KnK                  |
'|   To see if there are any new helpfull .bas's goto     |
' \               http://knk.tierranet.com/knk4o                 \
'  |         If you would like to submit somethin to this     |
' /            E-mail me at Bill@knk.tierranet.com              /
'|                                   KnK '98                                 |
' \________________________________________\
'/                                                                                    \
'\__________________________________________/
'This .bas works with others.  This was tested with Jolt32.bas
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global iniPath


Function AdminLoad(list As ListBox)
Open "admin.ini" For Input As #1 ' Open to write file.
Do While Not EOF(1)
Input #1, filedata
For l0072 = 0 To list.ListCount - 1
DoEvents
l007C = list.list(l0072)
l0080 = InStr(1, l007C, filedata, 1)
If l0080 Then
l0084 = Len(l007C)
l0088 = Len(filedata)
If l0084 = l0088 Then
GoTo 900
End If
End If
Next l0072
list.AddItem filedata
900:
Loop
Close #1
Exit Function
openerror:
' Cancel the error trapping
On Error GoTo 0
Exit Function
End Function
Function NickINI(First$, Second$, Third$)
r% = WritePrivateProfileString(First$, Second$, Third$, App.Path + "\nick.ini")
End Function

Function NickLoad(list As ListBox)
Open "nicklist.ini" For Input As #1 ' Open to write file.
Do While Not EOF(1)
Input #1, filedata
For l0072 = 0 To list.ListCount - 1
DoEvents
l007C = list.list(l0072)
l0080 = InStr(1, l007C, filedata, 1)
If l0080 Then
l0084 = Len(l007C)
l0088 = Len(filedata)
If l0084 = l0088 Then
GoTo 900
End If
End If
Next l0072
list.AddItem filedata
900:
Loop
Close #1
Exit Function
openerror:
' Cancel the error trapping
On Error GoTo 0
Exit Function
End Function

Function NickSave(list As ListBox)
Open "nicklist.ini" For Output As #1 ' Open to write file.
For i = 0 To list.ListCount - 1
Print #1, list.list(i)
Next i
Close #1
Exit Function
openerror1:
On Error GoTo 0
Exit Function
End Function

Function ScoreSave(list As ListBox)
Open "waspscores.txt" For Output As #1 ' Open to write file.
For i = 0 To list.ListCount - 1
Print #1, list.list(i)
Next i
Close #1
Exit Function
openerror1:
On Error GoTo 0
Exit Function
End Function

Function AdminSave(list As ListBox)
Open "admin.ini" For Output As #1 ' Open to write file.
For i = 0 To list.ListCount - 1
Print #1, list.list(i)
Next i
Close #1
Exit Function
openerror1:
On Error GoTo 0
Exit Function
End Function


Public Sub MoveForm(FRM As Form)
ReleaseCapture
X = SendMessage(FRM.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

End Sub


Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'To write to an ini type this
'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.ini")

'To read do this
'Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
'If Color$ = "bbb" Then

'*Note* an .ini must be in the the same foder as the prog with these examples
'For more info read the ini_Help.txt that was included with this
End Function


Public Function Random(Index As Integer)
Randomize
result = Int((Index * Rnd) + 1)
Random = result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function

Function WriteINI(First$, Second$, Third$)
r% = WritePrivateProfileString(First$, Second$, Third$, App.Path + "\server.ini")

End Function




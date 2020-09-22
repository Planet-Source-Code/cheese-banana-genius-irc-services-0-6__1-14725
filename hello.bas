Attribute VB_Name = "Module1"
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type






Function HideNZ()
NZ% = FindWindow("AwtDialog", vbNullString)
Call ShowWindow(NZ%, 0)
End Function

Function RHost()
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 51)
Select Case l003A
Case 1: RHost = "152"
Case 2: RHost = "325"
Case 3: RHost = "12"
Case 4: RHost = "444"
Case 5: RHost = "325"
Case 6: RHost = "351"
Case 7: RHost = "33"
Case 8: RHost = "412"
Case 9: RHost = "15"
Case 10: RHost = "333"
Case 11: RHost = "151"
Case 12: RHost = "456"
Case 13: RHost = "132"
Case 14: RHost = "465"
Case 15: RHost = "279"
Case 16: RHost = "111"
Case 17: RHost = "11"
Case 18: RHost = "63"
Case 19: RHost = "181"
Case 20: RHost = "231"
Case 21: RHost = "156"
Case 22: RHost = "24"
Case 23: RHost = "36"
Case 24: RHost = "132"
Case 25: RHost = "126"
Case 26: RHost = "175"
Case 27: RHost = "158"
Case 28: RHost = "314"
Case 29: RHost = "124"
Case 30: RHost = "206"
Case 31: RHost = "222"
Case 32: RHost = "132"
Case 33: RHost = "186"
Case 34: RHost = "197"
Case 35: RHost = "199"
Case 36: RHost = "34"
Case 37: RHost = "37"
Case 38: RHost = "167"
Case 39: RHost = "101"
Case 40: RHost = "102"
Case 41: RHost = "133"
Case 42: RHost = "122"
Case 43: RHost = "144"
Case 44: RHost = "165"
Case 45: RHost = "109"
Case 46: RHost = "106"
Case 47: RHost = "105"
Case 48: RHost = "137"
Case 49: RHost = "173"
Case 50: RHost = "149"
Case Else: RHost = "166"
End Select
End Function






Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function CenterMe(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Function

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function


Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function

Function RNick()
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 26)
Select Case l003A
Case 1: RNick = "a"
Case 2: RNick = "b"
Case 3: RNick = "c"
Case 4: RNick = "d"
Case 5: RNick = "e"
Case 6: RNick = "f"
Case 7: RNick = "g"
Case 8: RNick = "h"
Case 9: RNick = "i"
Case 10: RNick = "j"
Case 11: RNick = "k"
Case 12: RNick = "l"
Case 13: RNick = "m"
Case 14: RNick = "n"
Case 15: RNick = "o"
Case 16: RNick = "p"
Case 17: RNick = "q"
Case 18: RNick = "r"
Case 19: RNick = "s"
Case 20: RNick = "t"
Case 21: RNick = "u"
Case 22: RNick = "v"
Case 23: RNick = "w"
Case 24: RNick = "x"
Case 25: RNick = "y"
Case Else: RNick = "z"
End Select
End Function

Function RNum()
Randomize

RNum = Int((30000 - 1 + 1) * Rnd + 1)
End Function


Function RTNick()
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 15)
Select Case l003A
Case 1: RTNick = "blah"
Case 2: RTNick = "warez"
Case 3: RTNick = "stupid"
Case 4: RTNick = "guy"
Case 5: RTNick = "poopyhead"
Case 6: RTNick = "jon11"
Case 7: RTNick = "irc"
Case 8: RTNick = "yousuck"
Case 9: RTNick = "ack"
Case 10: RTNick = "flooder"
Case 12: RTNick = "Hacker"
Case 13: RTNick = "Demon"
Case 14: RTNick = "Satan"
Case Else: RTNick = "Tpyo"

End Select
End Function
Function RTSaying()
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 50)
Select Case l003A
Case 1: RTSaying = "Hi"
Case 2: RTSaying = "Sup"
Case 3: RTSaying = "heh is so cool!"
Case 4: RTSaying = "yeah, heh kicks ass!"
Case 5: RTSaying = "Hi Jon11!"
Case 6: RTSaying = "Hi Satan"
Case 7: RTSaying = "Hey."
Case 8: RTSaying = "Welcome to #Gaming!"
Case 9: RTSaying = "Who wants to kiss my ass?"
Case 10: RTSaying = "Screw you!"
Case 11: RTSaying = "This is the best network ever!"
Case 12: RTSaying = "I hate this server."
Case 13: RTSaying = "DIE!"
Case 14: RTSaying = "KoRn Rules"
Case 15: RTSaying = "pffffFFFft"
Case 16: RTSaying = "Message me if you got mp3s!!!"
Case 17: RTSaying = "Nobody ever talks.  Talk you fags!"
Case 18: RTSaying = "heh is god!"
Case 19: RTSaying = "Eat shit"
Case 20: RTSaying = "lol"
Case 21: RTSaying = "LMAO"
Case 22: RTSaying = "heh, so whats new?"
Case 23: RTSaying = "I'm not the one who said rotflmaoaoalalolololo or what the fuck ever you said, now THAT was lame"
Case 24: RTSaying = "haha,, you REALLA funny"
Case 25: RTSaying = "Everyone on this network is LAME!"
Case 26: RTSaying = "That was really lame blah!"
Case 27: RTSaying = "Hacker keeps sending me porn!"
Case 28: RTSaying = "What channel is the main channel?"
Case 29: RTSaying = "I need an IRCop to help me."
Case 30: RTSaying = "How do I register my nick?"
Case 31: RTSaying = "How do I register a channel?"
Case 32: RTSaying = "Sigh."
Case 33: RTSaying = "*Cough* *Cough*"
Case 34: RTSaying = "AOL Sucks"
Case 35: RTSaying = "Who is this peer that keeps resetting my connection?"
Case 36: RTSaying = "yoink"
Case 37: RTSaying = "erkies"
Case 38: RTSaying = "I'm bored"
Case 40: RTSaying = "heh"
Case 41: RTSaying = "thanks a lot!"
Case 42: RTSaying = "Lot of channels getting registered.."
Case 43: RTSaying = "Who are these people?"
Case 44: RTSaying = "Shadow!!"
Case 45: RTSaying = "Hey heh, can I have IRCop?"
Case 46: RTSaying = "join #teens!"
Case 47: RTSaying = "blah"
Case 48: RTSaying = "join #teen"
Case 49: RTSaying = "This network sucks"
Case Else: RTSaying = "This sucks"

End Select


End Function


Function ShowNZ()
NZ% = FindWindow("AwtDialog", vbNullString)
Call ShowWindow(NZ%, 5)
End Function

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop


End Sub





Function trimfront(Text)
asdftext$ = Text
asdftrim$ = Left$(asdftext$, 11)
For z = 1 To 11
    If Mid$(asdftrim$, z, 1) = "#" Then
        SN = Left$(asdftrim$, z - 1)
    End If
Next z
trimfront = SN
End Function

Sub StayOnTop(FRM As Form)
'Allows your form to stay on top of all other windows
Dim ontop%
ontop% = SetWindowPos(FRM.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub




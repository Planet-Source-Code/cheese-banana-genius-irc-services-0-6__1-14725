VERSION 4.00
Begin VB.Form Genius 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genius Services 0.6 Beta"
   ClientHeight    =   2445
   ClientLeft      =   1815
   ClientTop       =   1845
   ClientWidth     =   8460
   Height          =   3135
   Icon            =   "Genius.frx":0000
   Left            =   1755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   8460
   Top             =   1215
   Width           =   8580
   Begin VB.TextBox Text21 
      Height          =   495
      Left            =   6600
      TabIndex        =   43
      Text            =   "Text21"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Height          =   495
      Left            =   6720
      TabIndex        =   42
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   45000
      Left            =   7920
      Top             =   1560
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   840
      TabIndex        =   41
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   120
      TabIndex        =   40
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Text            =   "Text19"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   4320
      TabIndex        =   38
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   3000
      TabIndex        =   37
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   495
      Left            =   1320
      TabIndex        =   36
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   2520
      TabIndex        =   35
      Text            =   "Text17"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   3225
      TabIndex        =   33
      Text            =   "Port"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   1560
      TabIndex        =   29
      Text            =   "Text15"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   6225
      TabIndex        =   26
      Text            =   "Key for channel"
      Top             =   1710
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   4665
      TabIndex        =   25
      Text            =   "Spy Channel"
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   285
      Left            =   4515
      MouseIcon       =   "Genius.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   930
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Save"
      Height          =   255
      Left            =   3600
      MouseIcon       =   "Genius.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Remove Admin"
      Height          =   270
      Left            =   7170
      MouseIcon       =   "Genius.frx":05AE
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   1065
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Add Admin"
      Height          =   270
      Left            =   7200
      MouseIcon       =   "Genius.frx":0700
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   735
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   5880
      TabIndex        =   17
      Top             =   705
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Text            =   "Text12"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2010
      TabIndex        =   13
      Text            =   "Server Name"
      Top             =   780
      Width           =   2190
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2010
      TabIndex        =   12
      Text            =   "C/N password"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Raw:"
      Height          =   285
      Left            =   3885
      MouseIcon       =   "Genius.frx":0852
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2100
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   45
      TabIndex        =   10
      Text            =   "Nick"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Text            =   "#channel"
      Top             =   2115
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Text            =   "Put message to send here."
      Top             =   1485
      Width           =   3720
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Normal"
      Height          =   375
      Left            =   1290
      MouseIcon       =   "Genius.frx":09A4
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1905
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Notice"
      Height          =   375
      Left            =   2550
      MouseIcon       =   "Genius.frx":0AF6
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1905
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   4680
      Top             =   3960
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "Blah"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5130
      TabIndex        =   4
      Text            =   "Enter Command Here."
      Top             =   2100
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   285
      Left            =   4515
      MouseIcon       =   "Genius.frx":0C48
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   585
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6330
      TabIndex        =   2
      Top             =   3795
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Text            =   "Server name to link to"
      Top             =   480
      Width           =   2190
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin MsghookLib.Msghook Msghook1 
      Left            =   6360
      Top             =   2640
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   15
      Top             =   15
      Width           =   8430
   End
   Begin VB.Label rgfgf 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3015
      TabIndex        =   34
      Top             =   75
      Width           =   1155
   End
   Begin VB.Shape Shape4 
      Height          =   990
      Left            =   5820
      Top             =   420
      Width           =   2625
   End
   Begin VB.Shape Shape3 
      Height          =   990
      Left            =   0
      Top             =   420
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   8430
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Shape Shape2 
      Height          =   990
      Left            =   3840
      Top             =   1455
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      Height          =   990
      Left            =   15
      Top             =   1440
      Width           =   3795
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Connection"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4215
      TabIndex        =   32
      Top             =   75
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "ON"
      Height          =   495
      Left            =   840
      TabIndex        =   31
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "On"
      Height          =   495
      Left            =   480
      TabIndex        =   30
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Key to spy channel:"
      Height          =   255
      Left            =   6225
      TabIndex        =   28
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Spy Channel:"
      Height          =   255
      Left            =   4665
      TabIndex        =   27
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name:"
      Height          =   255
      Left            =   105
      TabIndex        =   24
      Top             =   810
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "C/N password:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1095
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name to link to:"
      Height          =   255
      Left            =   105
      TabIndex        =   22
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admins:"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   465
      Width           =   615
   End
   Begin VB.Menu server 
      Caption         =   "Server"
      Begin VB.Menu Connect 
         Caption         =   "Connect"
      End
      Begin VB.Menu Disconnect 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu fakes 
      Caption         =   "Fake Users"
      Begin VB.Menu goodfakes 
         Caption         =   "Make Good Fakes"
      End
      Begin VB.Menu add 
         Caption         =   "Add 100"
      End
      Begin VB.Menu add2 
         Caption         =   "Add 1000"
      End
      Begin VB.Menu setops 
         Caption         =   "Set Ops"
      End
      Begin VB.Menu rtalkon 
         Caption         =   "Random talk on"
      End
      Begin VB.Menu rtalkoff 
         Caption         =   "Random talk off"
      End
   End
   Begin VB.Menu spy 
      Caption         =   "Services"
      Begin VB.Menu turnon 
         Caption         =   "Turn on OperServ"
      End
      Begin VB.Menu turnoffspy 
         Caption         =   "Turn off OperServ"
      End
      Begin VB.Menu turnoncolor 
         Caption         =   "Turn on ColorBot"
      End
      Begin VB.Menu turnoffcolor 
         Caption         =   "Turn off ColorBot"
      End
      Begin VB.Menu turnonnick 
         Caption         =   "Turn on Nick"
      End
      Begin VB.Menu turnoffnick 
         Caption         =   "Turn off Nick"
      End
      Begin VB.Menu turnonall 
         Caption         =   "Turn on all"
      End
      Begin VB.Menu turnoffall 
         Caption         =   "Turn off all"
      End
      Begin VB.Menu turnpriv 
         Caption         =   "Turn on PRIVMSG spy"
      End
      Begin VB.Menu turnoffpriv 
         Caption         =   "Turn off PRIVMSG spy"
      End
      Begin VB.Menu turnonallspy 
         Caption         =   "Turn on all spy"
      End
      Begin VB.Menu turnoffallspy 
         Caption         =   "Turn off all spy"
      End
   End
   Begin VB.Menu save 
      Caption         =   "Save"
      Begin VB.Menu save1 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu proghelp 
         Caption         =   "Program Help"
      End
   End
End
Attribute VB_Name = "Genius"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Option Explicit

Dim MySock As Long

Function SendDataM(datatosend As String)
    Dim datat() As Byte
    Dim i As Integer
    
    ReDim datat(Len(datatosend))
    datat = StrConv(datatosend, vbFromUnicode)
    
    SendDataM = send(MySock, datat(0), Len(datatosend), 0)
End Function


Private Sub about_Click()
Form1.Show
End Sub

Private Sub add_Click()
text11.Text = "0"
Do Until text11.Text = "100"
Debug.Print SendData(MySock, "NICK " & RNick & RNum & RNick & RNum & RNick & RNum & " 1 936881700 " & RNick & " " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :blah" & Chr$(10) & Chr$(13));
text11.Text = CInt(text11.Text) + 1
Loop
End Sub

Private Sub add2_Click()
text11.Text = "0"
Do Until text11.Text = "1000"
Debug.Print SendData(MySock, "NICK " & RNick & RNum & RNick & RNum & RNick & RNum & " 1 936881700 " & RNick & " " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :blah" & Chr$(10) & Chr$(13));
text11.Text = CInt(text11.Text) + 1
Loop

End Sub

Private Sub Command1_Click()
 StartWinsock ("")
    Text1.Text = GetLocalHostName
    Msghook1.HwndHook = Genius.hwnd
    Msghook1.message(1025) = True
  Command1.Enabled = False
Command2.Enabled = True
connect.Enabled = False
Disconnect.Enabled = True
text2.Enabled = False
text10.Enabled = False
Text16.Enabled = False
Text9.Enabled = False
Label9.Caption = "Connecting..."
  text3.Text = AddrToIP(text2.Text)
    ConnectSock text2.Text, Text16.Text, 0, Genius.hwnd, True
    TimeOut 2
    Debug.Print SendData(MySock, "PASS " & Text9.Text & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "SERVER " & " " & text10.Text & " 1 blah" & Chr$(10) & Chr$(13))
Timer1.Enabled = False

End Sub



Private Sub Command12_Click()
Dim message
Dim title
Dim default
Dim myvalue
message = "Enter Nick to add"   ' Set prompt.
title = "" ' Set title.
default = ""   ' Set default.
' Display message, title, and default value.
myvalue = InputBox(message, title, default)
List1.AddItem myvalue
End Sub

Private Sub Command13_Click()
If List1.ListIndex < 0 Then Exit Sub
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command14_Click()

WriteINI "Server", "ServerToLink", text2.Text
WriteINI "Server", "ServerName", text10.Text
WriteINI "Server", "Password", Text9.Text
WriteINI "Server", "Port", Text16.Text
WriteINI "Spy", "Channel", Text13.Text
WriteINI "Spy", "Password", Text14.Text
NickINI "Nick", "List", text18.Text
Call AdminSave(List1)
Call NickSave(list2)
End Sub

Private Sub Command2_Click()
EndWinsock
Label9.Caption = "No Connection"
goodfakes.Enabled = True
Command2.Enabled = False
Command1.Enabled = True
connect.Enabled = True
Disconnect.Enabled = False
text2.Enabled = True
text10.Enabled = True
Text16.Enabled = True
Text9.Enabled = True
End Sub
Private Sub Command3_Click()
    Debug.Print SendData(MySock, Text4.Text & Chr$(10) & Chr$(13))
    Text4.Text = ""
End Sub

Private Sub Command4_Click()
Debug.Print SendData(MySock, ":" & Text8.Text & " notice " & Text7.Text & " :" & Text6.Text & Chr$(10) & Chr$(13))
Text6.Text = ""
End Sub

Private Sub Command5_Click()
 Debug.Print SendData(MySock, ":" & Text8.Text & " privmsg " & Text7.Text & " :" & Text6.Text & Chr$(10) & Chr$(13))
 Text6.Text = ""
 
End Sub


Private Sub Command6_Click()
On Error GoTo poops
Dim mypos
Dim mypos2
Dim mypos3
Dim asdf
Dim asdf2
Dim b
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
mypos3 = InStr(1, Text17.Text, ":identify", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
asdf2 = Mid(Text17.Text, mypos3 + 10)

If LCase(text18.Text) Like LCase("*@" & asdf & ":" & Left(asdf2, Len(asdf2) - 2) & "*") Then

Debug.Print SendData(MySock, ":nick notice " & asdf & " :Password accepted, you are now recognized." & Chr$(10) & Chr$(13))
Else:
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Incorrect Password for " & asdf & Chr$(10) & Chr$(13))
End If
poops:
Exit Sub
End Sub

Private Sub Command7_Click()
On Error GoTo poops
Dim mypos
Dim mypos2
Dim mypos3
Dim mypos4
Dim asdf
Dim asdf2
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)

mypos3 = InStr(1, Text17.Text, ":register", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
asdf2 = Mid(Text17.Text, mypos3 + 10)
If LCase(text18.Text) Like LCase("*@" & asdf & "*") Then
Debug.Print SendData(MySock, ":nick notice " & asdf & " :The nickname " & asdf & " is already registered!" & Chr$(10) & Chr$(13))
Else
text18.Text = text18.Text & "@" & asdf & ":" & asdf2
list2.AddItem asdf
NickINI "Nick", "List", text18.Text
Call NickSave(list2)
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Your nickname has been registered under the password " & asdf2 & "" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Please remember your password for later use." & Chr$(10) & Chr$(13))
End If
poops:
Exit Sub
End Sub

Private Sub Command8_Click()
On Error GoTo poops
Dim mypos
Dim mypos2
Dim mypos3
Dim mypos4
Dim mypos5
Dim asdf
Dim asdf2
Dim asdf3
Dim asdf4
Dim asdf5
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
mypos3 = InStr(1, Text17.Text, ":ghost", 1)
asdf2 = Mid(Text17.Text, mypos3 + 6)
mypos4 = InStr(1, asdf2, ":", 1)
asdf3 = Mid(asdf2, mypos4 + 1, 999)
mypos5 = Len(asdf3)
asdf4 = Left(asdf2, Len(asdf2) - mypos5 - 1)

'MsgBox "@" & Right(asdf4, Len(asdf4) - 1) & ":" & Left(asdf3, Len(asdf3) - 1)
If LCase(text18.Text) Like "*@" & Right(asdf4, Len(asdf4) - 1) & ":" & Left(asdf3, Len(asdf3) - 2) & "*" Then
Debug.Print SendData(MySock, "KILL " & asdf4 & " :Ghost Command used by " & asdf & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Your Ghost has been killed." & Chr$(10) & Chr$(13))
Else:
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Incorrect Password for " & asdf4 & Chr$(10) & Chr$(13))
End If
poops:
Exit Sub

End Sub



Private Sub connect_Click()
 StartWinsock ("")
    Text1.Text = GetLocalHostName
    Msghook1.HwndHook = Genius.hwnd
    Msghook1.message(1025) = True
  Command1.Enabled = False
Command2.Enabled = True
connect.Enabled = False
Disconnect.Enabled = True
text2.Enabled = False
text10.Enabled = False
Text16.Enabled = False
Text9.Enabled = False
Label9.Caption = "Connecting..."
  text3.Text = AddrToIP(text2.Text)
    ConnectSock text2.Text, 6667, 0, Genius.hwnd, True
    TimeOut 1
    Debug.Print SendData(MySock, "PASS " & Text9.Text & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "SERVER " & " " & text10.Text & " 1 blah" & Chr$(10) & Chr$(13))
Timer1.Enabled = False

End Sub

Private Sub Disconnect_Click()
EndWinsock
Label9.Caption = "No Connection"
Command2.Enabled = False
Command1.Enabled = True
connect.Enabled = True
Disconnect.Enabled = False
text2.Enabled = True
text10.Enabled = True
Text16.Enabled = True
Text9.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next

If App.PrevInstance = True Then End
If Me.WindowState <> 0 Then Me.WindowState = 0
CenterMe Genius
Dim ServerToLink$
Dim ServerName$
Dim Password$
Dim SpyPass$
Dim SpyChan$
Dim ConnectPort$
Dim NickNames$
ServerToLink$ = GetFromINI("Server", "ServerToLink", App.Path + "\server.ini")
ServerName$ = GetFromINI("Server", "ServerName", App.Path + "\server.ini")
ConnectPort$ = GetFromINI("Server", "Port", App.Path + "\server.ini")
Password$ = GetFromINI("Server", "Password", App.Path + "\server.ini")
SpyChan$ = GetFromINI("Spy", "Channel", App.Path + "\server.ini")
SpyPass$ = GetFromINI("Spy", "Password", App.Path + "\server.ini")
NickNames$ = GetFromINI("Nick", "List", App.Path + "\nick.ini")
text2.Text = ServerToLink$
text10.Text = ServerName$
Text9.Text = Password$
Text13.Text = SpyChan$
Text14.Text = SpyPass$
Text16.Text = ConnectPort$
text18.Text = NickNames$
Call AdminLoad(List1)
Call NickLoad(list2)
    Dim i

    Timer1.Enabled = False
    StayOnTop Me
    text12.Enabled = False
    text15.Enabled = False
    Text17.Enabled = False
    'text18.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    rtalkoff.Enabled = False
    Disconnect.Enabled = False
    Command2.Enabled = False
If Text13.Text = "" Then Text13.Text = "#Info"
If Text14.Text = "" Then Text14.Text = "cool"
If Text16.Text = "" Then Text16.Text = "6667"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Command14_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
    EndWinsock
     

End Sub


Private Sub goodfakes_Click()
goodfakes.Enabled = False
Debug.Print SendData(MySock, "NICK blah " & " 1 936881700 blah " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :blah" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":blah join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Warez " & " 1 936881700 Warez " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :warez" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Warez join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Stupid " & " 1 936881700 Stupid " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Stupid" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Stupid join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Guy " & " 1 936881700 Guy " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Guy" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Guy join #gaming" & Chr$(10) & Chr$(13))

Debug.Print SendData(MySock, "NICK Poopyhead " & " 1 936881700 Poopyhead " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :poopyhead" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Poopyhead join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Jon11 " & " 1 936881700 Jon11 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Jon11" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Jon11 join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK IRC " & " 1 936881700 IRC " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :IRC" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":IRC join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Yousuck " & " 1 936881700 Yousuck " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :yousuck" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Yousuck join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Ack " & " 1 936881700 Ack " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Ack" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Ack join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK asdf " & " 1 936881700 asdf " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :asdf" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":asdf join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Haha " & " 1 936881700 Haha " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :asfasdf" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Haha join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Bob " & " 1 936881700 Bob " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :bob" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Bob join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Me " & " 1 936881700 Me " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Me" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Me join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK mIRC " & " 1 936881700 mIRC " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :mIRC" & Chr$(10) & Chr$(13));
Debug.Print SendData(MySock, ":mIRC join #mIRC" & Chr$(10) & Chr$(13))
 
 
  Debug.Print SendData(MySock, ":mIRC join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Wasp " & " 1 936881700 Wasp " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Wasp" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Wasp join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Moron " & " 1 936881700 Moron " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Moron" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Moron join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Linux " & " 1 936881700 Linux " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Linux" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Linux join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Unix " & " 1 936881700 Unix " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Unix" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Unix join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Oper " & " 1 936881700 Oper " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Oper" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Oper join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK ooh " & " 1 936881700 ooh " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :ooh" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":ooh join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK z " & " 1 936881700 z " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :z" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":z join #mirc" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Butt " & " 1 936881700 Butt " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :butt" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Butt join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Gaming " & " 1 936881700 Gaming " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Gaming" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Gaming join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Car " & " 1 936881700 Car " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Car" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Car join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Helper " & " 1 936881700 Helper " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Helper" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Helper join #mirc" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Stinky " & " 1 936881700 Stinky " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Stinky" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Stinky join #hack" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK N64 " & " 1 936881700 N64 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :N64" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":N64 join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Zelda " & " 1 936881700 Zelda " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Zelda" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Zelda join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Chatter " & " 1 936881700 Chatter " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Chatter" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Chatter join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Bug " & " 1 936881700 Bug " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Bug" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Bug join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Hi31 " & " 1 936881700 Hi " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Hi31" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Hi31 join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Money " & " 1 936881700 Money " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Money" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Money join #hack" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Cash " & " 1 936881700 Cash " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Cash" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Cash join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Cat " & " 1 936881700 Cat " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Cat" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Cat join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Dog " & " 1 936881700 Dog " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Dog" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Dog join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Jack " & " 1 936881700 Jack " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Jack" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Jack join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Code " & " 1 936881700 Code " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Code" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Code join #mp3" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Click " & " 1 936881700 Click " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Click" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Click join #mp3" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Fart " & " 1 936881700 Fart " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Fart" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Fart join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Gamez " & " 1 936881700 Gamez " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Gamez" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Gamez join #warez" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Guest35811 " & " 1 936881700 Cool " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Guest35811" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Guest35811 join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Jamie " & " 1 936881700 heh " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Jamie" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Jamie join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Pepsi " & " 1 936881700 Pepsi " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Pepsi" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Pepsi join #mp3" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK User " & " 1 936881700 User " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :User" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":User join #help" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Secret " & " 1 936881700 Secret " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Secret" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Secret join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Die " & " 1 936881700 Die " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Die" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Die join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK God " & " 1 936881700 God " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :God" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":God join #help" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Killer " & " 1 936881700 Killer " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Killer" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Killer join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Jacob " & " 1 936881700 Sucks " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Jacob" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Jacob join #help" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Unknown " & " 1 936881700 Unknown& " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Unknown" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Unknown join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Typo " & " 1 936881700 Typo " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Typo" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Typo join #teens" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Password " & " 1 936881700 Password " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Password" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Password join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Remote " & " 1 936881700 remote " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Remote" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Remote join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Internet " & " 1 936881700 Internet " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Internet" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Internet join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Cup " & " 1 936881700 Cup " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Cup" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Cup join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Me33 " & " 1 936881700 geMe33nius " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Me33" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Me33 join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK IRCsucks " & " 1 936881700 IRC " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :IRCsucks" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":IRCsucks join #teens" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Marker " & " 1 936881700 Marker " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Marker" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Marker join #teens" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Doom " & " 1 936881700 Doom " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Doom" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Doom join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Quake " & " 1 936881700 Quake " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Quake" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Quake join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Psychic " & " 1 936881700 Psychic " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Psychic" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Psychic join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Psycho " & " 1 936881700 Psycho " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Psycho" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Psycho join #teen" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Credit " & " 1 936881700 Credit " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Credit" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Credit join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Buddy " & " 1 936881700 Buddy " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Buddy" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Buddy join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Pokemon " & " 1 936881700 Pokemon " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Pokemon" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Pokemon join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Pentium " & " 1 936881700 Pentium " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Pentium" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Pentium join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Enter " & " 1 936881700 Enter " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Enter" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Enter join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Citrus " & " 1 936881700 citrus " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Citrus" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Citrus join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Flooder " & " 1 936881700 Flooder " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Flooder" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Flooder join #teen" & Chr$(10) & Chr$(13))
   Debug.Print SendData(MySock, ":Flooder join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Legos " & " 1 936881700 Legos " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Legos" & Chr$(10) & Chr$(13));
 
 
  Debug.Print SendData(MySock, ":Legos join #chatzone" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Nuker " & " 1 936881700 Nuker " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Nuker" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Nuker join #chatzone" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Hacker " & " 1 936881700 Hacker " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Hacker" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Hacker join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Hacker join #hack" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Useless " & " 1 936881700 Useless " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Useless" & Chr$(10) & Chr$(13));
 
 
  Debug.Print SendData(MySock, ":Useless join #chatzone" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK sdhfhjuwekjh " & " 1 936881700 asdf " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :sdhfhjuwekjh" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":sdhfhjuwekjh join #chatzone" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK ICP " & " 1 936881700 ICP " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :ICP" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":ICP join #chatzone" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Starcraft " & " 1 936881700 Starcraft " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Starcraft" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Starcraft join #chatzone" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Hexen " & " 1 936881700 Gamez " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Hexen" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Hexen join #chatzone" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Heretic " & " 1 936881700 Heretic " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Heretic" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Heretic join #yo" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Phonez " & " 1 936881700 Phonez " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Phonez" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Phonez join #yo" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Teen " & " 1 936881700 Teen " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Teen" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Teen join #yo" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":Teen join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Teeny " & " 1 936881700 Teeny " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Teeny" & Chr$(10) & Chr$(13));
 
 
  Debug.Print SendData(MySock, ":Teeny join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Zip " & " 1 936881700 Zip " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Zip" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Zip join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Dumb35 " & " 1 936881700 Dumb35 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Dumb35" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Dumb35 join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Module " & " 1 936881700 Module " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Module" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Module join #yo" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK Chris " & " 1 936881700 chris " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Chris" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Chris join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK StarWars " & " 1 936881700 Star " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :StarWars" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":StarWars join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK R2D2 " & " 1 936881700 R2D2 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :R2D2" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":R2D2 join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Sock " & " 1 936881700 Sock " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Sock" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Sock join #yo" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Demon " & " 1 936881700 Demon " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Demon" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Demon join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Satan " & " 1 936881700 Satan " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Satan" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Satan join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Tpyo " & " 1 936881700 Typo " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Tpyo" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Tpyo join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Guest15551 " & " 1 936881700 MEMEME " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Guest15551" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Guest15551 join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Druggy " & " 1 936881700 Druggy " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Druggy" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Druggy join #mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Kitten " & " 1 936881700 Kitten " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Kitten" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Kitten join #mirc" & Chr$(10) & Chr$(13))
 
Debug.Print SendData(MySock, "NICK IRCD " & " 1 936881700 IRCD " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :IRCD" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":IRCD join #gaming" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Hi389 " & " 1 936881700 Hi389 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Hi389" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Hi389 join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Godz " & " 1 936881700 Godz " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Godz" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Godz join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK Idiot2 " & " 1 936881700 Idiot2 " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :Idiot2" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":Idiot2 join #teen" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "NICK hmmmmmmm " & " 1 936881700 hmmmmmmm " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :hmmmmmmm" & Chr$(10) & Chr$(13));
 
  Debug.Print SendData(MySock, ":hmmmmmmm join #hack" & Chr$(10) & Chr$(13))


End Sub

Private Sub Msghook1_Message(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
    Dim X, a As String, i
    Dim ReadBuffer(1000) As Byte
    Dim SendD As String
    Dim asdf
    Dim asdf2
    Dim asdf3
Dim mypos As Integer
Dim mypos2 As Integer
Dim mypos3 As Integer
Dim mypos4 As Integer
Dim mypos5 As Integer
Dim mypos6 As Integer
    Debug.Print msg, wp, lp, result
    If lp = FD_CONNECT Then
        MySock = wp
       ' SendD = "NICK " & Text5.Text & Chr$(10) & Chr$(13)
        'Debug.Print SendData(MySock, SendD)
        'SendD = "USER " & Text5.Text & " BLA BLA BLA BLA :BLA BLA" & Chr$(10) & Chr$(13)
        'Debug.Print SendData(MySock, SendD)
    End If

    If lp = FD_READ Then
        X = recv(MySock, ReadBuffer(0), 1000, 0)
        If X > 0 Then
            a = StrConv(ReadBuffer, vbUnicode)
              Debug.Print a
              If Text17.Enabled = True Then
              Text17.Text = ""
              Text17.Text = a
              End If
              
              If text15.Enabled = True Then
           text15.Text = ""
          text15.Text = a
          End If
                  If text12.Enabled = True Then
          text12.Text = ""
          text12.Text = a
          
                 
          End If
   If LCase(a) Like LCase("PING :*") Then
   Debug.Print SendData(MySock, "PONG " & text10.Text & Chr$(10) & Chr$(13))
   End If
   
    If LCase(a) Like LCase("*PRIVMSG IRCop :God 8029032heh4765*") Then
 mypos = InStr(1, text15.Text, ":", 1)
mypos2 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf = Mid(text15.Text, mypos + 1, mypos2 - 3)
List1.AddItem asdf
Exit Sub
End If

  If LCase(a) Like LCase("*GLOBOPS :Link with " & text10.Text & "*") Then
  Debug.Print SendData(MySock, "GNOTICE :Connection Established." & Chr$(10) & Chr$(13))
  Label9.Caption = "Connected"
  End If

If LCase(a) Like LCase("* :PING *") Then
mypos = InStr(1, a, ":", 1)
mypos2 = InStr(mypos, a, "P", 1)
mypos3 = InStr(1, a, "NG", 1)
mypos4 = InStr(1, a, "", 1)
mypos5 = InStr(1, a, "PRIVMSG", 1)
mypos6 = InStr(1, a, " :PING", 1)

asdf = Mid(a, mypos + 1, mypos2 - 3)
asdf2 = Mid(a, mypos3 + 3, mypos4 - 2)
asdf3 = Mid(a, mypos5 + 8, mypos6 - 14)
Debug.Print SendData(MySock, ":" & asdf3 & " NOTICE " & asdf & "           :PING " & asdf2 & Chr$(10) & Chr$(13))

End If
If LCase(text12.Text) Like LCase("* 451 * :You have not registered *") Then Exit Sub
If LCase(text12.Text) Like LCase("* :Not enough parameters *") Then Exit Sub
        If text12.Enabled = True And Label7.Caption = "OFF" And LCase(text12.Text) Like LCase("*PRIVMSG*") Then Exit Sub
           If text12.Enabled = True And Label8.Caption = "OFF" Then Exit Sub
   If text12.Enabled = True Then Debug.Print SendData(MySock, ":IRCop privmsg " & Text13.Text & " :[" & Time & "] " & a & Chr$(10) & Chr$(13))
        

         ' Else:
         ' If spyonchat.Checked = False And text12.Enabled = True And A = LCase("* PRIVMSG # *") Then
         ' Exit Sub
         ' Else:
         ' Debug.Print SendData(MySock, ":IRCop privmsg #info : " & A & Chr$(10) & Chr$(13))
     End If
   '  End If




End If

            


End Sub



Private Sub St_Click()

End Sub


Private Sub proghelp_Click()
HelpForm.Show
End Sub

Private Sub rtalkoff_Click()
rtalkoff.Enabled = False
rtalkon.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub rtalkon_Click()
rtalkon.Enabled = False
rtalkoff.Enabled = True
Timer2.Enabled = True

End Sub

Private Sub save1_Click()
Call Command14_Click
End Sub

Private Sub setops_Click()
Debug.Print SendData(MySock, "mode #help +o bob" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "mode #mirc +o mirc" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, "mode #help +o mirc" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #mirc +o z" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #mirc +o helper" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #hack +o Stinky" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #hack +o money" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #mp3 +o Code" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #mp3 +o Click" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #warez +o gamez" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #help +o user" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #help +o god" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #teens +o Typo" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #teens +o ircsucks" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #teen +o psycho" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, "mode #chatzone +o legos" & Chr$(10) & Chr$(13))
   Debug.Print SendData(MySock, "mode #chatzone +o nuker" & Chr$(10) & Chr$(13))
   Debug.Print SendData(MySock, "mode 3yo +o Heretic" & Chr$(10) & Chr$(13))
   Debug.Print SendData(MySock, "mode #yo +o phonez" & Chr$(10) & Chr$(13))
     Debug.Print SendData(MySock, "mode #yo +o module" & Chr$(10) & Chr$(13))
     Debug.Print SendData(MySock, "mode #mirc +o kitten" & Chr$(10) & Chr$(13))

End Sub


Private Sub text12_Change()
On Error GoTo poop:

If LCase(text12.Text) Like LCase("* JOIN *") Then

'1:
Dim asdf4
Dim asdf
Dim asdf2
Dim asdf3
Dim mypos As Integer
Dim mypos2 As Integer
Dim mypos3 As Integer
Dim mypos4 As Integer
Dim X
mypos = InStr(1, text12.Text, "#", 1)
mypos2 = InStr(mypos, text12.Text, " ", 1)
asdf = Mid(text12.Text, mypos, mypos2)
Debug.Print SendData(MySock, ":ircop join " & asdf & Chr$(10) & Chr$(13))

poop:
Exit Sub
Debug.Print SendData(MySock, ":ircop privmsg #Info :Warning: Error 1 at " & Time & ", continuing.." & Chr$(10) & Chr$(13))
End If
'2




If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help RAW*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Raw:  Use this to send raw irc commands.  There is no additional help for this command." & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help JUPE*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :JUPE:  Add a fake server to the links list to prevent a server of that name to link to yours." & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Syntax:  /msg IRCop JUPE [server name] [reason]" & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Example:  /msg IRCop JUPE irc.example.net blah" & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help FAKEUSER*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :FakeUser:  Creates a fake user." & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Syntax:  /msg IRCop FakeUser [nickname]" & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Example:  /msg IRCop FakeUser blah" & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help UNJUPE*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :UNJUPE:  Remove a server from the Juped list." & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Syntax:  /msg IRCop UNJUPE [server name to unjupe]" & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Example:  /msg IRCop UNJUPE irc.example.net" & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help SVSNICK*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)

asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :SVSNICK:  This changes a nick you specify to another nick." & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Syntax: /msg IRCop SVSNICK [nick] [newnick] :[reason]" & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Example:  /msg IRCop SVSNICK heh blah :mwahaha" & Chr$(10) & Chr$(13))
Next X
End If
If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :raw*") Then

mypos = InStr(1, text12.Text, ":raw", 1)
mypos3 = InStr(1, text12.Text, ":", 1)
mypos4 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 4, 999)
asdf2 = Mid(text12.Text, mypos3 + 1, mypos4 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf2) <> 0 Then Debug.Print SendData(MySock, asdf & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :SVSNICK*") Then
mypos = InStr(1, text12.Text, "SVSNICK", 1)
mypos3 = InStr(1, text12.Text, ":", 1)
mypos4 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf4 = Mid(text12.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text12.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf4) <> 0 Then Debug.Print SendData(MySock, "SVSNICK" & asdf2 & " " & asdf3 & Chr$(10) & Chr$(13))
Next X
End If
If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :JUPE*") Then
mypos = InStr(1, text12.Text, "JUPE", 1)
mypos3 = InStr(1, text12.Text, ":", 1)
mypos4 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf4 = Mid(text12.Text, mypos3 + 1, mypos4 - 3)

asdf = Mid(text12.Text, mypos + 5, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf4) <> 0 Then Debug.Print SendData(MySock, "SERVER " & asdf2 & " :" & asdf3 & Chr$(10) & Chr$(13))

Next X
End If
If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :UNJUPE*") Then
mypos = InStr(1, text12.Text, "UNJUPE", 1)
mypos3 = InStr(1, text12.Text, ":", 1)
mypos4 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf4 = Mid(text12.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text12.Text, mypos + 7, 999)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf4) <> 0 Then Debug.Print SendData(MySock, "SQUIT " & asdf & Chr$(10) & Chr$(13))

Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :FakeUser*") Then
mypos = InStr(1, text12.Text, "FAKEUSER", 1)
mypos3 = InStr(1, text12.Text, ":", 1)
mypos4 = InStr(1, text12.Text, "PRIVMSG", 1)

asdf4 = Mid(text12.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text12.Text, mypos + 9, 999)
asdf3 = Left(asdf, Len(asdf) - 2)

For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf4) <> 0 Then Debug.Print SendData(MySock, "NICK  " & asdf3 & " 1 936881700 " & asdf3 & " " & RHost & "." & RHost & "." & RHost & "." & RHost & " " & text10.Text & " 0 :" & asdf3 & Chr$(10) & Chr$(13));
If InStr(List1.list(X), asdf4) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf4 & " :Fake user " & asdf3 & " has been created." & Chr$(10) & Chr$(13))
Next X
End If

If LCase(text12.Text) Like LCase("*PRIVMSG IRCop :help*") Then
mypos = InStr(1, text12.Text, ":", 1)
mypos2 = InStr(1, text12.Text, "PRIVMSG", 1)
asdf = Mid(text12.Text, mypos + 1, mypos2 - 3)
For X = 0 To List1.ListCount - 1
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :My commands are: help, raw, SVSNICK, JUPE, UNJUPE, FakeUser " & Chr$(10) & Chr$(13))
If InStr(List1.list(X), asdf) <> 0 Then Debug.Print SendData(MySock, ":ircop notice " & asdf & " :Type /msg IRCop help [command] for additional information." & Chr$(10) & Chr$(13))
Next X
End If
End Sub


Private Sub text15_Change()
Dim mypos As Integer
Dim mypos2 As Integer
Dim mypos3 As Integer
Dim mypos4 As Integer
Dim asdf4
Dim asdf
Dim asdf2
Dim asdf3

If LCase(text15.Text) Like LCase("*PRIVMSG Color :help*") Then
mypos = InStr(1, text15.Text, ":", 1)
mypos2 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf = Mid(text15.Text, mypos + 1, mypos2 - 3)

Debug.Print SendData(MySock, ":color notice " & asdf & " :Hi!  I'm ColorBot!  I can color your nickname for you. " & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":color notice " & asdf & " :To get colored, type: /msg Color ColorMe [colorhere]  " & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":color notice " & asdf & " :My colors are: BLACK, DBLUE, DGREEN, RED, BROWN, PURPLE, ORANGE, YELLOW, GREEN, TEAL, LBLUE, BLUE, PINK, GRAY, DGRAY    " & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":color notice " & asdf & " :Example: /msg Color ColorMe RED " & Chr$(10) & Chr$(13))

End If

If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe BLACK*") Then

mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 5, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)
Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 1" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If

If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe DBLUE*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 2" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe DGREEN*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 3" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe RED*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)


Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 4" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe BROWN*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)


Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 5" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe PURPLE*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)


Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 6" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe ORANGE*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 7" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe YELLOW*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 8" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe GREEN*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 9" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe TEAL*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 10" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe LBLUE*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 11" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe BLUE*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 12" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe PINK*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 13" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe GRAY*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 15" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))
End If
If LCase(text15.Text) Like LCase("*PRIVMSG Color :ColorMe DGRAY*") Then
mypos = InStr(1, text15.Text, "ColorMe", 1)
mypos3 = InStr(1, text15.Text, ":", 1)
mypos4 = InStr(1, text15.Text, "PRIVMSG", 1)
asdf4 = Mid(text15.Text, mypos3 + 1, mypos4 - 3)
asdf = Mid(text15.Text, mypos + 7, 999)
mypos2 = InStr(1, asdf, " ", 1)
asdf2 = Mid(asdf, 1, mypos2 - 1)
asdf3 = Mid(asdf, mypos2 + 1, 999)

Debug.Print SendData(MySock, "SVSNICK " & asdf4 & " 14" & asdf4 & "" & " :ASDF" & Chr$(10) & Chr$(13))



End If

End Sub

Private Sub Text17_Change()

On Error GoTo poops:

Dim mypos
Dim mypos2
Dim mypos3
Dim mypos4
Dim asdf
Dim asdf2
Dim asdf3
Dim X


If LCase(Text17.Text) Like LCase("*PRIVMSG Nick :register*") Then
Call Command7_Click

End If

If LCase(Text17.Text) Like LCase("*PRIVMSG Nick :identify*") Then
Call Command6_Click
End If
If LCase(Text17.Text) Like LCase("*PRIVMSG Nick :ghost*") Then
Call Command8_Click
End If


If LCase(Text17.Text) Like LCase("*PRIVMSG nick :help Register*") Then
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Register:  Protects your nickname." & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Syntax:  /msg Nick Register [password]" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Example:  /msg Nick Register abc123" & Chr$(10) & Chr$(13))
End If

If LCase(Text17.Text) Like LCase("*PRIVMSG nick :help Identify*") Then
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Identify:  Identifies your nickname." & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Syntax:  /msg Nick Identify [password]" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Example:  /msg Nick Identify abc123" & Chr$(10) & Chr$(13))
End If

If LCase(Text17.Text) Like LCase("*PRIVMSG nick :help Ghost*") Then
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Ghost:  Kills a ghost nickname." & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Syntax:  /msg Nick ghost [nickname]:[password]" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Example:  /msg Nick ghost blah:abc123" & Chr$(10) & Chr$(13))
End If
If LCase(Text17.Text) Like LCase("*NICK*" & "1 *" & "0*" & ":*") Then
mypos = InStr(1, Text17.Text, "NICK", 1)
mypos2 = InStr(1, Text17.Text, "1", 1)
asdf = Mid(Text17.Text, mypos + 5, mypos2 - 7)
For X = 0 To List1.ListCount - 1
If LCase(text18.Text) Like LCase("*@" & asdf & ":*") And InStr(list2.list(X), asdf) <> 0 Then

Debug.Print SendData(MySock, ":nick notice " & asdf & " :This nickname is owned by someone else." & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Please Identify for this nickname." & Chr$(10) & Chr$(13))
End If

Next X
End If
If LCase(Text17.Text) Like LCase("*PRIVMSG Nick :help*") Then
mypos = InStr(1, Text17.Text, ":", 1)
mypos2 = InStr(1, Text17.Text, "PRIVMSG", 1)
asdf = Mid(Text17.Text, mypos + 1, mypos2 - 3)

Debug.Print SendData(MySock, ":nick notice " & asdf & " :My commands are:  Register, Identify, Ghost" & Chr$(10) & Chr$(13))
Debug.Print SendData(MySock, ":nick notice " & asdf & " :Type /msg Nick help [command] for additional information." & Chr$(10) & Chr$(13))
End If
'If LCase(Text17.Text) Like LCase("*:*" & "*NICK*" & "*:*") Then
'mypos = InStr(Text17.Text, "NICK", 1)
'mypos2 = InStr(Text17.Text, ":", 1)
'asdf = Mid(Text17.Text, mypos, mypos2)
'MsgBox asdf
'End If
poops:
Exit Sub

End Sub

Private Sub Text4_Click()
If Text4.Text = "Enter Command Here." Then Text4.Text = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command3_Click
End If
End Sub


Private Sub Text6_Click()
If Text6.Text = "Put message to send here." Then Text6.Text = ""
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command5_Click
End If
End Sub


Private Sub Timer1_Timer()
Debug.Print SendData(MySock, "PING wasp.dynip.com" & Chr$(10) & Chr$(13))

End Sub


Private Sub Timer2_Timer()
a:
Text20.Text = RTNick
If Text20.Text = text21.Text Then
GoTo a
Else
text21.Text = Text20.Text
Debug.Print SendData(MySock, ":" & text21.Text & " PRIVMSG #Gaming :" & RTSaying & Chr$(10) & Chr$(13))
End If
End Sub

Private Sub turnoffall_Click()
text12.Enabled = False
text15.Enabled = False
Text17.Enabled = False
End Sub

Private Sub turnoffallspy_Click()
Label8.Caption = "OFF"
Label7.Caption = "OFF"
End Sub

Private Sub turnoffcolor_Click()
text15.Enabled = False
End Sub

Private Sub turnoffnick_Click()
Text17.Enabled = False
End Sub

Private Sub turnoffpriv_Click()
Label7.Caption = "OFF"
End Sub

Private Sub turnoffspy_Click()
text12.Enabled = False
End Sub

Private Sub turnon_Click()
Debug.Print SendData(MySock, "NICK IRCop " & " 1 936881700  IRCop " & text10.Text & " " & text10.Text & " 0 :IRCop" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":IRCop mode IRCop +aASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":IRCop join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":IRCop join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":IRCop mode " & Text13.Text & " +o IRCop" & Chr$(10) & Chr$(13))
        Debug.Print SendData(MySock, ":IRCop mode " & Text13.Text & " +smntk " & Text14.Text & Chr$(10) & Chr$(13))
        Debug.Print SendData(MySock, ":IRCop TOPIC " & Text13.Text & " :Network Information Channel" & Chr$(10) & Chr$(13))
text12.Enabled = True
End Sub


Private Sub turnonall_Click()
Debug.Print SendData(MySock, "NICK IRCop " & " 1 936881700  IRCop " & text10.Text & " " & text10.Text & " 0 :IRCop" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":IRCop mode IRCop +aASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":IRCop join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":IRCop join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":IRCop mode " & Text13.Text & " +o IRCop" & Chr$(10) & Chr$(13))
        Debug.Print SendData(MySock, ":IRCop mode " & Text13.Text & " +smntk " & Text14.Text & Chr$(10) & Chr$(13))
        Debug.Print SendData(MySock, ":IRCop TOPIC " & Text13.Text & " :Network Information Channel" & Chr$(10) & Chr$(13))
text12.Enabled = True

Debug.Print SendData(MySock, "NICK Color " & " 1 936881700  Color " & text10.Text & " " & text10.Text & " 0 :Color Services" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":Color mode Color +aASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":Color join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Color join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Color mode " & Text13.Text & " +o Color" & Chr$(10) & Chr$(13))
    text15.Enabled = True
    
    Debug.Print SendData(MySock, "NICK Nick " & " 1 936881700  Nick " & text10.Text & " " & text10.Text & " 0 :NickName Services" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":Nick mode Nick +abcdefghjklmnopqrstuvwxaASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":Nick join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Nick join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Nick mode " & Text13.Text & " +o Nick" & Chr$(10) & Chr$(13))
    Text17.Enabled = True
End Sub

Private Sub turnonallspy_Click()
Label8.Caption = "ON"
Label7.Caption = "ON"
End Sub

Private Sub turnoncolor_Click()
Debug.Print SendData(MySock, "NICK Color " & " 1 936881700  Color " & text10.Text & " " & text10.Text & " 0 :Color Services" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":Color mode Color +abcdefghjklmnopqrstuvwxaASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":Color join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Color join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Color mode " & Text13.Text & " +o Color" & Chr$(10) & Chr$(13))
    text15.Enabled = True
       End Sub


Private Sub turnonnick_Click()
    
    Debug.Print SendData(MySock, "NICK Nick " & " 1 936881700  Nick " & text10.Text & " " & text10.Text & " 0 :NickName Services" & Chr$(10) & Chr$(13));
   Debug.Print SendData(MySock, ":Nick mode Nick +abcdefghjklmnopqrstuvwxaASoOrRNCTwgskcfxbWZ" & Chr$(10) & Chr$(13))
  Debug.Print SendData(MySock, ":Nick join #gaming" & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Nick join " & Text13.Text & Chr$(10) & Chr$(13))
    Debug.Print SendData(MySock, ":Nick mode " & Text13.Text & " +o Nick" & Chr$(10) & Chr$(13))
    Text17.Enabled = True
End Sub

Private Sub turnpriv_Click()
Label7.Caption = "ON"
End Sub



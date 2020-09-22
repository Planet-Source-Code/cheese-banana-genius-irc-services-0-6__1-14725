VERSION 4.00
Begin VB.Form Genius1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   1965
   ClientTop       =   1545
   ClientWidth     =   5760
   Height          =   5445
   Left            =   1905
   LinkTopic       =   "Form1"
   Picture         =   "genius1.frx":0000
   ScaleHeight     =   5040
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Top             =   1200
   Width           =   5880
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   "0.6"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   4680
      Width           =   3135
   End
   Begin MsghookLib.Msghook Msghook1 
      Left            =   480
      Top             =   960
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "Genius1"
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


Private Sub Command1_Click()
 
   
End Sub



Private Sub Form_Click()
Unload Me
Genius.Show
End Sub

Private Sub Form_Load()
Label2.FontSize = 15
CenterMe Genius1
StayOnTop Me
Label1.Caption = "Click this picture to start Genius Services."
  '  Dim i
   ' StartWinsock ("")
  
    'Msghook1.HwndHook = Genius1.hwnd
    'Msghook1.message(1025) = True
    'Label1.Caption = "Connecting to verification server..."
     'ConnectSock "game-pimp.com", 80, 0, Genius1.hwnd, True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'EndWinsock
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    EndWinsock
End Sub


Private Sub Msghook1_Message(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
'    Dim X, A As String, i
 '   Dim ReadBuffer(1000) As Byte
  '  Dim SendD As String
    
   ' Debug.Print msg, wp, lp, result
   ' If lp = FD_CONNECT Then
   ' Label1.Caption = "Connected, verifying..."
   '     MySock = wp
   '     SendD = "GET lavaisreallycold/" & Chr$(10)
   '     Debug.Print SendData(MySock, SendD)
        
   '        End If
    
   ' If lp = FD_READ Then
   '     X = recv(MySock, ReadBuffer(0), 1000, 0)
   '     If X > 0 Then
   '         A = StrConv(ReadBuffer, vbUnicode)
   '         If LCase(A) Like LCase("*12347=DISABLED*") Then
   '         MsgBox "This software has been disabled, please contact heh"
   '         Unload Me
   '         Else:
   '         TimeOut 2
   '
   '         EndWinsock
          '  Unload Me
  '           Genius.Show
  '          End If
  '          Debug.Print A
  '      End If
  '  End If
End Sub


'<Topaz> ReDim RecvBuffer(0 To 4095) 'create a 4k recieve buffer
'<Topaz> lDummy = recv(m_Socket, RecvBuffer(0), 4096, 0)
'<Topaz> If m_LineMode Then
'<Topaz> sTemp = StrConv(RecvBuffer, vbUnicode)
'<Topaz> If lDummy > 0 Then
'<Topaz> m_ReadBuffer = m_ReadBuffer & Left(sTemp, lDummy)
'<Topaz> then I do processing on m_ReadBuffer
'<Topaz> lDummy == the bytes you recieved
'<Topaz> make sure recvbuffer is a byte array


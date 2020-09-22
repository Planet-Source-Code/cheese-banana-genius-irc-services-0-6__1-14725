VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   2550
   ClientTop       =   2685
   ClientWidth     =   3585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4005
   Left            =   2490
   LinkTopic       =   "Form1"
   Picture         =   "about.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Top             =   2340
   Width           =   3705
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version .6 Beta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      MouseIcon       =   "about.frx":6108
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      MouseIcon       =   "about.frx":6412
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2760
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe Form1
StayOnTop Me
End Sub


Private Sub Label1_Click()
Call Shell("explorer http://genius.gmbt.net", 1)
Unload Form1
End Sub


Private Sub Label2_Click()
Call Shell("explorer http://members.aol.com/programmers", 1)
Unload Form1
End Sub



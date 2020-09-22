VERSION 4.00
Begin VB.Form HelpForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genius Services Help"
   ClientHeight    =   1815
   ClientLeft      =   1425
   ClientTop       =   3510
   ClientWidth     =   7230
   Height          =   2220
   Icon            =   "Help.frx":0000
   Left            =   1365
   LinkTopic       =   "Help"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7230
   Top             =   3165
   Width           =   7350
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   1815
      Left            =   2520
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Help.frx":030A
      Top             =   0
      Width           =   4740
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Help.frx":0320
      Left            =   0
      List            =   "Help.frx":0322
      TabIndex        =   0
      Top             =   0
      Width           =   2490
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterMe HelpForm
StayOnTop Me

List1.AddItem "BeforeNote", 0
List1.AddItem "Starting the Program", 1
List1.AddItem "Server to Link to", 2
List1.AddItem "C/N Password", 3
List1.AddItem "Server Name", 4
List1.AddItem "Spy Channel", 5
List1.AddItem "Spy Channel Key", 6
List1.AddItem "Admins", 7
List1.AddItem "Saving", 8
List1.AddItem "Connecting", 9
List1.AddItem "Disconnecting", 10
List1.AddItem "OperServ", 11
List1.AddItem "ColorBot", 12
List1.AddItem "Nick", 13
List1.AddItem "All", 14
List1.AddItem "PRIVMSG", 15
List1.AddItem "All Spy", 16
List1.AddItem "Service Help", 17
List1.AddItem "Make Good Fakes", 18
List1.AddItem "Add 100", 19
List1.AddItem "Add 1000", 20
List1.AddItem "Set Ops", 21
List1.AddItem "Normal, Notice", 22
List1.AddItem "Raw Commands", 23



End Sub


Private Sub List1_Click()
If List1.ListIndex = 0 Then Text1.Text = "I am still programming this software, so features change very rapidly, I can however provide a limited list of features and tell you how to get things started. Please send all questions, bugs, or comments to binary0100@yahoo.com."
If List1.ListIndex = 1 Then Text1.Text = "After you run it, you will see a graphic, click it.  The graphic should disappear, and the services should appear."
If List1.ListIndex = 2 Then Text1.Text = "This is where you put the ip address of the server you wish to link the services to."
If List1.ListIndex = 3 Then Text1.Text = "This is the password that you put in your C/N lines in your ircd.conf"
If List1.ListIndex = 4 Then Text1.Text = "This is where you put the name of the services server."
If List1.ListIndex = 5 Then Text1.Text = "This is the channel where information picked up by the services will be relayed through IRCop."
If List1.ListIndex = 6 Then Text1.Text = "This is the key for the spy channel.  (It keeps unwanted people from joining.)"
If List1.ListIndex = 7 Then Text1.Text = "There should be a listbox on the upper right, this is where you add admins.  IRCop will only respond to people on this list!  To remove an admin click on the admin's nick in the listbox and click the remove admin button."
If List1.ListIndex = 8 Then Text1.Text = "To save the data you have input thus far, click Save from the menu."
If List1.ListIndex = 9 Then Text1.Text = "After you have set things up, click the connect button, and the services will attempt to connect to the server you entered on the port specified.  If everything is working good, and you have the lines put in right, it should connect.  It should say: Connected in the bottom right corner after it connects.  Make sure you are connected to your server already before connecting to services."
If List1.ListIndex = 10 Then Text1.Text = "To disconnect, click disconnect."
If List1.ListIndex = 11 Then Text1.Text = "Turn on OperServ - Turns on IRCop (I put OperServ because people are more familiar with that name).  Turn off OperServ - Disables IRCop and all his functions."
If List1.ListIndex = 12 Then Text1.Text = "Turn on ColorBot - Turns on Color(Color Services).  Turn off ColorBot - Disables Color and all his functions."
If List1.ListIndex = 13 Then Text1.Text = "Turn on Nick - Turns on Nick(Nickname Services).  Turn off Nick - Disables Nick and all his functions."
If List1.ListIndex = 14 Then Text1.Text = "Turn on all - Turns on IRCop, Color and Nick.  Turn off all - Disables IRCop, Color and Nick"
If List1.ListIndex = 15 Then Text1.Text = "Turn on PRIVMSG Spy - Turns on the ability for IRCop to listen to PRIVMSG and relay it to the spy channel.(Automatically on).  Turn off PRIVMSG Spy - Turns off the ability for IRCop to listen to PRIVMSG and relay it to the spy channel."
If List1.ListIndex = 16 Then Text1.Text = "Turn on all spy - Turns on all the stuff IRCop spies on.(Automatically ON).  Turn off all spy - Turns off IRCop spying on stuff."
If List1.ListIndex = 17 Then Text1.Text = "Type /msg Color help for help with color services.  Type /msg IRCop help for help with Admin services (He will only respond if you are added to the admin list)"
If List1.ListIndex = 18 Then Text1.Text = "Creates 100 Good fake users.  Good meaning their nicknames makes sense and they join channels."
If List1.ListIndex = 19 Then Text1.Text = "Adds 100 fake users.  These will have random nicknames and will not join channels."
If List1.ListIndex = 20 Then Text1.Text = "Adds 1000 fake users.  These will also have random nicknames and will not join channels."
If List1.ListIndex = 21 Then Text1.Text = "Gives some of the good fakes Channel operator status in the channels they join."
If List1.ListIndex = 22 Then Text1.Text = "Enter the nick to talk and the channel and click to make it talk."
If List1.ListIndex = 23 Then Text1.Text = "Send a command to the server."
End Sub



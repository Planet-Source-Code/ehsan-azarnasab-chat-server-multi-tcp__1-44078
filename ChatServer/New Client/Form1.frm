VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Client"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFriends 
      Height          =   1815
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txt2Friends 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DisConnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox lstFriends 
      Height          =   840
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtReceived 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   2880
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Chat With Friends"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Send Message 2 Friend"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Friend List"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Send Message 2 Server"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Received From Server"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If txtUserName.Text = "" Then
        MsgBox "You need to type your username!", vbCritical, "Unable to complete"
        Exit Sub
    End If
    'we make it close by ourselves because peer may be closing it by
    'winsock state goes 8 itself and connection willbe refused then
    'and we must close it by client side for connecting another time
    wsMain.Close
    wsMain.Connect
    Do Until wsMain.State = 7 'means do while connected
        ' 0 is closed, 9 is error
        If wsMain.State = 0 Or wsMain.State = 9 Then
            MsgBox "Error in connecting!", vbCritical, "Winsock Error"
            ' there was an error, so let's leave
            Exit Sub
        End If
        DoEvents  'don't freeze the system!
    Loop
    ' "log-in":
    wsMain.SendData "U" & Chr(1) & txtUserName.Text
    Me.Caption = "Client " + txtUserName.Text
    txtUserName.Enabled = False
    txtMessage.Enabled = True
    txt2Friends.Enabled = True
    Command1.Enabled = False
    Command2.Enabled = True
    
End Sub

Private Sub Command2_Click()
 txtUserName.Enabled = True
 txtMessage.Enabled = False
 txt2Friends.Enabled = False
 Command1.Enabled = True
 Command2.Enabled = False
 Me.Caption = "Client"
 wsMain.Close
 lstFriends.Clear
End Sub

Private Sub Text1_Change()

End Sub

'send msg to a friend
Private Sub txt2Friends_KeyDown(KeyCode As Integer, Shift As Integer)
Dim User As String
    'First, check to make sure someone is online
    If lstFriends.ListCount = 0 And KeyCode = 13 Then
    
        'Display popup
        MsgBox "Nobody to send to!", vbExclamation, "Cannot send"
        
        'Clear input
        txtSendMessage.Text = ""
        Exit Sub
    End If

    ' If it was enter and shift wasn't pressed, then...
    If KeyCode = 13 Then
        ' Get Username to send to
        User = lstFriends.Text
        ' RetrieveUser returns -1 if the user wasn't found
        If User = "" Then
            'Display popup
            MsgBox "You have selected nobody to send to!", vbExclamation, "Cannot send"
        
            Exit Sub
        End If
        ' format the message
        ' Format is Command:chr(1):friendName:chr(2):msg
        wsMain.SendData "f" & Chr(1) & txt2Friends.Text & Chr(2) & User
        'We wanna sea our msg!
        txtFriends.SelStart = Len(txtFriends.Text)
        txtFriends.SelText = txtUserName & "->" & User & ":" & txt2Friends & vbCrLf
        
        DoEvents
        ' Blank the input
        txt2Friends.Text = ""
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        wsMain.SendData "t" & Chr(1) & txtMessage.Text
        txtMessage.Text = ""
        KeyAscii = 0
    End If
End Sub

Private Sub wsMain_Close()
 txtUserName.Enabled = True
 txtMessage.Enabled = False
 Command1.Enabled = True
 Command2.Enabled = False
 Me.Caption = "Client"
 lstFriends.Clear
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
Dim Data As String, CtrlChar As String
Dim EachNode As String
    wsMain.GetData Data
    CtrlChar = Left(Data, 1) ' Let's get the first char
    Data = Mid(Data, 3)      ' Then cut it off
    Select Case LCase(CtrlChar)   ' Check what it is
        Case "y"  ' message from a friend with the name of him
            txtFriends.SelStart = Len(txtFriends.Text)
            Mid(Data, InStr(1, Data, Chr(2)), 1) = ":"
            txtFriends.SelText = Data & vbCrLf
        Case "m"   ' Do stuff depending on it
            MsgBox Data, vbInformation, "Msg from server"
        Case "c"
            Me.Caption = "Client - " & Data
        Case "r"
            'After first login to server
            'We Retrieve Friend list from a simple line of text
            While InStr(1, Data, Chr(2)) <> 0
              EachNode = Mid(Data, 1, InStr(1, Data, Chr(2)) - 1)
              lstFriends.AddItem (EachNode)
              Data = Mid(Data, InStr(1, Data, Chr(2)) + 1)
            Wend
        Case "a" 'choose another name to login
            txtUserName.Enabled = True
            txtMessage.Enabled = False
            Command1.Enabled = True
            Me.Caption = "Client"
        Case "n" 'New user on the door!
            lstFriends.AddItem (Data)
        Case "o" 'One user goes offline
            ' Let's cycle through the list, looking for their
            ' name
            For X = 0 To lstFriends.ListCount - 1
    
                ' Check to see if it matches
                If lstFriends.List(X) = Data Then
        
                    ' It matches, so let's remove it form the
                    ' list and the array
                    lstFriends.RemoveItem X
                    Exit For
        
                End If
            Next X

        Case Else '"t"
            txtReceived.SelStart = Len(txtReceived.Text)
            txtReceived.SelText = Data & vbCrLf
    End Select
End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: " & Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub

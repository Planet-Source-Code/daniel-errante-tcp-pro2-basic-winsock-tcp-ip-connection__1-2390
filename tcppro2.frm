VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Pro2"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   Icon            =   "tcppro2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Directions"
      TabPicture(0)   =   "tcppro2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "TCP/IP Connection Setup"
      TabPicture(1)   =   "tcppro2.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Option1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Option2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblstatus"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Image4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "TCP Pro2 Chat"
      TabPicture(2)   =   "tcppro2.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtCHAT"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "CommonDialog1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtstatus"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtmsg"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Options"
      TabPicture(3)   =   "tcppro2.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtlocal"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Command1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Command2"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Command4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtversion"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtcomp"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Inet1"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label13"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label12"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label11"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Image5"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.CommandButton Command6 
         Caption         =   "< Back to chat"
         Height          =   375
         Left            =   -74880
         TabIndex        =   31
         Top             =   4860
         Width           =   1455
      End
      Begin VB.TextBox txtlocal 
         BackColor       =   &H80000000&
         Height          =   3255
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1500
         Width           =   7335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close Current Connection"
         Height          =   375
         Left            =   -69720
         TabIndex        =   29
         Top             =   4860
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save Chat"
         Height          =   375
         Left            =   -73320
         TabIndex        =   28
         Top             =   4860
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Check for new version"
         Height          =   375
         Left            =   -71880
         TabIndex        =   27
         Top             =   4860
         Width           =   2055
      End
      Begin VB.TextBox txtversion 
         Height          =   285
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtcomp 
         Height          =   285
         Left            =   -71640
         MaxLength       =   20
         TabIndex        =   25
         Text            =   "Friend"
         Top             =   540
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Server:"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   14
         Top             =   1260
         Width           =   3615
         Begin VB.TextBox txtserverport 
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Text            =   "5001"
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton cmdlisten 
            Caption         =   "&Listen"
            Height          =   375
            Left            =   840
            TabIndex        =   15
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label7 
            Caption         =   "Port to listen on:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Client:"
         Enabled         =   0   'False
         Height          =   3015
         Left            =   -71160
         TabIndex        =   8
         Top             =   1260
         Width           =   3615
         Begin VB.TextBox txtserverip 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtclientport 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Text            =   "5001"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdconnect 
            Caption         =   "&Connect"
            Height          =   375
            Left            =   840
            TabIndex        =   9
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Server IP address:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Port to connect to:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Server"
         Height          =   255
         Left            =   -70920
         TabIndex        =   7
         Top             =   900
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Client"
         Height          =   255
         Left            =   -69960
         TabIndex        =   6
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtmsg 
         Height          =   285
         Left            =   -73440
         TabIndex        =   5
         Top             =   660
         Width           =   4935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Send"
         Default         =   -1  'True
         Height          =   405
         Left            =   -68400
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtstatus 
         BackColor       =   &H80000004&
         Height          =   590
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4695
         Width           =   7335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Click to start"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   2940
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -74760
         Top             =   4140
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox txtCHAT 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   3
         Top             =   1140
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"tcppro2.frx":037A
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   -68400
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label13 
         Caption         =   "What is my..."
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "TCP Pro current version:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   33
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "The other computer is known as:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   32
         Top             =   540
         Width           =   2415
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74760
         Picture         =   "tcppro2.frx":044F
         Top             =   540
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   240
         Picture         =   "tcppro2.frx":0759
         Top             =   540
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   $"tcppro2.frx":0A63
         Height          =   615
         Left            =   840
         TabIndex        =   24
         Top             =   540
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   $"tcppro2.frx":0B4D
         Height          =   855
         Left            =   240
         TabIndex        =   23
         Top             =   1140
         Width           =   7215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7440
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label4 
         Caption         =   $"tcppro2.frx":0CF1
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   7335
      End
      Begin VB.Label Label5 
         Caption         =   "Please enter the appropriate information below:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   21
         Top             =   600
         Width           =   6615
      End
      Begin VB.Label Label6 
         Caption         =   "Are you going to be the server or the client?"
         Height          =   255
         Left            =   -74160
         TabIndex        =   20
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label lblstatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"tcppro2.frx":0DA5
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   4380
         Width           =   7335
      End
      Begin VB.Label Label10 
         Caption         =   "Message:"
         Height          =   255
         Left            =   -74160
         TabIndex        =   18
         Top             =   660
         Width           =   735
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74760
         Picture         =   "tcppro2.frx":0E6B
         Top             =   540
         Width           =   540
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   -74760
         Picture         =   "tcppro2.frx":1175
         Top             =   540
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Winsock3.LocalPort = txt1.Text
Winsock3.Listen
Frame3.Visible = False
t1.Visible = True
t2.Visible = True
t3.Visible = True
t4.Visible = True
t5.Visible = True
t6.Visible = True
t7.Visible = True
t8.Visible = True
t9.Visible = True
t10.Visible = True
t11.Visible = True
t12.Visible = True
t13.Visible = True
t14.Visible = True

End Sub

Private Sub cmd2_Click()
Winsock3.RemoteHost = txt2.Text
Winsock3.RemotePort = txt3.Text
Winsock3.Connect
Frame3.Visible = False
t1.Visible = True
t2.Visible = True
t3.Visible = True
t4.Visible = True
t5.Visible = True
t6.Visible = True
t7.Visible = True
t8.Visible = True
t9.Visible = True
t10.Visible = True
t11.Visible = True
t12.Visible = True
t13.Visible = True
t14.Visible = True

End Sub

Private Sub Command1_Click()
On Error GoTo err:
Winsock1.SendData "Other user disconnected!"
Winsock1.Close
Frame1.Enabled = True
Frame2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
SSTab1.Tab = 0
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub cmdconnect_Click()
On Error GoTo err:
Winsock1.RemoteHost = txtserverip.Text
Winsock1.RemotePort = txtclientport.Text
Winsock1.Connect
SSTab1.Tab = 2
txtmsg.SetFocus
Frame1.Enabled = False
Frame2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub cmdlisten_Click()
On Error GoTo err:
Winsock1.LocalPort = txtserverport.Text
Winsock1.Listen
SSTab1.Tab = 2
txtmsg.SetFocus
Frame1.Enabled = False
Frame2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command2_Click()
If txtCHAT.Text <> "" Then
    CommonDialog1.Filter = "Text files (*.txt)|*.txt"
    CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            Open CommonDialog1.FileName For Output As #1
            Print #1, txtCHAT.Text
            Close #1
        End If
End If

End Sub

Private Sub Command3_Click()
If txtmsg.Text <> "" Then
    On Error GoTo err:
    Winsock1.SendData txtmsg.Text
    txtCHAT.Text = txtCHAT.Text & "Me: " & txtmsg.Text & vbCrLf
    txtmsg.Text = ""
    txtCHAT.SelStart = Len(txtCHAT.Text)
End If
Exit Sub
err:
    txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
    txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Command4_Click()
On Error GoTo err:
If Inet1.OpenURL("http://www.nowresources.com/directchat/version.txt") > App.Major & App.Minor & App.Revision Then
    Dim prompt As String
    Dim reply As Integer
    prompt = MsgBox("There is an upgrade to a new version available on the NOW Resources server.", vbExclamation)
    MsgBox "Do you want to download it now?", vbYesNo, "Download now?"
    If reply = vbNo Then
    Exit Sub
    Else
    'get download from server...
    'start download
    Dim strURL As String
    Dim bData() As Byte      ' Data variable
    Dim intFile As Integer   ' FreeFile variable
    strURL = _
    "www.nowresources.com/danoph/setup.zip"
    intFile = FreeFile()      ' Set intFile to an unused
                            ' file.
    ' The result of the OpenURL method goes into the Byte
    ' array, and the Byte array is then saved to disk.
    bData() = Inet1.OpenURL(strURL, icByteArray)
    Open App.Path & "setup.zip" For Binary Access Write _
    As #intFile
    Put #intFile, , bData()
    Close #intFile

    End If
Else
MsgBox "There are no upgrades to your current version.", vbExclamation
End If
err:

End Sub

Private Sub Command5_Click()
SSTab1.Tab = 1
Option1.Value = True
Option2.Value = False
On Error GoTo err:
txtserverport.SetFocus
txtserverport.SelStart = 0
txtserverport.SelLength = Len(txtserverport.Text)
err:

End Sub

Private Sub Command6_Click()
SSTab1.Tab = 2
txtmsg.SetFocus

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Load()
txtversion.Text = App.Major & "." & App.Minor & "." & App.Revision
txtlocal.Text = "Local Host Name (networking name): " & Winsock1.LocalHostName _
& vbCrLf & "Local IP Address: " & Winsock1.LocalIP & vbCrLf _
& "Local Port: " & Winsock1.LocalPort & vbCrLf _
& vbCrLf & "---------What is---------" & vbCrLf & vbCrLf _
& "An IP address?" & vbCrLf & "An IP address is the address you use when you are online; your computer's online address.  Other computers " _
& "can talk to you using this IP address.  Your LOCAL IP address is listed above, but you need to check your internet connection settings in your ISP's dialer for your IP address." _
& vbCrLf & vbCrLf & "A Port?" & vbCrLf & "A port is space in your computer reserved for connecting to other computers.  Most computers have 5000+ ports on their computer.  For example, TCP Pro asks the server for the port to 'listen' on, and you have to connect to that port so you can chat using TCP Pro." & vbCrLf & vbCrLf & "TCP Pro2 Copyright © 1999 Daniel Errante.  All rights reserved.  Any questions or comments should be sent to danoph@hotmail.com and will get a response within 48 hours."
CommonDialog1.FileName = App.Path & "\knownas.txt"
On Error GoTo toobig:
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo toobig:    'set error handler
        Do Until EOF(1)          'then read lines from file
            Line Input #1, LineOfText$
            alltext$ = alltext$ & LineOfText$
        Loop
        txtcomp.Text = alltext$  'display file
        Close #1                 'close file
If Winsock1.State <> sckClosed Then
On Error GoTo err:
Winsock1.SendData "Other user connected!"
Winsock1.Close
End If
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)
toobig:

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Winsock1.State <> sckClosed Then
On Error GoTo err:
Winsock1.SendData "Other user disconnected!"
Winsock1.Close
End If
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Option1_Click()
Frame1.Enabled = True
Frame2.Enabled = False
lblstatus.Caption = "After you type in the port you will listen on, click listen to listen for users to connect.  Then click on the tab TCP Pro Chat to chat when the other user connects."
txtserverport.SetFocus
txtserverport.SelStart = 0
txtserverport.SelLength = Len(txtserverport.Text)

End Sub

Private Sub Option2_Click()
Frame2.Enabled = True
Frame1.Enabled = False
lblstatus.Caption = "After you type in the server IP address and the port you will connect to, click connect to connect to the server that is listening for you to connect.  NOTE:  There must be a server listening for clients to connect first or TCP connection will not work!"
txtserverip.SetFocus

End Sub

Private Sub Option3_Click()
Option4.Value = False
lbl1.Enabled = True
lbl2.Enabled = True
txt1.Enabled = True
cmd1.Enabled = True
lbl3.Enabled = False
lbl4.Enabled = False
lbl5.Enabled = False
txt2.Enabled = False
txt3.Enabled = False
cmd2.Enabled = False
End Sub

Private Sub Option4_Click()
Option3.Value = False
lbl3.Enabled = True
lbl4.Enabled = True
lbl5.Enabled = True
txt2.Enabled = True
txt3.Enabled = True
cmd2.Enabled = True
lbl1.Enabled = False
lbl2.Enabled = False
txt1.Enabled = False
cmd1.Enabled = False
End Sub

Private Sub Text1_Click()
text1.SelStart = 0
text1.SelLength = Len(text1.Text)

End Sub

Private Sub Text3_Click()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub txtclientport_Click()
With txtclientport
.SelStart = 0
.SelLength = Len(.Text)
End With

End Sub

Private Sub txtcomp_Change()
If txtcomp.Text <> "" Then
    CommonDialog1.FileName = App.Path & "\knownas.txt"
    On Error GoTo err:
    Open CommonDialog1.FileName For Output As #1
    Print #1, txtcomp.Text
    Close #1
End If
Exit Sub
err:

End Sub

Private Sub txtcomp_Click()
txtcomp.SelStart = 0
txtcomp.SelLength = Len(txtcomp.Text)

End Sub

Private Sub txtcomp_DblClick()
txtcomp.SelStart = Len(txtcomp.Text)

End Sub

Private Sub txtversion_Click()
txtversion.SelStart = 0
txtversion.SelLength = Len(txtversion.Text)
End Sub

Private Sub Winsock1_Close()
MsgBox "Connection Closed.", vbExclamation

End Sub

Private Sub Winsock1_Connect()
On Error GoTo err:
MsgBox "User Connected!", vbExclamation
Winsock1.SendData "Other user connected!"
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error GoTo err:
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Me.WindowState = 1 Then Me.WindowState = 0
Dim strdata As String
On Error GoTo err:
Winsock1.GetData strdata
txtCHAT.Text = txtCHAT.Text & txtcomp.Text & ": " & strdata & vbCrLf
txtCHAT.SelStart = Len(txtCHAT.Text)
Exit Sub
err:
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtstatus.Text = txtstatus.Text & err.Description & " - Error number: " & err.Number & vbCrLf
txtstatus.SelStart = Len(txtstatus.Text)

End Sub


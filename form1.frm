VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client2Client"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   600
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtOutput 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    Call tcpClient.SendData("REMOTE >>> " & txtSend.Text)
    txtOutput.Text = txtOutput.Text & _
        "YOUR MESSAGE >>> " & txtSend.Text & vbCrLf & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.Text = ""
    txtSend.SetFocus
End Sub

Private Sub Form_Load()
    cmdSend.Enabled = False
    
    tcpClient.RemoteHost = _
        InputBox("Enter the remote host IP address", _
            "IP Address", "127.0.0.1")
    
    If tcpClient.RemoteHost = "" Then
        tcpClient.RemoteHost = "127.0.0.1"
    End If
    tcpClient.RemotePort = 5000
    Call tcpClient.Connect

End Sub


Private Sub Form_Terminate()
    End
End Sub


Private Sub tcpClient_Close()
    cmdSend.Enabled = False
    Call tcpClient.Close
    txtOutput.Text = _
        txtOutput.Text & "Remote Host closed connection." & vbCrLf & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
    tcpClient.LocalPort = 5000
    tcpClient.Listen
End Sub

Private Sub tcpClient_Connect()
    cmdSend.Enabled = True
    txtOutput.Text = "*** Connected to IP Address:" & tcpClient.RemoteHostIP & " . Port #:" & _
        tcpClient.RemotePort & vbCrLf & vbCrLf
End Sub

Private Sub tcpClient_ConnectionRequest(ByVal requestID As Long)
    If tcpClient.State <> sckClosed Then
        Call tcpClient.Close
    End If
    
    tcpClient.Accept (requestID)
    txtOutput = txtOutput.Text & "*** " & _
        "Request From IP:" & tcpClient.RemoteHostIP & _
        ". Remote Port: " & tcpClient.RemotePort & vbCrLf & vbCrLf
    cmdSend.Enabled = True

End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)

    Dim message As String
    Call tcpClient.GetData(message)
    txtOutput.Text = txtOutput.Text & message & vbCrLf & vbCrLf
    txtOutput.SelStart = Len(txtOutput.Text)
    
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim result As Integer
    
    If Number = 10061 Then
        txtOutput.Text = "Cannot Connect to RomoteHost" & vbCrLf & vbCrLf
    Else
        result = MsgBox(Source & ": " & Description, _
            vbOKOnly, "TCP/IP Error")
    End If
    tcpClient.Close
    tcpClient.LocalPort = 5000
    Call tcpClient.Listen
End Sub


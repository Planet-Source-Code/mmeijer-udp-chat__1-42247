VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UDPChat"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   285
      Left            =   3930
      TabIndex        =   3
      Top             =   2745
      Width           =   435
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   3915
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3780
      TabIndex        =   1
      Top             =   0
      Width           =   570
   End
   Begin VB.TextBox txtHostRemote 
      Height          =   285
      Left            =   585
      TabIndex        =   0
      Top             =   0
      Width           =   2610
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   2040
      Top             =   1425
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtResult 
      Height          =   2430
      Left            =   -15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   300
      Width           =   4365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PORT"
      Height          =   285
      Left            =   3210
      TabIndex        =   6
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HOST"
      Height          =   285
      Left            =   15
      TabIndex        =   5
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const INT_PORT_MAX As Long = 65535
Private Const INT_PORT_START As Long = 2235

Private Sub cmdGo_Click()
    Call Send
End Sub

Private Sub txtPort_LostFocus()
    If Not isPort(txtPort.Text) Then
        MsgBox "Invalid port specified!", vbExclamation + vbOKOnly, "Failure"
        txtPort.Text = INT_PORT_START
    Else
        If txtPort.Text <> Socket.LocalPort Then
            Bind txtPort.Text
        End If
    End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Send
    End If
End Sub

Private Sub Send()
    If txtSend.Text <> "" Then
        If txtHostRemote.Text <> "" Then
            If isPort(txtPort.Text) Then
                On Error Resume Next
                With Socket
                    .RemoteHost = txtHostRemote.Text
                    .RemotePort = txtPort.Text
                    .SendData txtSend.Text
                End With
                txtSend.Text = ""
                If Err.Number <> 0 Then
                    MsgBox "Error occured: " & Err.Description, vbExclamation + vbOKOnly, "Failure"
                End If
            Else
                MsgBox "Invalid port specified!", vbExclamation + vbOKOnly, "Failure"
                txtPort.Text = INT_PORT_START
            End If
        Else
            MsgBox "Host not specified!", vbExclamation + vbOKOnly, "Failure"
        End If
    Else
        MsgBox "Some text please.", vbExclamation + vbOKOnly, "Failure"
    End If
End Sub

Private Sub Form_Load()
    Call Bind(INT_PORT_START)
    txtPort.Text = INT_PORT_START
End Sub

Private Function Bind(ByVal lPort As Long)
    On Error Resume Next
    If Socket.State <> sckClosed Then Socket.Close
    Socket.Bind lPort
    If Err.Number <> 0 Then
        MsgBox "Error occured: " & vbNewLine & Err.Description, vbExclamation + vbOKOnly, "Failure"
    End If
End Function

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    On Error Resume Next
    Socket.GetData strData
    If Len(strData) > 0 Then
        txtResult.Text = txtResult.Text & Socket.RemoteHostIP & ":" & strData & vbNewLine
        txtResult.SelStart = Len(txtResult.Text)
    End If
    If Err.Number <> 0 Then
        MsgBox "Error occured: " & vbNewLine & Err.Description, vbExclamation + vbOKOnly, "Failure"
    End If
End Sub

Private Function isPort(ByVal vData) As Boolean
    If IsNumeric(vData) Then
        If vData > 0 And vData < INT_PORT_MAX Then
            isPort = True
        Else
            isPort = False
        End If
    Else
        isPort = False
    End If
End Function

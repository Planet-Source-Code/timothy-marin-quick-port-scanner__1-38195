VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   6840
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Time Out"
      Height          =   2055
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5415
      Begin MSComctlLib.ListView LV1 
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Sock"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info"
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   4560
      Width           =   5415
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Text            =   "5"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton ClearB 
         Caption         =   "Clear Results"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar P 
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   1320
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton StopB 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Scan 
         Caption         =   "Scan"
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "50"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Text            =   "65000"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Out :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ports Found :"
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   16
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ports Per Second :"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   14
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Socks :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "To "
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ports :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Address :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open"
      Height          =   6255
      Left            =   5520
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      Begin MSComctlLib.ListView List2 
         Height          =   5895
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Receaved"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bytes"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Closed"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   5415
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim times As Long 'Variable For Ports Per Min
Dim stops As Integer 'Variable To Stop Scan
Dim a As Integer


Private Sub Scan_Click()
    'make sure port selections are scanable
    If Text2.Text >= Text3.Text Then
        Text2.Text = Text3.Text - 1
    End If
    
    'set Progrs Bar
    P.Max = Text3.Text
    P.Min = Text2.Text
    
    'set stop var to 0 aka no stop
    stops = 0
    
    'enable/disable things for gui
    StopB.Enabled = True
    Scan.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    
    'make sure max sockets isnt greater then ports bein scanned
    If Text4.Text >= ((Text3.Text - Text2.Text) - 1) Then
        Text4.Text = (Text3.Text - Text2.Text) - 1
    End If
    
    Dim i As Integer ' dim loop Var
    i = 1 ' Set loop Var
    
    'load Sockets Text4.Text = max socket
    Do Until i >= Int(Text4.Text) + 1
        DoEvents ' to prevent freezing
        If stops = 1 Then Exit Sub
        Load Winsock1(i) 'load socket
        Winsock1(i).Connect Text1.Text, Text2.Text ' connect to current port
        Text2.Text = Text2.Text + 1 ' increase port number
        'add to Timeout
        LV1.ListItems.Add , , Text1.Text
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(1) = Winsock1(i).RemotePort
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(2) = i
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(3) = Text5.Text
        i = i + 1 ' goto next i
    Loop
    
End Sub

Private Sub StopB_Click()
    'STOP THE SCAN
    'CLEAR everything and set all GUI stuff back
    LV1.ListItems.Clear
    P.Value = P.Min
    stops = 1
    Scan.Enabled = True
    StopB.Enabled = False
    'UNLOAD THE SOCKETS
    For i = 1 To Winsock1.UBound
        DoEvents
        Winsock1(i).Close
        Unload Winsock1(i)
    Next
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End Sub



Private Sub ClearB_Click()
    'CLEAR RESULTS
    List1.Clear
    List2.ListItems.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    stops = 1 ' STOP SCAN
    Me.Caption = "Unloading..."
    'UNLOAD SOCKETS
    For i = 1 To Winsock1.UBound
        DoEvents
        Winsock1(i).Close
        Unload Winsock1(i)
    Next
    End ' MAKE SURE PROG CLOSES
End Sub
Private Sub Text2_Change()
'ONLY ALLOW CERTIN THINGS FOR THIS TEXT BOX
    If Int(Text2.Text) > Int(Text3.Text) Then
    Text2.Text = "0"
    End If
    If Not IsNumeric(Text2.Text) Then
        Text2.Text = "0"
    End If
    If Text2.Text > 65530 Then
        Text2.Text = "65000"
    End If
End Sub

Private Sub Text3_Change()
'ONLY ALLOW CERTIN THINGS FOR THIS TEXT BOX
    If Text2.Text > Text3.Text Then
    Text2.Text = "0"
    End If
    If Not IsNumeric(Text3.Text) Then
        Text3.Text = "65000"
    End If
    If Text3.Text > 65530 Then
        Text3.Text = "65000"
    End If
End Sub

Private Sub Text4_Change()
'ONLY ALLOW CERTIN THINGS FOR THIS TEXT BOX
    If Not IsNumeric(Text4.Text) Then
        Text4.Text = "50"
    End If
    If Text4.Text > 175 Then
        Text4.Text = "175"
    End If
End Sub

Private Sub Text5_Change()
'ONLY ALLOW CERTIN THINGS FOR THIS TEXT BOX
    If Not IsNumeric(Text5.Text) Then
        Text5.Text = "5"
    End If
    If Text5.Text > 20 Then
        Text5.Text = "20"
    End If
End Sub

Private Sub Timer1_Timer()
    On Error GoTo err 'THERE WAS AN ERROR .. EXIT SUB
    a = 0
    'SET PORT PER MIN / FOUND PORTS
    Label3.Caption = List2.ListItems.Count
    Label2.Caption = Text2.Text - times
    times = Text2.Text 'SET PORTS FOR NEXT TIME
    Dim i, Y As Integer 'DIM VARS
    
    'SOCKET TIME OUT
    If LV1.ListItems.Count < 1 Then Exit Sub 'IF THERE IS NONE EXIT
    For i = 1 To LV1.ListItems.Count
        DoEvents
        LV1.ListItems.Item(i).SubItems(3) = LV1.ListItems.Item(i).SubItems(3) - 1 'SET IT BACK 1 SEC
        If LV1.ListItems.Item(i).SubItems(3) < 1 Then ' IF ITS AT 0 PREPARE TO REMOVE IT
            'IF SOCKET CONNECTED AND WAITING FOR INFO
            For Y = 1 To List2.ListItems.Count
                DoEvents
                If List2.ListItems.Item(Y) = Text1.Text Then
                    If List2.ListItems.Item(Y).SubItems(1) = Winsock1(LV1.ListItems.Item(i).SubItems(2)).RemotePort Then
                        List2.ListItems.Item(Y).SubItems(2) = "..."
                        List2.ListItems.Item(Y).SubItems(3) = "0"
                    End If
                End If
            Next
            'K NOW CLOSE IT AND REUSE THE SOCKET
            
            Winsock1(LV1.ListItems.Item(i).SubItems(2)).Close
            List1.AddItem Winsock1(LV1.ListItems.Item(i).SubItems(2)).RemotePort & " : " & "Closed"
            If stops = 1 Then Exit Sub
            If Int(Text2.Text) > (Int(Text3.Text) - 1) Then
                StopB_Click
                Exit Sub
            End If
            Winsock1(LV1.ListItems.Item(i).SubItems(2)).Connect Text1.Text, Text2.Text
                LV1.ListItems.Add , , Text1.Text
                LV1.ListItems.Item(LV1.ListItems.Count).SubItems(1) = Winsock1(LV1.ListItems.Item(i).SubItems(2)).RemotePort
                LV1.ListItems.Item(LV1.ListItems.Count).SubItems(2) = LV1.ListItems.Item(i).SubItems(2)
                LV1.ListItems.Item(LV1.ListItems.Count).SubItems(3) = Text5.Text
            Text2.Text = Text2.Text + 1
            P.Value = Text2.Text
            LV1.ListItems.Remove i
        End If
    Next
err:

End Sub

Private Sub Winsock1_Close(Index As Integer)
    On Error Resume Next
    'REMOVE IT FROM TIMEOUT
    For i = 1 To LV1.ListItems.Count
    DoEvents
        If LV1.ListItems.Item(i) = Text1.Text Then
            If LV1.ListItems.Item(i).SubItems(1) = Winsock1(Index).RemotePort Then
                LV1.ListItems.Remove i
            End If
        End If
    Next
    'K NOW CLOSE IT AND REUSE THE SOCKET
    List1.AddItem Winsock1(Index).RemotePort & " : " & "Closed"
    If stops = 1 Then Exit Sub
    Winsock1(Index).Close
    If Int(Text2.Text) > (Int(Text3.Text) - 1) Then
        StopB_Click
        Exit Sub
    End If
    Winsock1(Index).Connect Text1.Text, Text2.Text
        LV1.ListItems.Add , , Text1.Text
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(1) = Winsock1(Index).RemotePort
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(2) = Index
        LV1.ListItems.Item(LV1.ListItems.Count).SubItems(3) = Text5.Text
        Text2.Text = Text2.Text + 1
        P.Value = Text2.Text
        
    'FORCE UPDATE
    If List2.ListItems.Count > Label3.Caption Then
        Timer1_Timer
    End If
    a = a + 1
    If a >= Text4.Text Then
        a = 0
        Timer1_Timer
    End If
    
End Sub

Private Sub Winsock1_Connect(Index As Integer)
    'SOCKET CONNECTED GET INFO BEFOR CLOSING
    List2.ListItems.Add , , Text1.Text
    List2.ListItems.Item(List2.ListItems.Count).SubItems(1) = Winsock1(Index).RemotePort
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'SUMTHIN CAME ON A SOCKET
    Dim sata As String 'DIM STRING TO GET
    If Winsock1(Index).State = 7 Then ' MAKE SURE UR STILL CONNECTED AND DIDNT TIMEOUT
    Winsock1(Index).GetData sata ' GET INFO
    On Error Resume Next
    'LOOp THROUGH CONNECTED SOCKS AND ADD THIS INFO TO THE LIST
    For i = 1 To List2.ListItems.Count
        If List2.ListItems.Item(i) = Text1.Text Then
            If List2.ListItems.Item(i).SubItems(1) = Winsock1(Index).RemotePort Then
                List2.ListItems.Item(i).SubItems(2) = sata
                List2.ListItems.Item(i).SubItems(3) = bytesTotal
            End If
        End If
    Next
    End If
    Winsock1_Close Index 'CLOSE AND REUSE SOCKET
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1_Close Index 'ERROR CLOSE AND REUSE SOCKET
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMSIntranet"
   ClientHeight    =   5145
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   10200
      Top             =   4200
   End
   Begin VB.Timer Timer4 
      Interval        =   1500
      Left            =   8640
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4860
      Top             =   4020
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   4380
      Top             =   4020
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "SMS Service Status "
      Height          =   1755
      Left            =   420
      TabIndex        =   6
      Top             =   1740
      Width           =   3375
      Begin VB.Label Label7 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   9
         Top             =   1260
         Width           =   2730
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   8
         Top             =   840
         Width           =   2670
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   7
         Top             =   420
         Width           =   2595
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3720
      Top             =   840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   1058
      ButtonWidth     =   2725
      ButtonHeight    =   953
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SMS Service"
            Key             =   "Service"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Start"
                  Text            =   "&Start SMS Service"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stop"
                  Text            =   "&Stop SMS Service"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Message Settings  "
            Key             =   "Message"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Phone Settings"
            Key             =   "Phone"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Users"
            Key             =   "Users"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Database"
            Key             =   "Database"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   4755
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17568
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2002-03-30"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "08:30"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2595
      Left            =   4380
      TabIndex        =   10
      Top             =   1860
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From / To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text"
         Object.Width           =   10125
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   4320
      Picture         =   "frmMain.frx":34CC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Stop SMS Service"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   5
      Top             =   3900
      Width           =   2115
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Start SMS Service"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2190
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Checking for new messages in the database !!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   5970
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Sending the messages from the database !!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5340
      TabIndex        =   2
      Top             =   1140
      Width           =   6300
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   11520
      Picture         =   "frmMain.frx":65DC
      Stretch         =   -1  'True
      Top             =   660
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Public flag As Integer
    Dim int_i As Integer
    Dim J, jj, mess_num As Integer
    Dim read(2000)
    Dim spoji
    Dim str_PDUInput As String
    
    Dim lv As ListItem
    Public T1, T2 As String
    
    Public multisend As Boolean
    
'    Public SMSSend As SMSSend
    
    Public ErrorSMS As Integer
    Public Slanje As Boolean
'    Dim a
    
    



Private Sub Form_Load()

On Error Resume Next
    Image2.Visible = False
    Label2.Visible = False
    Image1.Visible = False
    Label1.Visible = False
    
    If MSComm1.PortOpen = True Then
        Label5.Caption = "Communication Port Open"
    Else
        Label5.Caption = "Communication Port Closed"
    End If
    
    
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set a = fs.CreateTextFile("c:\Temp\SMSIntranet.txt", True)
'a.WriteLine ("SMS Intranet log at: " & Time)
'a.WriteBlankLines (2)



    
    Timer1.Enabled = False
    
    int_i = 0
    flag = 0
    J = 1
    
    
    
                    Dim inbox As ADODB.Connection
                    Dim str_inbox As String
                    Dim str_rs As String
                    Dim rs As ADODB.Recordset

                    Set inbox = New ADODB.Connection
                    str_inbox = "DSN=SMSIntranet;UID=sa;pwd="
                    Set rs = New ADODB.Recordset

                    inbox.CommandTimeout = 6
                    str_rs = "SELECT * FROM Buffer"

                    inbox.Open str_inbox

                        rs.CursorLocation = adUseClient
                        rs.CursorType = adOpenKeyset
                        rs.LockType = adLockOptimistic

                    rs.Open str_rs, inbox, , , adCmdText
                    If rs.EOF = True Then
                    Else
                    rs.MoveFirst
                    T1 = rs.RecordCount
                    End If
                    rs.Close

                    inbox.Close

    
    
    multisend = False
    Timer2.Enabled = False
    Timer4.Enabled = False
    ErrorSMS = 0
    Slanje = False
    
    
    
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.Font.Underline = False
    Label3.ForeColor = &H812
    Label4.Font.Underline = False
    Label4.ForeColor = &H812

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call forms
'a.Close

End
'    Dim i As Integer
'
'    'close all sub forms
'    For i = forms.Count - 1 To 1 Step -1
'        Unload forms(i)
'    Next
End Sub

Private Sub Label3_Click()
        flag = 0
        
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        Else
        End If
        
        MSComm1.InputLen = 0
        
        MSComm1.CommPort = GetSetting(App.Title, "Settings", "PortNumber", "1")
        MSComm1.Settings = GetSetting(App.Title, "Settings", "DataSpeed", "9600") & "," _
        & Left(GetSetting(App.Title, "Settings", "Parity", "None"), 1) & "," _
        & GetSetting(App.Title, "Settings", "DataBits", "8") & "," _
        & GetSetting(App.Title, "Settings", "StopBit", "1")
        MSComm1.RTSEnable = GetSetting(App.Title, "Settings", "RTS", False)
        Select Case GetSetting(App.Title, "Settings", "FlowControl", "RTS")
            Case "None"
                MSComm1.Handshaking = comNone
            Case "RTS"
                MSComm1.Handshaking = comRTS
            Case "XOnXOff"
                MSComm1.Handshaking = comXOnXoff
            Case "RTSXOnXOff"
                MSComm1.Handshaking = comRTSXOnXOff
        End Select
        
        MSComm1.PortOpen = True
        Label5.Caption = "Communication Port Open"
        Label5.ForeColor = &HFF
        MSComm1.Output = "AT" & vbCr
        Timer1.Enabled = True

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.Font.Underline = True
    Label3.ForeColor = &HFF
End Sub

Private Sub Label4_Click()
Timer2.Enabled = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Font.Underline = True
    Label4.ForeColor = &HFF
End Sub

Private Sub mnuExit_Click()
    Call forms
End
End Sub

Private Sub MSComm1_OnComm()
'each time when character is received in input buffer MSComm control is
'raising the event.
'MSComm logic is the folloving:
'When first character is received start the timer with offset 0.5 sec.
'After 0.5 sec read the port.

    Select Case MSComm1.CommEvent
        Case comEvReceive
        If flag = 11 Then
'Timer 3 mi sluzi za brisanje smeca koje dodje u input buffer!
'            offset3 = Timer + 3
            Timer3.Enabled = True
            MSComm1.InputLen = 1
            J = J + 1
            Call ulaz1(MSComm1.Input)
        Else
        End If

        int_i = int_i + 1
        
        If int_i = 1 Then
            Timer1.Enabled = True
        Else
        End If
    
    End Select

End Sub

Private Sub Timer1_Timer()

    Timer1.Enabled = False

    int_i = 0

    Select Case flag
    
    Case 0
'########################################################
'If we get OK respond then ask for network registration.
'Else - increase offset and ask again. Up to 8 times.
'Else - Error message box
'########################################################

        
'#######################################################
'If you have connected GSM phone to your com port then
' coment following code
'
        Label6.Caption = "Simulation work"
        read(J) = MSComm1.Input
        spoji = Join(read, "")
        
        Label7.Caption = "Check the code for more info !"
        flag = 11
        MSComm1.InputLen = 1
        Label7.ForeColor = &HFF
        Timer2.Enabled = True
' up to here
'and uncoment the following code
            
        
'            If InStr(1, spoji, "OK", vbTextCompare) <> 0 Then
'                For J = 1 To J + 1
'                read(J) = ""
'                Next
'                spoji = ""
'                flag = 1
'                J = 1
'                MSComm1.Output = "AT+CREG?" & vbCr
'
'                Exit Sub
'            Else
'                J = J + 1
'
'                If J > 4 Then
'                    sbStatusBar.Panels(2).Text = "Configuring Error!"
'                    MsgBox "GSM terminal is not responding!" & vbCr _
'                    & "Please make shore that you have connect GSM terminal to right port.", _
'                    vbExclamation, "GSM Terminal"
'                    For J = 1 To J + 1
'                    read(J) = ""
'                    Next
'                    J = 1
'                    spoji = ""
'                    Exit Sub
'                Else
'
'                    flag = 0
'                    MSComm1.Output = "AT" & vbCr
'                    Timer1.Enabled = True
'                End If
'            End If
'up to here
'######################################################################
            
            
            
            
                Case 1
'#######################################################################
'If the mobile is registarted to home network ask for the operator name.
'Else - ask for a pin
'#######################################################################
        read(J) = MSComm1.Input
        spoji = Join(read, "")

            If InStr(1, spoji, "+CREG", vbTextCompare) <> 0 Then

                If InStr(1, spoji, "1", vbTextCompare) <> 0 Then
                    If InStr(1, spoji, "OK", vbTextCompare) <> 0 Then
                        For J = 1 To J + 1
                        read(J) = ""
                        Next
                        spoji = ""
                        MSComm1.Output = "AT+COPS?" & vbCr
                        flag = 6
                    
                        Exit Sub
                    Else
                    End If
                Else
                End If

                If InStr(1, spoji, "ERROR", vbTextCompare) <> 0 Then
                        spoji = ""
                        For J = 1 To J + 1
                        read(J) = ""
                        Next
                        MSComm1.Output = "AT+CPIN?" & vbCr
                        flag = 2
           
                        J = 1
                        Exit Sub
                Else
                End If
'            Else



                If J > 30 Then
'                    MsgBox "GSM terminal is not responding!" & vbCr _
'                    & "Please make shore that you have connect GSM terminal to right port.", _
'                    vbExclamation, "GSM Terminal"
                        spoji = ""
                        For J = 1 To J + 1
                        read(J) = ""
                        Next
                        flag = 2
                        MSComm1.Output = "AT+CPIN?" & vbCr
                 
                        J = 1
                        Exit Sub

                Else
                        J = J + 1
                        spoji = ""
                        For J = 1 To J + 1
                        read(J) = ""
                        Next
                MSComm1.Output = "AT+CREG?" & vbCr
                flag = 1
                Exit Sub
                End If
            Else
            End If
            
            
            
    Case 6
'################################################
'Read the name of the GSM network provider and
'Configure the mobile phone
'###############################################
            read(J) = MSComm1.Input
            spoji = Join(read, "")


            If InStr(1, spoji, "OK", vbTextCompare) <> 0 Then
                Dim b, c
                b = Split(spoji, Chr(34), -1, vbTextCompare)
                c = b(1)
                    Label6.Caption = c
                    Label6.ForeColor = &HFF
                    Label7.Caption = "Configuring"
                
                For J = 1 To J + 1
                read(J) = ""
                Next
                spoji = ""
                If Len(GetSetting(App.Title, "Settings", "IString", "")) > 0 Then
                    MSComm1.Output = GetSetting(App.Title, "Settings", "IString", "") & vbCr
                    flag = 7
                    Timer1.Enabled = True
                    Exit Sub
                Else
                    flag = 8
   
                    Timer1.Enabled = True
                    Exit Sub
                End If

            Else

                J = J + 1
             
                flag = 6
                Timer1.Enabled = True

            End If
            
            
            
            
            
            
              Case 7
'##############################################
'Chack for a configuring Error
'else finish configuration and start the SMS service
'##############################################

        
        
        read(J) = MSComm1.Input
        spoji = Join(read, "")
 
            
            If InStr(1, spoji, "OK", vbTextCompare) <> 0 Then
                    For J = 1 To J + 1
                    read(J) = ""
                    Next
                    spoji = ""
               flag = 8
               Timer1.Enabled = True
                
                
            Else
            
                J = J + 1
                If J > 8 Then
                    MsgBox "GSM terminal configuring Error", _
                    vbExclamation, "GSM Terminal"
                  
                For J = 1 To J + 1
                read(J) = ""
                Next
                spoji = ""
                J = 1
                Else
                
            
                    flag = 7
                    Timer1.Enabled = True
                End If
            End If

    
        Case 8
        
'#############################################
'Timer 2 - Start of the SMS Service
'#############################################

            flag = 11
            MSComm1.InputLen = 1
            Label7.Caption = "Ready"
            Label7.ForeColor = &HFF

            Timer2.Enabled = True
    
    

            
              
            
                Case 13
                
'Donja dva reda treba vidjeti mogu li se izbrisati!
        read(J) = MSComm1.Input
        spoji = read(J)
                flag = 11
                J = J + 1

                MSComm1.Output = SMSSend.PDUOutputMessage & Chr(26)
                Timer1.Enabled = True
    
            
            
End Select
End Sub

Private Sub Timer2_Timer()

            Timer4.Enabled = True
            Image2.Visible = True
            Label2.Visible = True

                    Dim inbox As ADODB.Connection
                    Dim str_inbox As String
                    Dim str_rs As String
                    Dim str_delete As String
                    Dim rs As ADODB.Recordset

                    Set inbox = New ADODB.Connection
                    str_inbox = "DSN=SMSIntranet;UID=sa;pwd="
                    str_delete = "DELETE FROM Buffer"
                    Set rs = New ADODB.Recordset

                    inbox.CommandTimeout = 6
                    str_rs = "SELECT * FROM Buffer"

                    inbox.Open str_inbox

                        rs.CursorLocation = adUseClient
                        rs.CursorType = adOpenKeyset
                        rs.LockType = adLockOptimistic

                    rs.Open str_rs, inbox, , , adCmdText
                    
                    
                        If rs.EOF = True Then
                        
                            rs.Close
                            inbox.Close
                            Exit Sub
                                         
                        
                        Else
                        
                        rs.MoveFirst
                        
                                   
                        
                        Do While rs.EOF = False
                        
                            Set lv = ListView2.ListItems.Add
                            Dim MyString, MyArray, Msg
                            MyString = rs(1).Value
                            MyArray = Split(MyString, "+", -1, 1)
                            ' MyArray(0) contains "VBScript".
                            ' MyArray(1) contains "is".
                            ' MyArray(2) contains "fun!".
                            

                            lv.Text = Replace(MyArray(1), " ", "", 1, -1, vbTextCompare)
                            lv.SubItems(1) = rs(2).Value
'                            a.WriteLine (Replace(MyArray(1), " ", "", 1, -1, vbTextCompare) & "/" & rs(2).Value)

                        rs.MoveNext

                        Loop
                    
                        
                        rs.Close
                        inbox.Execute str_delete
                        inbox.Close
                        
                        End If

                        
        
        
        multisend = True
        Timer2.Enabled = False
        Image1.Visible = True
        Label1.Visible = True
        
'################################################
'Calling the procedure for sending SMS message in PDU format.
'
'        Call send
'################################################


Timer5.Enabled = True
        

 

        
        

End Sub

Private Sub Timer3_Timer()
        Timer3.Enabled = False

'###########################################
'Flushing the input buffer
'##########################################

'        MsgBox "Èišæenje input buffera!" & spoji

        If Slanje = True Then
        MSComm1.Output = vbKeyEscape
        Else
        End If
                    For J = 1 To J + 1
                    read(J) = ""
                    Next
                    spoji = ""

End Sub

Private Sub Timer4_Timer()
    Image2.Visible = False
    Label2.Visible = False
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
    Timer5.Enabled = False
    Image1.Visible = False
    Label1.Visible = False
    ListView2.ListItems.Clear
    Timer2.Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call forms
    Select Case Button.Key
        
        Case "Service"
        
        Case "Message"
        
            frmMessage.Visible = True
            frmMessage.SetFocus
        Case "Phone"
            frmPhone.Visible = True
            frmPhone.SetFocus
        Case "Users"
            frmUsers.Visible = True
            frmUsers.SetFocus
        
        Case "Database"
            frmInbox.Visible = True
            frmInbox.SetFocus
            
        Case "dbBuffer"
        
    End Select
End Sub

Public Function forms()
    
    If frmMessage.Visible = True Then
        frmMessage.Visible = False
    Else
    End If
    
    If frmPhone.Visible = True Then
        frmPhone.Visible = False
    Else
    End If
    
    If frmUsers.Visible = True Then
        frmUsers.Visible = False
    Else
    End If
    
    
End Function

Public Function ulaz1(kar As String)

            read(J) = kar
            spoji = Join(read, "")
                
'Ovdje ugasi sliku i text Sending away
                If InStr(1, spoji, "ERROR", vbTextCompare) <> 0 Then
                    For J = 1 To J + 1
                    read(J) = ""
                    Next
                    
                            Call fjamultisend
                Else
                End If












End Function

Function DecodeSMS()

    Set inbox = New ADODB.Connection
    str_inbox = "DSN=SMS;UID=sa;pwd="
    inbox.CommandTimeout = 6
    Set rsu = New ADODB.Recordset
    rsu.CursorLocation = adUseClient
    rsu.CursorType = adOpenKeyset
    rsu.LockType = adLockOptimistic

'                    Set SMS = New SMSSend
                    SMS.PDUInputMessage = str_PDUInput
                    SMS.Decode
     inbox.Open str_inbox
          
    str_input = "INSERT INTO Inbox( MobileNumber,SendTime,Content,MRead,Status)" _
    & "VALUES ('" & SMS.DTPOriginatedAddress & "','" & SMS.DTPSendTime & "','" _
    & SMS.DTPUserData & "','No','Inbox')"
    
    
    inbox.Execute str_input
    
    If lvview <> 2 Then
    
    Set lv = ListView1.ListItems.Add(1, , "", , "NewA")

            Dim str_rsu As String
            str_rsu = "SELECT MobilePhone,Name FROM Users"

            rsu.Open str_rsu, inbox, , , adCmdText
            
            rsu.MoveFirst
            Do While rsu.EOF = False
            If StrComp(SMS.DTPOriginatedAddress, rsu(0).Value, vbTextCompare) = "0" Then
            lv.Text = rsu(1).Value
            Exit Do
            Else
            End If
            rsu.MoveNext
            Loop
            If lv.Text = "" Then
                lv.Text = SMS.DTPOriginatedAddress
            Else
            End If

            lv.SubItems(2) = Format(SMS.DTPSendTime, "ddd, hh:mm:ss") ' & S.SendTime
            lv.SubItems(3) = SMS.DTPSendTime
            lv.SubItems(1) = SMS.DTPUserData
            inbox.Close
    int_i = 0
    Else
    End If
End Function

Public Function send()
            
'        Set SMSSend = New SMSSend
        SMSSend.STPDestinationAddress = ListView2.ListItems(1).Text
        SMSSend.STPUserData = ListView2.ListItems(1).ListSubItems(1).Text
        If GetSetting(App.Title, "Settings", "Default Service Centre Address", "1") = 1 Then
            SMSSend.STPServiceCentreAddress = "00"
        Else
        SMSSend.STPServiceCentreAddress = GetSetting(App.Title, "Settings", "Service Centre Address", "")
        End If
        SMSSend.STPMessageReferenceNumber = GetSetting(App.Title, "Settings", "Message Reference Number", "1")
        SMSSend.STPRejectDuplicates = GetSetting(App.Title, "Settings", "Reject Duplicates", "0")
        SMSSend.STPStatusReportRequest = GetSetting(App.Title, "Settings", "Status Report Requested", "0")
        SMSSend.STPReplyPath = GetSetting(App.Title, "Settings", "Reply Path", "0")
     
        SMSSend.Code
        flag = 13
        Slanje = True
        MSComm1.Output = "AT+CMGS=" & SMSSend.STPNumberOfOctets & vbCr
        
End Function

Function fjamultisend()
    
    If ListView2.ListItems.Count = 0 Then
        multisend = False
        Image1.Visible = False
        Label1.Visible = False
        Timer2.Enabled = True
        Exit Function
    Else
        multisend = True
        Call send
    End If

End Function

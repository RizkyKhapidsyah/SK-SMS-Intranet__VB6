VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   3915
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7740
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Height          =   240
      Left            =   3150
      TabIndex        =   34
      Top             =   675
      Width           =   315
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current SMS Settings "
      Height          =   2490
      Left            =   5250
      TabIndex        =   26
      Top             =   675
      Width           =   2265
      Begin VB.CheckBox Check3 
         Height          =   240
         Left            =   225
         TabIndex        =   32
         Top             =   2025
         Width           =   1440
      End
      Begin VB.CheckBox Check2 
         Height          =   240
         Left            =   225
         TabIndex        =   31
         Top             =   1725
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Height          =   240
         Left            =   225
         TabIndex        =   30
         Top             =   1425
         Width           =   240
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   225
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   225
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   225
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   7740
      TabIndex        =   18
      Top             =   3165
      Width           =   7740
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1425
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   75
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   4050
         TabIndex        =   22
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2700
         TabIndex        =   21
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1425
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   59
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7740
      TabIndex        =   12
      Top             =   3615
      Width           =   7740
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   6000
         Picture         =   "frmMessage.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   5400
         Picture         =   "frmMessage.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   750
         Picture         =   "frmMessage.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   75
         Picture         =   "frmMessage.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1425
         TabIndex        =   17
         Top             =   0
         Width           =   3810
      End
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Reject_Duplicates"
      Height          =   285
      Index           =   5
      Left            =   3090
      TabIndex        =   11
      Top             =   2640
      Width           =   1725
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Status_Report_Requested"
      Height          =   285
      Index           =   4
      Left            =   3090
      TabIndex        =   9
      Top             =   2310
      Width           =   1875
   End
   Begin VB.CheckBox chkFields 
      DataField       =   "Reply_Path"
      Height          =   285
      Index           =   3
      Left            =   3090
      TabIndex        =   7
      Top             =   1995
      Width           =   1800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Validy_Period"
      Height          =   285
      Index           =   2
      Left            =   3090
      TabIndex        =   5
      Top             =   1680
      Width           =   1800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Message_Reference_Number"
      Height          =   285
      Index           =   1
      Left            =   3090
      TabIndex        =   3
      Top             =   1350
      Width           =   1800
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Service_Centre_Address_Number"
      Height          =   285
      Index           =   0
      Left            =   3090
      TabIndex        =   1
      Top             =   1050
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "Use default Service Centre Address:"
      ForeColor       =   &H80000018&
      Height          =   240
      Left            =   450
      TabIndex        =   35
      Top             =   675
      Width           =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "Set new SMS settings !"
      Height          =   240
      Left            =   2025
      MouseIcon       =   "frmMessage.frx":0D08
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   225
      Width           =   4665
   End
   Begin VB.Label lblLabels 
      Caption         =   "Reject_Duplicates:"
      Height          =   255
      Index           =   5
      Left            =   420
      TabIndex        =   10
      Top             =   2640
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      Caption         =   "Status_Report_Requested:"
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   8
      Top             =   2310
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      Caption         =   "Reply_Path:"
      Height          =   255
      Index           =   3
      Left            =   450
      TabIndex        =   6
      Top             =   2025
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      Caption         =   "Validy_Period:"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   4
      Top             =   1680
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      Caption         =   "Message_Reference_Number:"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1350
      Width           =   2640
   End
   Begin VB.Label lblLabels 
      Caption         =   "Service_Centre_Address_Number:"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   1050
      Width           =   2640
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()


  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDASQL;dsn=SMSIntranet;uid=sa;pwd=;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Service_Centre_Address_Number,Message_Reference_Number,Validy_Period,Reply_Path,Status_Report_Requested,Reject_Duplicates from Message Order by Service_Centre_Address_Number", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Dim oCheck As CheckBox
  'Bind the check boxes to the data provider
  For Each oCheck In Me.chkFields
    Set oCheck.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
  
Text1.Text = GetSetting(App.Title, "Settings", "Service Centre Address", "Parameter Not Set !")
Text2.Text = GetSetting(App.Title, "Settings", "Message Reference Number", "Parameter Not Set !")
Text3.Text = GetSetting(App.Title, "Settings", "Validy Period", "Parameter Not Set !")
Check1.Value = GetSetting(App.Title, "Settings", "Reply Path", "0")
Check2.Value = GetSetting(App.Title, "Settings", "Status Report Requested", "0")
Check3.Value = GetSetting(App.Title, "Settings", "Reject Duplicates", "0")
Check4.Value = GetSetting(App.Title, "Settings", "Default Service Centre Address", "1")






  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Font.Underline = False
End Sub

Private Sub Form_Resize()
'  On Error Resume Next
'  lblStatus.Width = Me.Width - 1500
'  cmdNext.Left = lblStatus.Width + 700
'  cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub

Private Sub Label1_Click()
            SaveSetting App.Title, "Settings", "Service Centre Address", txtFields(0).Text
            SaveSetting App.Title, "Settings", "Default Service Centre Address", Check4.Value
            SaveSetting App.Title, "Settings", "Message Reference Number", txtFields(1).Text
            SaveSetting App.Title, "Settings", "Validy Period", txtFields(2).Text
            SaveSetting App.Title, "Settings", "Reply Path", chkFields(3).Value
            SaveSetting App.Title, "Settings", "Status Report Requested", chkFields(4).Value
            SaveSetting App.Title, "Settings", "Reject Duplicates", chkFields(5).Value
    
    
    
Text1.Text = GetSetting(App.Title, "Settings", "Service Centre Address", "Parameter Not Set !")
Text2.Text = GetSetting(App.Title, "Settings", "Message Reference Number", "Parameter Not Set !")
Text3.Text = GetSetting(App.Title, "Settings", "Validy Period", "Parameter Not Set !")
Check1.Value = GetSetting(App.Title, "Settings", "Reply Path", "0")
Check2.Value = GetSetting(App.Title, "Settings", "Status Report Requested", "0")
Check3 = GetSetting(App.Title, "Settings", "Reject Duplicates", "0")
Check4.Value = GetSetting(App.Title, "Settings", "Default Service Centre Address", "1")

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbCustom
    Label1.Font.Underline = True
End Sub


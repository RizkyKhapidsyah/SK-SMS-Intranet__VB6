VERSION 5.00
Begin VB.Form frmPhone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Phone"
   ClientHeight    =   5550
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Current Comm. Settings "
      Height          =   3090
      Left            =   4200
      TabIndex        =   34
      Top             =   600
      Width           =   2190
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   41
         Text            =   "Text7"
         Top             =   2625
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   40
         Text            =   "Text6"
         Top             =   2175
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   38
         Text            =   "Text4"
         Top             =   1425
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   37
         Text            =   "Text3"
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   150
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RTS"
      Height          =   465
      Left            =   225
      TabIndex        =   24
      Top             =   3150
      Width           =   3540
      Begin VB.OptionButton Option2 
         Caption         =   "Disable"
         Height          =   240
         Left            =   1800
         TabIndex        =   26
         Top             =   150
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Enable"
         Height          =   240
         Left            =   600
         TabIndex        =   25
         Top             =   150
         Width           =   1515
      End
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   2100
      TabIndex        =   23
      Text            =   "Combo6"
      Top             =   2775
      Width           =   1665
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2100
      TabIndex        =   22
      Text            =   "Combo5"
      Top             =   2400
      Width           =   1665
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2100
      TabIndex        =   21
      Text            =   "Combo4"
      Top             =   2025
      Width           =   1665
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2100
      TabIndex        =   20
      Text            =   "Combo3"
      Top             =   1650
      Width           =   1665
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2100
      TabIndex        =   19
      Text            =   "Combo2"
      Top             =   1275
      Width           =   1665
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2100
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   900
      Width           =   1665
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6720
      TabIndex        =   10
      Top             =   4725
      Width           =   6720
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1425
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   75
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5550
         TabIndex        =   15
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2850
         TabIndex        =   13
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1425
         TabIndex        =   12
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   59
         TabIndex        =   11
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
      ScaleWidth      =   6720
      TabIndex        =   4
      Top             =   5250
      Width           =   6720
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   6225
         Picture         =   "frmPhone.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   5550
         Picture         =   "frmPhone.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   750
         Picture         =   "frmPhone.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   75
         Picture         =   "frmPhone.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   420
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1425
         TabIndex        =   9
         Top             =   0
         Width           =   3360
      End
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H80000018&
      DataField       =   "InitString"
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   4200
      Width           =   4875
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      DataField       =   "Type"
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   1
      Top             =   3900
      Width           =   4875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "Set new communication settings !"
      Height          =   240
      Left            =   375
      MouseIcon       =   "frmPhone.frx":0D08
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   150
      Width           =   5940
   End
   Begin VB.Label Label1 
      Caption         =   "Flow Control:"
      Height          =   315
      Index           =   6
      Left            =   225
      TabIndex        =   32
      Top             =   2775
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Stop Bit:"
      Height          =   315
      Index           =   5
      Left            =   225
      TabIndex        =   31
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Parity:"
      Height          =   315
      Index           =   4
      Left            =   225
      TabIndex        =   30
      Top             =   2025
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Data Bits:"
      Height          =   315
      Index           =   3
      Left            =   225
      TabIndex        =   29
      Top             =   1650
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Data Speed:"
      Height          =   315
      Index           =   2
      Left            =   225
      TabIndex        =   28
      Top             =   1275
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Port Number:"
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   27
      Top             =   900
      Width           =   1740
   End
   Begin VB.Label lblLabels 
      Caption         =   "InitString:"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   4200
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "Type:"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   3885
      Width           =   765
   End
End
Attribute VB_Name = "frmPhone"
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


Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"

Combo2.AddItem "1200"
Combo2.AddItem "2400"
Combo2.AddItem "4800"
Combo2.AddItem "9600"
Combo2.AddItem "14400"
Combo2.AddItem "19200"
Combo2.AddItem "28800"

Combo3.AddItem "6"
Combo3.AddItem "7"
Combo3.AddItem "8"

Combo4.AddItem "None"
Combo4.AddItem "Even"
Combo4.AddItem "Odd"
Combo4.AddItem "Mark"

Combo5.AddItem "1"
Combo5.AddItem "1.5"
Combo5.AddItem "2"

Combo6.AddItem "None"
Combo6.AddItem "XOnXOff"
Combo6.AddItem "RTS"
Combo6.AddItem "RTSXOnXOff"


Combo1.Text = GetSetting(App.Title, "Settings", "PortNumber", "1")
Combo2.Text = GetSetting(App.Title, "Settings", "DataSpeed", "9600")
Combo3.Text = GetSetting(App.Title, "Settings", "DataBits", "8")
Combo4.Text = GetSetting(App.Title, "Settings", "Parity", "None")
Combo5.Text = GetSetting(App.Title, "Settings", "StopBit", "1")
Combo6.Text = GetSetting(App.Title, "Settings", "FlowControl", "None")
Option1.Value = GetSetting(App.Title, "Settings", "RTS", False)

Call RTS

Text1.Text = GetSetting(App.Title, "Settings", "PortNumber", "Parameter Not Set !")
Text2.Text = GetSetting(App.Title, "Settings", "DataSpeed", "Parameter Not Set !")
Text3.Text = GetSetting(App.Title, "Settings", "DataBits", "Parameter Not Set !")
Text4.Text = GetSetting(App.Title, "Settings", "Parity", "Parameter Not Set !")
Text5.Text = GetSetting(App.Title, "Settings", "StopBit", "Parameter Not Set !")
Text6.Text = GetSetting(App.Title, "Settings", "FlowControl", "Parameter Not Set !")
Dim str_rts As String
str_rts = GetSetting(App.Title, "Settings", "RTS", False)
If str_rts = "True" Then
    Text7.Text = "Enable"
Else
    Text7.Text = "Disable"
End If





 



  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=MSDASQL;dsn=SMSIntranet;uid=sa;pwd=;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Type,InitString from Phone Order by Type", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  mbDataChanged = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Font.Underline = False
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

Function RTS()
    
    If Option1.Value = True Then
        Option2.Value = False
    Else
        Option2.Value = True
    End If
End Function

Private Sub Label2_Click()
    SaveSetting App.Title, "Settings", "PortNumber", Combo1.Text
    SaveSetting App.Title, "Settings", "DataSpeed", Combo2.Text
    SaveSetting App.Title, "Settings", "DataBits", Combo3.Text
    SaveSetting App.Title, "Settings", "Parity", Combo4.Text
    SaveSetting App.Title, "Settings", "StopBit", Combo5.Text
    SaveSetting App.Title, "Settings", "FlowControl", Combo6.Text
    SaveSetting App.Title, "Settings", "RTS", Option1.Value
    SaveSetting App.Title, "Settings", "IString", txtFields(1).Text
    
    
    Text1.Text = GetSetting(App.Title, "Settings", "PortNumber", "Parameter Not Set !")
    Text2.Text = GetSetting(App.Title, "Settings", "DataSpeed", "Parameter Not Set !")
    Text3.Text = GetSetting(App.Title, "Settings", "DataBits", "Parameter Not Set !")
    Text4.Text = GetSetting(App.Title, "Settings", "Parity", "Parameter Not Set !")
    Text5.Text = GetSetting(App.Title, "Settings", "StopBit", "Parameter Not Set !")
    Text6.Text = GetSetting(App.Title, "Settings", "FlowControl", "Parameter Not Set !")
    
    Dim str_rts As String
    
    str_rts = GetSetting(App.Title, "Settings", "RTS", False)
    If str_rts = "True" Then
        Text7.Text = "Enable"
    Else
        Text7.Text = "Disable"
    End If

    
    
    
    
    
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbCustom
    Label2.Font.Underline = True
End Sub

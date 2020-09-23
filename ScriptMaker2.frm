VERSION 5.00
Begin VB.Form ScriptMaker 
   Caption         =   "Script Maker"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Transition Setup"
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   3960
         TabIndex        =   21
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtTransitionType 
         DataField       =   "TransitionType"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ScriptMaker2.frx":0000
         Left            =   1320
         List            =   "ScriptMaker2.frx":0061
         TabIndex        =   19
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox cmbDrawMode 
         DataField       =   "DrawModeType"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ScriptMaker2.frx":01EF
         Left            =   1320
         List            =   "ScriptMaker2.frx":0217
         TabIndex        =   18
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox cmbHoldImage 
         DataField       =   "HoldImage"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ScriptMaker2.frx":02B6
         Left            =   1320
         List            =   "ScriptMaker2.frx":02D8
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cmbTransitionSteps 
         DataField       =   "DelaySteps"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "ScriptMaker2.frx":02FB
         Left            =   1320
         List            =   "ScriptMaker2.frx":031D
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ListBox lstLoopSound 
         DataField       =   "LoopSound"
         DataSource      =   "Data2"
         Height          =   255
         ItemData        =   "ScriptMaker2.frx":0340
         Left            =   3720
         List            =   "ScriptMaker2.frx":034A
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.ListBox lstLoopAnimation 
         DataField       =   "LoopAnimation"
         DataSource      =   "Data2"
         Height          =   255
         ItemData        =   "ScriptMaker2.frx":0354
         Left            =   1440
         List            =   "ScriptMaker2.frx":035E
         TabIndex        =   13
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtSoundName 
         DataField       =   "SoundName"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   3840
         Width           =   4695
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "E:\Microsoft Visual Studio\VB98\New001.tsf"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AnimationDetails"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtPicName 
         DataField       =   "PictureName"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label8 
         Caption         =   "Loop Sound"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Loop Animation"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Play Sound"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Transition Steps"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Hold Image for "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Draw Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Transition Type"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Picture Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Select File -> New to Start a new script. Select File -> Edit to edit an existing script"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   6495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuMakeNew 
         Caption         =   "Make New Script"
      End
      Begin VB.Menu mnuEditScript 
         Caption         =   "Edit Script "
      End
   End
End
Attribute VB_Name = "ScriptMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
Variables.SelectedFileName = ""
Dim f As New Selection
f.Label1.Caption = "Picture File Name:"
f.Show (vbModal)
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Data1.Recordset.Delete
If Data1.Recordset.EOF Then
    Exit Sub
Else
    Data1.Recordset.MoveNext
End If
End Sub

Private Sub cmdNew_Click()
'On Error Resume Next
Data1.Recordset.AddNew
'Data2.Recordset.MoveFirst
ScriptMaker.cmdNew.Enabled = False
ScriptMaker.cmdSave.Enabled = True
ScriptMaker.cmdDelete.Enabled = False
End Sub

Private Sub cmdRefresh_Click()
Data1.DatabaseName = Variables.DatabaseName
Data2.DatabaseName = Variables.DatabaseName

'MsgBox Data1.DatabaseName
Data1.RecordSource = "Select * From TransitionDetails"
Data2.RecordSource = "Select * From AnimationDetails"
Data1.Refresh
Data2.Refresh

End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Data1.Recordset.Update
ScriptMaker.cmdNew.Enabled = True
ScriptMaker.cmdSave.Enabled = True
ScriptMaker.cmdDelete.Enabled = True
Data2.Recordset.Edit
End Sub

Private Sub Combo1_LostFocus()
txtTransitionType.Text = Combo1.ListIndex + 1
End Sub

Private Sub Command1_Click()
Variables.SelectedFileName = ""
Dim f As New Selection
f.Label1.Caption = "Sound File Name:"
f.File1.Pattern = "*.mid;*.wav"
f.Show (vbModal)
End Sub



Private Sub Form_Paint()
Gradient Me, 255, 0, 0, False
End Sub

Private Sub mnuEditScript_Click()
Dim dataname As String
Selection.File1.Visible = True
Selection.Image1.Visible = False
Selection.File1.Pattern = "*.tsf"
Selection.Label1.Caption = "Open for Edit"
Selection.Show (vbModal)
dataname = Variables.DatabaseName
'MsgBox dataname
cmdRefresh_Click
Frame1.Visible = True
End Sub

Private Sub mnuMakeNew_Click()
Dim f As New Selection
Load f
f.Image1.Visible = False
f.File1.Visible = False
f.Label1.Caption = "New Database Name:"
f.Show (vbModal)
Frame1.Visible = True
End Sub

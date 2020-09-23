VERSION 5.00
Begin VB.Form ScriptMaker 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cmbTransitionSteps 
      DataField       =   "DelaySteps"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "ScriptMaker.frx":0000
      Left            =   1440
      List            =   "ScriptMaker.frx":0022
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cmbHoldImage 
      DataField       =   "HoldImage"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "ScriptMaker.frx":0045
      Left            =   1440
      List            =   "ScriptMaker.frx":0067
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbDrawMode 
      DataField       =   "DrawModeType"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "ScriptMaker.frx":008A
      Left            =   1440
      List            =   "ScriptMaker.frx":0094
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "ScriptMaker.frx":00B0
      Left            =   1440
      List            =   "ScriptMaker.frx":010E
      TabIndex        =   8
      Top             =   960
      Width           =   3375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "s"
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transition Setup"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtTransitionType 
         DataField       =   "TransitionType"
         DataSource      =   "Data1"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   4800
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
Data1.Recordset.AddNew
ScriptMaker.cmdNew.Enabled = False
ScriptMaker.cmdSave.Enabled = True
ScriptMaker.cmdDelete.Enabled = False
End Sub

Private Sub cmdRefresh_Click()
Data1.DatabaseName = Variables.DatabaseName
'MsgBox Data1.DatabaseName
Data1.RecordSource = "Select * From TransitionDetails"
'Data1.Recordset.Requery
Data1.Refresh

End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Data1.Recordset.Update
ScriptMaker.cmdNew.Enabled = True
ScriptMaker.cmdSave.Enabled = True
ScriptMaker.cmdDelete.Enabled = True
End Sub

Private Sub mnuEditScript_Click()
Dim dataname As String
Selection.File1.Visible = True
Selection.Image1.Visible = False
Selection.File1.Pattern = "*.tsf"
Selection.Label1.Caption = "Open for Edit"
Selection.Show (vbModal)
dataname = Variables.DatabaseName
MsgBox dataname
cmdRefresh_Click
End Sub

Private Sub mnuMakeNew_Click()
Dim f As New Selection
Load f
f.Image1.Visible = False
f.File1.Visible = False
f.Label1.Caption = "New Database Name:"
f.Show (vbModal)
End Sub

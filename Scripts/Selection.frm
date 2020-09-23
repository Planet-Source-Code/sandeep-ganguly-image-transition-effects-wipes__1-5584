VERSION 5.00
Begin VB.Form Selection 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   4935
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   3480
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
 Me.Hide
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmdOK_Click()

On Error Resume Next
Dim db As Database

Dim td As TableDef
Dim fldPicName As Field
Dim fldTranstype As Field
Dim fldDrawMode As Field
Dim fldHoldImage As Field
Dim fldDelayTrans As Field
Dim dbname As String

If Label1.Caption = "New Database Name:" Then
'if label indicate that a new database is to be created
'then create database
    dbname = File1.Path & "\" & txtFileName.Text
    'add extension
    dbname = dbname & ".tsf"
    
    If MsgBox("Create?", vbYesNo + vbQuestion, dbname) = vbYes Then
        'Create Database
        Kill dbname
        Set db = CreateDatabase(dbname, dbLangGeneral, dbVersion30)
        Variables.DatabaseName = dbname 'for future reference from ScriptMaker form
        
        'create table structure
        Set td = db.CreateTableDef("TransitionDetails")
        Set fldPicName = td.CreateField("PictureName", dbText, 255)
        Set fldTranstype = td.CreateField("TransitionType", dbText, 50)
        Set fldDrawMode = td.CreateField("DrawModeType", dbText, 20)
        Set fldHoldImage = td.CreateField("HoldImage", dbInteger)
        Set fldDelayTrans = td.CreateField("DelaySteps", dbInteger)
        
        'append fields to table
        td.Fields.Append fldPicName
        td.Fields.Append fldTranstype
        td.Fields.Append fldDrawMode
        td.Fields.Append fldHoldImage
        td.Fields.Append fldDelayTrans
                        
        db.TableDefs.Append td
        
    End If
ScriptMaker.cmdNew.Enabled = True
ScriptMaker.cmdSave.Enabled = False
ScriptMaker.cmdDelete.Enabled = False
ScriptMaker.Data1.DatabaseName = Variables.DatabaseName
'MsgBox Data1.DatabaseName
ScriptMaker.Data1.RecordSource = "Select * From TransitionDetails"
'Data1.Recordset.Requery
ScriptMaker.Data1.Refresh
Unload Me
End If
    
    If Label1.Caption = "Picture File Name:" Then
        'Variables.SelectedFileName = File1.Path & "\" & txtFileName.Text
        ScriptMaker.txtPicName.Text = txtFileName.Text
        Unload Me
    End If
        
    If Label1.Caption = "Open for Edit" Then
        Variables.DatabaseName = txtFileName.Text
        ScriptMaker.cmdNew.Enabled = True
        ScriptMaker.cmdSave.Enabled = True
        ScriptMaker.cmdDelete.Enabled = True

        Unload Me
    End If
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
If File1.Pattern <> "*.tsf" Then
    strFilename = Dir1.Path & "\" & File1.FileName
    Image1.Picture = LoadPicture(strFilename)
End If

End Sub

Private Sub File1_DblClick()
Dim strFilename As String
strFilename = Dir1.Path & "\" & File1.FileName
txtFileName = strFilename
End Sub

Private Sub Form_Activate()
txtFileName.SetFocus
txtFileName.SelText = "Untiled"
txtFileName.SelStart = 0
txtFileName.SelLength = Len(txtFileName.Text)
End Sub


Private Sub txtFileName_GotFocus()
If Label1.Caption = "PicFile" Then
    txtFileName.Text = ""
    txtFileName.SelText = "Untiled"
    txtFileName.SelStart = 0
    txtFileName.SelLength = Len(txtFileName.Text)
End If
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtFileName_LostFocus()
'If Label1.Caption = "PicFile" Then
'    If Len(txtFileName) > 0 Then
'        txtFileName.Text = txtFileName.Text & ".tsf"
'    End If
'End If
End Sub

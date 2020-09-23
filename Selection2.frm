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
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H00808000&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   6015
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2625
      Left            =   3480
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4935
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
Dim dbname As String

Dim td As TableDef
Dim fldPicName As Field
Dim fldTranstype As Field
Dim fldDrawMode As Field
Dim fldHoldImage As Field
Dim fldDelayTrans As Field

Dim td2 As TableDef
Dim fldAnimationLoop As Field
Dim fldSoundName As Field
Dim fldSoundLoop As Field

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
        
        Set td2 = db.CreateTableDef("AnimationDetails")
        Set fldAnimationLoop = td2.CreateField("LoopAnimation", dbText, 1)
        Set fldSoundName = td2.CreateField("SoundName", dbText, 255)
        Set fldSoundLoop = td2.CreateField("LoopSound", dbText, 1)
        
        td2.Fields.Append fldAnimationLoop
        td2.Fields.Append fldSoundName
        td2.Fields.Append fldSoundLoop
                        
        db.TableDefs.Append td2
    End If
ScriptMaker.cmdNew.Enabled = True
ScriptMaker.cmdSave.Enabled = False
ScriptMaker.cmdDelete.Enabled = False

ScriptMaker.Data1.DatabaseName = Variables.DatabaseName
ScriptMaker.Data2.DatabaseName = Variables.DatabaseName

ScriptMaker.Data1.RecordSource = "Select * From TransitionDetails"
ScriptMaker.Data2.RecordSource = "Select * from AnimationDetails"

ScriptMaker.Data1.Refresh
ScriptMaker.Data2.Refresh

ScriptMaker.Data2.Recordset.AddNew
ScriptMaker.lstLoopAnimation.Text = "N"
ScriptMaker.lstLoopSound.Text = "N"
ScriptMaker.txtSoundName.Text = ""
ScriptMaker.Data2.Recordset.Update
ScriptMaker.Data2.Refresh
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
    
    If Label1.Caption = "Load Script" Then
        ScriptMaker.Data1.DatabaseName = txtFileName.Text
        ScriptMaker.Data2.DatabaseName = txtFileName.Text

        'Refresh method for data1
        ScriptMaker.Data1.RecordSource = "Select * From TransitionDetails"
        ScriptMaker.Data2.RecordSource = "Select * From AnimationDetails"
        ScriptMaker.Data1.Refresh
        ScriptMaker.Data2.Refresh
        ScriptMaker.Data2.Recordset.MoveFirst
        ScriptMaker.Data1.Refresh
        Variables.SelectedSoundName = ScriptMaker.txtSoundName.Text
        ScriptMaker.Hide
    Unload Me
    End If
    
    If Label1.Caption = "Sound File Name:" Then
        Variables.SelectedSoundName = txtFileName.Text
        ScriptMaker.txtSoundName.Text = txtFileName.Text
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
If File1.Pattern = "*.jpg;*.gif;*.bmp" Then
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


Private Sub Form_Paint()
Gradient Me, 0, 127, 0, True
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

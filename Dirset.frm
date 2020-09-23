VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preset Animation: Options"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFadeSteps 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Text            =   "5"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtRandomBlocks 
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Text            =   "9"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtDelaySteps 
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Text            =   "1"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtDelay 
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Text            =   "1"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CheckBox chkQuickClear 
      Caption         =   "Clear previous picture before next transition - no animation here - just a quick clear"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Drive & Directory"
      Height          =   2895
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5535
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CheckBox chkInfinite 
      Caption         =   "Check if you want the transtion to loop infinitely"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   3975
   End
   Begin VB.CheckBox chkClearDest 
      Caption         =   "Check if you want each image to be wiped out (using the previous transition method) befor the next transition"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "In how many frames shold faing be completed?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "When showing random blocks, how many blocks should be shown per row and col (row=col)"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Specify the transiton speed. Higher numbers will lead to rapid transitions"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Specify (in secs) how long the image should be displayed before next transiton begins"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddone_Click()
Me.Hide
End Sub

Private Sub Dir1_Change()
WipesForm.File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
End Sub

Private Sub Form_Paint()
Gradient Form1, 0, 0, 255, True
End Sub


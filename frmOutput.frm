VERSION 5.00
Begin VB.Form frmOutput 
   Caption         =   "Output"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   3495
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim outputpath

If Dir1.Path = "C:\" Then
Form1.txtOutput.Text = Dir1.Path
ElseIf Dir1.Path = "D:\" Then
Form1.txtOutput.Text = Dir1.Path
ElseIf Dir1.Path = "E:\" Then
Form1.txtOutput.Text = Dir1.Path
ElseIf Dir1.Path = "F:\" Then
Form1.txtOutput.Text = Dir1.Path
Else
outputpath = Dir1.Path & "\"
Form1.txtOutput.Text = outputpath
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive

End Sub


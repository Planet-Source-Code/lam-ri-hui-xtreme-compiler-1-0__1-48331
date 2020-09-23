VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Xtreme Compiler 1.0"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd2 
      Left            =   9360
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9360
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Compile Now"
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
      Left            =   8280
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compilation Info"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton Command3 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse..."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtVB 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   7095
      End
      Begin VB.TextBox Project 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   7095
      End
      Begin VB.TextBox txtOutput 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "Visual Basic Executable :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Project :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Output Path :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Click 'Compile Now' to start compilation."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   3675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit

Private Sub Command1_Click()
On Error Resume Next
cd1.Filter = "Visual Basic Project (*.vbp)|*.vbp|All Files (*.*)|*.*"""
cd1.CancelError = True
cd1.Action = 1
Project.Text = cd1.Filename
End Sub

Private Sub Command2_Click()
On Error Resume Next
cd2.Filter = "Visual Basic Executable (VB6.exe)|VB6.exe"
cd2.CancelError = True
cd2.Action = 1
txtVB.Text = cd2.Filename
End Sub

Private Sub Command3_Click()
frmOutput.Show 1
End Sub

Private Sub Command4_Click()

On Error GoTo Compile_Error
If txtOutput.Text = "" Then
MsgBox "You didn't enter a valid output path!", , "Xtreme Compiler"
Exit Sub
ElseIf Project.Text = "" Then
MsgBox "You didn't select a project!", , "Xtreme Compiler"
Exit Sub
ElseIf txtVB.Text = "" Then
MsgBox "You didn't specific the Visual Basic 6 Executable file!", , "Xtreme Compiler"
Command4.Enabled = False
        Form1.Refresh
        DoEvents
        Label6.Caption = "Compiling project... Please wait."
        CompileProject (Project.Text)
Command4.Enabled = True
Label6.Caption = "Compilation completed. Waiting for another compilation to begin."
Exit Sub
Compile_Error:
MsgBox "An error " & Err.Number & " had occured." & Err.Description
Command4.Enabled = True
Exit Sub
End Sub

Public Sub CompileProject(Project As String)
    Dim cmd As String
    
    cmd = """" + txtVB.Text + """ /m """ + Project + """ /outdir """ + txtOutput.Text + """"
    Call Spawn(cmd, True)
End Sub



VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Archive Pro!"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   7560
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Rename files!"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Frame frameStep3 
      Caption         =   "Step Three - (3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   6375
      Begin VB.TextBox txtDefaultFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "FileDefault"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtNumberPad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtStartingNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblFile 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Standard File Name:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "'Padded' Number:"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Number:"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "How do you want to rename files?"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frameStep2 
      Caption         =   "Step Two - (2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   6375
      Begin VB.OptionButton opText 
         Caption         =   "Text Files (*.txt, *.doc)"
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.OptionButton opGraphic 
         Caption         =   "Graphic Files (*.jpg, *.gif, *.bmp)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton opAll 
         Caption         =   "All Files in Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "What type of Files do you want to rename?"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame frameStep1 
      Caption         =   "Step One - (1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.FileListBox FileList 
         Appearance      =   0  'Flat
         Archive         =   0   'False
         Enabled         =   0   'False
         Height          =   2370
         Left            =   3600
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.DriveListBox DriveList 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3375
      End
      Begin VB.DirListBox FolderList 
         Appearance      =   0  'Flat
         Height          =   2340
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblDirectory 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   6135
      End
      Begin VB.Label lblArray 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose the Directory to Archive:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
Dim Files As String

If FileList.ListCount = 0 Then Exit Sub

If opAll = True Then Files = "All"
If opGraphic = True Then Files = "Graphic"
If opText = True Then Files = "Text"

Me.Enabled = False

Bar.Max = FileList.ListCount

RenameFiles FolderList.Path, txtDefaultFile.Text, txtNumberPad.Text, _
    txtStartingNumber.Text, Files

Me.Enabled = True
FileList.Refresh
MsgBox "Successfully renamed all files. You're Welcome!", vbExclamation, "Success!"
End Sub

Private Sub DriveList_Change()
    On Error GoTo Err_Init
FolderList.Path = DriveList.Drive
    Exit Sub
    
Err_Init:
HandleError "frmMain", "DriveList_Change()", Err.Number, Err.Description
End Sub

Private Sub FolderList_Change()
    On Error GoTo Err_Init
FileList.Path = FolderList.Path
lblDirectory.Caption = FolderList.Path
    Exit Sub

Err_Init:
HandleError "frmMain", "DriveList_Change()", Err.Number, Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo Err_Init
FolderList.Path = DriveList.Drive
lblDirectory.Caption = FolderList.Path

FormatExample
    Exit Sub
    
Err_Init:
HandleError "frmMain", "DriveList_Change()", Err.Number, Err.Description
End Sub

Private Sub txtDefaultFile_Validate(Cancel As Boolean)
If txtDefaultFile.Text = "" Then txtDefaultFile.Text = "NotBlank"
FormatExample
End Sub

Private Sub txtNumberPad_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Dim Msg As String

Numbers = KeyAscii

If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
    Msg = MsgBox("You can only use numeric values here.", vbCritical, "Whoops!")
    KeyAscii = 0
ElseIf ((Numbers > 52) And (Numbers <> 8)) Then
    Msg = MsgBox("Please only pad up to 4 zero's, Thank You.", vbCritical, "Whoops!")
    KeyAscii = 0
End If

End Sub

Private Sub txtNumberPad_Validate(Cancel As Boolean)
If txtNumberPad.Text = "" Then txtNumberPad.Text = "0"
FormatExample
End Sub

Private Sub txtStartingNumber_KeyPress(KeyAscii As Integer)
Dim Numbers As Integer
Dim Msg As String

Numbers = KeyAscii

If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
    Msg = MsgBox("You can only use numeric values here.", vbCritical, "Whoops!")
    KeyAscii = 0
End If
End Sub

Private Sub txtStartingNumber_Validate(Cancel As Boolean)
If txtStartingNumber.Text = "" Then txtStartingNumber.Text = "0"
FormatExample
End Sub

Public Sub FormatExample()
lblFile.Caption = txtDefaultFile.Text
Select Case txtNumberPad.Text
    Case "0"
        lblFile.Caption = lblFile.Caption
    Case "1"
        lblFile.Caption = lblFile.Caption & "0"
    Case "2"
        lblFile.Caption = lblFile.Caption & "00"
    Case "3"
        lblFile.Caption = lblFile.Caption & "000"
    Case "4"
        lblFile.Caption = lblFile.Caption & "0000"
End Select
lblFile.Caption = lblFile.Caption & txtStartingNumber.Text
lblFile.Caption = lblFile.Caption & ".jpg"
End Sub

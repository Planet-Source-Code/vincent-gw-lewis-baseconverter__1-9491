VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "BaseConverter"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame fraTo 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2400
      TabIndex        =   12
      Top             =   840
      Width           =   2055
      Begin VB.OptionButton optToBase 
         Caption         =   "Base 2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optToBase 
         Caption         =   "Base 16"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optToBase 
         Caption         =   "Base 10"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame fraFrom 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   2055
      Begin VB.OptionButton optFromBase 
         Caption         =   "Base 10"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optFromBase 
         Caption         =   "Base 16"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optFromBase 
         Caption         =   "Base 2"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox txtGiven 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblAnswer 
      AutoSize        =   -1  'True
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label lblGiven 
      AutoSize        =   -1  'True
      Caption         =   "Given:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fromBase As Integer
Public toBase As Integer

Private Function StrInBag(str As String, bag As String) As Boolean
    Dim i As Integer
    StrInBag = True
    For i = 1 To Len(str)
        If InStr(1, bag, Mid(str, i, 1)) = 0 Then
            StrInBag = False
            Exit Function
        End If
    Next i
End Function

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdConvert_Click()
    Select Case fromBase
    Case 2
        If Not StrInBag(txtGiven.Text, "01") Then
            MsgBox "Base 2 number can only contain '0' and '1'."
            Exit Sub
        End If
    Case 10
        If Not StrInBag(txtGiven.Text, "0123456789") Then
            MsgBox "Base 10 number can only contain '0' to '9'."
            Exit Sub
        End If
    Case 16
        If Not StrInBag(txtGiven.Text, "0123456789ABCDEFabcdef") Then
            MsgBox "Base 16 number can only contain '0' to '9' and 'A' to 'F' or 'a' to 'f'."
            Exit Sub
        End If
    End Select
    txtAnswer.Text = ConvertNum(txtGiven.Text, fromBase, toBase)
    txtAnswer.SetFocus
End Sub

Private Sub Form_Load()
    fromBase = 10
    toBase = 16
End Sub

Private Sub optFromBase_Click(Index As Integer)
    fromBase = Index
End Sub

Private Sub optToBase_Click(Index As Integer)
    toBase = Index
End Sub

Private Sub txtAnswer_GotFocus()
    txtAnswer.SelStart = 0
    txtAnswer.SelLength = Len(txtAnswer.Text)
End Sub

Private Sub txtGiven_GotFocus()
    txtGiven.SelStart = 0
    txtGiven.SelLength = Len(txtGiven.Text)
End Sub

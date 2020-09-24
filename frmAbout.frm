VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About BG Change"
   ClientHeight    =   3570
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5535
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2464.077
   ScaleMode       =   0  'User
   ScaleWidth      =   5197.651
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "URL:"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "E-mail:"
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtICQ 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "ICQ:"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblDesc 
      Caption         =   "This application converts numbers from one base to another. Namely binary, decimal and hexadecimal numbers."
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Caption = "About " & App.Title
  lblTitle.Caption = App.Title
  lblAuthor.Caption = "Coded by Twister of Twisted Media"
  txtEmail.Text = "E-mail: vincent_gw_lewis@hotmail.com"
  txtICQ.Text = "ICQ: 12674360"
  txtURL.Text = "URL: http://www.twistedmedia.f2s.com"
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblDesc.Caption = "This application converts numbers from one base to another. Namely binary, decimal and hexadecimal numbers."
End Sub

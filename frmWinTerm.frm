VERSION 5.00
Begin VB.Form frmWinTerm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5505
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWinTerm.frx":0000
   ScaleHeight     =   2910
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton btOK 
      BackColor       =   &H0000FFFF&
      Caption         =   "I will..."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2212
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   2250
      Width           =   1080
   End
   Begin VB.Line ln 
      BorderColor     =   &H80000002&
      BorderWidth     =   3
      X1              =   0
      X2              =   5520
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pool"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   480
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   -15
      Width           =   885
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Please stop playing before terminating Windows."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1170
      Width           =   5280
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pool"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   15
      Width           =   885
   End
End
Attribute VB_Name = "frmWinTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btOK_Click()

    Unload Me
    frmPool.Enabled = True
    frmPool.SetFocus

End Sub

Private Sub Form_Paint()

    btOK.SetFocus

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 1  Code: 17  Total: 18 Lines
':) CommentOnly: 2 (11,1%)  Commented: 0 (0%)  Empty: 7 (38,9%)  Max Logic Depth: 1

VERSION 5.00
Begin VB.Form frmStop 
   BackColor       =   &H00004000&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1500
      Width           =   225
   End
   Begin VB.OptionButton optYesNo 
      BackColor       =   &H00008000&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   465
      Index           =   1
      Left            =   2340
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   825
      Width           =   1005
   End
   Begin VB.OptionButton optYesNo 
      BackColor       =   &H00008000&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   465
      Index           =   0
      Left            =   495
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   825
      Width           =   1005
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00008000&
      Height          =   1485
      Left            =   0
      Shape           =   4  'Gerundetes Rechteck
      Top             =   0
      Width           =   3825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to stop playing?"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   255
      TabIndex        =   3
      Top             =   195
      Width           =   3330
   End
End
Attribute VB_Name = "frmStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

    Beep

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case Chr$(KeyAscii)
      Case "Y", "y"
        optYesNo(0) = True
        Hide
      Case "N", "n"
        optYesNo(1) = True
        Hide
      Case Else
        Beep
    End Select

End Sub

Private Sub lbl_Click()

    Beep

End Sub

Private Sub optYesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Hide

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 1  Code: 36  Total: 37 Lines
':) CommentOnly: 2 (5,4%)  Commented: 0 (0%)  Empty: 13 (35,1%)  Max Logic Depth: 2

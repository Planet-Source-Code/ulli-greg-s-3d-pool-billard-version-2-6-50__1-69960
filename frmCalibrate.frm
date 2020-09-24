VERSION 5.00
Begin VB.Form frmCalibrate 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5580
   ControlBox      =   0   'False
   ForeColor       =   &H00004000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Image Image2 
      Height          =   1200
      Left            =   135
      Picture         =   "frmCalibrate.frx":0000
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   195
      Top             =   225
      Width           =   765
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "To calibrate your Joystick: Move to all four corners;   then move to center and  press Fire Button #1."
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   990
      Left            =   1485
      TabIndex        =   0
      Top             =   180
      Width           =   3975
   End
End
Attribute VB_Name = "frmCalibrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form has no code

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 14:15)  Decl: 6  Code: 0  Total: 6 Lines
':) CommentOnly: 3 (50%)  Commented: 0 (0%)  Empty: 2 (33,3%)  Max Logic Depth: 0

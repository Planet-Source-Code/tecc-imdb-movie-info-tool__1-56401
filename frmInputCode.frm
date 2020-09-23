VERSION 5.00
Begin VB.Form frmInputCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input IMDB Movie ID"
   ClientHeight    =   1305
   ClientLeft      =   5295
   ClientTop       =   5460
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInputCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Get Movie Info"
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   660
      Width           =   1635
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Text            =   "tt0133093"
      Top             =   300
      Width           =   4035
   End
End
Attribute VB_Name = "frmInputCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim XXX As String
AddQueueEntry txtCode.Text
Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmInCorrect 
   Caption         =   "Warning"
   ClientHeight    =   1950
   ClientLeft      =   6135
   ClientTop       =   4695
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3720
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "One or more required fields left empty, details will not be stored"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   360
      Picture         =   "frmInCorrect.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmInCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to inform the user that the details that he or she
'has just entered, are not complete

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'None

'Objects Used -
'None

Private Sub cmdOk_Click()
'Hides this screen
frmInCorrect.Visible = False
End Sub

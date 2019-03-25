VERSION 5.00
Begin VB.Form frmNotFound 
   Caption         =   "Warning"
   ClientHeight    =   2070
   ClientLeft      =   6060
   ClientTop       =   3615
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1170
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   240
      Picture         =   "frmNotFound.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Supplier not found, make sure spelling is correct and that the supplier exists"
      Height          =   450
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2685
   End
End
Attribute VB_Name = "frmNotFound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to inform the user that the supplier that he or she was
'looking for, was not found

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'None

'Objects Used -
'None

Private Sub cmdOk_Click()
'hides this screen
frmNotFound.Visible = False
End Sub

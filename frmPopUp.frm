VERSION 5.00
Begin VB.Form frmPopUp 
   Caption         =   "Data Save Successful"
   ClientHeight    =   1830
   ClientLeft      =   6810
   ClientTop       =   4185
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3120
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   600
      Picture         =   "frmPopUp.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "The data has been saved"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1845
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to inform the user that the details that he or she
'has just entered, has been saved.

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'None

'Objects Used -
'None

Private Sub cmdOk_Click()
'hides the screen and empties the textboxes in the "Add Supplier" screen
frmPopUp.Visible = False
frmAddSupplier.txtSuppName.Text = ""
frmAddSupplier.txtAdd1.Text = ""
frmAddSupplier.txtAdd2.Text = ""
frmAddSupplier.txtAdd3.Text = ""
frmAddSupplier.txtTelNum.Text = ""
frmAddSupplier.txtMobNum.Text = ""
frmAddSupplier.txtEmail.Text = ""
frmAddSupplier.txtSuppName.SetFocus
End Sub


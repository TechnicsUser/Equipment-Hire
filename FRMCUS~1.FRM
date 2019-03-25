VERSION 5.00
Begin VB.Form frmCustomerFileProcessing 
   Caption         =   "Customer Maintenance Menu"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdViewAmend 
      Caption         =   "&View/Amend Customer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete Customer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Customer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdFileMaintenance 
      Cancel          =   -1  'True
      Caption         =   "&File Maintenance Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9060
      TabIndex        =   2
      Top             =   9840
      Width           =   2775
   End
   Begin VB.CommandButton cmdMM 
      Caption         =   "&Main Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   1
      Top             =   9840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Customer Maintenance Menu"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   3210
      TabIndex        =   0
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "frmCustomerFileProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub cmdAdd_Click()
frmAddCustomer.Show
Unload Me
End Sub

Private Sub cmdDel_Click()
frmDeleteCustomer.Show             'Display the delete a supplier screen
Unload Me
End Sub

Private Sub cmdFileMaintenance_Click()
frmFileMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdMM_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdViewAmend_Click()
frmAmendViewCustomer.Show
Unload Me
End Sub

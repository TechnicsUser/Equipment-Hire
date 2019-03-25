VERSION 5.00
Begin VB.Form frmFileMaintenanceMenu 
   Caption         =   "File Maintenance Menu"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
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
      Left            =   13320
      TabIndex        =   6
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSupplierFP 
      Caption         =   "Supplier File Processing"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   5160
      Width           =   3615
   End
   Begin VB.CommandButton cmdEquipmentFP 
      Caption         =   "Equipment File Processing "
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
      Left            =   5880
      TabIndex        =   4
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton cmdCustomerFP 
      Caption         =   "Customer File Processing"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdStaffFP 
      Caption         =   "Staff File Processing"
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
      Left            =   5880
      TabIndex        =   2
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdMainmenu 
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
      Left            =   11520
      TabIndex        =   1
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "File Maintenance Menu"
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
      Left            =   4155
      TabIndex        =   0
      Top             =   240
      Width           =   7245
   End
End
Attribute VB_Name = "frmFileMaintenanceMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Derek Stafford com2026
'The file maintenance menu is used to acccess/open the :
'1 Customer Maintenance Menu screen
'2 Equipment Maintenance Menu screen
'3 Staff Maintenance Menu screen
'4 Supplier Maintenance Menu screen

Private Sub cmdCustomerFP_Click()
frmCustomerFileProcessing.Show
Unload Me
End Sub

Private Sub cmdEquipmentFP_Click()
frmEquipmentFileProcessing.Show
Unload Me
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdStaffFP_Click()
frmStaffMenu.Show
Unload Me
End Sub

Private Sub cmdSupplierFP_Click()
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

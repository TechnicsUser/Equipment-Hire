VERSION 5.00
Begin VB.Form frmStaffMenu 
   Caption         =   "Staff Maintenance Menu"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "File Maintenance Menu"
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
      Left            =   7920
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "To File Maintenence Menu"
      Top             =   9000
      Width           =   2655
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "Main Menu"
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
      Left            =   10920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddStaff 
      Caption         =   "Add Staff"
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
      TabIndex        =   2
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelStaff 
      Caption         =   "Delete Staff"
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
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdViewStaff 
      Caption         =   "Amend/View Staff"
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
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Staff Maintenance Menu"
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
      Left            =   3540
      TabIndex        =   3
      Top             =   600
      Width           =   7575
   End
End
Attribute VB_Name = "frmStaffMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'allows the user access to staff processing screens such as, Delete a Staff Member and
'Amend/View Staff.

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'None

'Objects Used -
'None

Option Explicit
Private Sub cmdBack_Click()
'Unloads this screen and shows the File Maintenance Menu
frmFileMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdDelStaff_Click()
'Unloads this screen and shows the Delete Staff screen
frmDelStaff.Show
Unload Me
End Sub

Private Sub cmdMainMenu_Click()
'Unloads this screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdViewStaff_Click()
'Unloads this screen and shows the Amend/View Staff screen
frmAmendViewStaff.Show
Unload Me
End Sub

VERSION 5.00
Begin VB.Form frmReportsMenu 
   Caption         =   "Reports Menu Screen"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form2"
   ScaleHeight     =   8550
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMainMenu 
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
      Height          =   615
      Left            =   11760
      TabIndex        =   8
      Top             =   9840
      Width           =   2055
   End
   Begin VB.CommandButton cmdPayments 
      Caption         =   "&Payments Report"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   4520
      Width           =   3135
   End
   Begin VB.CommandButton cmdOrders 
      Caption         =   "&Orders Report"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   5680
      Width           =   3135
   End
   Begin VB.CommandButton cmdCustDetails 
      Caption         =   "Customer Details &Report"
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
      Left            =   6000
      TabIndex        =   5
      Top             =   6840
      Width           =   3135
   End
   Begin VB.CommandButton cmdEqptHired 
      Caption         =   "&Equipment Hired Report"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13320
      TabIndex        =   2
      Top             =   12225
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   12225
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   """Sub-Menu"""
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
      Left            =   11160
      TabIndex        =   0
      Top             =   12225
      Width           =   1575
   End
   Begin VB.Label lblReportsMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Reports Menu"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   4425
   End
End
Attribute VB_Name = "frmReportsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to select and open any one of 4
'               Report screens
'Student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02



Private Sub cmdCustDetails_Click()

frmCustReport.Show
Unload Me

End Sub

Private Sub cmdEqptHired_Click()

frmEquipmentHiredReport.Show
Unload Me

End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdOrders_Click()

frmOrdersReport.Show
Unload Me

End Sub

Private Sub cmdPayments_Click()

frmPaymentsReport.Show
Unload Me

End Sub

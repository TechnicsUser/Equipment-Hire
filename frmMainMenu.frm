VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "Main Menu"
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
   Begin VB.CommandButton cmdClock 
      Caption         =   "&Staff Clock - in"
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
      TabIndex        =   8
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdReports 
      Caption         =   "Repor&ts"
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
      TabIndex        =   7
      Top             =   4717
      Width           =   2175
   End
   Begin VB.CommandButton cmdMaintenance 
      Caption         =   "&File Maintenance"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   2955
      Width           =   2175
   End
   Begin VB.CommandButton cmdReOrder 
      Caption         =   "R&eceive an Order"
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
      TabIndex        =   6
      Top             =   2955
      Width           =   2175
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "&Place an &Order"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "&Customer Status"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   4717
      Width           =   2175
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "&Payment Details"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdHire 
      Caption         =   "Hire E&quipment"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   2955
      Width           =   2175
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "&Return Equipment"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   4717
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      TabIndex        =   9
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Main Menu"
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
      Left            =   5985
      TabIndex        =   10
      Top             =   240
      Width           =   3585
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the screen from which all processing can be carried out

Option Explicit

Private Sub cmdClock_Click()
frmClockIn.Show
Unload Me
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdHire_Click()
frmRental.Show
Unload Me
End Sub

Private Sub cmdMaintenance_Click()
frmFileMaintenanceMenu.Show
Unload Me

End Sub

Private Sub cmdPay_Click()
frmMakeaPayment.Show
Unload Me
End Sub

Private Sub cmdOrder_Click()
frmPlaceanOrder.Show
Unload Me
End Sub

Private Sub cmdReports_Click()
frmReportsMenu.Show
Unload Me
End Sub

Private Sub cmdReturn_Click()
frmReturnEquipment.Show
Unload Me
End Sub





VERSION 5.00
Begin VB.Form frmRental 
   Caption         =   "Rental Form"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   23551.5
   ScaleMode       =   0  'User
   ScaleWidth      =   25079.68
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "&Main Menu"
      CausesValidation=   0   'False
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
      Left            =   13080
      TabIndex        =   36
      Top             =   10200
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      CausesValidation=   0   'False
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
      Left            =   13080
      TabIndex        =   35
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Data datEmp 
      Caption         =   "Emp"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmdNewCust 
      Caption         =   "Ne&xt Customer"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   33
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdNumAvail 
      Caption         =   "Repeat same T&ype"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   29
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "&Pay Now"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   34
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton cmdRepeatAny 
      Caption         =   "Repea&t any Type"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   30
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   31
      Top             =   6600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   13080
      TabIndex        =   32
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Frame fmeEmpID 
      Height          =   615
      Left            =   600
      TabIndex        =   78
      Top             =   240
      Width           =   4335
      Begin VB.TextBox txtEmpID 
         Height          =   405
         Left            =   3000
         TabIndex        =   0
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter your Employee ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   2445
      End
   End
   Begin VB.Data datRent 
      Caption         =   "Rent"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Frame frmRent 
      Caption         =   "Hire out Equiptment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5415
      Left            =   6240
      TabIndex        =   71
      Top             =   4920
      Width           =   6735
      Begin VB.TextBox txtTotal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   27
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtNewBalance 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   28
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txtValueItem 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   2
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   26
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtTimeDue 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtDayDue 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   24
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox cboPeriod 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":0000
         Left            =   4200
         List            =   "frmHireEquipment.frx":000D
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton cmdDateDue 
         Caption         =   "Calc&ulate Date and time due back"
         Enabled         =   0   'False
         Height          =   855
         Left            =   1560
         TabIndex        =   38
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtTimeOut 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   4
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cboID 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":002B
         Left            =   3360
         List            =   "frmHireEquipment.frx":002D
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtDateReturn 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   25
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtRentPeriod 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtDateOut 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   3
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total Cost of Rentals"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1320
         TabIndex        =   82
         Top             =   4440
         Width           =   2145
      End
      Begin VB.Label lblNewBalance 
         AutoSize        =   -1  'True
         Caption         =   "New Account Balance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   77
         Top             =   4920
         Width           =   2190
      End
      Begin VB.Label lblValueItem 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Rental"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   76
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblTimeOut 
         AutoSize        =   -1  'True
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   75
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblDateOut 
         AutoSize        =   -1  'True
         Caption         =   "Date Out"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   74
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label lblRentPeriod 
         AutoSize        =   -1  'True
         Caption         =   "Enter Rental Period"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   73
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblEquipID 
         AutoSize        =   -1  'True
         Caption         =   "Select Equipment ID number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   72
         Top             =   480
         Width           =   2880
      End
   End
   Begin VB.Data DatSQL4 
      Caption         =   "SQL4"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datSQL3 
      Caption         =   "SQL3"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data datSQL2 
      Caption         =   "SQL2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10755
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data datSQL1 
      Caption         =   "SQL1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datType 
      Caption         =   "Type"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datEquipment 
      Caption         =   "Equipment"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10755
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame frmNameNumOption 
      Caption         =   "Customer Name/ID Number Option"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9135
      Left            =   360
      TabIndex        =   42
      Top             =   1440
      Width           =   5775
      Begin VB.Frame fmeMove 
         Height          =   3495
         Left            =   3960
         TabIndex        =   83
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
         Begin VB.TextBox txtMatches 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   480
            TabIndex        =   84
            Top             =   2280
            Width           =   615
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Display Previous"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Display Next"
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblNumMatchs 
            Caption         =   "Number of Customers matching the search critera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   240
            TabIndex        =   85
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdAddCust 
         Caption         =   "Add ne&w Customer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   8040
         Width           =   1575
      End
      Begin VB.TextBox txtCustID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   80
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtMobileNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   53
         Top             =   5700
         Width           =   1935
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   52
         Top             =   6300
         Width           =   1935
      End
      Begin VB.TextBox txtCreditLimit 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   51
         Top             =   6900
         Width           =   1935
      End
      Begin VB.TextBox txtAddress1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   50
         Top             =   3300
         Width           =   1935
      End
      Begin VB.TextBox txtAddress2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   49
         Top             =   3900
         Width           =   1935
      End
      Begin VB.TextBox txtAddress3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   48
         Top             =   4500
         Width           =   1935
      End
      Begin VB.TextBox txtPhoneNo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   47
         Top             =   5100
         Width           =   1935
      End
      Begin VB.TextBox txtNameFound 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   46
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   45
         Top             =   7560
         Width           =   1935
      End
      Begin VB.TextBox txtBalance 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   44
         Top             =   8100
         Width           =   1935
      End
      Begin VB.ComboBox cboNumber 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":002F
         Left            =   3240
         List            =   "frmHireEquipment.frx":0031
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSeekCustomer 
         Caption         =   "Disp&lay Customer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":0033
         Left            =   2640
         List            =   "frmHireEquipment.frx":0035
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton optNumber 
         Caption         =   "Search using Customers Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.OptionButton optName 
         Caption         =   "Search useing Customers Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.Label lblCustID 
         AutoSize        =   -1  'True
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   480
         TabIndex        =   81
         Top             =   2325
         Width           =   1305
      End
      Begin VB.Label lblMobileNo 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   63
         Top             =   5700
         Width           =   1590
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "e-mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   62
         Top             =   6300
         Width           =   615
      End
      Begin VB.Label lblCreditLimit 
         AutoSize        =   -1  'True
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   600
         TabIndex        =   61
         Top             =   6900
         Width           =   1215
      End
      Begin VB.Label lblAddress1 
         AutoSize        =   -1  'True
         Caption         =   "Address1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         TabIndex        =   60
         Top             =   3345
         Width           =   945
      End
      Begin VB.Label lblAddress2 
         AutoSize        =   -1  'True
         Caption         =   "Address2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         TabIndex        =   59
         Top             =   3900
         Width           =   945
      End
      Begin VB.Label lblAddress3 
         AutoSize        =   -1  'True
         Caption         =   "Address3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         TabIndex        =   58
         Top             =   4500
         Width           =   945
      End
      Begin VB.Label lblPhoneNo 
         AutoSize        =   -1  'True
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Top             =   5100
         Width           =   1485
      End
      Begin VB.Label lblNameFound 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   56
         Top             =   2805
         Width           =   600
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   55
         Top             =   7605
         Width           =   630
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         TabIndex        =   54
         Top             =   8145
         Width           =   810
      End
      Begin VB.Label lblNameNumber 
         AutoSize        =   -1  'True
         Caption         =   "Enter Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   2250
      End
   End
   Begin VB.Data datCustomer 
      Caption         =   "Customer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10755
      Visible         =   0   'False
      Width           =   1695
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
      TabIndex        =   40
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
      TabIndex        =   39
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
      TabIndex        =   37
      Top             =   12225
      Width           =   1575
   End
   Begin VB.Frame fmeStatus 
      Caption         =   "Check Status of Required Equipment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   6240
      TabIndex        =   64
      Top             =   1320
      Width           =   9015
      Begin VB.CommandButton cmdNewItem 
         Caption         =   "C&ount Availability"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4800
         TabIndex        =   17
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cboTypeID 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":0037
         Left            =   2760
         List            =   "frmHireEquipment.frx":0039
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboEquipModel 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":003B
         Left            =   7320
         List            =   "frmHireEquipment.frx":0042
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cboEquipDetails 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":004B
         Left            =   7320
         List            =   "frmHireEquipment.frx":0052
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtNoInStock 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   16
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox cboEquipMake 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":005B
         Left            =   7320
         List            =   "frmHireEquipment.frx":0062
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cboEquipDescription 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmHireEquipment.frx":006B
         Left            =   7320
         List            =   "frmHireEquipment.frx":006D
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optEquipName 
         Caption         =   "Search for equipment using equipment details"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   4095
      End
      Begin VB.OptionButton optEqptNum 
         Caption         =   "Search for equipment useing ID number"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.Label lblEquipDetails 
         AutoSize        =   -1  'True
         Caption         =   "Enter Equipment Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4440
         TabIndex        =   70
         Top             =   1440
         Width           =   2460
      End
      Begin VB.Label lblNoInStock 
         AutoSize        =   -1  'True
         Caption         =   "Number available to hire"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1320
         TabIndex        =   69
         Top             =   2880
         Width           =   2475
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Enter Equipment Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         TabIndex        =   68
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblMake 
         AutoSize        =   -1  'True
         Caption         =   "Enter Equipment Make"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4800
         TabIndex        =   67
         Top             =   1920
         Width           =   2355
      End
      Begin VB.Label lblModel 
         AutoSize        =   -1  'True
         Caption         =   "Enter Equipment Model"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4800
         TabIndex        =   66
         Top             =   2400
         Width           =   2400
      End
      Begin VB.Label lblEquipTypeID 
         AutoSize        =   -1  'True
         Caption         =   "Enter Equipment Type ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   2580
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Hire Equipment"
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
      Left            =   5265
      TabIndex        =   41
      Top             =   240
      Width           =   5025
   End
End
Attribute VB_Name = "frmRental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to select a search for a Custemers details and
'               then update his file with the details of a rental

'student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02

'Variables used
'strMake        String
'strModel       String
'strDescription String
'strDetails     String
'strAvailabilty String
'sqlDescription String
'strDate        String
'sqlTime          String
'sqlStatus      String
'sqlEqptID      String
'sqlCount       String
'intID          Integer
'intResponse    Integer
'intResponse2   Integer
'intNumber      Integer
'intCount       Integer
'intIndex       Integer
'intEqID        Integer
'varCust        Variant
'varID          Variant
'varEmp         Variant
'varDate        Variant
'varPeriod      Variant
'varIn          Variant
'varCos         Variant
'dteDateOut     Date
'dteDueBack     Date
'dblTime        Double
'dblHour        Double
'curRate        Currancey

    
    

Private Sub cboEquipDescription_KeyPress(KeyAscii As Integer)
If KeyAscii Then
    If MsgBox("To ensure accurate selection from the equiptment in stock please select from the drop down list", vbOKOnly, "Accurate Input") = vbOK Then
         cboEquipDescription.Text = ""
    End If
    cboEquipDescription.Text = ""
End If

    

End Sub


'This section is used to fill the Equipment Details drop down
'List with All the unique details of equiptment where the
'Description match those already chosen

Private Sub cboEquipDescription_Click()
Dim sqlDescription, strDescription As String

cboEquipDetails.Clear
cboEquipDetails.AddItem ("All")


strDescription = cboEquipDescription.Text
sqlDescription = "SELECT DISTINCT Details FROM [Equipment Type] WHERE Description = '" & strDescription & "'"

datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = sqlDescription
datSQL1.Refresh

While Not datSQL1.Recordset.EOF
    cboEquipDetails.AddItem datSQL1.Recordset("Details")
    datSQL1.Recordset.MoveNext
Wend

End Sub

'This section is used to fill the Equipment Make drop down
'List with any make of equiptment where the Description
'and details match those already chosen

Private Sub cboEquipDetails_Click()

Dim sqlDetails, sqlMake, strDetails, sqlDescriptionAll As String

cboEquipMake.Clear
cboEquipMake.AddItem ("All")

strDescription = cboEquipDescription.Text
strDetails = cboEquipDetails.Text

If cboEquipDetails.Text = "All" Then
    sqlDescription = "SELECT DISTINCT Make FROM [Equipment Type] WHERE  Description = '" & strDescription & "'"
Else
    sqlDescription = "SELECT DISTINCT Make FROM [Equipment Type] WHERE Details = '" & strDetails & "' AND Description = '" & strDescription & "'"
End If

    datSQL1.DatabaseName = strThePath
    datSQL1.RecordSource = sqlDescription
    datSQL1.Refresh

    While Not datSQL1.Recordset.EOF
        cboEquipMake.AddItem datSQL1.Recordset("Make")
        datSQL1.Recordset.MoveNext
    Wend

End Sub

'Using the infromation stored in the text part of the
'Equipment description,details,make and model comboboxes
'the unique Type ID can be queried from the equipment
'table this is then stored in the equipment number combo box
Private Sub cboEquipModel_Click()
Dim sqlDescription, strDescription, strModel, strMake, strDetails As String

cmdNewItem.Enabled = True

strDetails = cboEquipDetails.Text
strModel = cboEquipModel.Text
strMake = cboEquipMake.Text
strDescription = cboEquipDescription.Text

If strMake = "All" And strDetails = "All" Then
   sqlDescription = "SELECT Type_ID FROM [Equipment Type] WHERE Model = '" & strModel & "' AND Description = '" & strDescription & "'"
    
ElseIf strMake = "All" Then
    sqlDescription = "SELECT Type_ID FROM [Equipment Type] WHERE Model = '" & strModel & "' AND Details = '" & strDetails & "' AND Description = '" & strDescription & "'"
    
ElseIf strDetails = "All" Then
    sqlDescription = "SELECT Type_ID FROM [Equipment Type] WHERE Model = '" & strModel & "' AND Make = '" & strMake & "' And Description = '" & strDescription & "'"

Else
   sqlDescription = "SELECT Type_ID FROM [Equipment Type] WHERE Model = '" & strModel & "' AND Make = '" & strMake & "' And Description = '" & strDescription & "' AND Details = '" & strDetails & "'"
End If
    datSQL1.DatabaseName = strThePath
    datSQL1.RecordSource = sqlDescription
    datSQL1.Refresh
    
    While Not datSQL1.Recordset.EOF
        cboTypeID.Text = datSQL1.Recordset("Type_ID")
        datSQL1.Recordset.MoveNext
    Wend
End Sub


'This section is used to fill the Equipment Make drop down
'List with any Model of equiptment where the Description
'details and make match those already chosen

Private Sub cboEquipMake_Click()
Dim strDescription, sqlDescription, strMake, strDetails As String

cboEquipModel.Clear
strDetails = cboEquipDetails.Text
strDescription = cboEquipDescription.Text
strMake = cboEquipMake.Text


If cboEquipMake.Text = "All" And cboEquipDetails = "All" Then
    sqlDescription = "SELECT Model FROM [Equipment Type] WHERE Description = '" & strDescription & "'"
    
ElseIf cboEquipMake.Text = "All" Then
    sqlDescription = "SELECT Model FROM [Equipment Type] WHERE Description = '" & strDescription & "' AND Details = '" & strDetails & "'"

ElseIf cboEquipDetails.Text = "All" Then
    sqlDescription = "SELECT Model FROM [Equipment Type] WHERE Description = '" & strDescription & "' AND Make = '" & strMake & "'"

Else
    sqlDescription = "SELECT Model FROM [Equipment Type] WHERE Make = '" & strMake & "' AND Description = '" & strDescription & "' AND Details = '" & strDetails & "'"

End If

datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = sqlDescription
datSQL1.Refresh

While Not datSQL1.Recordset.EOF
    cboEquipModel.AddItem datSQL1.Recordset("Model")
    datSQL1.Recordset.MoveNext
Wend
End Sub



'iniated when a new Equipment ID number is chosen automatically sets the date and time
'to now as a logical default it also enables the date and time text
'text input boxes to allow the user to change them if necessary
Private Sub cboID_Click()
Dim strDate, strTime As String
txtDateOut = FormatDateTime(Now, vbShortDate)
txtTimeOut = FormatDateTime(Now, vbShortTime)
txtDateOut.Enabled = True
txtTimeOut.Enabled = True
txtRentPeriod.Enabled = True
cboPeriod.Enabled = True
txtRentPeriod.SetFocus

End Sub

Private Sub cboName_Change()
cboName_Click
fmeMove.Visible = False
End Sub

Private Sub cboName_Click()
fmeMove.Visible = False
cmdSeekCustomer.Enabled = True
End Sub

'Uses the validate procedure to check if the name of the supplier has been input
'before the user can proceed
Private Sub cboName_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(cboName.Text) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please Enter Data")
    If intResponse = 1 Then
        cboName.SetFocus
    End If
    
End If
End Sub

Private Sub cboNumber_Change()
cboNumber_Click
End Sub

Private Sub cboNumber_Click()
fmeMove.Visible = False
cmdSeekCustomer.Enabled = True

End Sub
'Uses the validate procedure to check if the number of the supplier has been input
'before the user can proceed
Private Sub cboNumber_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(cboNumber.Text) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please Enter Data")
    If intResponse = 1 Then
        cboNumber.SetFocus
    End If
    
End If
End Sub

Private Sub cboPeriod_Click()
cmdDateDue.Enabled = True
End Sub

Private Sub cboPeriod_Validate(Cancel As Boolean)
Dim intResponse As Integer
'A validation check to ensure that the user has filled in the rental period
'before trying to calculate the return time
If Len(cboPeriod.Text) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill enter the Rental period")
    If intResponse = 1 Then
        cboPeriod.SetFocus
    End If
End If
'A validation check for logical input ( as rental by the hour is
'more expensive the a maximum number of hours exceptable is set)
If Val(txtRentPeriod) > 8 And cboPeriod.Text = "Hour(s)" Then
    Cancel = True
    intResponse = MsgBox("It is not logical to enter more than 8 hours please use the day option", vbOKOnly, "Ilogical input")
    If intResponse = 1 Then
        cboPeriod.SetFocus
    End If
End If

End Sub

Private Sub cboTypeID_Click()
cmdNewItem.Enabled = True
BlacklistCheck  'Procedure to check if a customer has been black listed
fmeMove.Visible = False
cmdSeekCustomer.Enabled = False
cmdAddCust.Enabled = False
End Sub
'Allows the user direct access to the the add customer screen where
'they can add a new customers details and then only return to this
'screen
Private Sub cmdAddCust_Click()

frmAddCustomer.Show
frmAddCustomer.Tag = "frmRental"
frmAddCustomer.cmdBack.Enabled = True
frmAddCustomer.cmdMainMenu.Enabled = False


End Sub
'Allows you to exit this screen and prompts you to save data before you exit also
'checks if you need to return to Main Menu or Hire Equiptment
Private Sub cmdBack_Click()
Dim intResponse As Integer

If txtTotal = "" Then
    If frmRental.Tag <> "" Then
        frmReturnEquipment.Show
        Unload Me
        frmHireEquipment.Tag = ""
    Else
        frmMainMenu.Show
        Unload Me
    End If
Else
    intResponse = MsgBox("Do you wish to save details before you exit this screen", vbOKCancel, "Unsaved Input")
    If intResponse = vbOK Then
        cmdSave_Click
        If frmHireEquipment.Tag <> "" Then
            frmReturnEquipment.Show
            Unload Me
            frmHireEquipment.Tag = ""
        Else
            frmMainMenu.Show
            Unload Me
        End If
    Else
        cmdCancel_Click
        If frmHireEquipment.Tag <> "" Then
            frmReturnEquipment.Show
            Unload Me
            frmHireEquipment.Tag = ""
        Else
            frmMainMenu.Show
            Unload Me
        End If
    End If
End If
    
        

End Sub

'The cancel button clears all the data stored in text boxes and also
'clears all the data stored in the tags of textboxes(Data has been saved
'in the relevent textbox tags so the tables can be updated at a later stage
'if the user selects the save button
Private Sub cmdCancel_Click()
Dim intindex, intEqID As Integer
Dim varID As Variant  'varID is declared as a Varient as it is to
                      'be used in conjunction with the VB split function

'It has been necessary to mark a peice of equiptments record as out
'(Even though the the rental has not been saved) this is to allow the
'accurate searching of the Equiptment and equipment type tables. But
'if the rental has been canceled the status field of the relevent
'needs to be reset to available
If txtCustID.Tag <> "" And InStr(txtCustID.Tag, Chr(9)) <> 0 Then
'If anyone of the tags used to hold unsaved data contain
'a string and if that string contains a tab then there is more than
'one record saved to the tags so the tags need to be split back into
'values which then canbe used to identified the records altered
    varID = Split(cboID.Tag, Chr(9))
 
    For intindex = 0 To (Val(lblHeader.Tag) - 1)
        intEqID = Val(varID(intindex))
        datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""

        datEquipment.Recordset.Edit
        datEquipment.Recordset("Status") = "Available"
        datEquipment.Recordset.Update

    Next intindex
    
ElseIf txtCustID.Tag <> "" Then
'if the string contained in the tag is only one record
    intEqID = Val(cboID.Tag)
        
    datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""

    datEquipment.Recordset.Edit
    datEquipment.Recordset("Status") = "Available"
    datEquipment.Recordset.Update
End If
'Clear all the relevent input boxes
cboID.Clear
cboID.Text = ""
txtDateOut = ""
txtTimeOut = ""
txtRentPeriod = ""
cboPeriod.Text = ""
txtDateReturn = ""
txtDayDue = ""
txtTimeDue = ""
txtNoInStock = ""
cboEquipModel.Clear
cboEquipModel.Text = ""
cboEquipMake.Clear
cboEquipMake.Text = ""
cboEquipDetails.Clear
cboEquipDetails.Text = ""
cboEquipDescription.Text = ""
cboTypeID.Text = ""
cboID.Enabled = False
txtDateOut.Enabled = False
txtTimeOut.Enabled = False
txtRentPeriod.Enabled = False
cboPeriod.Enabled = False
txtDateReturn.Enabled = False
txtDayDue.Enabled = False
txtTimeDue.Enabled = False
txtNoInStock.Enabled = False
cmdNewItem.Enabled = False
cmdDateDue.Enabled = False
cboEquipDescription.Enabled = False
cmdSave.Enabled = False
cmdNumAvail.Enabled = False
cmdRepeatAny.Enabled = False
cmdPay.Enabled = False
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
txtValueItem = ""
txtTotal = ""
txtNewBalance = ""
txtCustID.Tag = ""
cboID.Tag = ""
txtEmpId.Tag = ""
txtDateOut.Tag = ""
txtDateReturn.Tag = ""
cboPeriod.Tag = ""
txtValueItem.Tag = ""
lblHeader.Tag = ""
cmdNewCust.Enabled = True

End Sub
'The following code works out the date and time a peice of
'equipment should be returned. It takes the time out and
'date out and adds on the the rental period BUT
'-if the return day falls on a Sunday the the return day moves
' to the following Monday
'-if the return time falls outside opening Hours then the
' the amount of hours after closeing today(6:00 pm) are
' added on to opening time(9:00AM) tommorow
'The due back time does not effect the charge
Private Sub cmdDateDue_Click()
Dim dteDateOut, dteDueBack, dteTimeOut, dteTimeBack As Date
Dim intDateOut As Integer
Dim dblTime, dblHour As Double
Dim curRate As Currency
Dim intID As Integer
Dim SQLEqptID, strDayDue As String

Const conWeek As Integer = 7

dteDateOut = txtDateOut

If Not cboPeriod = "Hour(s)" Then
    If cboPeriod = "Week(s)" Then
        dblTime = txtRentPeriod * conWeek
    Else
        dblTime = txtRentPeriod
    End If
    
    dteDueBack = DateAdd("d", dblTime, dteDateOut) 'Add rental period to Out Date
    If Weekday(dteDueBack) = 1 Then               'Must presume that the business closes
        dteDueBack = dteDueBack + 1               'On a Sunday so the due back date moves Monday
    End If
    txtDateReturn = dteDueBack
    txtTimeDue = txtTimeOut
Else
    dteDueBack = txtDateOut
    dteTimeOut = txtTimeOut
    dblHour = Val(txtRentPeriod)
    dteTimeBack = DateAdd("H", dblHour, dteTimeOut) 'Add rental hours to Time out
    If dteTimeBack > TimeSerial(18, 0, 0) Then      'if after closing time
        dteTimeBack = DateAdd("H", 15, dteTimeBack) '15 hours between open and close
        dteDueBack = DateAdd("D", 1, dteDueBack)
    End If
    If Weekday(dteDueBack) = 1 Then               'Must presume that the business closes
        dteDueBack = dteDueBack + 1               'On a Sunday so the due back date moves Monday
    End If
    txtDateReturn = dteDueBack
    txtTimeDue = FormatDateTime(dteTimeBack, vbShortTime)
End If

txtDayDue = WeekdayName(Weekday(dteDueBack - 1))

intID = Val(cboID.Text)
If cboPeriod.Text <> "Hour(s)" Then 'Querry the rental charge from the Equipment Table
    SQLEqptID = "SELECT [Hire Price per Day] FROM  Equipment  WHERE Equipment_ID = " & intID & " ;"

    DatSQL4.DatabaseName = strThePath
    DatSQL4.RecordSource = SQLEqptID
    DatSQL4.Refresh
    
    With DatSQL4.Recordset
        While Not .EOF
            curRate = .Fields("Hire Price per Day")
            .MoveNext
        Wend
    End With
Else
    SQLEqptID = "SELECT [Hire Price per Hour] FROM  Equipment  WHERE Equipment_ID = " & intID & " ;"
    
    
    DatSQL4.DatabaseName = strThePath
    DatSQL4.RecordSource = SQLEqptID
    DatSQL4.Refresh
   
    With DatSQL4.Recordset
        While Not .EOF
            curRate = .Fields("Hire Price per Hour")
            .MoveNext
        Wend
    End With
End If

If cboPeriod.Text = "Week(s)" Then
    curRate = curRate * 7
End If

curRate = curRate * Val(txtRentPeriod)

txtValueItem = curRate
txtTotal = Val(txtTotal) + curRate
txtNewBalance = Val(txtBalance) + Val(txtTotal)

cmdCancel.Enabled = True

cmdRepeatAny.Enabled = True
cmdSave.Enabled = True
cmdDateDue.Enabled = False
cmdNumAvail.Enabled = True
If Val(txtNoInStock) < 2 Then
    cmdNumAvail.Enabled = False
    txtNoInStock = 0
End If

End Sub
'Allows you to exit this screen and prompts you to save data before you exit
Private Sub cmdMainMenu_Click()
Dim intResponse As Integer

If cboID.Text = "" And cboID.Tag = "" Then
    frmMainMenu.Show
    Unload Me
Else
    intResponse = MsgBox("Do you wish to save details before you exit this screen", vbOKCancel, "Unsaved Input")
    If intResponse = vbOK Then
        cmdSave_Click
        frmMainMenu.Show
        Unload Me
    Else
        cmdCancel_Click
        frmMainMenu.Show
        Unload Me
    End If
End If
    
End Sub

'Count the number of peices of equipment that match the selected
'Type of equipment and the load their Equipment_ID numbers into
'a combo box to allow the user to select an available number
Private Sub cmdNewItem_Click()
Dim SqlCount As String
Dim intID As Integer
Dim intCount As Integer
Dim SQLEqptID As String

intID = Val(cboTypeID)
cboID.Clear

SqlCount = "SELECT Count(*)AS IDCount FROM  Equipment  WHERE Type_ID= " & intID & " AND Status =  'Available' AND Deletion = False;"
SQLEqptID = "SELECT Equipment_ID FROM  Equipment  WHERE Type_ID= " & intID & " AND Status =  'Available' AND Deletion = False;"

DatSQL4.DatabaseName = strThePath
DatSQL4.RecordSource = SqlCount
DatSQL4.Refresh

datSQL3.DatabaseName = strThePath
datSQL3.RecordSource = SQLEqptID
datSQL3.Refresh

intCount = DatSQL4.Recordset.Fields("IDCount")
txtNoInStock.Text = Str(intCount)


While Not datSQL3.Recordset.EOF
    cboID.AddItem datSQL3.Recordset("Equipment_ID")
    datSQL3.Recordset.MoveNext
    Wend


cboID.Enabled = True
End Sub
'Clear the screen and reset initial settings
Private Sub cmdNewCust_Click()
cboID.Clear
cboID.Text = ""
txtDateOut = ""
txtTimeOut = ""
txtRentPeriod = ""
cboPeriod.Text = ""
txtDateReturn = ""
txtDayDue = ""
txtTimeDue = ""
txtNoInStock = ""
cboEquipModel.Clear
cboEquipModel.Text = ""
cboEquipMake.Clear
cboEquipMake.Text = ""
cboEquipDetails.Clear
cboEquipDetails.Text = ""
cboEquipDescription.Text = ""
cboTypeID.Text = ""
txtCustID = ""
txtNameFound = ""
txtAddress1 = ""
txtAddress2 = ""
txtAddress3 = ""
txtPhoneNo = ""
txtMobileNo = ""
txtEmail = ""
txtCreditLimit = ""
txtStatus = ""
txtBalance = ""
cboID.Enabled = False
txtDateOut.Enabled = False
txtTimeOut.Enabled = False
txtRentPeriod.Enabled = False
cboPeriod.Enabled = False
txtDateReturn.Enabled = False
txtDayDue.Enabled = False
txtTimeDue.Enabled = False
txtNoInStock.Enabled = False
cboEquipModel.Enabled = False
cboEquipMake.Enabled = False
cboEquipDetails.Enabled = False
cboEquipDescription.Enabled = False
cmdNewItem.Enabled = False
cmdDateDue.Enabled = False
cboTypeID.Enabled = False
optEqptNum.Enabled = False
optEquipName.Enabled = False
cboEquipDescription.Enabled = False
cboTypeID.Enabled = False
cmdSave.Enabled = False
cmdNumAvail.Enabled = False
cmdRepeatAny.Enabled = False
cmdPay.Enabled = False
cmdSave.Enabled = False
cmdCancel.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cboName.Text = ""
cboNumber.Text = ""
cmdSeekCustomer.Enabled = True
End Sub

Private Sub cmdNext_Click()
Dim strName As String

strName = cboName.Text
With datCustomer.Recordset
    .FindNext ("Name like '*" & strName & "*'")
    
End With
txtCustID = datCustomer.Recordset("Cust_ID")
txtNameFound = datCustomer.Recordset("Name")
txtAddress1 = datCustomer.Recordset("Address 1")
txtAddress2 = datCustomer.Recordset("Address 2")
txtAddress3 = datCustomer.Recordset("Address 3")
txtPhoneNo = datCustomer.Recordset("Phone No")
txtMobileNo = datCustomer.Recordset("Mobile No")
txtEmail = datCustomer.Recordset("E-Mail")
txtCreditLimit = datCustomer.Recordset("Credit Limit")
txtStatus = datCustomer.Recordset("Status")
txtBalance = datCustomer.Recordset("Balance owed")
    
End Sub
'selecting this button allows you to select another item
'of equiptment where the Equipment type matches the type
'last selected. The inputs are not saved at this stage but
'instead moved to the tags of the relevent text boxes, the
'text boxes are then cleared to make way for the next inputs
Private Sub cmdNumAvail_Click()
Dim SqlCount As String
Static intID, intNumber As Integer
Dim intCount As Integer
Dim SQLEqptID As String
Dim intResponse As Integer
Dim intEqID As Integer

intEqID = Val(cboID.Text)

If cboID.Tag = "" Then
'If there is nothing stored in the tag the the contents of
'the text boxes is moved into the tag
    lblHeader.Tag = 1
    txtCustID.Tag = txtCustID
    cboID.Tag = cboID.Text
    txtEmpId.Tag = txtEmpId
    txtDateOut.Tag = DateValue(txtDateOut) + TimeValue(txtTimeOut)
    txtDateReturn.Tag = DateValue(txtDateReturn) + TimeValue(txtTimeDue)
    cboPeriod.Tag = cboPeriod.Text
    txtValueItem.Tag = Val(txtValueItem)
    
Else
'if there is something in the tag then the contents of the
'text boxes are concatinated on to the tag
    lblHeader.Tag = Val(lblHeader.Tag) + 1
    txtCustID.Tag = txtCustID.Tag & Chr(9) & txtCustID
    cboID.Tag = cboID.Tag & Chr(9) & cboID.Text
    txtEmpId.Tag = txtEmpId.Tag & Chr(9) & txtEmpId
    txtDateOut.Tag = txtDateOut.Tag & Chr(9) & DateValue(txtDateOut) + TimeValue(txtTimeOut)
    txtDateReturn.Tag = txtDateReturn.Tag & Chr(9) & DateValue(txtDateReturn) + TimeValue(txtTimeDue)
    cboPeriod.Tag = cboPeriod.Tag & Chr(9) & cboPeriod.Text
    txtValueItem.Tag = txtValueItem.Tag & Chr(9) & Val(txtValueItem)
End If
'The on screen view changes to reflect the fact that a peice
'of equipment has been selected and therefore can not be
'reselected. All relevent input boxes are cleared to allow
'the next input
txtNoInStock = Val(txtNoInStock) - 1
If txtNoInStock <> "0" Then
    cboID.RemoveItem (cboID.ListIndex)
    cboID.Text = ""
    txtDateOut = ""
    txtTimeOut = ""
    txtDateReturn = ""
    txtDayDue = ""
    txtTimeDue = ""
    cboID.SetFocus
    cmdDateDue.Enabled = True
Else
'if there are no more of the selected peices of equipment
'left in stock amessage is sent to the screen to inform
'the user
    intResponse = MsgBox("There are no further matching items left in stock", vbOKOnly, "You Can not use this option")
    cboID.RemoveItem (cboID.ListIndex)
    cboID.Text = ""
    txtDateOut = ""
    txtTimeOut = ""
    txtRentPeriod = ""
    cboPeriod.Text = ""
    txtDateReturn = ""
    txtDayDue = ""
    txtTimeDue = ""
End If
cmdNumAvail.Enabled = False
cmdRepeatAny.Enabled = False
'But to allow any future searchs of equipment for the present customer
'The status field of the Equipment table nust be set to "OUT"
datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""

datEquipment.Recordset.Edit
datEquipment.Recordset("Status") = "Out"
datEquipment.Recordset.Update

    

End Sub
'Opens the make a payment screen from this screen
Private Sub cmdPay_Click()
frmMakeaPayment.Show
frmMakeaPayment.Tag = txtCustID
End Sub

Private Sub cmdPrevious_Click()
'Allows the user to search through the
Dim strName As String

strName = cboName.Text
With datCustomer.Recordset
    .FindPrevious ("Name like '*" & strName & "*'")
    
End With
txtCustID = datCustomer.Recordset("Cust_ID")
txtNameFound = datCustomer.Recordset("Name")
txtAddress1 = datCustomer.Recordset("Address 1")
txtAddress2 = datCustomer.Recordset("Address 2")
txtAddress3 = datCustomer.Recordset("Address 3")
txtPhoneNo = datCustomer.Recordset("Phone No")
txtMobileNo = datCustomer.Recordset("Mobile No")
txtEmail = datCustomer.Recordset("E-Mail")
txtCreditLimit = datCustomer.Recordset("Credit Limit")
txtStatus = datCustomer.Recordset("Status")
txtBalance = datCustomer.Recordset("Balance owed")
End Sub

'Allows the user to search for another peice of equipment
'where the Equipment type does not matche the type
'last selected. The inputs are not saved at this stage but
'instead moved to the tags of the relevent text boxes, the
'text boxes are then cleared to make way for the next inputs
Private Sub cmdRepeatAny_Click()
intEqID = Val(cboID.Text)


If cboID.Tag = "" Then
'If there is nothing stored in the tag the the contents of
'the text boxes is moved into the tag
    lblHeader.Tag = 1
    txtCustID.Tag = txtCustID
    cboID.Tag = cboID.Text
    txtEmpId.Tag = txtEmpId
    txtDateOut.Tag = Str(DateValue(txtDateOut) + TimeValue(txtTimeOut))
    txtDateReturn.Tag = Str(DateValue(txtDateReturn) + TimeValue(txtTimeDue))
    cboPeriod.Tag = cboPeriod.Text
    txtValueItem.Tag = Val(txtValueItem)
    
Else
'if there is something in the tag then the contents of the
'text boxes are concatinated on to the tag
    lblHeader.Tag = Val(lblHeader.Tag) + 1
    txtCustID.Tag = txtCustID.Tag & Chr(9) & txtCustID
    cboID.Tag = cboID.Tag & Chr(9) & cboID.Text
    txtEmpId.Tag = txtEmpId.Tag & Chr(9) & txtEmpId
    txtDateOut.Tag = txtDateOut.Tag & Chr(9) & Str(DateValue(txtDateOut) + TimeValue(txtTimeOut))
    txtDateReturn.Tag = txtDateReturn.Tag & Chr(9) & Str(DateValue(txtDateReturn) + TimeValue(txtTimeDue))
    cboPeriod.Tag = cboPeriod.Tag & Chr(9) & cboPeriod.Text
    txtValueItem.Tag = txtValueItem.Tag & Chr(9) & Val(txtValueItem)
End If


cboID.Clear
cboID.Text = ""
txtDateOut = ""
txtTimeOut = ""
txtRentPeriod = ""
cboPeriod.Text = ""
txtDateReturn = ""
txtDayDue = ""
txtTimeDue = ""
txtNoInStock = ""
cboEquipModel.Clear
cboEquipModel.Text = ""
cboEquipMake.Clear
cboEquipMake.Text = ""
cboEquipDetails.Clear
cboEquipDetails.Text = ""
cboEquipDescription.Text = ""
cboTypeID.Text = ""

cboID.Enabled = False
txtDateOut.Enabled = False
txtTimeOut.Enabled = False
txtRentPeriod.Enabled = False
cboPeriod.Enabled = False
txtDateReturn.Enabled = False
txtDayDue.Enabled = False
txtTimeDue.Enabled = False
txtNoInStock.Enabled = False
cboEquipModel.Enabled = False
cboEquipMake.Enabled = False
cboEquipDetails.Enabled = False
cboEquipDescription.Enabled = False
cmdNewItem.Enabled = False
cmdDateDue.Enabled = False
optEqptNum = True
cmdNumAvail.Enabled = False
cmdRepeatAny.Enabled = False
'But to allow any future searchs of equipment for the present customer
'The status field of the Equipment table nust be set to "OUT"

datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""

datEquipment.Recordset.Edit
datEquipment.Recordset("Status") = "Out"
datEquipment.Recordset.Update

End Sub
'Saves the information presently displayed in the relevant
'input boxes as well as the information stored in the tags
'A new unique ID number is created for each new record
'added to the rental table. The total
Private Sub cmdSave_Click()
Dim intResponse As Integer
If Val(txtNewBalance) < Val(txtCreditLimit) Then
    SaveAll 'Procedure to save all in tags and text boxes
    
Else
    intResponse = MsgBox("Please inform the customer that he must be willing to bring the Balance of his account back in line with his credit limit, if he is willing to make a payment now select OK otherwise you must Cancel the Rental", vbOKCancel, "Credit limit exceeded")
    If intResponse = vbOK Then
        SaveAll
        cmdPay_Click
    Else
        cmdCancel_Click
    End If
End If
cmdNewCust.Enabled = True
End Sub
Private Sub SaveAll()
Dim SqlCount As String
Static intID, intNumber As Integer
Dim intCount, intindex As Integer
Dim SQLEqptID As String
Dim intResponse, intResponse2 As Integer
Dim intEqID As Integer
Dim varCust, varID, varEmp, varDate, varPeriod, varIn, varCost As Variant

intEqID = Val(cboID.Text)
If Not datRent.Recordset.BOF Then
    datRent.Recordset.MoveLast
    If intID >= datRent.Recordset("Rental_ID") Then
            intID = intID + 1
    Else
        intID = datRent.Recordset("Rental_ID")
        intID = intID + 1
    End If
Else
    intID = 1
End If
    
    
intResponse2 = MsgBox("Click OK to save  all inputted Data", vbOKCancel, "SAVE NOW")
If intResponse2 = 1 Then
        
    If cboID.Text <> "" Then
        datRent.Recordset.AddNew
        datRent.Recordset("Rental_ID") = intID
        datRent.Recordset("Cust_ID") = txtCustID
        datRent.Recordset("Equipment_ID") = cboID.Text
        datRent.Recordset("EmpNumHire") = txtEmpId
        datRent.Recordset("Date/Time out") = DateValue(txtDateOut) + TimeValue(txtTimeOut)
        datRent.Recordset("Date/Time Due") = DateValue(txtDateReturn) + TimeValue(txtTimeDue)
        datRent.Recordset("Charge Type") = cboPeriod.Text
        datRent.Recordset("Charge Cost") = Val(txtValueItem)
        datRent.Recordset.Update
    End If
        
    If txtCustID.Tag <> "" And InStr(txtCustID.Tag, Chr(9)) <> 0 Then
            
        varCust = Split(txtCustID.Tag, Chr(9))
        varID = Split(cboID.Tag, Chr(9))
        varEmp = Split(txtEmpId.Tag, Chr(9))
        varDate = Split(txtDateOut.Tag, Chr(9))
        varIn = Split(txtDateReturn.Tag, Chr(9))
        varPeriod = Split(cboPeriod.Tag, Chr(9))
        varCost = Split(txtValueItem.Tag, Chr(9))
           
        For intindex = 0 To (Val(lblHeader.Tag) - 1)
            intID = intID + 1
            datRent.Recordset.AddNew
            datRent.Recordset("Rental_ID") = intID
            datRent.Recordset("Cust_ID") = varCust(intindex)
            datRent.Recordset("Equipment_ID") = varID(intindex)
            datRent.Recordset("EmpNumHire") = varEmp(intindex)
            datRent.Recordset("Date/Time out") = varDate(intindex)
            datRent.Recordset("Date/Time Due") = varIn(intindex)
            datRent.Recordset("Charge Type") = varPeriod(intindex)
            datRent.Recordset("Charge Cost") = varCost(intindex)
            datRent.Recordset.Update
            
        Next intindex
            
    ElseIf txtCustID.Tag <> "" Then
        intID = intID + 1
        datRent.Recordset.AddNew
        datRent.Recordset("Rental_ID") = intID
        datRent.Recordset("Cust_ID") = txtCustID.Tag
        datRent.Recordset("Equipment_ID") = cboID.Tag
        datRent.Recordset("EmpNumHire") = txtEmpId.Tag
        datRent.Recordset("Date/Time out") = DateValue(txtDateOut.Tag)
        datRent.Recordset("Date/Time Due") = DateValue(txtDateReturn.Tag)
        datRent.Recordset("Charge Type") = cboPeriod.Tag
        datRent.Recordset("Charge Cost") = Val(txtValueItem.Tag)
        datRent.Recordset.Update
    End If
        
    If cboID.Text <> "" Then
        datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""
        
        datEquipment.Recordset.Edit
        datEquipment.Recordset("Status") = "Out"
        datEquipment.Recordset.Update
    End If
            
    intNumber = Val(txtCustID)
    datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
    datCustomer.Recordset.Edit
    datCustomer.Recordset("Balance owed") = Val(txtNewBalance)
    If Val(txtNewBalance) > Val(txtCreditLimit) Then
        datCustomer.Recordset("Status") = "Blacklisted"
        txtStatus = "Blacklisted"
    End If
    datCustomer.Recordset.Update
    
    If txtCustID.Tag <> "" And InStr(txtCustID.Tag, Chr(9)) <> 0 Then
        varID = Split(cboID.Tag, Chr(9))
 
        For intindex = 0 To (Val(lblHeader.Tag) - 1)
            intEqID = Val(varID(intindex))
            datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""
    
            datEquipment.Recordset.Edit
            datEquipment.Recordset("Status") = "Available"
            datEquipment.Recordset.Update
    
        Next intindex
    
    ElseIf txtCustID.Tag <> "" Then
        intEqID = Val(cboID.Tag)
        
        datEquipment.Recordset.FindFirst "Equipment_ID = " & intEqID & ""

        datEquipment.Recordset.Edit
        datEquipment.Recordset("Status") = "Available"
        datEquipment.Recordset.Update
    End If
End If
cmdCancel_Click
End Sub
'Using the Name/Number/letters entered a search is made of the
'Customer Table to return the first match, the coresponding
'record is then displayed. If a search is made using a name
'and their is more than one match, the user is then allowed
'step through the records of those with matching names
Private Sub cmdSeekCustomer_Click()
Dim strName, strID As String
Dim intNumber, intID As Integer


cmdPrevious.Enabled = True
cmdNext.Enabled = True
cboTypeID.Enabled = True
optEqptNum.Enabled = True
optEquipName.Enabled = True
cmdAddCust.Enabled = True

strID = cboName.Text

SqlCount = "SELECT Count(*)AS IDCount FROM  Customer  WHERE Name Like '*" & strID & "*';"


datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = SqlCount
datSQL1.Refresh


intCount = datSQL1.Recordset.Fields("IDCount")

If intCount > 1 And cboName.Visible = True Then  'Displays matching records if their is more than 1
    fmeMove.Visible = True
    txtMatches = intCount
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
End If
If optName = True Then
    strName = cboName.Text
    If Not Len(strName) = 0 Then
        datCustomer.Recordset.FindFirst "Name like '*" & strName & "*'"
        
        txtCustID = datCustomer.Recordset("Cust_ID")
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
        txtPhoneNo = datCustomer.Recordset("Phone No")
        txtMobileNo = datCustomer.Recordset("Mobile No")
        txtEmail = datCustomer.Recordset("E-Mail")
        txtCreditLimit = datCustomer.Recordset("Credit Limit")
        txtStatus = datCustomer.Recordset("Status")
        txtBalance = datCustomer.Recordset("Balance owed")
    End If
Else
    intNumber = Int(Val(cboNumber.Text))
    
        datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
 
        txtCustID = datCustomer.Recordset("Cust_ID")
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
        txtPhoneNo = "" & datCustomer.Recordset("Phone No")
        txtMobileNo = "" & datCustomer.Recordset("Mobile No")
        txtEmail = "" & datCustomer.Recordset("E-Mail")
        txtCreditLimit = datCustomer.Recordset("Credit Limit")
        txtStatus = datCustomer.Recordset("Status")
        txtBalance = "" & datCustomer.Recordset("Balance owed")
        
        
End If
BlacklistCheck  'Procedure to check if a customer has been black listed

End Sub
'If a customer has been black listed then a message is displayed to the screen
'and his record is removed from the screen
Private Sub BlacklistCheck()

If txtStatus = "Blacklisted" Then
    If MsgBox("This Customer has been BlackListed and the company has discontinued doing business with him", vbOKOnly, "Important message") = vbOK Then
        txtCustID = ""
        txtNameFound = ""
        txtAddress1 = ""
        txtAddress2 = ""
        txtAddress3 = ""
        txtPhoneNo = ""
        txtMobileNo = ""
        txtEmail = ""
        txtCreditLimit = ""
        txtStatus = ""
        txtBalance = ""
        End If
End If
End Sub

'When the form is loaded 4 combo boxes lists are loaded
Private Sub Form_Load()

datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datCustomer.Refresh

datEquipment.DatabaseName = strThePath
datEquipment.RecordSource = "Equipment"
datEquipment.Refresh

datType.DatabaseName = strThePath
datType.RecordSource = "Equipment Type"
datType.Refresh

datRent.DatabaseName = strThePath
datRent.RecordSource = "Rental/Return"
datRent.Refresh

datEmp.DatabaseName = strThePath
datEmp.RecordSource = "Employee"
datEmp.Refresh

While Not datCustomer.Recordset.EOF
    cboName.AddItem datCustomer.Recordset("Name")
    cboNumber.AddItem datCustomer.Recordset("Cust_ID")
    datCustomer.Recordset.MoveNext
Wend

While Not datType.Recordset.EOF
    cboTypeID.AddItem datType.Recordset("Type_ID")
    datType.Recordset.MoveNext
Wend

Dim sqlDescription As String

sqlDescription = "SELECT DISTINCT Description FROM [Equipment Type]"


datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = sqlDescription
datSQL1.Refresh

While Not datSQL1.Recordset.EOF
    cboEquipDescription.AddItem datSQL1.Recordset("Description")
    datSQL1.Recordset.MoveNext
Wend


                    
End Sub


Private Sub optEqptNum_Click()

cboTypeID.Enabled = True
cboEquipDescription.Enabled = False
cboEquipMake.Enabled = False
cboEquipModel.Enabled = False
cboEquipDetails.Enabled = False

End Sub

Private Sub optEquipName_Click()

cboTypeID.Enabled = False
cboEquipDescription.Enabled = True
cboEquipMake.Enabled = True
cboEquipModel.Enabled = True
cboEquipDetails.Enabled = True
BlacklistCheck  'Procedure to check if a customer has been black listed
fmeMove.Visible = False
cmdSeekCustomer.Enabled = False
cmdAddCust.Enabled = False
End Sub

Private Sub optName_Click()
fmeMove.Visible = False
lblNameNumber = "Enter Customers Name"
cboName = ""
cboName.Visible = True
cboNumber.Visible = False
cboName.SetFocus
fmeMove.Visible = False
End Sub

Private Sub optNumber_Click()
fmeMove.Visible = False
lblNameNumber = "Enter Customers ID Number"
cboName.Visible = False
cboNumber.Visible = True
cboNumber = ""
cboNumber.Enabled = True
fmeMove.Visible = False
End Sub
'Ensures that this feild is filled before the user moves on with a date of a correct format
Private Sub txtDateOut_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtDateOut) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill enter the Rental period")
    If intResponse = 1 Then
        txtDateOut.SetFocus
    End If
ElseIf Not IsDate(txtDateOut) Then
    Cancel = True
    intResponse = MsgBox("You must enter a Valid date in this text box in the format dd/mm/yyyy", vbOKOnly, "Incorrect data format")
    If intResponse = 1 Then
        txtDateOut.SetFocus
        txtDateOut = ""
    End If
End If

End Sub
'Ensures that the user fills in his employee number before
'the form can be filled

Private Sub txtEmpId_Validate(Cancel As Boolean)
Dim intResponse, intID, intCount As Integer
Dim SqlCount As String
intID = Val(txtEmpId)

SqlCount = "SELECT Count(*)AS IDCount FROM  Employee  WHERE Employee_ID= " & intID & ";"

datEmp.DatabaseName = strThePath
datEmp.RecordSource = SqlCount
datEmp.Refresh

intCount = datEmp.Recordset.Fields("IDCount")

If intCount = 0 Then
    Cancel = True
    intResponse = MsgBox("You have entered a invalid ID", vbOKOnly, "Incorect Data Entry")
    If intResponse = 1 Then
        txtEmpId.SetFocus
    End If
cmdAddCust.Enabled = True
End If
End Sub
'Ensures that this feild is filled before the user moves on
Private Sub txtRentPeriod_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtRentPeriod) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill enter the Rental period")
    If intResponse = 1 Then
        txtRentPeriod.SetFocus
    End If
ElseIf Not IsNumeric(txtRentPeriod) Then
    Cancel = True
    intResponse = MsgBox("You must enter a number in this box", vbOKOnly, "Numders only")
    If intResponse = 1 Then
        txtRentPeriod.SetFocus
    End If
End If
End Sub
'Ensures that this feild is filled before the user moves on
Private Sub txtTimeOut_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtTimeOut) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill enter the Rental period")
    If intResponse = 1 Then
        txtTimeOut.SetFocus
ElseIf Not IsDate(txtDateOut) Then
    Cancel = True
    intResponse = MsgBox("You must enter a Valid Time in this text box in the format hh:mm", vbOKOnly, "Incorrect data format")
    If intResponse = 1 Then
        txtTimeOut.SetFocus
        txtTimeOut = ""
    End If
End If
End If
End Sub

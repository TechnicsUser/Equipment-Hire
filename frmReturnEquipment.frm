VERSION 5.00
Begin VB.Form frmReturnEquipment 
   Caption         =   "Return Equipment"
   ClientHeight    =   7965
   ClientLeft      =   1470
   ClientTop       =   570
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data datEquip2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datEquipmentType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdmakePayment 
      Caption         =   "Make Payment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox txtCustNumber 
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
      Left            =   5520
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Data datEmpNumReturn 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdReHire 
      Caption         =   "&Re-Hire"
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
      Left            =   11880
      TabIndex        =   7
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Data datEquipment 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame myframe 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   975
      Left            =   13320
      TabIndex        =   30
      Top             =   6960
      Width           =   1575
      Begin VB.OptionButton OptNoPenalty 
         Height          =   255
         Left            =   840
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton OptYesPenalty 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.TextBox txtEmpNumReturn 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   7800
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtIndexEquipment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CmbboxEquipment 
      DataSource      =   "datReturn"
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
      Height          =   405
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Data datReturn 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
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
      Left            =   10200
      TabIndex        =   6
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   8520
      TabIndex        =   9
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Data datCustomer 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtCustomerName 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox txtBalanceowned 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   9960
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6360
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5760
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8760
      Width           =   2295
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9360
      Width           =   2295
   End
   Begin VB.ComboBox CmbCustomer 
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
      Height          =   405
      ItemData        =   "frmReturnEquipment.frx":0000
      Left            =   360
      List            =   "frmReturnEquipment.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtAmountDue 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   10920
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7650
      Width           =   2055
   End
   Begin VB.TextBox txtReturnTime 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   12600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2055
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2400
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox txtCost 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   7080
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox txtDateOut 
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
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   12600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox txtDateDue 
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
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   12600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5040
      Width           =   2010
   End
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
      Height          =   495
      Left            =   13560
      TabIndex        =   10
      Top             =   9840
      Width           =   1575
   End
   Begin VB.TextBox txtEquipmentID 
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
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   7800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6615
      Left            =   120
      TabIndex        =   35
      Top             =   3960
      Width           =   4815
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Balance Owed"
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
         TabIndex        =   45
         Top             =   6120
         Width           =   1440
      End
      Begin VB.Label Label14 
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
         Left            =   120
         TabIndex        =   44
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   43
         Top             =   4920
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
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
         TabIndex        =   42
         Top             =   4320
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Mobile No"
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
         TabIndex        =   41
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Phone No"
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
         TabIndex        =   40
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
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
         TabIndex        =   36
         Top             =   600
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Equipment Hire Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   5160
      TabIndex        =   46
      Top             =   3960
      Width           =   9975
      Begin VB.TextBox txtEmpNumHire 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2640
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "EmpNumReturn"
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
         TabIndex        =   52
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "EmpNumHire"
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
         TabIndex        =   51
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Equipment ID"
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
         TabIndex        =   50
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Date/Time Returned"
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
         Left            =   4920
         TabIndex        =   49
         Top             =   1560
         Width           =   2070
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Date/Time Due"
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
         Left            =   4920
         TabIndex        =   48
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Date/Time Out"
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
         Left            =   4920
         TabIndex        =   47
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Charge Details"
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
      Height          =   2655
      Left            =   5160
      TabIndex        =   53
      Top             =   6720
      Width           =   9975
      Begin VB.TextBox txtChargeType 
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
         Height          =   375
         Left            =   1920
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtPenalty 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   5760
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "Apply Penalty"
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
         Height          =   375
         Left            =   8280
         TabIndex        =   65
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Amount Due   £"
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
         Height          =   405
         Left            =   4080
         TabIndex        =   58
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Penalty            £"
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
         Height          =   405
         Left            =   4080
         TabIndex        =   57
         Top             =   480
         Width           =   1590
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Charge Cost   £"
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
         Height          =   405
         Left            =   240
         TabIndex        =   56
         Top             =   1080
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Charge Type"
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
         TabIndex        =   55
         Top             =   480
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Customer/Equipment Selection"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   60
      Top             =   1200
      Width           =   15015
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   9360
         TabIndex        =   70
         Top             =   240
         Width           =   5535
         Begin VB.Label lblEqSelection 
            Caption         =   "Select Equipment Currently On Hired To Customer"
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
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Display"
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
         Left            =   7440
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OptCustName 
         Caption         =   "Search Customer Name"
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
         TabIndex        =   54
         Top             =   480
         Width           =   2895
      End
      Begin VB.OptionButton OpCustID 
         Caption         =   "Search Customer ID"
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
         Height          =   255
         Left            =   5280
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.Frame Frame5 
         Height          =   1335
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   4575
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   4800
         TabIndex        =   69
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lbleqsel 
         Caption         =   "Equipment On Hire By This Customer"
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
         Height          =   375
         Left            =   10200
         TabIndex        =   67
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Customer ID:"
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
         Left            =   5640
         TabIndex        =   61
         Top             =   1080
         Width           =   1500
      End
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Penalty"
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
      Left            =   1800
      TabIndex        =   64
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label lblcustomer 
      Caption         =   "Label6"
      Height          =   375
      Left            =   5400
      TabIndex        =   59
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblPenaltyApplied 
      AutoSize        =   -1  'True
      Caption         =   "Apply Penalty"
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
      Left            =   11160
      TabIndex        =   29
      Top             =   8520
      Width           =   1395
   End
   Begin VB.Label Lblheader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Return Equipment"
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
      Left            =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   5805
   End
End
Attribute VB_Name = "frmReturnEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose of screen:

'The purose of this scren is to allow the user to add back returned equipment
'to the sytem in order for it to be rehired again.
'when the equipment is returned by the customer the user can select
'the the customer from a list of customer ofr by simmply entering the customers
'id number and clicking display,once the customer is selected the user can
'see all the equipment on hire by that particular customer,the user can than select the
'on the equipment being returned by that customer,when selected all the relevant
'details on that piece of equipment are displayed,the user also has the option
'of appling penanties for late returns which the system has calculated,in this
'screen it is also possible for the user to rehire the equipment back out to the
'customer, also the user can deal with payments from this by press the make payment
'button.


'Author: Derek Stafford com2026
'15/03/2002

'vaiables used                          Purpose

'strSQL                 string;         holds SQL query result
'dtTime                 String;         Holds the time part of a date split
'dtDate                 String;         holds the date part of a date split
'strReturnTime          string;         temperarty holds date
'Isvalid                Boolean;        used to see if entry number is valid
'HourDiff               integer;        holds the different in hours for penalty calulations
'WeekDiff               integer;        holds the different in weeks for penalty calulations
'DayDiff                integer;        holds the different in days for penalty calulations
'intindex               integer;        used as a counter variable in loops etc
'intResult              integer;        holds result of answer from message box
'intEmpReturn           integer;        holds employee id number returning equipment
'intEmpHire             integer;        holds employee id number who rented out the item
'IntCustID              integer;        holds the selected customer id
'intRentalID            integer;        holds the selected rental id
'
'objects used
'
'datEquipmentType       datacontrol     retrieves data from the equipment type table
'datEquipment           datacontrol     retrieves data from the equipment table
'datEquip2              datacontrol     retrieves  data from the equipment table
'datcustomer            datacontrol     retrieves data from customer table
'datEmpNumReturn        datacontrol     retrieves data from find employees table
'datEquipment           datacontrol     retrieves data fom equipment table
'datReturn              datacontrol     retrieves locate rented table









Option Explicit
Dim counter, intRentalID, intCustID, index, intEquipID, intEmpReturn, intEmpHire As Integer, HourlyRate, DailyRate As Currency
Private Sub CheckRental()

Dim strmySQl As String 'query which checks current rentals

datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
strmySQl = "SELECT Customer.Cust_ID, [Rental/Return].Equipment_ID, [Rental/Return].[Date/Time returned], [Rental/Return].Rental_ID " & _
"FROM Customer INNER JOIN [Rental/Return] ON Customer.Cust_ID = [Rental/Return].Cust_ID " & _
"WHERE (((Customer.Cust_ID)= " & intCustID & ") AND (([Rental/Return].[Date/Time returned])is Null));"
datReturn.RecordSource = strmySQl
datReturn.Refresh

End Sub
Private Sub cmbCustomer_Click()
'empties unneeded  data from selected textboxes and suplies new customer details details
intCustID = CmbCustomer.ItemData(CmbCustomer.ListIndex)
cmdmakePayment.Enabled = True
CheckRental
display
If datReturn.Recordset.EOF Then
    Call MsgBox("The selected customer currently has no equipment on hire", vbInformation, "No Equipment On Hire")
Else
    lblEqSelection.Visible = True
    cmdClear.Enabled = True
    CmbboxEquipment.Enabled = True
    txtEmpNumHire.Text = ""
    txtEmpNumReturn.Text = ""
    txtEquipmentID.Text = ""
    txtDateOut.Text = ""
    txtDateDue.Text = ""
    txtChargeType.Text = ""
    txtReturnTime.Text = ""
    txtPenalty.Text = ""
    txtCost.Text = ""
    txtAmountDue.Text = ""
End If
End Sub
Private Sub cmdBack_Click()
    frmEquipmentFileProcessing.Show
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Clear
    cmdmakePayment.Enabled = False
End Sub
Private Sub Clear()
'clears out all textboxes on the screen
CmbCustomer.Text = ""
OptNoPenalty.Enabled = False
OptYesPenalty.Enabled = False
cmdReHire.Enabled = False
cmdReHire.Enabled = False
cmdClear.Enabled = False
cmdReturn.Enabled = False
lblEqSelection.Visible = False
txtCustomerName = ""
txtAddress1 = ""
txtAddress2 = ""
txtAddress3 = ""
txtPhoneNo = ""
txtMobileNo = ""
txtEmail = ""
txtEmpNumHire = ""
txtEmpNumReturn = ""
txtEquipmentID = ""
txtDateOut = ""
txtDateDue = ""
txtChargeType = ""
txtCost = ""
txtAmountDue = ""
txtReturnTime = ""
txtStatus = ""
txtCreditLimit = ""
txtBalanceowned = ""
txtPenalty = ""
CmbboxEquipment.Text = ""
End Sub

Private Sub display()
'procedure used to display customer info

datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
txtCustomerName.Text = datCustomer.Recordset("Name")
txtAddress1.Text = datCustomer.Recordset("Address 1")
txtAddress2.Text = datCustomer.Recordset("Address 2")
txtAddress3.Text = datCustomer.Recordset("Address 3")
txtPhoneNo.Text = datCustomer.Recordset("Phone No")
txtMobileNo.Text = datCustomer.Recordset("Mobile No")
txtEmail.Text = datCustomer.Recordset("E-Mail")
txtStatus.Text = datCustomer.Recordset("Status")
txtCreditLimit.Text = datCustomer.Recordset("Credit Limit")
txtBalanceowned.Text = datCustomer.Recordset("Balance owed")
    
End Sub
Private Sub cmdDisplay_Click()
Dim strmySQl As String 'displays all current customers
cmdDisplay.Enabled = False
intCustID = Val(txtCustNumber)
    
strmySQl = "SELECT Customer.*" & _
"From Customer " & _
"WHERE (((Customer.Deletion)=False));"
CmbCustomer.Clear
datCustomer.RecordSource = strmySQl
datCustomer.Refresh
    
datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
If datCustomer.Recordset.NoMatch Then
    Call MsgBox("No Match Found", vbOKOnly, "No Match")
    txtCustNumber = ""
    
Else
    cmdmakePayment.Enabled = True
    display
    CheckRental
    If datReturn.Recordset.EOF Then
        Call MsgBox("The selected customer currently has no equipment on hire", vbInformation, "No Equipment On Hire")
    Else
        CmbboxEquipment.Enabled = True
    End If
End If
End Sub

Private Sub cmdMakePayment_Click()
frmMakeaPayment.Tag = intCustID
frmMakeaPayment.Show
End Sub
Private Sub cmdReHire_Click()
'sents the relavant info to hire screen so the customer can rehire the same piece of
'equipment, it also enables and diables the needed buttons on the hire screen

Dim dtTime, dtDate, strReturnTime As String
If txtReturnTime.Text = "" Then
    Call MsgBox("This Equipment Has Not Been Returned", vbInformation, "Equipment Still On Hire")
Else
    strReturnTime = txtReturnTime
    dtTime = Mid(strReturnTime, 11, Len(strReturnTime))
    dtDate = Mid(strReturnTime, 1, 10)
    frmRental.Tag = "frmReturnEquipment"
    frmRental.Show
    frmRental.txtCustID = intCustID
    frmRental.txtEmpID = intEmpReturn
    frmRental.cboName.Text = txtCustomerName.Text
    frmRental.txtNameFound = txtCustomerName.Text
    frmRental.txtAddress1 = txtAddress1.Text
    frmRental.txtAddress2 = txtAddress1.Text
    frmRental.txtAddress3 = txtAddress1.Text
    frmRental.txtPhoneNo = txtPhoneNo.Text
    frmRental.txtMobileNo = txtMobileNo.Text
    frmRental.txtCreditLimit = txtCreditLimit.Text
    frmRental.txtStatus = txtStatus.Text
    frmRental.txtBalance = txtBalanceowned.Text
    frmRental.txtDateOut = dtDate
    frmRental.txtTimeOut = dtTime
    frmRental.cboID.Text = intEquipID
    frmRental.txtRentPeriod.TabIndex = 1
    frmRental.cmdBack.Enabled = False
    frmRental.cmdDateDue.Enabled = False
    frmRental.cboID.Enabled = False
    frmRental.txtRentPeriod.Enabled = True
    frmRental.cboPeriod.Enabled = True
    frmRental.txtEmpID.Enabled = False
    frmRental.cmdSeekCustomer.Enabled = False
    frmRental.cmdAddCust.Enabled = False
    frmRental.optName.Enabled = False
    frmRental.optNumber.Enabled = False
    frmRental.cboName.Enabled = False
    Unload Me
End If
End Sub

'function used to check that inputted characters are  valid numbers

Function IsValid(StrEQID As Variant) As Boolean
Dim intindex As Integer, temp As Variant, blnValid As Boolean
blnValid = False
For intindex = 1 To Len(StrEQID)
    temp = Mid(StrEQID, intindex, 1)
    If (temp > 0 And temp < 9) Then
    blnValid = True
    Else
    intindex = Len(StrEQID)
    blnValid = False
      
    End If
Next
IsValid = blnValid
End Function

'checks employee id and if its valid it allow the user to
'to return equipment back into the data base,by updating the rental tables


Private Sub cmdReturn_Click()
Dim intEmpNumReturn, intResult As Integer
Dim srtmySQL As String

datEmpNumReturn.RecordSource = "Employee"
datEmpNumReturn.Refresh
datReturn.RecordSource = "Rental/Return"
datReturn.Refresh


txtEmpNumReturn = InputBox("Enter Your Employee_ID Number Please", "Enter Employee ID")
If IsValid(txtEmpNumReturn.Text) = False Or (Val(txtEmpNumReturn) > 9999) Then
    Call MsgBox("You have entered invalid data or no data , a number exceeeding 9999", vbCritical, "Invalid Data Entry")
    CmbboxEquipment.SetFocus
Else
    intEmpNumReturn = Val(txtEmpNumReturn)
    'searches the employee table for match
        
    datEmpNumReturn.Recordset.FindFirst "Employee_ID = " & intEmpNumReturn & ""
            
    If datEmpNumReturn.Recordset("Employee_ID") = intEmpNumReturn Then
        intResult = MsgBox("Are You Sure?", vbYesNo, "Returning Equipment")
                
        If intResult = vbYes Then
            cmdmakePayment.Enabled = True
                    
            intEquipID = Val(txtEquipmentID.Text)
            intEmpNumReturn = Val(txtEmpNumReturn.Text)
            datEmpNumReturn.Recordset.FindFirst "Employee_ID =" & intEmpNumReturn & ""
            txtReturnTime = FormatDateTime(Now, vbShortDate) + "  " + FormatDateTime(Now, vbShortTime)
            
            datEquipment.Recordset.FindFirst "Equipment_ID = " & intEquipID & ""
            'updates the relevant field in the equipment table
            datEquipment.Recordset.Edit
            datEquipment.Recordset("Status") = "Available"
            datEquipment.Recordset.Update
            'updates the relevant field in the rental table
            datReturn.Recordset.FindFirst "Equipment_ID = " & intEquipID & ""
            datReturn.Recordset.Edit
            datReturn.Recordset("EmpNumReturn") = txtEmpNumReturn.Text
            datReturn.Recordset("Date/Time returned") = txtReturnTime
            datReturn.Recordset.Update
            datReturn.Refresh
                    
            txtAmountDue = datReturn.Recordset("Amount Due")
            txtDateDue = datReturn.Recordset("Date/Time Due")
            
            If txtDateDue.Text >= txtReturnTime.Text Then
                datReturn.Recordset.FindFirst "Equipment_ID =" & intEquipID & ""
                OptNoPenalty.Enabled = False
                        
                datReturn.Recordset.Edit
                datReturn.Recordset("Penalty") = CCur(Val(txtPenalty.Text))                'FormatDateTime(Now, vbShortDate) + "  " + FormatDateTime(Now, vbShortTime)
                txtPenalty = datReturn.Recordset("Penalty")
                datReturn.Recordset.Update
                    
            Else
                
                OptYesPenalty.Enabled = True
                datReturn.Recordset.Edit
                datReturn.Recordset("Penalty") = ccurPenality  'function
                txtPenalty = datReturn.Recordset("Penalty")
                datReturn.Recordset.Update
            End If
                    
                    
            Call MsgBox("Return Complete", vbInformation, "System Response")
            
            datEmpNumReturn.Recordset.FindFirst "Employee_ID =" & txtEmpNumReturn & ""
            intEmpReturn = txtEmpNumReturn
            txtEmpNumReturn.Text = datEmpNumReturn.Recordset("Employee Name") 'puts out employee name
                  
            
            
            
            cmdReturn.Enabled = False
            CheckRental
            If datReturn.Recordset.EOF Then
                Call MsgBox("The selected customer currently has no equipment on hire", vbInformation, "No Equipment On Hire")
            End If
      
        Else
            Call MsgBox("No Return ", vbInformation, "System Response")
            txtEmpNumReturn = ""
        End If
    Else
        txtEmpNumReturn = ""
        Call MsgBox("Wrong Or Incorrect ID Number Entered", vbCritical, "Wrong ID Number Entered")
    End If
End If
End Sub
Private Sub cmdMainMenu_Click()
    frmMainMenu.Show
Unload Me
End Sub
Private Sub Form_Activate()
'sets up screen layout
cmdDisplay.Enabled = False
OptCustName.SetFocus
cmdmakePayment.Enabled = False
lblEqSelection.Visible = False
txtChargeType.Enabled = False
txtCustNumber.Enabled = False
CmbCustomer.Enabled = False
cmdReHire.Enabled = False
cmdReHire.Enabled = False
cmdClear.Enabled = False
cmdReturn.Enabled = False
OptNoPenalty.Enabled = False
OptYesPenalty.Enabled = False
cmdDisplay.ToolTipText = "Displays customer details"
cmdMainMenu.ToolTipText = "Back to main menu"
cmdClear.ToolTipText = "Clears current contains from the screen"
cmdReHire.ToolTipText = "Rehire the selected equipment"
CmbboxEquipment.ToolTipText = "Selects the equipment on hire from the selected customer"
cmdReturn.ToolTipText = "Allows employee to return the selected equipment onto the system"
cmdmakePayment.ToolTipText = "Allows Customer to make a payment"
End Sub
Private Sub cmbCustomer_DropDown()
'fills combo box with current customers

Dim strmySQl As String
Clear
strmySQl = "SELECT Customer.* " & _
"From Customer " & _
"WHERE (((Customer.Deletion)=False));"
CmbCustomer.Clear
datCustomer.RecordSource = strmySQl
datCustomer.Refresh
While Not datCustomer.Recordset.EOF
    CmbCustomer.AddItem (datCustomer.Recordset("Name")) + ", " + (datCustomer.Recordset("Address 1")) + ", " + (datCustomer.Recordset("Address 2"))
    CmbCustomer.ItemData(CmbCustomer.NewIndex) = datCustomer.Recordset("Cust_ID")
    datCustomer.Recordset.MoveNext
Wend
End Sub
Private Sub Form_Load()
'assigns the relevant data base tables to data controls

Dim strmySQl As String
CmbboxEquipment.Enabled = False

datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"

datEmpNumReturn.DatabaseName = strThePath
datEmpNumReturn.RecordSource = "Employee"  '"select*from supplier where [deletion]= false;"

datEquipment.DatabaseName = strThePath
datEquipment.RecordSource = "Equipment"

datReturn.DatabaseName = strThePath
datReturn.RecordSource = "Rental/Return"

datEquip2.DatabaseName = strThePath
datEquip2.RecordSource = "Equipment"

datEquipmentType.DatabaseName = strThePath
datEquipmentType.RecordSource = "Equipment Type"
End Sub
Private Sub CmbboxEquipment_DropDown()
'fills equipment combo box with all the equipment the selected customer has on hire

Dim strmySQl As String, intEquipType  As Integer
    
    
    
strmySQl = "SELECT Customer.Cust_ID, [Rental/Return].Equipment_ID, [Rental/Return].[Date/Time returned], [Rental/Return].Rental_ID " & _
"FROM Customer INNER JOIN [Rental/Return] ON Customer.Cust_ID = [Rental/Return].Cust_ID " & _
"WHERE (((Customer.Cust_ID)= " & intCustID & ") AND (([Rental/Return].[Date/Time returned])Is Null));"
datReturn.RecordSource = strmySQl

datReturn.Refresh
datEquipmentType.RecordSource = "Equipment Type"
CmbboxEquipment.Clear
While Not datReturn.Recordset.EOF
        
    intEquipID = datReturn.Recordset("Equipment_ID")
        
    datEquip2.Recordset.FindFirst "Equipment_ID = " & intEquipID & ""
    intEquipType = datEquip2.Recordset("Type_ID")
        
    datEquipmentType.Recordset.FindFirst "Type_ID = " & intEquipType & ""
        
    CmbboxEquipment.AddItem (datEquipmentType.Recordset("Make")) + ",  " + datEquipmentType.Recordset("Model") + ",  " + datEquipmentType.Recordset("Details")
    datEquipmentType.Recordset.MoveFirst
    datEquip2.Recordset.MoveFirst
        
    CmbboxEquipment.ItemData(CmbboxEquipment.NewIndex) = datReturn.Recordset("Equipment_ID")
    datReturn.Recordset.MoveNext
Wend
End Sub

Private Sub CmbboxEquipment_Click()
Dim Cost, AmountDue As Currency
'find the relevant info on the selected equipment and places it in combo boxes

intEquipID = CmbboxEquipment.ItemData(CmbboxEquipment.ListIndex)
datReturn.RecordSource = "Rental/Return"
datReturn.Refresh
datEmpNumReturn.Refresh
datReturn.Recordset.FindFirst "Equipment_ID = " & intEquipID & ""

datEmpNumReturn.Recordset.FindFirst "Employee_ID =" & intEmpHire & ""

txtEmpNumHire.Text = datEmpNumReturn.Recordset("Employee Name")

txtEquipmentID.Text = datReturn.Recordset("Equipment_ID")
txtDateOut.Text = datReturn.Recordset("Date/Time Out")
txtDateDue.Text = datReturn.Recordset("Date/Time Due")


txtChargeType.Text = datReturn.Recordset("Charge Type")
Cost = CCur(datReturn.Recordset("Charge Cost"))
txtCost.Text = Str(Cost)
txtPenalty = ccurPenality
If txtPenalty = "" Then
    txtPenalty = "0"
End If
If txtPenalty = "0" Then
    OptYesPenalty.Enabled = False
Else
    OptYesPenalty.Enabled = True
End If
AmountDue = CCur(datReturn.Recordset("Amount Due"))
txtAmountDue.Text = Str(AmountDue)
OptNoPenalty.Enabled = False
cmdReHire.Enabled = True
cmdReHire.Enabled = True
cmdReturn.Enabled = True
End Sub
Private Sub optCustName_Click()
'resets screen layout
cmdDisplay.Enabled = False
txtCustNumber.Enabled = False
CmbCustomer.Enabled = True
txtCustNumber.Text = ""
txtCustNumber.Enabled = False
Clear
End Sub

Private Sub OpCustID_Click()
'reset screen layout
Clear
cmdmakePayment.Enabled = False
CmbCustomer.Text = ""
CmbCustomer.Enabled = False
txtCustNumber.Enabled = True
txtCustNumber.SetFocus
cmdDisplay.Enabled = True
txtCustNumber.Enabled = True
End Sub
Private Sub OptNoPenalty_Click()
'takes added penalty away from the amount due
    
    If OptNoPenalty.Value = True Then
        OptYesPenalty.Value = False
        txtAmountDue.Text = Str(CCur(txtAmountDue.Text) - CCur(Val(txtPenalty.Text)))
    End If
End Sub
Private Sub OptYesPenalty_Click()
'adds penalty away from the amount due

If txtPenalty.Text = "0" Then
    Call MsgBox("There is No Penalty ON The Selected Equipment", vbInformation, "No Penalty")
End If
If OptYesPenalty.Value = True Then
    OptNoPenalty.Value = False
    OptNoPenalty.Enabled = True
End If
txtAmountDue = Str(CCur(txtAmountDue.Text) + CCur(Val(txtPenalty.Text)))
End Sub
Public Function ccurPenality()

'works oout the penalty due if any

Dim total As Currency, HourDiff, DayDiff, WeekDiff As Integer
Const PentRate As Currency = 1.5

DayDiff = DateDiff("d", txtDateDue, Now)
HourDiff = DateDiff("h", txtDateDue, Now) 'datdiff function
WeekDiff = DateDiff("ww", txtDateDue, Now) 'reference Peter Nortons  Guide to Visual Basic 6

If DayDiff > 0 Or HourDiff > 0 Or WeekDiff > 0 Then
    If txtChargeType = "Hour(s)" Then
        ccurPenality = Str(HourDiff * Val(txtCost.Text) * PentRate)
    Else
        If txtChargeType = "Day(s)" Then
            ccurPenality = Str(DayDiff * Val(txtCost.Text) * PentRate)
        Else
            If txtChargeType = "Week(s)" Then
                ccurPenality = Str(WeekDiff * Val(txtCost.Text) * PentRate)
            End If
        End If
    End If
End If
End Function
Private Sub txtCustNumber_Validate(Cancel As Boolean)
'checks for valid input
If IsNumeric(txtCustNumber.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtCustNumber.SetFocus
    Cancel = True
End If
End Sub


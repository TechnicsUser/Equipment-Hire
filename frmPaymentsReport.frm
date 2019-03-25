VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmPaymentsReport 
   Caption         =   "Payments Received Report screen"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
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
      Height          =   615
      Left            =   9480
      TabIndex        =   37
      Top             =   9960
      Width           =   1935
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
      Height          =   615
      Left            =   12000
      TabIndex        =   36
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Frame frmConfirm 
      Caption         =   "Confirmation"
      ForeColor       =   &H00800000&
      Height          =   4455
      Left            =   3360
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "Confirm"
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Frame fmeMove 
         Height          =   2535
         Left            =   4320
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   3015
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
            Left            =   2280
            TabIndex        =   10
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Display Previous"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Display Next"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblMatches 
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
            Height          =   855
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1815
         End
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1440
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
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
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
         Left            =   2040
         TabIndex        =   4
         Top             =   2880
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
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblName 
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
         Left            =   1065
         TabIndex        =   15
         Top             =   720
         Width           =   600
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
         Left            =   720
         TabIndex        =   14
         Top             =   2880
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
         Left            =   720
         TabIndex        =   13
         Top             =   2160
         Width           =   945
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
         Left            =   720
         TabIndex        =   12
         Top             =   1440
         Width           =   945
      End
   End
   Begin VB.Frame fmeName 
      Caption         =   "Customer Selection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   2880
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   9255
      Begin VB.ComboBox cboNumber 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPaymentsReport.frx":0000
         Left            =   4080
         List            =   "frmPaymentsReport.frx":0002
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdSeekCustomer 
         Caption         =   "Display Customer"
         Enabled         =   0   'False
         Height          =   495
         Left            =   6120
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "frmPaymentsReport.frx":0004
         Left            =   3480
         List            =   "frmPaymentsReport.frx":000B
         Sorted          =   -1  'True
         TabIndex        =   20
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
         Left            =   4920
         TabIndex        =   19
         Top             =   360
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
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
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
         Left            =   960
         TabIndex        =   23
         Top             =   1320
         Width           =   2250
      End
   End
   Begin VB.Frame fmeSort 
      Caption         =   "Sorting Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   10320
      TabIndex        =   32
      Top             =   2400
      Width           =   3855
      Begin VB.OptionButton optDateSort 
         Caption         =   "Sort by Date"
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
         Left            =   360
         TabIndex        =   34
         Top             =   480
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optCustomerSort 
         Caption         =   "Sort by Customer"
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
         Height          =   435
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Frame fmeCustomerOptions 
      Caption         =   "Customer Selection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   1200
      TabIndex        =   27
      Top             =   2400
      Width           =   3855
      Begin VB.OptionButton optSelectCust 
         Caption         =   "Select a Customer"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optAllCust 
         Caption         =   "All Customers"
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
         TabIndex        =   28
         Top             =   480
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.Frame fmeDateOptions 
      Caption         =   "Date Selection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   5760
      TabIndex        =   26
      Top             =   2400
      Width           =   3855
      Begin VB.OptionButton optSelectDate 
         Caption         =   "Select a Date"
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
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optAllDates 
         Caption         =   "All Dates"
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
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.Frame fmeDate 
      Caption         =   "Date Selection"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4815
      Left            =   2280
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton cmdConfirmDate 
         BackColor       =   &H00000000&
         Caption         =   "Con&firm"
         Height          =   495
         Left            =   3600
         TabIndex        =   35
         Top             =   4080
         Width           =   3015
      End
      Begin MSACAL.Calendar calCalander 
         Height          =   3615
         Left            =   600
         TabIndex        =   25
         Top             =   360
         Width           =   8055
         _Version        =   524288
         _ExtentX        =   14208
         _ExtentY        =   6376
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2002
         Month           =   3
         Day             =   12
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fxgPayments 
      Bindings        =   "frmPaymentsReport.frx":0016
      Height          =   5055
      Left            =   1080
      TabIndex        =   1
      Top             =   4560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8916
      _Version        =   393216
      ForeColor       =   8388608
      AllowUserResizing=   3
      MousePointer    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data datGrid 
      Caption         =   "Grid Control"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data datCustomer 
      Caption         =   "Customer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data datSQL1 
      Caption         =   "SQL1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Payments Received Report"
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
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "frmPaymentsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to Display a number of different Reports
'               depending on the options selected
'Student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02

'Variables used
'strName        String
'strID          String
'intNumber      Integer
'intID          Integer
'intResponse    Integer
'dateDate       Date
'sqlDescription String
 



Private Sub cboName_Click()
cmdSeekCustomer.Enabled = True

End Sub
'Returns to the reports menu
Private Sub cmdBack_Click()
frmReportsMenu.Show
Unload Me
End Sub

'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub cmdConfirm_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
EnableOptions
fmeName.Visible = False
frmConfirm.Visible = False

End Sub
'A function to enable the select and sort option buttons
Private Sub EnableOptions()

optAllDates.Enabled = True
optAllCust.Enabled = True
optSelectCust.Enabled = True
optDateSort.Enabled = True
optCustomerSort.Enabled = True
optSelectDate.Enabled = True
End Sub
'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub cmdConfirmDate_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
EnableOptions
fmeDate.Visible = False

End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub optAllCust_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
End Sub
'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub optAllDates_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
End Sub
'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub optCustomerSort_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
End Sub
'Runs the function "functSQL" which selects the appropiate sql statement
'and returns the value to the data grid
Private Sub optDateSort_Click()

datGrid.RecordSource = functSQL
datGrid.Refresh
End Sub

Private Sub optName_Click()

lblNameNumber = "Enter Customers Name"
cboName = ""
cboName.Visible = True
cboNumber.Visible = False
cboName.SetFocus

End Sub

Private Sub optNumber_Click()

cboNumber.Text = cboNumber.List(0)
lblNameNumber = "Enter Customers ID Number"
cboName.Visible = False
cboNumber.Visible = True
cboNumber = ""
cboNumber.Enabled = True

End Sub
Private Sub cboName_Change()
cboName_Click
End Sub
'This function checks to see which search and sort options have been selected
'then selects the correct SQL querry
Private Function functSQL() As String
Dim dateDate As Date
Dim intID As Integer

If Not optAllDates Then
    dateDate = Format(calCalander.Value, "mm dd yyyy")
End If
If Len(cboNumber.Text) <> 0 Then
    intID = Val(cboNumber.Text)
End If

If optAllCust = True And optAllDates = True And optDateSort = True Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID ORDER BY Payment.Date;"
ElseIf optAllCust = False And optAllDates = True And optDateSort = True Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Customer.Cust_ID = " & intID & " ORDER BY Payment.Date;"
ElseIf optAllCust = False And optAllDates = False And optDateSort = True Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Customer.Cust_ID = " & intID & " AND Payment.Date = #" & dateDate & "# ORDER BY Payment.Date;"
ElseIf optAllCust = False And optAllDates = False And optDateSort = False Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Customer.Cust_ID = " & intID & " AND Payment.Date = #" & dateDate & "# ORDER BY Customer.Name;"
ElseIf optAllCust = False And optAllDates = True And optDateSort = False Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Customer.Cust_ID = " & intID & " ORDER BY Customer.Name;"
ElseIf optAllCust = True And optAllDates = True And optDateSort = False Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID ORDER BY Customer.Name;"
ElseIf optAllCust = True And optAllDates = False And optDateSort = True Then
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Payment.Date = #" & dateDate & "# ORDER BY Payment.Date;"
Else
    functSQL = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID WHERE Payment.Date = #" & dateDate & "# ORDER BY Customer.Name;"
End If
End Function
'Ensures that the cboName input section has been filled

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
cmdSeekCustomer.Enabled = True
End Sub

Private Sub cboNumber_Click()
cmdSeekCustomer.Enabled = True

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

strID = cboName.Text

SqlCount = "SELECT Count(*)AS IDCount FROM  Customer  WHERE Name Like '*" & strID & "*';"

datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = SqlCount
datSQL1.Refresh

intCount = datSQL1.Recordset.Fields("IDCount")

If intCount > 1 Then  'Displays matching records if their is more than 1
    fmeMove.Visible = True
    txtMatches = intCount
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
End If
If optName = True Then
    strName = cboName.Text
    If Not Len(strName) = 0 Then
        datCustomer.Recordset.FindFirst "Name like '*" & strName & "*'"
        
        cboNumber.Text = datCustomer.Recordset("Cust_ID")
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
        
    End If
Else
    intNumber = Int(Val(cboNumber.Text))
    
        datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
 
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
   
        
End If
frmConfirm.Visible = True 'A confirm option is used to ensure that the correct
                          'information has been selected

End Sub
'This only becomes visable if there is a match in names selected and allows you to
'move forward through each possible customer until you can confirm you have
'the correct one from the extra details shown
Private Sub cmdNext_Click()
Dim strName As String

strName = cboName.Text
With datCustomer.Recordset
    .FindNext ("Name like '*" & strName & "*'")
    
End With
cboNumber.Text = datCustomer.Recordset("Cust_ID")
txtNameFound = datCustomer.Recordset("Name")
txtAddress1 = datCustomer.Recordset("Address 1")
txtAddress2 = datCustomer.Recordset("Address 2")
txtAddress3 = datCustomer.Recordset("Address 3")

    
End Sub
'This only becomes visable if there is a match in names selected and allows you to
'move backwards through each possible customer until you can confirm you have
'the correct one from the extra details shown
Private Sub cmdPrevious_Click()
Dim strName As String

strName = cboName.Text
With datCustomer.Recordset
    .FindPrevious ("Name like '*" & strName & "*'")
    
End With
cboNumber.Text = datCustomer.Recordset("Cust_ID")
txtNameFound = datCustomer.Recordset("Name")
txtAddress1 = datCustomer.Recordset("Address 1")
txtAddress2 = datCustomer.Recordset("Address 2")
txtAddress3 = datCustomer.Recordset("Address 3")

    
End Sub
'Fills the customer name and id number combo box's with information stored in the
'Customer table also an SQL query which is correct for the initial set up of
'the option buttons is run and the Data Grid is filled

Private Sub Form_Load()
Dim sqlDescription As String
cboName.Text = cboName.List(0)
datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datCustomer.Refresh

While Not datCustomer.Recordset.EOF
    cboName.AddItem datCustomer.Recordset("Name")
    cboNumber.AddItem datCustomer.Recordset("Cust_ID")
    datCustomer.Recordset.MoveNext
Wend

sqlDescription = "SELECT Payment.Payment_ID, Customer.Cust_ID, Customer.Name, Customer.[Address 1], Customer.[Address 2], Customer.[Address 3], Payment.Date, Payment.Amount, Payment.[Payment Method] FROM Customer INNER JOIN Payment ON Customer.Cust_ID = Payment.Cust_ID ORDER BY Payment.Date;"

datGrid.DatabaseName = strThePath
datGrid.RecordSource = sqlDescription
datGrid.Refresh

datGrid.Refresh
End Sub
'Disables all the options while a selection process is taking place
Private Sub DisableOptions()

optAllDates.Enabled = False
optAllCust.Enabled = False
optSelectDate.Enabled = False
optSelectCust.Enabled = False
optDateSort.Enabled = False
optCustomerSort.Enabled = False

End Sub

Private Sub optSelectCust_Click()

DisableOptions
fmeName.Visible = True

End Sub

Private Sub optSelectDate_Click()

DisableOptions
fmeDate.Visible = True

End Sub

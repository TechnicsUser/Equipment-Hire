VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOrdersReport 
   Caption         =   "Orders Report"
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
   Begin VB.Data datTotal 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid flxTotal 
      Bindings        =   "frmOrdersReport.frx":0000
      Height          =   5055
      Left            =   12000
      TabIndex        =   14
      Top             =   4680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8916
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      ForeColor       =   8388608
      WordWrap        =   -1  'True
      FocusRect       =   0
      MergeCells      =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Default         =   -1  'True
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
      Left            =   7080
      TabIndex        =   2
      ToolTipText     =   "Show order report"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Data datOrder 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrder 
      Bindings        =   "frmOrdersReport.frx":0017
      Height          =   5055
      Left            =   240
      TabIndex        =   13
      Top             =   4680
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8916
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      FixedCols       =   0
      ForeColor       =   8388608
      WordWrap        =   -1  'True
      FocusRect       =   0
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   $"frmOrdersReport.frx":002E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort By"
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
      Height          =   2055
      Left            =   10320
      TabIndex        =   12
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton optSortOrder 
         Caption         =   "Sort by Order &ID"
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
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optSortSupp 
         Caption         =   "Sort by Su&pplier ID"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&<Back"
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
      Left            =   11280
      TabIndex        =   7
      Top             =   10200
      Width           =   1575
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
      Left            =   13320
      TabIndex        =   8
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "List By"
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
      Height          =   2055
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   7455
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
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
         Left            =   3600
         TabIndex        =   15
         ToolTipText     =   "Show order report"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtList 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Left            =   4560
         TabIndex        =   0
         ToolTipText     =   "Enter date"
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optAll 
         Caption         =   "List all &Orders"
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
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optSupplier 
         Caption         =   "List by S&upplier"
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
         TabIndex        =   3
         Top             =   1050
         Width           =   1575
      End
      Begin VB.OptionButton optDate 
         Caption         =   "List by &Date"
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
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblList 
         AutoSize        =   -1  'True
         Caption         =   "Enter Date"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   645
         Width           =   1110
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Orders Report"
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
      Left            =   5505
      TabIndex        =   9
      Top             =   240
      Width           =   4545
   End
End
Attribute VB_Name = "frmOrdersReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to show the user an order report.

'The user has the option of listing orders by date, by supplier, or all the orders.
'The user also has the option of sorting by order ID or supplier ID.
'When the user clicks show the details of every order is shown in one flex grid and each order cost is
'shown in another flex grid.


'Author Fergal Purcell
'Date   13/03/2002

'Variables used
'intCounter     :   Integer; This variable is used as a control variable in a For Loop
'strSql         :   String; This string variable is used to store an SQL
'strSort        :   String; Holds the field by which the user wishes to sort the report
'strSql2        :   String; Stores the second SQL which finds the cost of each order


'Objects used
'datOrder   :   Data Control; Holds the result of the first SQL
'datTotal   :   Data Control; Holds the result of the second SQL

Option Explicit
    
Private Function CheckNum() As Boolean
'Bug in IsNumeric says that 6d6 is numeric
Dim intCounter As Integer
    For intCounter = 1 To Len(txtList.Text)
        If Not IsNumeric(Mid(txtList, intCounter, 1)) Then        'Check each character
            CheckNum = True
        End If
    Next
End Function

Private Sub cmdBack_Click()
frmReportsMenu.Show
Unload Me
End Sub

Private Sub cmdBrowse_Click()
frmBrowse.Tag = "Report"
frmBrowse.Show
End Sub

Private Sub cmdShow_Click()
Dim strSQL As String, strSort As String, strSql2 As String
datOrder.DatabaseName = strThePath
datTotal.DatabaseName = strThePath
If optSortOrder.Value = True Then
    strSort = "Order.Order_ID"
Else
    strSort = "Order.Supplier_ID"
End If
If Optdate.Value = True Then                'List orders by date
    If IsDate(txtList.Text) = True Then     'Check to see if user has entered the correct date
        'strSql holds the SQL that finds all the order(s) placed since a particular date
        strSQL = "SELECT Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price] " & _
                    "FROM ([Equipment Type] INNER JOIN ([Order] INNER JOIN [Order/Equipment] ON Order.Order_ID = [Order/Equipment].Order_ID) ON [Equipment Type].Type_ID = [Order/Equipment].Type_ID) INNER JOIN Equipment ON [Equipment Type].Type_ID = Equipment.Type_ID " & _
                    "GROUP BY Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price], Equipment.Supplier_ID, Order.CancelOrder " & _
                    "Having (((Order.DateOrdered) >= #" & Format(txtList.Text, "MM, DD,YYYY") & "#) And ((Equipment.Supplier_ID) = [Order].[Supplier_ID]) And ((Order.CancelOrder) <> True)) " & _
                    "ORDER BY " & strSort & ""
        'strSql2 holds the SQL that finds the cost of order(s) placed since a particular date
        '"OrderReport" is a query table produced from the query thats in strSql
        strSql2 = "SELECT OrderReport.Order_ID, Sum([Purchase Price]*[QuantityOrdered]) AS [Order Cost] " & _
                    "From OrderReport " & _
                    "GROUP BY OrderReport.Order_ID, OrderReport.Supplier_ID, OrderReport.DateOrdered " & _
                    "Having (((OrderReport.DateOrdered) >= #" & Format(txtList.Text, "MM, DD, YYYY") & "#)) " & _
                    "ORDER BY " & strSort & ""
    Else
        MsgBox "Please enter a valid date", vbExclamation, "Invalid date"
        txtList.Text = ""
    End If
    txtList.SetFocus
ElseIf optSupplier.Value = True Then        'List by supplier
    If CheckNum = True Or txtList.Text = "" Then            'Call procedure CheckNum to see if the ID that the user has entered is numeric
        MsgBox "Please enter a valid supplier number", vbExclamation, "Invalid number"
        txtList.Text = ""
    Else
        'strSql holds the SQL that finds all the order(s) placed with a particular supplier
        strSQL = "SELECT Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price] " & _
                    "FROM ([Equipment Type] INNER JOIN ([Order] INNER JOIN [Order/Equipment] ON Order.Order_ID = [Order/Equipment].Order_ID) ON [Equipment Type].Type_ID = [Order/Equipment].Type_ID) INNER JOIN Equipment ON [Equipment Type].Type_ID = Equipment.Type_ID " & _
                    "GROUP BY Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price], Equipment.Supplier_ID, Order.CancelOrder " & _
                    "Having (((Order.Supplier_ID) = " & txtList.Text & ") And ((Equipment.Supplier_ID) = [Order].[Supplier_ID]) And ((Order.CancelOrder) <> True)) " & _
                    "ORDER BY " & strSort & ""
        'strSql2 holds the SQL that finds the cost each order placed with a particular supplier
        '"OrderReport" is a query table produced from the query thats in strSql
        strSql2 = "SELECT OrderReport.Order_ID, Sum([Purchase Price]*[QuantityOrdered]) AS [Order Cost]" & _
                    "From OrderReport " & _
                    "GROUP BY OrderReport.Order_ID, OrderReport.Supplier_ID " & _
                    "Having (((OrderReport.Supplier_ID) = " & txtList.Text & ")) " & _
                    "ORDER BY OrderReport.Order_ID;"
        datOrder.RecordSource = strSQL
        datOrder.Refresh
        If datOrder.Recordset.RecordCount = 0 Then
            MsgBox "There is no match", vbExclamation, "No Match"
        End If
    End If
    txtList.SetFocus
Else
    'strSql holds the SQL that finds all the order(s)
    strSQL = "SELECT Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price] " & _
                "FROM ([Equipment Type] INNER JOIN ([Order] INNER JOIN [Order/Equipment] ON Order.Order_ID = [Order/Equipment].Order_ID) ON [Equipment Type].Type_ID = [Order/Equipment].Type_ID) INNER JOIN Equipment ON [Equipment Type].Type_ID = Equipment.Type_ID " & _
                "GROUP BY Order.Order_ID, Order.Supplier_ID, [Equipment Type].Type_ID, Order.DateOrdered, Order.DateDelivered, [Equipment Type].Description, [Equipment Type].Make, [Equipment Type].Model, [Equipment Type].Details, [Order/Equipment].QuantityOrdered, Equipment.[Purchase Price], Equipment.Supplier_ID, Order.CancelOrder " & _
                "Having (((Equipment.Supplier_ID) = [Order].[Supplier_ID]) And ((Order.CancelOrder) <> True)) " & _
                "ORDER BY " & strSort & ""
    'strSql2 holds the SQL that finds the cost each order
    '"OrderReport" is a query table produced from the query thats in strSql
    strSql2 = "SELECT OrderReport.Order_ID, Sum([Purchase Price]*[QuantityOrdered]) AS [Order Cost] " & _
                "From OrderReport " & _
                "GROUP BY OrderReport.Order_ID " & _
                "ORDER BY OrderReport.Order_ID "
End If
datOrder.RecordSource = strSQL
datOrder.Refresh
datTotal.RecordSource = strSql2
datTotal.Refresh
End Sub

Private Sub cmdMM_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub optAll_Click()
lblList.Visible = False
txtList.Visible = False
cmdBrowse.Visible = False
End Sub

Private Sub optDate_Click()
cmdBrowse.Visible = False
txtList.ToolTipText = "Enter date"
lblList.Visible = True
txtList.Visible = True
txtList.SetFocus
txtList.Text = ""
lblList.Caption = "Enter Date"
End Sub

Private Sub optSupplier_Click()
cmdBrowse.Visible = True
txtList.ToolTipText = "Enter ID"
lblList.Visible = True
txtList.Visible = True
txtList.SetFocus
txtList.Text = ""
lblList.Caption = "Enter ID"
End Sub

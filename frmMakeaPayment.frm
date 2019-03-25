VERSION 5.00
Begin VB.Form frmMakeaPayment 
   Caption         =   "Make a Payment"
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
   Begin VB.Frame Frame3 
      Caption         =   "Add Payment"
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
      Height          =   1095
      Left            =   1560
      TabIndex        =   38
      Top             =   8520
      Width           =   11895
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   9480
         TabIndex        =   8
         ToolTipText     =   "Click to add Payment"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboType 
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
         Left            =   5400
         TabIndex        =   0
         ToolTipText     =   "Enter payment type"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtPay 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """€""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   6153
            SubFormatType   =   2
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
         Left            =   2640
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Enter a payment"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Enter Payment"
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
         TabIndex        =   40
         Top             =   420
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Details"
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
      Height          =   3975
      Left            =   1560
      TabIndex        =   15
      Top             =   4320
      Width           =   11895
      Begin VB.TextBox txtStatus 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtBal 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtCredit 
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtName 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtID 
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtMob 
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1935
      End
      Begin VB.TextBox txtMail 
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtPhone 
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
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   885
         Width           =   1935
      End
      Begin VB.TextBox txtAdd3 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtAdd2 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1935
      End
      Begin VB.TextBox txtAdd1 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   885
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Left            =   4320
         TabIndex        =   37
         Top             =   3420
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   35
         Top             =   2820
         Width           =   810
      End
      Begin VB.Label Label12 
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
         Left            =   6240
         TabIndex        =   34
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Supplier ID"
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
         Left            =   6240
         TabIndex        =   32
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label15 
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
         Left            =   6240
         TabIndex        =   31
         Top             =   1560
         Width           =   1590
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "E - Mail"
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
         Left            =   6240
         TabIndex        =   30
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label Label17 
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
         Left            =   6240
         TabIndex        =   29
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Address 3"
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
         TabIndex        =   28
         Top             =   2220
         Width           =   1005
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Address 2"
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
         TabIndex        =   27
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Address 1"
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
         TabIndex        =   26
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search/Browse"
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
      Height          =   2775
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   11895
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Previous"
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
         Left            =   7800
         TabIndex        =   5
         ToolTipText     =   "Click to display previous match"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
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
         Left            =   9840
         TabIndex        =   6
         ToolTipText     =   "Click to display next match"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
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
         Left            =   9000
         TabIndex        =   7
         ToolTipText     =   "Click to browse"
         Top             =   1440
         Width           =   1575
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
         Left            =   5040
         TabIndex        =   4
         ToolTipText     =   "Click to show details "
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton OptName 
         Caption         =   "Search by &Name"
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
         Left            =   4080
         TabIndex        =   2
         ToolTipText     =   "Select to search by name"
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optID 
         Caption         =   "Search by &ID"
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
         Left            =   6120
         TabIndex        =   1
         ToolTipText     =   "Select to search by ID"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         ToolTipText     =   "Enter Name or ID"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblSearch 
         Caption         =   "Enter Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Data datCustomer 
      Caption         =   "Customer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   10440
      TabIndex        =   10
      ToolTipText     =   "Click to exit screen and ignore payments"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      TabIndex        =   9
      ToolTipText     =   "Click to exit screen "
      Top             =   9840
      Width           =   1935
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
      Left            =   12960
      TabIndex        =   11
      ToolTipText     =   "Go to Main Menu"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Make a Payment"
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
      Left            =   5175
      TabIndex        =   12
      Top             =   240
      Width           =   5205
   End
End
Attribute VB_Name = "frmMakeaPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to allow the user to make a payment.

'The user enters either the customers name or primary key, or they can browse for the
'customer.
'When the user clicks "Show" the system finds the matching criteria in the customer
'table using the data control "datCustomer" and prints the customers details into text boxes.
'The user then enters the payment amount and selects the payment method from a combo box.
'When the user clicks "Add" the customers balance is updated and the payment details
'are stored in tags.
'If the user clicks "Cancel" the customers balance is stored to its original value.
'If the user clicks "Ok" the payment table is updated

'Author Fergal Purcell
'Date   06/03/2002

'Variables used
'strSql     :   String ; This variable stores the Sql statement.
'intSpace   :   Integer; Holds the position of the space as it marks the end of one primary key.
'intCounter :   Integer; This variable is used as a control variable in a For Loop
'intPayID   :   Integer; Stores the system generated payment ID

'Objects used
'datCustomer    :   Data Control; This data control retrieves information from the "Customer" table and updates their balance owed

Option Explicit
Private Function CheckNum(ByVal strBox) As Boolean
'Bug in IsNumeric says that 6d6 is numeric
Dim intCounter As Integer
    For intCounter = 1 To Len(strBox)
        If Not IsNumeric(Mid(strBox, intCounter, 1)) Then        'Check each character
            CheckNum = True
        End If
    Next
End Function
Private Sub cmdAdd_Click()
'This procedure updates the customers balance if he/she makes a payment, but it doesnt update
'the "Payment" table
Dim strSQL As String
If txtPay.Text = "" Then
    MsgBox "Please enter a payment amount", vbExclamation, "Enter amount"
ElseIf CheckNum(txtPay.Text) = True Then
    MsgBox "Please enter a valid payment amount", vbExclamation, "Invalid data"
ElseIf cboType.ListIndex = -1 Then
    MsgBox "Please select a payment type", vbExclamation, "Enter type"
Else
    datCustomer.DatabaseName = strThePath
    datCustomer.RecordSource = "Customer"
    datCustomer.Refresh
    datCustomer.Recordset.FindFirst "Cust_ID = " & txtID.Text
    txtPay.Tag = txtPay.Tag & txtPay.Text & Chr(9)              'Store payments in case the user makes a mistake
    cmdAdd.Tag = cmdAdd.Tag & txtID.Text & Chr(9)               'Store the customers ID
    cboType.Tag = cboType.Tag & cboType.Text & Chr(9)           'Store the type of payment
    txtStatus.Tag = txtStatus.Tag & txtStatus.Text & Chr(9)
    datCustomer.Recordset.Edit
    datCustomer.Recordset.Fields("Balance Owed") = datCustomer.Recordset.Fields("Balance owed") - txtPay.Text    'Update the Customers balance owed
    If datCustomer.Recordset.Fields("Balance owed") <= datCustomer.Recordset.Fields("Credit Limit") And datCustomer.Recordset.Fields("Balance owed") > 0 Then
        datCustomer.Recordset.Fields("Status") = "Normal"       'Change the customers status
        txtStatus.Text = "Normal"
    End If
    datCustomer.Recordset.Update
    txtBal.Text = datCustomer.Recordset.Fields("Balance owed")
End If
End Sub

Private Sub cmdCancel_Click()
If cmdAdd.Tag <> "" Then            'Display this message only if the user has added payments
    If MsgBox("Do you wish to save payments?", vbQuestion + vbYesNo, "Save payments") = vbNo Then
        CancelPay           'Call procedure that cancels the payments that a customer has made
    Else
        OkPay               'Call procedure that adds payments
    End If
End If
If frmMakeaPayment.Tag = "" Then
    frmMainMenu.Show
End If
Unload Me
End Sub

Private Sub cmdMM_Click()
If cmdAdd.Tag <> "" Then            'Display this message only if the user has added payments
    If MsgBox("Do you wish to save payments?", vbQuestion + vbYesNo, "Save Payments") = vbNo Then
        CancelPay
    Else
        OkPay
    End If
End If
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdNext_Click()
datCustomer.Recordset.MoveNext
If datCustomer.Recordset.EOF Then
    datCustomer.Recordset.MoveFirst
End If
AssignFields
End Sub

Private Sub cmdOk_Click()
OkPay
End Sub

Private Sub cmdPrevious_Click()
datCustomer.Recordset.MovePrevious         'Move to the previous match
If datCustomer.Recordset.BOF Then
    datCustomer.Recordset.MoveLast
End If
AssignFields
End Sub

Private Sub cmdShow_Click()
ShowDetails
End Sub

Private Sub cmdBrowse_Click()
cmdNext.Visible = False
cmdPrevious.Visible = False
optID.Value = True              'Set the search type to ID search since browse returns the primary key
lblSearch.Caption = "Enter ID"              'Change the labels
frmBrowse.Show                  'Display the browse screen
frmBrowse.Tag = "Customer"
End Sub

Private Sub Form_Activate()
If frmMakeaPayment.Tag <> "" Then               'Check to see if the form has been activated from "Hiring of Equipment" or "Return of Equipment"
    optID.Value = True
    cmdMM.Enabled = False
    cmdBrowse.Enabled = False
    txtSearch.Text = frmMakeaPayment.Tag        'frmMakeaPayment.Tag holds the customers ID
    ShowDetails
End If
txtSearch.SetFocus
cboType.Clear
cboType.AddItem ("Cash")                'Add payment methods to the combo box
cboType.AddItem ("Cheque")
cboType.AddItem ("Credit Card")
cboType.Text = "Enter Payment type"
End Sub

Private Sub OptID_Click()
lblSearch.Caption = "Enter ID"              'Change the labels
txtSearch.SetFocus
txtSearch.Text = ""
End Sub

Private Sub OptName_Click()
lblSearch.Caption = "Enter Name"
txtSearch.Text = ""
txtSearch.SetFocus
End Sub

Private Sub AssignFields()
'This procedure assigns the text boxes
txtName.Text = datCustomer.Recordset.Fields("Name")
txtID.Text = datCustomer.Recordset.Fields("Cust_ID")
txtAdd1.Text = datCustomer.Recordset.Fields("[Address 1]")
txtAdd2.Text = datCustomer.Recordset.Fields("[Address 2]")
txtAdd3.Text = datCustomer.Recordset.Fields("[Address 3]")
txtPhone.Text = datCustomer.Recordset.Fields("[Phone No]")
txtMob.Text = "" + datCustomer.Recordset.Fields("[Mobile No]")
txtMail.Text = "" + datCustomer.Recordset.Fields("E-Mail")
txtBal.Text = datCustomer.Recordset.Fields("Balance owed")
txtCredit.Text = datCustomer.Recordset.Fields("Credit limit")
txtStatus.Text = datCustomer.Recordset.Fields("Status")
cmdAdd.Enabled = True
End Sub

Private Sub ShowDetails()
'This procedure searches for the customer and displays there details if found
cmdNext.Visible = False
cmdPrevious.Visible = False
datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datCustomer.Refresh
If optName = True Then                  'Search for matching name
    Dim strSQL As String
    strSQL = "Select Count(Customer.Name) AS CountOfSup From Customer Where Customer.[Name] Like '" & txtSearch.Text & "*'" & _
                "And Customer.Deletion = false"
    datCustomer.RecordSource = strSQL
    datCustomer.Refresh
    If datCustomer.Recordset.Fields("CountOfSup") = 0 Or txtSearch = "" Then
        MsgBox "There is no match.", vbExclamation, "No Match"
        txtSearch.SetFocus
        txtSearch.Text = ""
    ElseIf datCustomer.Recordset.Fields("CountOfSup") = 1 Then
        strSQL = "Select * From Customer Where Customer.[Name] Like '" & txtSearch.Text & "*'" & _
                "And Customer.Deletion = False"
        datCustomer.RecordSource = strSQL
        datCustomer.Refresh
        AssignFields
    Else
        cmdNext.Visible = True              'Display these buttons if there is more than one match
        cmdPrevious.Visible = True
        strSQL = "Select * From Customer Where Customer.[Name] Like '" & txtSearch.Text & "*'" & _
                "And Customer.Deletion = False"
        datCustomer.RecordSource = strSQL
        datCustomer.Refresh
        AssignFields
    End If
Else                                                                'Search for primary keys
    If CheckNum(txtSearch.Text) = True Or txtSearch.Text = "" Then
        MsgBox "Invalid data.", vbExclamation, "Invalid data"
        txtSearch.SetFocus
        txtSearch.Text = ""
    Else
        datCustomer.Recordset.FindFirst "Cust_ID = " & txtSearch.Text
        If datCustomer.Recordset.NoMatch = True Or datCustomer.Recordset.Fields("Deletion") = True Then
            MsgBox "There is no match.", vbExclamation, "Invalid data"
            txtSearch.SetFocus
            txtSearch.Text = ""
        Else
            AssignFields
        End If
    End If
End If
End Sub

Private Sub CancelPay()
'This procedure cancels all the payments added to the customers balance
Dim intTab As Integer, intTab2 As Integer, intTab3 As Integer
intTab = 1
intTab2 = 1
intTab3 = 1
Do Until cmdAdd.Tag = ""                        'Restore the customers balance if the user presses cancel
    intTab = InStr(cmdAdd.Tag, vbTab)           'Find the location of the space because it marks the end of one primary key
    txtID.Text = Mid(cmdAdd.Tag, 1, intTab - 1)    'Find the primary key
    datCustomer.DatabaseName = strThePath
    datCustomer.RecordSource = "Customer"
    datCustomer.Refresh
    datCustomer.Recordset.FindFirst "Cust_ID = " & txtID.Text   'Find matching key field
    datCustomer.Recordset.Edit
    intTab2 = InStr(txtPay.Tag, vbTab)           'Find the location of the tab because it marks the end of every payment
    intTab3 = InStr(txtStatus.Tag, vbTab)
    datCustomer.Recordset.Fields("Balance owed") = (datCustomer.Recordset.Fields("Balance owed")) + (Mid(txtPay.Tag, 1, intTab2 - 1)) 'Restore the balance to the original settings
    datCustomer.Recordset.Fields("Status") = Mid(txtStatus.Tag, 1, intTab3 - 1)
    datCustomer.Recordset.Update
    If Len(cmdAdd.Tag) <= 2 Then
        cmdAdd.Tag = ""
    Else
        cmdAdd.Tag = Mid(cmdAdd.Tag, intTab + 1, (Len(cmdAdd.Tag) - (intTab - 1)))
        txtPay.Tag = Mid(txtPay.Tag, intTab2 + 1, (Len(txtPay.Tag) - (intTab2 - 1)))
        txtStatus.Tag = Mid(txtStatus.Tag, intTab3 + 1, (Len(txtStatus.Tag) - (intTab3 - 1)))
    End If
Loop
End Sub

Private Sub OkPay()
'This procedure updates the payment table with all the payments added
Dim intTab  As Integer, intTab2 As Integer, intTab3 As Integer, intPayID As Integer
intTab = 1
intTab2 = 1
intTab3 = 1
Do Until cmdAdd.Tag = ""
    intTab = InStr(cmdAdd.Tag, vbTab)           'Find the location of the space because it marks the end of one primary key
    txtID.Text = Mid(cmdAdd.Tag, 1, intTab - 1)    'Find the primary key
    datCustomer.DatabaseName = strThePath
    datCustomer.RecordSource = "Payment"
    datCustomer.Refresh
    If datCustomer.Recordset.RecordCount <> 0 Then
        datCustomer.Recordset.MoveLast
        intPayID = datCustomer.Recordset.Fields("Payment_ID") + 1       'Create a new payment ID
    Else
        intPayID = 1
    End If
    datCustomer.Recordset.AddNew
    intTab2 = InStr(txtPay.Tag, vbTab)           'Find the location of the space because it marks the end of one primary key
    intTab3 = InStr(cboType.Tag, vbTab)
    With datCustomer.Recordset
        .Fields("Payment_ID") = intPayID
        .Fields("Cust_ID") = txtID.Text
        .Fields("Date") = Date
        .Fields("Amount") = Mid(txtPay.Tag, 1, intTab2 - 1)             'txtPay.Tag holds the payment amount
        .Fields("Payment Method") = Mid(cboType.Tag, 1, intTab3 - 1)    'cboType.Tag holds the payment type
        datCustomer.Recordset.Update
    End With
    If Len(cmdAdd.Tag) <= 2 Then
        cmdAdd.Tag = ""
    Else
        cmdAdd.Tag = Mid(cmdAdd.Tag, intTab + 1, (Len(cmdAdd.Tag) - (intTab - 1)))          'Delete first ID since it has been processed
        txtPay.Tag = Mid(txtPay.Tag, intTab2 + 1, (Len(txtPay.Tag) - (intTab2 - 1)))        'Delete first payment amount since it has been processed
        cboType.Tag = Mid(cboType.Tag, intTab3 + 1, (Len(cboType.Tag) - (intTab3 - 1)))     'Delete first payment type since it has been processed
    End If
Loop
If frmMakeaPayment.Tag = "" Then                'Go back to main menu or else back to rental or return
    frmMainMenu.Show
End If
Unload Me
End Sub

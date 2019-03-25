VERSION 5.00
Begin VB.Form frmDeleteCustomer 
   Caption         =   "Delete Customer"
   ClientHeight    =   8490
   ClientLeft      =   1260
   ClientTop       =   -135
   ClientWidth     =   10500
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
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCustomer 
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
      ItemData        =   "frmDeleteCustomer.frx":0000
      Left            =   3240
      List            =   "frmDeleteCustomer.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   32
      Top             =   2880
      Width           =   4335
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
      Height          =   315
      Left            =   11160
      TabIndex        =   31
      Top             =   2880
      Width           =   1455
   End
   Begin VB.OptionButton optCustID 
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
      Left            =   8640
      TabIndex        =   27
      Top             =   2280
      Width           =   2535
   End
   Begin VB.OptionButton optCustName 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      Top             =   2280
      Width           =   2895
   End
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
      Height          =   495
      Left            =   11400
      TabIndex        =   25
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton Cmdcancel 
      Caption         =   "Cancel"
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
      TabIndex        =   24
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelCust 
      Caption         =   "Delete"
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
      Left            =   9720
      TabIndex        =   23
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Data datRentalReturn 
      Caption         =   "Data1"
      Connect         =   ";Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
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
      Left            =   11040
      TabIndex        =   22
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Data datCustomer 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
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
      Left            =   4920
      TabIndex        =   11
      Top             =   7560
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   10
      Top             =   6840
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   9
      Top             =   6360
      Width           =   2535
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
      Left            =   4920
      TabIndex        =   8
      Top             =   6360
      Width           =   2535
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
      Left            =   4920
      TabIndex        =   7
      Top             =   6960
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   6
      Top             =   5760
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   5
      Top             =   7440
      Width           =   2535
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
      Left            =   10560
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
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
      Left            =   4920
      TabIndex        =   3
      Top             =   5760
      Width           =   2535
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
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
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
      Left            =   13200
      TabIndex        =   1
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Search Options"
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
      Height          =   3015
      Left            =   2520
      TabIndex        =   28
      Top             =   1800
      Width           =   10815
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   5520
         TabIndex        =   34
         Top             =   240
         Width           =   5055
         Begin VB.Label lblID 
            Caption         =   "Enter ID"
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
            Left            =   1800
            TabIndex        =   35
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2535
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Customer ID"
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
         Left            =   7080
         TabIndex        =   30
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   9480
      TabIndex        =   36
      Top             =   8280
      Width           =   4095
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
      Left            =   4440
      TabIndex        =   29
      Top             =   2640
      Width           =   1500
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
      Left            =   8400
      TabIndex        =   21
      Top             =   6480
      Width           =   630
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
      Left            =   2640
      TabIndex        =   20
      Top             =   7080
      Width           =   945
   End
   Begin VB.Label LblCustName 
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
      Left            =   2640
      TabIndex        =   19
      Top             =   5280
      Width           =   1635
   End
   Begin VB.Label LblAddress2 
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
      Left            =   2640
      TabIndex        =   18
      Top             =   6480
      Width           =   945
   End
   Begin VB.Label lblPhonenumber 
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
      Left            =   2640
      TabIndex        =   17
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label LblCreditlimit 
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
      Left            =   8400
      TabIndex        =   16
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblBalanceOwned 
      AutoSize        =   -1  'True
      Caption         =   "Balance Owned"
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
      Left            =   8400
      TabIndex        =   15
      Top             =   7560
      Width           =   1560
   End
   Begin VB.Label lblmobile 
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
      Left            =   8400
      TabIndex        =   14
      Top             =   5280
      Width           =   1080
   End
   Begin VB.Label Lbladdress1 
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
      Left            =   2640
      TabIndex        =   13
      Top             =   5880
      Width           =   945
   End
   Begin VB.Label lblemail 
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
      Left            =   8400
      TabIndex        =   12
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Delete Customer"
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
      Left            =   5205
      TabIndex        =   0
      Top             =   240
      Width           =   5145
   End
End
Attribute VB_Name = "frmDeleteCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose of screen:

'This screen allows the user to delete an existing customer from the database
'the user can select the customer from a list or enter in the customer id number,
'when a customer is selected the rest of the customer details are displayed
'on the screen,the user can delete the cutomer from the screen by clicking the
'"Delete" button or they can cancel'the deletion or return to the main menu without
'deleting,if the user presses the "cancel" button all the deletios are undeleted

'Author Derek Stafford
'Date   015/03/2002


'varibales used                 Purpose
'intMaxoption        integer;     sores a count for exception deletion
'intCustID           integer;     store customer id number
'intCounter          integer;     stores counter in for loop for entrycode to exception deletion
'intResult           integer;     stores result of an answer
'strmySQl            string;      store sql string
'strTempName         string;      stores name of customer
'strTempAdd1         string;      stores address1 of customer
'strTempAdd2         string;      stores address2 of customer
'Delete              boolean;     stores true/false value
'Enterycode          const;       stores the exception entry code
'ainCustID   Array of Integers hold all deleted customer ids



'Objects used

'datCustomer;        Datacontrol          retrieves data from customer table
'datRentalReturn;    Datacontrol          retrieves data from rental/return table
















Option Explicit
Dim intCustID, intindex, aintCustIDs(1 To 30) As Integer
Private Sub cmdCancel_Click()
'find the last record entered into the cutomer table and deletes it
Dim intAmount, intResult, intDelCust As Integer
    
datCustomer.RecordSource = "Customer"
datCustomer.Refresh
intResult = MsgBox("Are you sure you want to undelete the " & intindex - 1 & " customer(s) you deleted ", vbYesNo, "System response")
If intResult = vbYes Then
    For intAmount = 1 To intindex
        intDelCust = aintCustIDs(intAmount)
        datCustomer.Recordset.FindFirst "Cust_ID = " & intDelCust & ""
        datCustomer.Recordset.Edit
        datCustomer.Recordset("Deletion") = "False"
        datCustomer.Recordset.Update
        datCustomer.Recordset.MoveNext
    Next intAmount
intindex = 1
Cmdcancel.Enabled = False
End If
End Sub

Private Sub cmdDisplay_Click()
'searchs for the a customer based on the customer id entered if found it calls display sub
'if not found it returns an error
Dim strmySQl, Result As String

intCustID = Val(txtCustNumber)

strmySQl = "SELECT Customer.Cust_ID, Customer.Name,Customer.[Address 1], Customer.[Address 2], Customer.Deletion " & _
"From Customer " & _
"WHERE (((Customer.Deletion)=False));"
datCustomer.RecordSource = strmySQl
datCustomer.Refresh
datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
If datCustomer.Recordset.NoMatch Then
    Call MsgBox("This customer migth not be listed or is already deleted", vbInformation, "Not Found")
Else
    display
Cmdcancel.Enabled = True
End If
End Sub



Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

Private Sub txtCustNumber_change()
cmdDisplay.Enabled = True
End Sub
Private Sub txtCustNumber_Validate(Cancel As Boolean)
'check valid id was entered

If IsNumeric(txtCustNumber.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtCustNumber.SetFocus
    Cancel = True
End If
End Sub
Private Sub optCustName_Click()
'resets screen
lblID.Visible = False
cmdDisplay.Enabled = False
txtCustNumber = ""
CmbCustomer.Enabled = True
txtCustNumber.Enabled = False
Clear
End Sub
Private Sub Clear()
'empties all textboxes
txtCustomerName = ""
txtAddress1 = ""
txtAddress2 = ""
txtAddress3 = ""
txtPhoneNo = ""
txtMobileNo = ""
txtEmail = ""
txtStatus = ""
txtCreditLimit = ""
txtBalanceowned = ""
End Sub
Private Sub optCustID_Click()
'resets screen

txtCustNumber = ""
lblID.Visible = True
CmbCustomer.Enabled = False
txtCustNumber.Enabled = True
txtCustNumber.SetFocus
CmbCustomer.Text = ""
Clear
End Sub
Private Sub cmbCustomer_Click()
'stores the cutomer id selected
'Dim intID As Integer
intCustID = CmbCustomer.ItemData(CmbCustomer.ListIndex)
datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
cmdDelCust.Enabled = True
display
End Sub

Private Sub cmbCustomer_DropDown()
'fills the customer combo box based on result of SQL
Dim strmySQl As String
strmySQl = "SELECT Customer.Cust_ID, Customer.Name,Customer.[Address 1], Customer.[Address 2], Customer.Deletion " & _
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

Private Sub cmdDelCust_Click()
'search for the selected customer record and finds out if that customer can be
'deleted or not based on having equipment still on hire or having an outstanding
'balance

Dim ccurID As Currency, intReturnValue As Integer, BalanceOwed As Currency
Dim strmySQl, strTempCustName, strTempAdd1, strTempAdd2 As String, Delete As Boolean
Delete = False

strmySQl = "SELECT Customer.Cust_ID, Customer.Name, Customer.[Address 1],Customer.[Address 2], [Rental/Return].[Date/Time returned],Customer.[Balance owed], [Rental/Return].Equipment_ID, Customer.Deletion " & _
"From Customer, [Rental/Return]" & _
"WHERE (((Customer.Cust_ID)= " & intCustID & "));"

datCustomer.RecordSource = strmySQl
datCustomer.Refresh
datCustomer.Recordset.MoveFirst
datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""

'checks if customer has an outstanding balance or still has equipment on hire
If datCustomer.Recordset("Balance Owed") > 0 Or datCustomer.Recordset("Date/Time returned") = "" Then
    strTempCustName = datCustomer.Recordset("Name")
    strTempAdd1 = datCustomer.Recordset("Address 1")
    strTempAdd2 = datCustomer.Recordset("Address 2")
    If datCustomer.Recordset("Deletion") = True Then
        Call MsgBox(strTempCustName & " From " & strTempAdd1 & ", " & strTempAdd2 & " Has Already Been Deleted ", vbInformation, "Information")
    Else
        Call MsgBox(strTempCustName & " From " & strTempAdd1 & ", " & strTempAdd2 & " Cannot Be Deleted As He/She Still Has An Equipment On Hire OR Has An OutStanding Balance ", vbInformation, "Information")
        intReturnValue = MsgBox("If You Wish To Delete " & strTempCustName & " From " & strTempAdd1 & ", " & strTempAdd2 & " Anyway Press Yes ", vbYesNo, "Exception Deletion")
        If intReturnValue = vbYes Then
            ExceptionDeletion
             Cmdcancel.Enabled = True
        End If
    End If
Else
    While Not datCustomer.Recordset.EOF
    datCustomer.Recordset.FindNext "Cust_ID = " & intCustID & ""
        If datCustomer.Recordset("Date/Time returned") = "" Then
            If datCustomer.Recordset("Deletion") = False Then
                Delete = True
            End If
        End If
        datCustomer.Recordset.MoveNext
    Wend
    If Delete Then
        Call MsgBox(strTempCustName & " From " & strTempAdd1 & "," & strTempAdd2 & " Has Already Been Deleted", vbInformation, "Information")
    Else
        Cmdcancel.Enabled = True
        Deletion
    End If
End If
End Sub

Private Sub ExceptionDeletion()
'allows the user to make exception deletions
Dim intCounter, intMaxoption, intResult As Integer
Const intEntrycode As Integer = 1111, intMax As Integer = 3
For intCounter = 1 To intMax
    If intEntrycode = Val(InputBox("Please Enter Security Code", "Security")) Then
        intCounter = intMax
        Deletion
    Else
        Call MsgBox("Incorrect Code! You Have " & intMax - intCounter & " Changes Left", vbCritical, "System Response")
        intMaxoption = intMax - intCounter
        If intMaxoption > 0 Then
            intResult = MsgBox("Do You Want To Try Again ", vbYesNo, "System Response")
            If intResult = vbNo Then
                intCounter = intMax
            End If
        End If
    End If
Next intCounter
End Sub

Private Sub cmdBack_Click()
Dim intAmount, intResult, intDelCust As Integer
'asks user to confirm deletion and shows another form
datCustomer.RecordSource = "Customer"
datCustomer.Refresh
If intindex > 1 Then
    intResult = MsgBox("Are you Sure you want to  save the deletions made ", vbYesNo, "System response")
    If intResult = vbNo Then
        For intAmount = 1 To intindex
            intDelCust = aintCustIDs(intAmount)
            datCustomer.Recordset.FindFirst "Cust_ID = " & intDelCust & ""
            datCustomer.Recordset.Edit
            datCustomer.Recordset("Deletion") = "False"
            datCustomer.Recordset.Update
            datCustomer.Recordset.MoveNext
        Next intAmount
        intindex = 1
    End If
End If
frmCustomerFileProcessing.Show
Unload Me
End Sub

Private Sub Deletion()
'this procedure deletes the selected customer from the customer table
Dim strTempCustName, strTempAdd1, strTempAdd2 As String
Dim intReturnValue, integ As Integer

datCustomer.RecordSource = "Customer"
datCustomer.Refresh
datCustomer.Recordset.FindFirst "Cust_ID = " & intCustID & ""
strTempCustName = datCustomer.Recordset("Name")
strTempAdd1 = datCustomer.Recordset("Address 1")
strTempAdd2 = datCustomer.Recordset("Address 2")
intReturnValue = MsgBox("Are You Sure You Want To Delete" & strTempCustName & " From " & strTempAdd1 & "," & strTempAdd2 & " From The System ", vbYesNo, "Response Required")
    If intReturnValue = vbYes Then
        datCustomer.Recordset.Edit
        datCustomer.Recordset("Deletion") = "True" 'marks the customer for deletion
        datCustomer.Recordset.Update
        Call MsgBox(strTempCustName & ", From " & strTempAdd1 & ", " & strTempAdd2 & " Has Sucessfully Been Deleted", vbInformation, "Sucessful Deletion")
        aintCustIDs(intindex) = intCustID
        intindex = intindex + 1
    End If
End Sub
Private Sub Form_Activate()
'sets screen when form is activated
lblID.Visible = False
CmbCustomer.Enabled = False
txtCustNumber.Enabled = False
cmdDelCust.Enabled = False
cmdDisplay.Enabled = False
Cmdcancel.Enabled = False
intindex = 1
End Sub
Private Sub Form_Load()
'assigns tables to data controls

datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datRentalReturn.RecordSource = strThePath
datRentalReturn.RecordSource = "Rental/Return"

End Sub

Private Sub display()
'fill all the textboxes with the selected customer details
Dim AmountDue, Cost, ChargeType As Currency
datCustomer.RecordSource = "Customer"
datCustomer.Refresh
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
Private Sub txtCustomerName_Change()
    
    cmdDelCust.Enabled = True
End Sub

VERSION 5.00
Begin VB.Form frmPlaceanOrder 
   Caption         =   "Place an order"
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
   Begin VB.Data datOrder 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order List"
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
      Height          =   4575
      Left            =   480
      TabIndex        =   18
      Top             =   4800
      Width           =   14415
      Begin VB.ListBox lstOrder 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2535
         ItemData        =   "frmPlaceanOrder.frx":0000
         Left            =   240
         List            =   "frmPlaceanOrder.frx":0002
         TabIndex        =   3
         ToolTipText     =   "Order List"
         Top             =   840
         Width           =   13935
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove from List"
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
         Left            =   7440
         TabIndex        =   5
         ToolTipText     =   "Click to remove from list"
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton cmdCAR 
         Caption         =   "&Change Order Qty"
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
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Click to change amount required"
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Supp ID"
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
         Left            =   2880
         TabIndex        =   28
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label17 
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
         Left            =   11400
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Order Qty"
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
         Left            =   12840
         TabIndex        =   26
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label15 
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
         Left            =   9720
         TabIndex        =   25
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label14 
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
         Left            =   4320
         TabIndex        =   24
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label13 
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
         Left            =   6240
         TabIndex        =   23
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label12 
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
         Left            =   8160
         TabIndex        =   22
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Type ID"
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
         TabIndex        =   21
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         TabIndex        =   20
         Top             =   480
         Width           =   1155
      End
   End
   Begin VB.Data datSQL 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   13440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data datEquipment 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Equipment"
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
      Height          =   3135
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   14415
      Begin VB.CommandButton cmdAdd 
         Caption         =   "O&rder"
         Default         =   -1  'True
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
         Left            =   9000
         TabIndex        =   2
         ToolTipText     =   "Click to add to order list"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4080
         TabIndex        =   1
         Text            =   "Select a Supplier"
         ToolTipText     =   "Select supplier"
         Top             =   2400
         Width           =   4095
      End
      Begin VB.ListBox lstEquip 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   960
         ItemData        =   "frmPlaceanOrder.frx":0004
         Left            =   600
         List            =   "frmPlaceanOrder.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Select equipment you wish to order"
         Top             =   840
         Width           =   13215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Details"
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
         Left            =   7440
         TabIndex        =   19
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ReOrder Qty"
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
         Left            =   12360
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Make"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Model"
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
         Left            =   5400
         TabIndex        =   14
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type ID"
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
         TabIndex        =   13
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Current Lvl"
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
         Left            =   9000
         TabIndex        =   12
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ReOrder Lvl"
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
         Left            =   10560
         TabIndex        =   11
         Top             =   480
         Width           =   1290
      End
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
      Left            =   11160
      TabIndex        =   7
      ToolTipText     =   "Click to ignore order made"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   9120
      TabIndex        =   6
      ToolTipText     =   "Click to exit screen and print orders"
      Top             =   9960
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
      Top             =   9960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Place an Order"
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
      Left            =   5445
      TabIndex        =   9
      Top             =   240
      Width           =   4665
   End
End
Attribute VB_Name = "frmPlaceanOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to allow orders to be placed with customers.

'When the form is activated the list box "lstEquip" is filled with details of every
'equipment type the business has, the equipment types ID is stored in the list boxes
'ItemData property.
'When the user selects the equipment type he wishs to order, the combo box "cboSupplier"
'is filled with names of every supplier that has supplied the equipment type and who is not
'marked for deletion.
'The user then selects the supplier he wishs to order from and clicks "Order".
'When "Order" is clicked the order is added to the list box "lstOrder".
'If the user wishes to change the order quantity he clicks "Change order Qty" and enters
'the amount required.
'If the user clicks "Remove from List" the order is removed from the list.
'When the user clicks "Ok" an order list is printed for every order and the table "Order" is
'updated with a seperate order for every supplier and the table "Order/Equipment" is updated
'with every equipment type ordered.

'Author Fergal Purcell
'Date   06/03/2002

'Variables used
'strSql         :   String ; This variable stores the Sql statement.
'intSpace       :   Integer; Holds the position of the space as it marks the end of one primary key.
'intId          :   Integer; Stores the equipment type Id.
'intOrder_ID    :   Integer; Holds a concatenation of Suppier ID and Equipment Type ID.
'intIndex       :   Integer; Control variable.
'blnExist       :   Boolean; Returns true if an order already exists.
'strRecord      :   String;  Holds the order to be changed.
'intMyPos       :   Integer; Stores to position of the last tab in strRecord.
'strNewAmt      :   Integer; The new amount required for a particular order of equipment.
'strSql         :   String;  Holds the SQL statement
'intCounter     :   Integer; Holds a count an equipment that is in stock.
'intSuppID      :   Integer; Holds the suppliers ID.
'strSuppName    :   String;  Holds the supplier name.
'strOrder       :   String;  Stores details of an order for processing.
'strOrderID     :   String;  Holds a system generated order ID.
'strSuppID      :   String;  Holds the supplier ID.
'strTemp        :   Variant; Assigned a split of the order for processing as it goes through the loop.


'Objects used
'datOrder       :   Data Control; This data control is used to add new record to the "Order" table.
'datEquipment   :   Data Control; This data control is used to retrieve information from the "Equipment Type" table and the "Supplier" table.
'datSQL         :   Data Control; This data control is used to view the results of SQL's, its also used to retrieve information from the "Supplier" table.

Option Explicit
'Use a DLL to place tab stops in a list box.
'This piece of code is taken from "Programming in Visual Basic 6.0, by Julia Case Bradley"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long
Const LB_SETTABSTOPS = &H192
Private Function CheckNum(strMyString) As Boolean
'Bug in IsNumeric says that 6d6 is numeric
Dim intCounter As Integer
    For intCounter = 1 To Len(strMyString)
        If Not IsNumeric(Mid(strMyString, intCounter, 1)) Then        'Check each character
            CheckNum = True
        End If
    Next
End Function

Private Sub cboSupplier_Click()
cmdAdd.Enabled = True
End Sub

Private Sub cmdAdd_Click()
'This procedure creates an order with the selected supplier
Dim intID As Integer, intOrder_ID As Integer, intindex As Integer, blnExist As Boolean
If cboSupplier.ListIndex = -1 Then          'Check that the user has selected a supplier
    MsgBox "Please select a supplier", vbExclamation, "Select a supplier"
Else
    intID = lstEquip.ItemData(lstEquip.ListIndex)       'Get the Type ID of the selected equipment type
    SetTabs2 lstOrder       'Set the tabs
    datEquipment.DatabaseName = strThePath
    datEquipment.RecordSource = "Equipment Type"
    datEquipment.Refresh
    datSQL.DatabaseName = strThePath
    datSQL.RecordSource = "Supplier"
    datSQL.Refresh
    datEquipment.Recordset.FindFirst "Type_ID = " & intID
    datSQL.Recordset.FindFirst "Supplier_ID = " & cboSupplier.ItemData(cboSupplier.ListIndex)   'Find the matching suppliers record for the ID
    intOrder_ID = datSQL.Recordset.Fields("Supplier_ID") & intID    'Create an order ID for every order in the list box so that only one specific order can be made
    intindex = 0
    blnExist = False
    For intindex = 0 To lstOrder.ListCount - 1              'Check to see if the order already exists
        If intOrder_ID = lstOrder.ItemData(intindex) Then
            blnExist = True
        End If
    Next intindex
    If blnExist = True Then
        MsgBox "Order already exists", vbExclamation, "Order exists"
    Else
        With datEquipment.Recordset
            lstOrder.AddItem (.Fields("Description") & Chr(9) & .Fields("Type_ID") & Chr(9) & (datSQL.Recordset.Fields("Supplier_ID")) & Chr(9) & (datSQL.Recordset.Fields("Supplier Name")) & Chr(9) & (datSQL.Recordset.Fields("Address 1")) & Chr(9) & (datSQL.Recordset.Fields("Address 2")) & Chr(9) & (datSQL.Recordset.Fields("Address 3")) & Chr(9) & (datSQL.Recordset.Fields("Phone No")) & Chr(9) & .Fields("ReOrder Quantity"))   'Add details of order
            lstOrder.ItemData(lstOrder.NewIndex) = intOrder_ID          'Add order ID
        End With
    End If
End If
End Sub
Private Sub cmdCancel_Click()
If lstOrder.ListCount <> 0 Then         'Only display this message if there are orders in the list
    If MsgBox("Do you wish to save orders?", vbQuestion + vbYesNo, "Save orders") = vbYes Then
        OkOrder
    End If
End If
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdCAR_Click()
'This procedure changes the amount required of a particular equipment type
Dim strRecord As String, intMyPos As Integer, strNewAmt As String
strRecord = lstOrder.List(lstOrder.ListIndex)       'Copy order details into string
intMyPos = InStrRev(strRecord, Chr(9))              'Find the last space in the order details
strNewAmt = InputBox("Please enter the new amount required.", "ReOrder Quantity")   'Get the new amount from the user
If CheckNum(strNewAmt) = True Or strNewAmt = "" Then            'Call function CheckNum to see if inputted value is numeric
    MsgBox "Invalid data", vbExclamation, "Invalid data"
ElseIf strNewAmt < 1 Then
        MsgBox "Invalid data, no zero or negative amounts", vbExclamation, "Invalid data"
Else
    strRecord = Mid(strRecord, 1, intMyPos)
    strRecord = strRecord + strNewAmt               'Insert new amount into order details
    lstOrder.List(lstOrder.ListIndex) = strRecord
End If
End Sub

Private Sub cmdMM_Click()
If lstOrder.ListCount <> 0 Then
    If MsgBox("Save Orders?", vbQuestion + vbYesNo, "Save orders") = vbYes Then
        OkOrder
    End If
End If
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdOk_Click()
OkOrder                         'Call procedure that prints order(s) and updates "Order" table
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdRemove_Click()
'This procedure removes an order from the list
If MsgBox("Are you sure you to remove this order from the list?", vbQuestion + vbYesNo, "Remove order") = vbYes Then
    lstOrder.RemoveItem (lstOrder.ListIndex)    'Remove order from list
End If
cmdRemove.Enabled = False
cmdCAR.Enabled = False
End Sub

Private Sub Form_Load()
Dim strSQL As String, intCounter As Integer, intID As Integer
SetTabs lstEquip        'Set the tabs in the list box
datEquipment.DatabaseName = strThePath
datEquipment.RecordSource = "Equipment Type"
datEquipment.Refresh
datSQL.DatabaseName = strThePath
While Not datEquipment.Recordset.EOF        'Display equipment types in stock
    If datEquipment.Recordset.Fields("Deletion") = False Then       'Only display the equipment types that are not deleted
        intID = datEquipment.Recordset.Fields("Type_ID")
        strSQL = "SELECT Count(Equipment.Type_ID) AS CountOfType_ID From Equipment Where Equipment.Type_ID= " & intID & ""      'Find how many equipment types are in stock
        datSQL.RecordSource = strSQL
        datSQL.Refresh
        intCounter = datSQL.Recordset.Fields("CountOfType_ID")
        If datEquipment.Recordset.Fields("ReOrder Level") > intCounter Then         'Place an "*" at the end of the string to indicate that the stock level is lower than the reorder level
            lstEquip.AddItem (datEquipment.Recordset("Description") & Chr(9) & datEquipment.Recordset("Type_ID") & Chr(9) & datEquipment.Recordset("Make") & Chr(9) & datEquipment.Recordset("Model") & Chr(9) & datEquipment.Recordset("Details") & Chr(9) & Str(intCounter) & Chr(9) & datEquipment.Recordset("[ReOrder Level]") & Chr(9) & datEquipment.Recordset("[ReOrder Quantity]") & " *")
        Else
            lstEquip.AddItem (datEquipment.Recordset("Description") & Chr(9) & datEquipment.Recordset("Type_ID") & Chr(9) & datEquipment.Recordset("Make") & Chr(9) & datEquipment.Recordset("Model") & Chr(9) & datEquipment.Recordset("Details") & Chr(9) & Str(intCounter) & Chr(9) & datEquipment.Recordset("[ReOrder Level]") & Chr(9) & datEquipment.Recordset("[ReOrder Quantity]"))
        End If
        lstEquip.ItemData(lstEquip.NewIndex) = (datEquipment.Recordset("Type_ID"))      'Store the type id in the ItemData property
    End If
    datEquipment.Recordset.MoveNext
Wend
End Sub

Private Sub SetTabs(lst As ListBox)
'Set the tab stops in a listbox
ReDim lngTabs(0 To 7) As Long           'Eight tab stops needed
Dim lngRtn As Long                      'DLL function returns a long variable
lngTabs(0) = 75                         'Twips measurement for 1st tab stop
lngTabs(1) = 120
lngTabs(2) = 180
lngTabs(3) = 260
lngTabs(4) = 340
lngTabs(5) = 400
lngTabs(6) = 460
lngRtn = SendMessage(lst.hwnd, LB_SETTABSTOPS, 7, lngTabs(0))
End Sub

Private Sub SetTabs2(lst As ListBox)
'Set the tab stops in a listbox
ReDim lngTabs(0 To 7) As Long           'Eight tab stops needed
Dim lngRtn As Long                      'DLL function returns a long variable
lngTabs(0) = 65                         'Twips measurement for 1st tab stop
lngTabs(1) = 110
lngTabs(2) = 150
lngTabs(3) = 230
lngTabs(4) = 305
lngTabs(5) = 360
lngTabs(6) = 425
lngRtn = SendMessage(lst.hwnd, LB_SETTABSTOPS, 7, lngTabs(0))
End Sub

Private Sub lstEquip_Click()
'This procedure finds all suppliers that supply the selected equipment type
Dim intID As Integer, strSQL As String, intSuppID As Integer, strSuppName As String
cboSupplier.Clear
cboSupplier.Text = "Select a Supplier"
cmdAdd.Enabled = False
intID = lstEquip.ItemData(lstEquip.ListIndex)  'Get the type ID of the selected equipment type
strSQL = "Select [Supplier/Type].Supplier_ID " & _
            "From [Supplier/Type] INNER JOIN Supplier ON [Supplier/Type].Supplier_ID = Supplier.Supplier_ID " & _
            " Where ((([Supplier/Type].Type_ID = " & intID & " AND ((Supplier.Deletion = False)))));"
datSQL.DatabaseName = strThePath
datEquipment.DatabaseName = strThePath
datEquipment.RecordSource = "Supplier"
datEquipment.Refresh
datSQL.RecordSource = strSQL
datSQL.Refresh
While Not datSQL.Recordset.EOF
    intSuppID = datSQL.Recordset.Fields("Supplier_ID")
    datEquipment.Recordset.FindFirst "Supplier_ID = " & intSuppID           'Find the supplier
    cboSupplier.AddItem (datEquipment.Recordset.Fields("[Supplier Name]"))
    cboSupplier.ItemData(cboSupplier.NewIndex) = intSuppID          'Store the supplier ID
    datSQL.Recordset.MoveNext
Wend
End Sub

Private Sub lstOrder_Click()
cmdCAR.Enabled = True
cmdRemove.Enabled = True
End Sub

Private Sub PrintHeader(ByVal Name As String, ByVal Add1 As String, ByVal Add2 As String, ByVal Add3 As String, ByVal ID As String, ByVal Phone As String, ByVal OrdID As String)
'This procedure prints every header for every page
With Printer        'Set the font
    .FontName = "Times New Roman"
    .Font.Size = 35
    .FontUnderline = True
    .ForeColor = &H800000
    .FontBold = False
    .FontItalic = True
End With
Printer.Print
Printer.Print Tab(9); "Equipment Hire System"
Printer.FontSize = 25
Printer.FontUnderline = False
Printer.Print
Printer.FontItalic = False
Printer.Print Tab(24); "Order List"
Printer.Print
Printer.FontSize = 17
Printer.Print Tab(7); Name
Printer.Print Tab(7); Add1
Printer.Print Tab(7); Add2
Printer.Print Tab(7); Add3
Printer.Print Tab(7); Phone
Printer.Print
Printer.Print Tab(7); "==============================================================="
Printer.Print Tab(7); "Supplier No.  : "; ID; Tab(35); "Order No.  : "; OrdID; Tab(65); "Date : "; Date
Printer.Print Tab(7); "==============================================================="
Printer.Print
Printer.FontUnderline = True
Printer.Print Tab(7); "ID"; Tab(15); "Make"; Tab(34); "Model"; Tab(50); "Details"; Tab(80); "Qty"
Printer.FontUnderline = False
Printer.Print
End Sub

Private Sub PrintDetails(ID As Integer, Make As String, Model As String, Details As String, ByVal Qty As String)
'This procedure prints the order details for every supplier
With Printer
    .FontName = "Times New Roman"
    .Font.Size = 15
    .FontUnderline = False
    .ForeColor = &H800000
End With
Printer.Print Tab(7); ID; Tab(17); Make; Tab(38); Model; Tab(56); Details; Tab(90); Qty
Printer.Print
End Sub

Private Sub OkOrder()
'This procedure prints out the order list and updates the "Order" table
Dim strOrder As String, intindex As Integer, strOrderID As String, strSuppID As String, strTemp As Variant
Do While lstOrder.ListCount <> 0            'Do until all orders are removed from list and sent to printer
    strOrder = lstOrder.List(intindex)
    strTemp = Split(strOrder, Chr(9))
    strSuppID = strTemp(2)                  'Get the supplier ID
    datEquipment.DatabaseName = strThePath
    datEquipment.RecordSource = "Equipment Type"
    datEquipment.Refresh
    datOrder.DatabaseName = strThePath
    datOrder.RecordSource = "Order"
    datOrder.Refresh
    With datOrder.Recordset                 'Create a new order for that supplier
        If .RecordCount <> 0 Then           'Create order ID
            .MoveLast
            strOrderID = .Fields("Order_ID") + 1
        Else
            strOrderID = 1
        End If
        .AddNew                             'Add a new order record
        .Fields("Order_ID") = strOrderID
        .Fields("Supplier_ID") = strSuppID
        .Fields("DateOrdered") = Date
        .Update
    End With
    PrintHeader strTemp(3), strTemp(4), strTemp(5), strTemp(6), strTemp(2), strTemp(7), strOrderID     'Print the header for the order sheet
    Do While (intindex <> lstOrder.ListCount) And (lstOrder.ListCount <> 0)     'Search the order list for matching suppliers ID
        strTemp = Split(lstOrder.List(intindex), Chr(9))
        If strTemp(2) = strSuppID Then                  'If true then add the current ordered euipment to the order list for that supplier
            datOrder.RecordSource = "Order/Equipment"
            datOrder.Refresh
            With datOrder.Recordset
                .AddNew                             'Create a new Order/Equipment Record
                .Fields("Order_ID") = strOrderID
                .Fields("Type_ID") = strTemp(1)
                .Fields("QuantityOrdered") = strTemp(8)
                .Update
            End With
            datEquipment.Recordset.FindFirst "Type_ID = " & strTemp(1)          'Find equipment details
            With datEquipment.Recordset
                PrintDetails .Fields("Type_ID"), .Fields("Make"), .Fields("Model"), .Fields("Details"), strTemp(8)     'Print the description of the equipment thats on order with the supplier
            End With
            lstOrder.RemoveItem (intindex)      'Remove ordered equipment from list since it has been processed
        Else
            intindex = intindex + 1         'Move to next item on the order list
        End If
    Loop
    Printer.NewPage         'Go to a new page for every supplier
    intindex = 0
Loop
Printer.EndDoc      'Start printing
End Sub

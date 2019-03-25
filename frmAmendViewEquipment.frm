VERSION 5.00
Begin VB.Form frmAmendViewEquipment 
   Caption         =   "Amend/View Equipment"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDescription 
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
      Left            =   5520
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5400
      Width           =   4935
   End
   Begin VB.ComboBox cmbStatus 
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
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6600
      Width           =   3735
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
      Left            =   11280
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ListBox lstEquipmentType 
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
      Height          =   1770
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   10215
   End
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8520
      Width           =   1575
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
      Left            =   13320
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton CmdBack 
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
      Left            =   11520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox txtHireHour 
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
      Left            =   5520
      TabIndex        =   5
      Top             =   8400
      Width           =   3735
   End
   Begin VB.TextBox txtHireDay 
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
      Left            =   5520
      TabIndex        =   4
      Top             =   7800
      Width           =   3735
   End
   Begin VB.TextBox txtPurchasePrice 
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
      Left            =   5520
      TabIndex        =   3
      Top             =   7200
      Width           =   3735
   End
   Begin VB.TextBox txtSerialNumber 
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
      Left            =   5520
      TabIndex        =   2
      Top             =   9000
      Width           =   3735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Data datEquipmentType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data datCategory 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAMend 
      Caption         =   "&Amend"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1575
   End
   Begin VB.ComboBox cmbSupplier 
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
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   6000
      Width           =   4935
   End
   Begin VB.Data datSupplier 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Equipment ID Selection"
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
      Height          =   3735
      Left            =   2280
      TabIndex        =   23
      Top             =   1320
      Width           =   10935
      Begin VB.ListBox lstEquipmentType2 
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
         Height          =   1770
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   10095
      End
      Begin VB.ComboBox cmbSerialNumb 
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
         Left            =   8400
         Sorted          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtEquipmentID 
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
         Height          =   465
         Left            =   7560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3150
         Width           =   1095
      End
      Begin VB.OptionButton optEquipmentID 
         Caption         =   "Search Equipment ID"
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
         Left            =   7920
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optEquipmentType 
         Caption         =   "Search By Equipment Description"
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
         Left            =   360
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblSerialNo 
         Caption         =   "Select  Serial No"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblEqID 
         Caption         =   "Enter Equipment ID"
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
         Left            =   5280
         TabIndex        =   27
         Top             =   3240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
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
      Height          =   2775
      Left            =   10560
      TabIndex        =   24
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Equipment Details"
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
      TabIndex        =   31
      Top             =   5640
      Width           =   1845
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Amend/View Equipment"
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
      Left            =   3960
      TabIndex        =   21
      Top             =   480
      Width           =   7635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number"
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
      TabIndex        =   20
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Supplier Details"
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
      TabIndex        =   19
      Top             =   6120
      Width           =   1605
   End
   Begin VB.Label lblSupplier 
      AutoSize        =   -1  'True
      Caption         =   "Type Details"
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
      Height          =   525
      Left            =   8640
      TabIndex        =   18
      Top             =   1800
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Hire Price Per Day"
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
      TabIndex        =   17
      Top             =   7920
      Width           =   1905
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Purchace Price"
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
      TabIndex        =   16
      Top             =   7320
      Width           =   1485
   End
   Begin VB.Label Label8 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   6720
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Hire Price Per Hour"
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
      TabIndex        =   14
      Top             =   8400
      Width           =   1995
   End
End
Attribute VB_Name = "frmAmendViewEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose of screen:

'the purpose of this screen is to allow the user to to select an item of
'equipment either by entering the equipment id number or by browsing threw
'equipment details until the item is found ,when the item is found the user
'can 'make changes to the equipment details by clicking on the amend button
'to which allows the user to to select the listed details any manually enter
'the rest,when the changes have been made the user has two options,either
'save or cancel the changes made to the equipment details,if the details
'are saved the relevant tables are updated ie the equipment type table and
'the equipment.
'
'
'Author: Derek Stafford com2026
'15/03/2002


'Variables Used                              Purpose of string
         
'strSerialNumb              string          stores a tempary string value from the serial number textbox
'strStatus,                 string          stores a tempary string value from the status textbox
'strDescription             string          stores a tempary string value from the descrption  textbox
'strmySQL                   string          stores the SQL queries
'curHirePDay                currency        stores a tempary currency value from the hire per day textbox
'strStatus,                 currency        stores a tempary currency value from the status textbox
'strDescription             currency        stores a tempary currency value from the descrption  textbox
'intindex                   integer         used as a counter variable in loops etc
'intTempEquipTypeID         integer         stores a tempary equipment type ID value
'intEmpSupplierID           integer         stores supplier ID

'objects used

'datEquipmentType           datacontrol     Retrieves info from equipmentType table
'datEquipment               datacontrol     retrieves info from equipment table
'datSupplier                datacontrol     retrieves info from supplier table
'datSupplier                datacontrol     retrieves info from category table
'datCategory









Option Explicit

Dim strSerialNumb, strStatus, strDescription As String
Dim curHirePDay, curHirePHour, curPurchPrice As Currency
Dim intTempEquipTypeID, intTempSupplierID, intTempEquipmentID As Integer
Dim intEquipTypeID, intEquipTypeID2, intSupplierID, intEquipmentID As Integer


'Empties textboxes on the screen

Private Sub Cleartextboxes()
cmbSupplier = ""
txtSerialNumber = ""
cmbStatus = ""
txtPurchasePrice = ""
txtHireDay = ""
txtHireHour = ""
txtDescription = ""
    
End Sub

'fills combo box with serial numbers of the equipment type


Private Sub cmbSerialNumb_click()
Dim strmySQl As String
lstEquipmentType2.Enabled = False
cmdAMend.Enabled = True
cmbSerialNumb.Enabled = False
strSerialNumb = cmbSerialNumb.ItemData(cmbSerialNumb.ListIndex)
If intEquipTypeID = "" Then
   Call MsgBox("you have not selected an item", vbCritical, "Equipment Not Selected")
 
Else
    cmdAMend.Enabled = True
    strmySQl = "SELECT Equipment.* From Equipment WHERE (((Equipment.Type_ID)=" & intEquipTypeID & ") AND ((Equipment.[S/N])='" & strSerialNumb & "'));"
    datEquipment.RecordSource = strmySQl
    datEquipment.Refresh
    intEquipmentID = datEquipment.Recordset("Equipment_ID")
    display
End If
End Sub

'puts data into combobox

Private Sub cmbStatus_DropDown()
cmbStatus.Clear
cmbStatus.AddItem ("Available")
cmbStatus.AddItem ("Broke")
cmbStatus.AddItem ("Out")
End Sub

'fills the supplier combo box
Private Sub cmbSupplier_DropDown()
cmbSupplier.Clear
datSupplier.Recordset.MoveFirst
While Not datSupplier.Recordset.EOF
    cmbSupplier.AddItem (datSupplier.Recordset("Supplier Name")) + ",  " + (datSupplier.Recordset("Address 1")) + ",  " + (datSupplier.Recordset("Address 2"))
    cmbSupplier.ItemData(cmbSupplier.NewIndex) = datSupplier.Recordset("Supplier_ID")
    datSupplier.Recordset.MoveNext
Wend
End Sub


'stores the selected supplier id number

Private Sub cmbSupplier_Click()
intSupplierID = cmbSupplier.ItemData(cmbSupplier.ListIndex)
End Sub

'stores the textbox data into variables for retieval if changes are cancelled
'and also resets screen layout

Private Sub cmdAMend_Click()
cmdAMend.Enabled = False
lstEquipmentType.Enabled = False
lstEquipmentType.Visible = False
lstEquipmentType.Visible = False
lstEquipmentType2.Visible = True
lstEquipmentType2.Enabled = False
lstEquipmentType2.Enabled = True
cmbSupplier.Enabled = True
cmbSupplier.SetFocus
txtSerialNumber.Enabled = True
txtSerialNumber.SetFocus
cmbStatus.Enabled = True
txtPurchasePrice.Enabled = True
txtHireDay.Enabled = True
txtHireHour.Enabled = True
cmdSave.Enabled = True
Cmdcancel.Enabled = True

intTempEquipmentID = datEquipment.Recordset("Equipment_ID")
 If intEquipTypeID2 = "" Then
    intTempEquipTypeID = intEquipmentID
Else
    intTempEquipTypeID = intEquipTypeID2
    End If
    intTempEquipTypeID = datEquipment.Recordset("Type_ID")
    intTempSupplierID = intSupplierID
    strSerialNumb = txtSerialNumber
    strStatus = cmbStatus.Text
    curPurchPrice = txtPurchasePrice
    curHirePDay = txtHireDay
    curHirePHour = txtHireHour
    strDescription = txtDescription
End Sub

'puts back the orignal data back into textboxes before changes were made

Private Sub cmdCancel_Click()
cmdAMend.Enabled = True
intSupplierID = datEquipment.Recordset("Supplier_ID")
datSupplier.Recordset.FindFirst "Supplier_ID =" & intTempSupplierID & ""
cmbSupplier.Text = datSupplier.Recordset("Supplier Name")

txtSerialNumber = strSerialNumb
cmbStatus = strStatus
txtPurchasePrice = curPurchPrice
txtHireDay = curHirePDay
txtHireHour = curHirePHour
lstEquipmentType.Visible = True
lstEquipmentType.Enabled = True
txtDescription = strDescription
Disabletextboxes
End Sub

'check for valid number

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


'check if equipment id entered is valid and if so it finds the info to be diplayed
'if not a an error message is displayed

Private Sub cmdDisplay_Click()
If IsValid(txtEquipmentID.Text) = False Then
    Call MsgBox("You have entered invalid data or no data ,you must enter a number", vbCritical, "Invalid Data Entry")
    txtEquipmentID = ""
    txtEquipmentID.SetFocus
Else
    If txtEquipmentID > 9999 Then
        Call MsgBox("The highest entry is 9999", vbCritical, "Entry Exceeds Limit")
        txtEquipmentID.Text = ""
        txtEquipmentID.SetFocus
    Else
        intEquipmentID = txtEquipmentID
        cmdAMend.Enabled = True
        datEquipment.RecordSource = "Equipment"
        datEquipment.Refresh
        datEquipment.Recordset.FindFirst "Equipment_ID=" & intEquipmentID & ""
        If datEquipment.Recordset.NoMatch Then
            Call MsgBox("This ID is not listed ", vbInformation, "Not Listed")
            txtEquipmentID.Text = ""
            txtEquipmentID.SetFocus
        Else
            intEquipTypeID = datEquipment.Recordset("Type_ID")
            cmdAMend.Enabled = True
            lstEquipmentType.Visible = False
            lstEquipmentType2.Visible = True
            lstEquipmentType.Enabled = False
            lstEquipmentType2.Enabled = False
            display
        End If
    End If
End If
End Sub

'checks if valid info was entered and if so searches for the the correct
'record in the equipment table  record to save the amended info to

Private Sub cmdSave_Click()
cmdAMend.Enabled = True
Dim intindex As Integer, blnInvalid As Boolean

If IsValid(txtSerialNumber) = False Then
    Call MsgBox("You have entered invalid serial number ,you must enter a number", vbCritical, "Invalid Data Entry")
    txtSerialNumber = ""
    txtSerialNumber.SetFocus
Else
    Disabletextboxes
    datEquipmentType.RecordSource = "Equipment Type"
    datEquipmentType.Refresh
    blnInvalid = False
    If ((cmbStatus.Text = "Broke") Or (cmbStatus.Text = "Out") Or (cmbStatus.Text = "Available")) = False Then
        blnInvalid = True
    Else
        blnInvalid = False
    End If
    If txtSerialNumber = "" Or txtPurchasePrice = "" Or txtHireDay = "" Or txtHireHour = "" Or (blnInvalid = True) Then
        Call MsgBox("All The Equipment Details Have Not Been Filled In Or The Status Entry Has Been Written in Not Selected", vbCritical, "Details InComplete")
        cmdAMend.Enabled = True
    Else
        datEquipment.Recordset.MoveFirst
        datEquipment.Recordset.FindFirst "Equipment_ID=" & intEquipmentID & ""
        intTempEquipmentID = intEquipmentID
        With datEquipment
            .Recordset.Edit
            If intEquipTypeID2 = "" Then
                .Recordset("Type_ID") = intEquipTypeID
                datEquipmentType.Recordset.FindFirst "Type_ID = " & intEquipTypeID & ""
                txtDescription = datEquipmentType.Recordset("Make") + "   " + datEquipmentType.Recordset("Model") & _
                "   " + datEquipmentType.Recordset("Description") + "   " + datEquipmentType.Recordset("Details")
            Else
            .Recordset("Type_ID") = intEquipTypeID2
            datEquipmentType.Recordset.FindFirst "Type_ID = " & intEquipTypeID2 & ""
            txtDescription = datEquipmentType.Recordset("Make") + "   " + datEquipmentType.Recordset("Model") & _
            "   " + datEquipmentType.Recordset("Description") + "   " + datEquipmentType.Recordset("Details")
            End If
            
            
            
            
            .Recordset("Supplier_ID") = intSupplierID
            .Recordset("S/N") = Val(txtSerialNumber.Text)
            .Recordset("Status") = cmbStatus
            .Recordset("[Purchase Price]") = CCur(Val(txtPurchasePrice.Text))
            .Recordset("[Hire Price Per Day]") = CCur(Val(txtHireDay.Text))
            .Recordset("[Hire Price Per Hour]") = CCur(Val(txtHireHour))
            .Recordset.Update
        
        End With
        optEquipmentType.Value = False
        optEquipmentID.Value = False
        lstEquipmentType.Enabled = False
        lstEquipmentType2.Enabled = False
    End If
End If
End Sub

'disables textboxes and combo boxes on the screen

Private Sub Disabletextboxes()
txtSerialNumber.Enabled = False
cmbSupplier.Enabled = False
cmbStatus.Enabled = False
txtPurchasePrice.Enabled = False
txtHireDay.Enabled = False
txtHireHour.Enabled = False
cmdSave.Enabled = False
Cmdcancel.Enabled = False
End Sub

'loads main menu screen and unloads the amend/view equipment screen

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

'loads equipmentfileprocessing screen and unloads the amend/view equipment screen

Private Sub cmdBack_Click()
frmEquipmentFileProcessing.Show
Unload Me
End Sub

'fills the two listboxes and sets up screen layout

Private Sub Form_Activate()
Dim srtDescription, strMake, strModel, strDetails As String, intindex As Integer
txtEquipmentID.Visible = True
cmdDisplay.Visible = False
cmbSerialNumb.Visible = False
optEquipmentType.SetFocus
lblSerialNo.Visible = False
lblEqID.Visible = False


cmbSerialNumb.Visible = False
lstEquipmentType.Visible = True
lstEquipmentType2.Visible = False
lstEquipmentType.Enabled = False
txtEquipmentID.Enabled = False
Disabletextboxes
cmdDisplay.Enabled = False
lblSerialNo.Visible = False
lstEquipmentType.Clear
lstEquipmentType2.Clear
'datEquipmentType.Recordset.MoveFirst
While Not datEquipmentType.Recordset.EOF
    With datEquipmentType
        srtDescription = .Recordset("Description")
        strMake = .Recordset("Make")
        strModel = .Recordset("Model")
        strDetails = .Recordset("Details")
        For intindex = Len(.Recordset("Description")) To 20
            srtDescription = srtDescription + " "
        Next intindex
        For intindex = Len(.Recordset("Make")) To 20
            strMake = strMake + " "
        Next intindex
            For intindex = Len(.Recordset("Model")) To 20
                strModel = strModel + " "
        Next intindex
        For intindex = Len(.Recordset("Details")) To 20
            strDetails = strDetails + " "
        Next intindex
                  
        lstEquipmentType.AddItem (srtDescription + vbTab + strMake + vbTab + strModel + vbTab + strDetails)
        lstEquipmentType.ItemData(lstEquipmentType.NewIndex) = .Recordset("Type_ID")
        lstEquipmentType2.AddItem (srtDescription + vbTab + strMake + vbTab + strModel + vbTab + strDetails)
        lstEquipmentType2.ItemData(lstEquipmentType2.NewIndex) = .Recordset("Type_ID")
        .Recordset.MoveNext
    End With
Wend
End Sub

'displays the selected item details to the screen

Private Sub display()
intSupplierID = datEquipment.Recordset("Supplier_ID")
datEquipmentType.RecordSource = "Equipment Type"
datEquipmentType.Refresh
datEquipmentType.Recordset.FindFirst "Type_ID = " & intEquipTypeID & ""
txtDescription = datEquipmentType.Recordset("Make") + "   " + datEquipmentType.Recordset("Model") & _
"   " + datEquipmentType.Recordset("Description") + "   " + datEquipmentType.Recordset("Details")

txtSerialNumber = datEquipment.Recordset("S/N")
cmbStatus = datEquipment.Recordset("Status")
txtPurchasePrice = datEquipment.Recordset("Purchase Price")
txtHireDay = datEquipment.Recordset("Hire Price Per Day")
txtHireHour = datEquipment.Recordset("Hire Price Per Hour")
datEquipmentType.RecordSource = "Equipment Type"
datEquipmentType.Refresh

datSupplier.Recordset.MoveFirst
datSupplier.Recordset.FindFirst "Supplier_ID =" & intSupplierID & ""
cmbSupplier.Text = datSupplier.Recordset("Supplier Name") + ",  " + datSupplier.Recordset("Address 1") + ",  " + datSupplier.Recordset("Address 2")


End Sub

'assigns the record sources to the datacontrols,from the selected database

Private Sub Form_Load()
Cmdcancel.Enabled = False
cmdAMend.Enabled = False

cmdSave.Enabled = False
datEquipmentType.DatabaseName = strThePath
datEquipmentType.RecordSource = "Equipment Type"
datEquipment.DatabaseName = strThePath
datEquipment.RecordSource = "Equipment"
datCategory.DatabaseName = strThePath
datCategory.RecordSource = "Category"
datSupplier.DatabaseName = strThePath
datSupplier.RecordSource = "Supplier"

End Sub

'takes in the selected equipment type details and gives out the and fills a combo box
'of all the serial numbers stored on that equipment type

Private Sub lstEquipmentType_Click()
Dim strmySQl As String


optEquipmentType.Value = False
cmbSerialNumb.Visible = True
cmbSerialNumb.Enabled = True
cmbSerialNumb.Visible = True
lblSerialNo.Visible = True


cmbSerialNumb.SetFocus
lblSerialNo.Visible = True
lstEquipmentType.Enabled = False
cmbSerialNumb.Enabled = True

intEquipTypeID = lstEquipmentType.ItemData(lstEquipmentType.ListIndex)
strmySQl = "SELECT Equipment.Type_ID, [Equipment Type].Type_ID, [Equipment Type].Model, Equipment.[S/N]" & _
"FROM [Equipment Type] INNER JOIN Equipment ON [Equipment Type].Type_ID = Equipment.Type_ID " & _
"WHERE ((([Equipment Type].Type_ID)=" & intEquipTypeID & "));"
datEquipmentType.RecordSource = strmySQl

datEquipmentType.Refresh
cmbSerialNumb.Clear
If datEquipmentType.Recordset.EOF Then
    Call MsgBox("There is no Items matching this Description stored", vbCritical, "No Match")
    lstEquipmentType.Enabled = True
Else
    cmbSerialNumb.Clear
    While Not datEquipmentType.Recordset.EOF
        cmbSerialNumb.AddItem (datEquipmentType.Recordset("S/N"))
        cmbSerialNumb.ItemData(cmbSerialNumb.NewIndex) = datEquipmentType.Recordset("S/N")
        datEquipmentType.Recordset.MoveNext
    Wend
End If
End Sub

'holds the new equipment type id that has been selected by the user

Private Sub lstEquipmentType2_Click()
intEquipTypeID2 = lstEquipmentType.ItemData(lstEquipmentType2.ListIndex)
End Sub

'sets up the screen layout for the enter equipment id option

Private Sub optEquipmentID_Click()
txtEquipmentID.Visible = True
cmdDisplay.Visible = True
cmbSerialNumb.Visible = False

Cleartextboxes
cmdAMend.Enabled = False
txtEquipmentID.Enabled = True
txtEquipmentID.Text = ""
cmdSave.Enabled = False
txtEquipmentID.Enabled = True
Cmdcancel.Enabled = False
intEquipTypeID = ""
intEquipTypeID2 = ""
Disabletextboxes
cmbSerialNumb.Visible = False
lblSerialNo.Visible = False
lblEqID.Visible = True
cmdDisplay.Enabled = True
txtEquipmentID.Enabled = True
txtEquipmentID.SetFocus
End Sub

'sets up the screen layout for the enter equipment browse option

Private Sub optEquipmentType_Click()
cmdAMend.Enabled = False

txtEquipmentID.Visible = False
cmdDisplay.Visible = False
'cmbSerialNumb.Visible = True

'lblSerialNo.Visible = True
txtEquipmentID.Enabled = False
txtEquipmentID.Text = ""
Cleartextboxes
intEquipTypeID = ""
intEquipTypeID2 = ""
Disabletextboxes
cmdSave.Enabled = False
Cmdcancel.Enabled = False
'cmdDisplay.Enabled = False
lblEqID.Visible = False
lstEquipmentType.Visible = True
lstEquipmentType2.Visible = False
lstEquipmentType.Enabled = True
lstEquipmentType2.Enabled = False
End Sub


'checks that serial number is valid

Private Sub txtSerialNumber_Validate(Cancel As Boolean)
If IsNumeric(txtSerialNumber.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtSerialNumber.SetFocus
    Cancel = True
End If
End Sub


'checks that serial number entered is valid

Private Sub txtPurchasePrice_Validate(Cancel As Boolean)
If IsNumeric(txtPurchasePrice.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtHireDay.SetFocus
    Cancel = True
End If
End Sub


'checks that hire day number entered is valid

Private Sub txtHireDay_Validate(Cancel As Boolean)
If IsNumeric(txtHireDay.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtHireDay.SetFocus
    Cancel = True
End If
End Sub


'checks that hire hour number entered is valid

Private Sub txtHireHour_Validate(Cancel As Boolean)
If IsNumeric(txtHireHour.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtHireHour.SetFocus
    Cancel = True
End If
End Sub


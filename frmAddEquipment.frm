VERSION 5.00
Begin VB.Form frmAddEquipment 
   Caption         =   "Add  Equipment"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCategory 
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
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5040
      Width           =   3735
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
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   5640
      Width           =   3735
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
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   9855
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
      Left            =   11280
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7560
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      Left            =   5400
      TabIndex        =   5
      Top             =   7440
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
      Left            =   5400
      TabIndex        =   4
      Top             =   6840
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
      Left            =   5400
      TabIndex        =   3
      Top             =   6240
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
      Left            =   5400
      TabIndex        =   2
      Top             =   8040
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
      Left            =   11280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6840
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
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   5175
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
      Caption         =   "Equipment Description Selection"
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
      Height          =   2895
      Left            =   2280
      TabIndex        =   20
      Top             =   1320
      Width           =   10935
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
      Left            =   10680
      TabIndex        =   21
      Top             =   5640
      Width           =   2655
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
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
         Left            =   600
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Category Name"
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
      TabIndex        =   22
      Top             =   5160
      Width           =   1590
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Add  Equipment"
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
      TabIndex        =   18
      Top             =   480
      Width           =   5145
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
      TabIndex        =   17
      Top             =   8160
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
      TabIndex        =   16
      Top             =   4560
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   6960
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
      TabIndex        =   13
      Top             =   6360
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
      TabIndex        =   12
      Top             =   5760
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
      TabIndex        =   11
      Top             =   7560
      Width           =   1995
   End
End
Attribute VB_Name = "frmAddEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose of screen:

'the purpose of this screen is to allow the user to add in new equipment
'to the database,the user can enter in the new equipment details by selecting
'on some of the listed details such as equipment type description
'supplier of the equipment and category of the equipment, the rest of the
'details are then manually entered in, when all the details are entered
'the user has two options exit without saving the new equipment
'details to the databse or he/she can save the details to the databse
'and he/she can then cancel/delete the record just entered from
'the databse,the user can then add another new record to the database
'by clicking the add new button


'Author: Derek Stafford com2026
'15/03/2002


'variables used                     purpose
'
'strMake            string;          holds make fields from the equipment type table
'strModel,          string;          holds model fields from the equipment type table
'strDetails         string;          holds details fields from the equipment type table
'strSQL             string;          holds SQL query result
'strDescription     string;          holds description fields form the equipment type table
'curHirePDay        currency;        stores a tempary currency value from the hire per day textbox
'strStatus,         currency;        stores a tempary currency value from the status textbox
'strDescription     currency;        stores a tempary currency value from the descrption  textbox
'
'intindex           integer;         used as a counter variable in loops etc
'intTempEquipTypeID integer;         holds equipment type ID value
'intEmpSupplierID   integer;         holds supplier ID
'intcategoryID      integer;         holds  category id
'intResult          integer;         holds result of answer from message box
'
'objects used
'
'datEquipmentType   datacontrol;     holds data from the equipment type table
'datEquipment       datacontrol;     holds data from the equipment table
'datSupplier        datacontrol;     holds data from the supplier table
'datCategory        datacontrol;     holds data from the category table


'Author: Derek Stafford com2026
'15/03/2002
'





Option Explicit
Dim intEquipTypeID, intCategoryID, intSupplierID As Integer

Private Sub cmbCategory_Click()
intCategoryID = cmbCategory.ItemData(cmbCategory.ListIndex)
datCategory.Recordset.FindFirst "Category_ID = " & intCategoryID & ""
End Sub
Private Sub cmdAddNew_Click()

  '  cmbEquipType.Text = ""
    cmbCategory.Text = ""
    cmbSupplier.Text = ""
    txtSerialNumber = ""
 '   txtStatus = ""
    txtPurchasePrice = ""
    txtHireDay = ""
    txtHireHour = ""
End Sub

Private Sub cmdCancel_Click()
Dim intResult As Integer
intResult = MsgBox("Are you sure you want to cancel,if you click yes you will delete the new record from the database", vbYesNo, "Cancel")
If intResult = vbYes Then
    With datEquipment
        .Recordset.MoveLast
        .Recordset.Delete
    End With
    
    
    
    
    
    
End If
cmdCancel.Enabled = False
cmdNew.Enabled = True
cmdSave.Enabled = True
End Sub

Private Sub cmdNew_Click()
cmdSave.Enabled = True
cmdCancel.Enabled = False
cmbSupplier = ""
cmbCategory = ""
cmbStatus = ""
txtPurchasePrice = ""
txtHireDay = ""
txtHireHour = ""
txtSerialNumber = ""
lstEquipmentType.SetFocus
End Sub

Private Sub cmdSave_Click()

Dim intindex As Integer
   If intEquipTypeID = "" Then
        Call MsgBox("An equipment description has not been selected", vbCritical, "Empty Field")
        cmdSave.Enabled = True
        cmbSupplier.SetFocus
    Else
    If cmbCategory.ListIndex = -1 Or cmbSupplier.ListIndex = -1 Or txtPurchasePrice = "" Or cmbStatus.ListIndex = -1 Or txtHireDay = "" Or txtHireHour = "" Then
        Call MsgBox("All/Some Of The Equipment Details Have Not Been Filled In ", vbCritical, "Details InComplete")
        cmdSave.Enabled = True
    Else
        datEquipment.Recordset.MoveLast
        intindex = datEquipment.Recordset("Equipment_ID") + 1
        With datEquipment
            .Recordset.AddNew
            .Recordset("Equipment_ID") = intindex
            .Recordset("Type_ID") = intEquipTypeID
            .Recordset("Supplier_ID") = intSupplierID
            .Recordset("S/N") = Val(txtSerialNumber.Text)
            .Recordset("Status") = cmbStatus.Text
            .Recordset("[Purchase Price]") = CCur(Val(txtPurchasePrice.Text))
            .Recordset("[Hire Price Per Day]") = CCur(Val(txtHireDay.Text))
            .Recordset("[Hire Price Per Hour]") = CCur(Val(txtHireHour))
            .Recordset.Update
            
    End With
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
    cmdNew.Enabled = True

End If
End If
End Sub

Private Sub cmbStatus_DropDown()
cmbStatus.Clear
cmbStatus.AddItem ("Available")
cmbStatus.AddItem ("Broke")
cmbStatus.AddItem ("Out")
End Sub

Private Sub cmbcategory_DropDown()
'fills the supplier combo box
cmbCategory.Clear
While Not datCategory.Recordset.EOF
    cmbCategory.AddItem (datCategory.Recordset("Category Name"))
    cmbCategory.ItemData(cmbCategory.NewIndex) = datCategory.Recordset("Category_ID")
   datCategory.Recordset.MoveNext
Wend

End Sub

Private Sub cmbSupplier_DropDown()
cmbSupplier.Clear
datSupplier.Recordset.MoveFirst
While Not datSupplier.Recordset.EOF
    cmbSupplier.AddItem datSupplier.Recordset("Supplier Name") + ",  " + datSupplier.Recordset("Address 1") + ",  " + datSupplier.Recordset("Address 2")
    cmbSupplier.ItemData(cmbSupplier.NewIndex) = datSupplier.Recordset("Supplier_ID")
    datSupplier.Recordset.MoveNext
Wend
End Sub


'stores the selected supplier id number

Private Sub cmbSupplier_Click()
intSupplierID = cmbSupplier.ItemData(cmbSupplier.ListIndex)
End Sub

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
cmdNew.Enabled = False
lstEquipmentType.SetFocus
cmdSave.Enabled = True

datEquipmentType.Recordset.MoveFirst
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
        .Recordset.MoveNext
    End With
Wend
End Sub


Private Sub Form_Load()
cmdCancel.Enabled = False
'cmdAMend.Enabled = False

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


Private Sub lstEquipmentType_Click()
Dim strmySQl As String

intEquipTypeID = lstEquipmentType.ItemData(lstEquipmentType.ListIndex)
End Sub

'checks that serial number entered is valid
'==========================================================
Private Sub txtSerialNumber_Validate(Cancel As Boolean)
If IsNumeric(txtSerialNumber.Text) Then
    Cancel = False
Else
    Call MsgBox("Sorry invalid data", vbOKOnly, "Warning")
    txtSerialNumber.SetFocus
    Cancel = True
End If
End Sub

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


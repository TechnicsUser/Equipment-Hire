VERSION 5.00
Begin VB.Form frmDeleteEquipment 
   Caption         =   "Delete Equipment Screen"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
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
      Left            =   6960
      TabIndex        =   27
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      TabIndex        =   26
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
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
      Left            =   13440
      TabIndex        =   25
      Top             =   10080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
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
      Left            =   11280
      TabIndex        =   24
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Data datSQL2 
      Caption         =   "datSQL2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data datSQL1 
      Caption         =   "datSQL1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data datType 
      Caption         =   "datType"
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
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtSN 
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
      Left            =   8880
      TabIndex        =   22
      Top             =   7920
      Width           =   2055
   End
   Begin VB.ComboBox cboID 
      Height          =   315
      ItemData        =   "frmDeleteEquipment.frx":0000
      Left            =   8880
      List            =   "frmDeleteEquipment.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Frame fmeStatus 
      Caption         =   "Search for equiptment"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2280
      Width           =   9015
      Begin VB.OptionButton optEqptNum 
         Caption         =   "Search for equipment useing ID number"
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
         TabIndex        =   13
         Top             =   480
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton optEquipName 
         Caption         =   "Search for equipment using equipment details"
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
         Left            =   4200
         TabIndex        =   12
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cboEquipDescription 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeleteEquipment.frx":0004
         Left            =   7320
         List            =   "frmDeleteEquipment.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboEquipMake 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeleteEquipment.frx":0008
         Left            =   7320
         List            =   "frmDeleteEquipment.frx":000F
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1920
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
         TabIndex        =   9
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox cboEquipDetails 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeleteEquipment.frx":0018
         Left            =   7320
         List            =   "frmDeleteEquipment.frx":001F
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cboEquipModel 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeleteEquipment.frx":0028
         Left            =   7320
         List            =   "frmDeleteEquipment.frx":002F
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cboTypeID 
         Height          =   315
         ItemData        =   "frmDeleteEquipment.frx":0038
         Left            =   2760
         List            =   "frmDeleteEquipment.frx":003A
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewItem 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   5
         Top             =   2760
         Width           =   1095
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
         TabIndex        =   19
         Top             =   960
         Width           =   2580
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
         Left            =   4680
         TabIndex        =   18
         Top             =   2400
         Width           =   2400
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
         TabIndex        =   17
         Top             =   1920
         Width           =   2355
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
         TabIndex        =   16
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblNoInStock 
         AutoSize        =   -1  'True
         Caption         =   "Number presently in Stock"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   2880
         Width           =   2670
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Width           =   2460
      End
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   12225
      Width           =   1575
   End
   Begin VB.Label lblSN 
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
      Left            =   7200
      TabIndex        =   23
      Top             =   7920
      Width           =   1455
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
      Left            =   5760
      TabIndex        =   21
      Top             =   6960
      Width           =   2880
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Delete Equipment"
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
      Left            =   5010
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmDeleteEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to select a peice of equiptment for
'               and mark its record for deletion. If the user knows the equipments
'               ID number he can input it in a text box and then confirm it's
'               serial number and then delete it (If the peice of equipment is
'               out on hire then it should not be deleted)
'               Alternatively you can search for the ID numbers of all stock
'               by selecting either their Type ID number or working your way down
'               a number of selection steps to search the a unique type. Each step
'               You take reduces the scope of your search. When you have chosen the
'               type you require you can calculate the number in stock and add the
'               ID numbers to drop down list of a combo box
'Student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02

'Variables used
'strMake        String
'strModel       String
'strDescription String
'strDetails     String
'strAvailabilty String
'sqlDescription String
'strMess        String
'sqlSN          String
'sqlStatus      String
'sqlEqptID      String
'sqlCount       String
'intID          Integer
'intResponse    Integer
'intNumber      Integer
'intCount       Integer





'Checks if a valid "Equipment_ID" has been entered and if so it then checks
'If the equipment is on hire
Private Sub cboID_Click()
Dim intID, intResponse As Integer
Dim sqlSN, strMess, sqlStatus, strAvailabilty As String

intID = Val(cboID.Text)
sqlSN = "SELECT [S/N],Status FROM  Equipment  WHERE Equipment_ID = " & intID & " "

datSQL2.DatabaseName = strThePath
datSQL2.RecordSource = sqlSN
datSQL2.Refresh


On Error GoTo NoIDError
txtSN = datSQL2.Recordset.Fields("S/N")
strAvailabilty = datSQL2.Recordset.Fields("Status")
If strAvailability = "Out" Then
    strMess = "This item of equipment is curently on hire and should not be Deleted"
    intResponse = MsgBox(strMess, vbOKOnly, "Equipment on hire")
    If intResponse = vbOK Then
        txtSN = ""
        cboID.Text = ""
        cboID.SetFocus
    End If
End If
Exit Sub

NoIDError:
If Err.Number = 3021 Then  'If an ID is entered that is not in the Table
    If Len(txtSN) > 0 Then
        strMess = "There is no matching ID"
        intResponse = MsgBox(strMess, vbOKOnly, "Inncorect input")
        If intResponse = vbOK Then
            txtSN = ""
            cboID.Text = ""
            cboID.SetFocus
        End If
    End If
    On Error GoTo 0
End If


End Sub

'Resets screen to original settings

Private Sub cmdClear_Click()
cboEquipDescription.Text = ""
cboEquipDetails.Text = ""
cboEquipMake.Text = ""
cboEquipModel.Text = ""
cboTypeID.Text = ""
cboID.Text = ""
txtSN = ""
txtNoInStock = ""

End Sub

'Deletes the record curently shown
Private Sub cmdDelete_Click()
Dim intNumber As Integer

intResponse = MsgBox("Click OK to Delete  all inputted Data", vbOKCancel, "Delete NOW")
If intResponse = 1 Then

    datType.DatabaseName = strThePath
    datType.RecordSource = "Equipment"
    datType.Refresh
    
    intNumber = Val(cboID.Text)
    datType.Recordset.FindFirst " Equipment_ID = " & intNumber & ""
    
    datType.Recordset.Edit
    datType.Recordset("Deletion") = True
    datType.Recordset.Update
    
    cmdClear_Click
Else
    cmdClear_Click
End If
End Sub

Private Sub Command4_Click()
frmEquipmentFileProcessing.Show
Unload Me
End Sub

Private Sub Command6_Click()
frmMainMenu.Show
Unload Me
End Sub

'When the form is loaded 2 combo boxes lists are loaded, cboTypeID contains
'a list of all the Equipment ID numbers, and the cboEquipDescription contains
'a list of Equipment descriptions (the first subset that a Type is devided into)
Private Sub Form_Load()
Dim sqlDescription As String

datType.DatabaseName = strThePath
datType.RecordSource = "Equipment Type"
datType.Refresh

While Not datType.Recordset.EOF
    cboTypeID.AddItem datType.Recordset("Type_ID")
    datType.Recordset.MoveNext
Wend

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

End Sub
Private Sub cboTypeID_Click()
cmdNewItem.Enabled = True
End Sub
'calculate the number in stock and add the ID numbers to drop down list of
'a combo box
Private Sub cmdNewItem_Click()
Dim SqlCount As String
Dim intID As Integer
Dim intCount As Integer
Dim SQLEqptID As String

intID = Val(cboTypeID)
cboID.Clear

SqlCount = "SELECT Count(*)AS IDCount FROM  Equipment  WHERE Type_ID= " & intID & " AND Status =  'Available' AND Deletion = False ;"
SQLEqptID = "SELECT Equipment_ID FROM  Equipment  WHERE Type_ID= " & intID & " AND Status =  'Available' AND Deletion = False;"

datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = SqlCount
datSQL1.Refresh

datSQL2.DatabaseName = strThePath
datSQL2.RecordSource = SQLEqptID
datSQL2.Refresh

intCount = datSQL1.Recordset.Fields("IDCount")
txtNoInStock.Text = Str(intCount)


While Not datSQL2.Recordset.EOF
    cboID.AddItem datSQL2.Recordset("Equipment_ID")
    datSQL2.Recordset.MoveNext
    Wend


cboID.Enabled = True
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

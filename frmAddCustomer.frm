VERSION 5.00
Begin VB.Form frmAddCustomer 
   Caption         =   "Add a Customer Screen"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      CausesValidation=   0   'False
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
      Left            =   7320
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9960
      Width           =   1575
   End
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
      Left            =   13320
      TabIndex        =   22
      Top             =   9960
      Width           =   1575
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
      Left            =   11400
      TabIndex        =   21
      Top             =   9960
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
      Left            =   10560
      TabIndex        =   11
      Top             =   11280
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
      Left            =   8520
      TabIndex        =   10
      Top             =   11280
      Width           =   1575
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
      Left            =   12720
      TabIndex        =   9
      Top             =   11280
      Width           =   1575
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
      Left            =   6600
      TabIndex        =   0
      Top             =   2235
      Width           =   1935
   End
   Begin VB.TextBox txtPhoneNo 
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
      Left            =   6600
      TabIndex        =   4
      Top             =   5535
      Width           =   1935
   End
   Begin VB.TextBox txtAddress3 
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
      Left            =   6600
      TabIndex        =   3
      Top             =   4815
      Width           =   1935
   End
   Begin VB.TextBox txtAddress2 
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
      Left            =   6600
      TabIndex        =   2
      Top             =   3855
      Width           =   1935
   End
   Begin VB.TextBox txtAddress1 
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
      Left            =   6600
      TabIndex        =   1
      Top             =   3015
      Width           =   1935
   End
   Begin VB.TextBox txtCreditLimit 
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
      Left            =   6600
      TabIndex        =   5
      Top             =   6375
      Width           =   1935
   End
   Begin VB.TextBox txtEmail 
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
      Left            =   6600
      TabIndex        =   7
      Top             =   8175
      Width           =   1935
   End
   Begin VB.TextBox txtMobileNo 
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
      Left            =   6600
      TabIndex        =   6
      Top             =   7335
      Width           =   1935
   End
   Begin VB.Data datCustomer 
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
      Top             =   8295
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   9360
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9975
      Width           =   1575
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   7
      Left            =   6360
      TabIndex        =   29
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   28
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   27
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   26
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   25
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblAstrix 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   24
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblRed 
      Caption         =   "Required feilds are marked with a astrix"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Add a Customer"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   720
      Width           =   5115
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
      Left            =   5640
      TabIndex        =   19
      Top             =   2280
      Width           =   600
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
      Left            =   4800
      TabIndex        =   18
      Top             =   5640
      Width           =   1485
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
      Left            =   5280
      TabIndex        =   17
      Top             =   4920
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
      Left            =   5280
      TabIndex        =   16
      Top             =   3960
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
      Left            =   5280
      TabIndex        =   15
      Top             =   3120
      Width           =   945
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
      Left            =   5160
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
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
      Left            =   5760
      TabIndex        =   13
      Top             =   8280
      Width           =   615
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
      Left            =   4680
      TabIndex        =   12
      Top             =   7440
      Width           =   1590
   End
End
Attribute VB_Name = "frmAddCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to Display a add a new customer to the
'               Customer table
'Student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02

'Variables used
'intID          Integer
'intResponse    Integer


'If the this form was opened from the Equipment Hire screen then a string
'will have been placed in the form tag else this form must have been opened
'from the Main Menu so return their
Private Sub cmdBack_Click()
If frmAddCustomer.Tag <> "" Then
    frmRental.optNumber = False
    frmRental.optName = True
    frmRental.cboName.Visible = True
    frmRental.cboNumber.Visible = False
    frmRental.optNumber.Enabled = False
    frmRental.optName.Enabled = False
    frmRental.cboName.Enabled = False
    frmRental.cboNumber.Visible = False
    frmRental.cmdAddCust.Enabled = False
    frmRental.lblNameNumber = "New Customers Name"
    frmRental.cboName = txtName
    frmRental.txtCustID = ""
    frmRental.txtNameFound = ""
    frmRental.txtAddress1 = ""
    frmRental.txtAddress2 = ""
    frmRental.txtAddress3 = ""
    frmRental.txtPhoneNo = ""
    frmRental.txtMobileNo = ""
    frmRental.txtEmail = ""
    frmRental.txtCreditLimit = ""
    frmRental.txtStatus = ""
    frmRental.txtBalance = ""
    frmRental.datCustomer.Refresh
Else
    frmCustomerFileProcessing.Show
End If

Unload Me

End Sub
'Empties text boxes
Private Sub cmdClear_Click()
txtName = ""
txtAddress1 = ""
txtAddress2 = ""
txtAddress3 = ""
txtPhoneNo = ""
txtMobileNo = ""
txtEmail = ""
txtCreditLimit = ""
txtName.SetFocus
cmdSave.Enabled = False
End Sub

Private Sub cmdMainMenu_Click()
If Len(txtCreditLimit) > 0 Then
    intResponse = MsgBox("Do you wish to save before you Exit", vbYesNo, "Save Now")
        If intResponse = vbYes Then
            cmdSave_Click
        Else
            cmdClear_Click
    End If
End If
frmMainMenu.Show
Unload Me
End Sub

'Create a new customers ID then save the contents the screen to a table then
'Clear out the text boxes
Private Sub cmdSave_Click()

Dim intID As Integer
Dim intResponse As Integer

If Not datCustomer.Recordset.BOF Then
        datCustomer.Recordset.MoveLast
        If intID >= datCustomer.Recordset("Cust_ID") Then
            intID = intID + 1
        Else
            intID = datCustomer.Recordset("Cust_ID")
            intID = intID + 1
        End If
    Else
        intID = 1
    End If


intResponse = MsgBox("Click OK to save inputted Data", vbOKCancel, "SAVE NOW")
If intResponse = 1 Then
    datCustomer.Recordset.AddNew
    datCustomer.Recordset("Cust_ID") = intID
    datCustomer.Recordset("Name") = txtName
    datCustomer.Recordset("Address 1") = txtAddress1
    datCustomer.Recordset("Address 2") = txtAddress2
    datCustomer.Recordset("Address 3") = txtAddress3
    datCustomer.Recordset("Phone No") = txtPhoneNo
    datCustomer.Recordset("Mobile No") = txtMobileNo
    datCustomer.Recordset("E-Mail") = txtEmail
    datCustomer.Recordset("Credit Limit") = txtCreditLimit
    datCustomer.Recordset("Status") = "Normal"
    datCustomer.Recordset("Balance owed") = 0
    datCustomer.Recordset.Update
    
Else
    cmdClear_Click
End If
cmdSave.Enabled = False

End Sub

Private Sub Form_Activate()
txtName.SetFocus
End Sub

Private Sub Form_Load()

datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datCustomer.Refresh
End Sub

Private Sub Label3_Click()
End Sub

'uses the validate function to ensure esential data is taken in

Private Sub txtAddress1_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtAddress1) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtAddress1.SetFocus
    End If
    
End If


End Sub



Private Sub txtAddress2_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtAddress2) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtAddress2.SetFocus
    End If
End If
End Sub




Private Sub txtCreditLimit_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtCreditLimit) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtCreditLimit.SetFocus
    End If
ElseIf Not IsNumeric(txtCreditLimit) Then
    Cancel = True
    intResponse = MsgBox("You must enter a number in this box", vbOKOnly, "Numders only")
    If intResponse = 1 Then
        txtCreditLimit.SetFocus
    End If
End If
cmdSave.Enabled = True
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtName) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtName.SetFocus
    End If
End If
End Sub

Private Sub txtAddress3_Validate(Cancel As Boolean)
Dim intResponse As Integer
If Len(txtAddress3) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtAddress1.SetFocus
    End If
End If
End Sub



Private Sub txtPhoneNo_Change()
Dim intResponse As Integer
If Len(txtPhoneNo) = 0 Then
    Cancel = True
    intResponse = MsgBox("You must enter data into this field", vbOKOnly, "Please fill out the form properaly")
    If intResponse = 1 Then
        txtPhoneNo.SetFocus
    End If
End If
End Sub

VERSION 5.00
Begin VB.Form frmAddSupplier 
   Caption         =   "Add Supplier"
   ClientHeight    =   8880
   ClientLeft      =   1845
   ClientTop       =   1140
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNoSave 
      Caption         =   "&Cancel"
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
      Left            =   8400
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Exit screen without saving"
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Data datAddSupplier 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.TextBox txtSuppName 
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
      Height          =   405
      Left            =   6870
      TabIndex        =   0
      Top             =   3285
      Width           =   2295
   End
   Begin VB.TextBox txtMobNum 
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
      Height          =   405
      Left            =   6870
      TabIndex        =   5
      Top             =   6045
      Width           =   2295
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6870
      TabIndex        =   6
      Top             =   6585
      Width           =   2295
   End
   Begin VB.TextBox txtTelNum 
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
      Height          =   405
      Left            =   6870
      TabIndex        =   4
      Top             =   5445
      Width           =   2295
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6870
      TabIndex        =   3
      Top             =   4845
      Width           =   2295
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6870
      TabIndex        =   2
      Top             =   4365
      Width           =   2295
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6870
      TabIndex        =   1
      Top             =   3885
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Main Menu"
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
      Left            =   10920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Ok"
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
      Left            =   5880
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Exit screen and save all new additions"
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddSupplier 
      Caption         =   "Add Supplier"
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
      Left            =   9600
      TabIndex        =   7
      ToolTipText     =   "Click to add supplier details to the database"
      Top             =   6495
      Width           =   2175
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Address3:"
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
      Left            =   5610
      TabIndex        =   23
      Top             =   4845
      Width           =   1020
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Address2:"
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
      Left            =   5610
      TabIndex        =   22
      Top             =   4365
      Width           =   1020
   End
   Begin VB.Label Label12 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   6630
      TabIndex        =   21
      Top             =   4845
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   6630
      TabIndex        =   20
      Top             =   4365
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   6630
      TabIndex        =   19
      Top             =   5445
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   6630
      TabIndex        =   18
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   135
      Left            =   6630
      TabIndex        =   17
      Top             =   3885
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(Required Fields marked with an asterisk)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5715
      TabIndex        =   16
      Top             =   2640
      Width           =   3705
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Add Supplier"
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
      Left            =   5460
      TabIndex        =   15
      Top             =   1320
      Width           =   4155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Mobile Num:"
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
      Left            =   5295
      TabIndex        =   14
      Top             =   6045
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Email:"
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
      Left            =   5985
      TabIndex        =   13
      Top             =   6645
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tele. Number:"
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
      Left            =   5175
      TabIndex        =   12
      Top             =   5445
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address1:"
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
      Left            =   5610
      TabIndex        =   11
      Top             =   3885
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
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
      Left            =   5955
      TabIndex        =   10
      Top             =   3285
      Width           =   675
   End
End
Attribute VB_Name = "frmAddSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to allow the user to enter the details of a new supplier.
'The user enters the details of the supplier in the relevant textboxes and then clicks
'"Add Supplier".
'If the user clicks "Cancel", any suppliers that were entered since this screen was last
'invoked, are completly deleted from the system.

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'MyString   :   String  (Holds the telephone number of a supplier)
'TheLastId  :   String  (Holds the primary ID of the last record in the current Supplier table)
'Counter    :   Integer (Control variable of a loop)
'Answer     :   Boolean (indicates whether the user wishes to save the details he has just entered)
'CountUp    :   Intger  (indicates how many characters in a string are numeric)

'Objects Used -
'datAddSupplier  :   Data Control    (Used to access the supplier table in the database)


Option Explicit

Private Sub cmdAddSupplier_Click()
'This procedure adds a supplier to the database if the supplier doesn't already exists
If (Len(txtSuppName.Text) > 0) And (Len(txtAdd1.Text) > 0) And (Len(txtAdd2.Text) > 0) And (Len(txtAdd3.Text) > 0) And (Len(txtTelNum.Text) > 0) Then
    Dim MyString, TheLastId As String
    datAddSupplier.RecordSource = "Select * From Supplier Order By Supplier_ID"
    datAddSupplier.Refresh
    datAddSupplier.Recordset.MoveLast
    TheLastId = datAddSupplier.Recordset("Supplier_ID")
    datAddSupplier.RecordSource = "Select * from Supplier where [Deletion] = False Order By Supplier_ID"
    datAddSupplier.Refresh
    MyString = txtTelNum.Text
    datAddSupplier.Recordset.FindFirst "[Phone No] = '" & MyString & "'"
    If datAddSupplier.Recordset.NoMatch Then
        datAddSupplier.Recordset.AddNew
        datAddSupplier.Recordset("Supplier_ID") = TheLastId + 1
        datAddSupplier.Recordset("Supplier Name") = txtSuppName.Text
        datAddSupplier.Recordset("Address 1") = txtAdd1.Text
        datAddSupplier.Recordset("Address 2") = txtAdd2.Text
        datAddSupplier.Recordset("Address 3") = txtAdd3.Text
        datAddSupplier.Recordset("Phone No") = txtTelNum.Text
        datAddSupplier.Recordset("Mobile No") = txtMobNum.Text
        datAddSupplier.Recordset("E-mail") = txtEmail.Text
        datAddSupplier.Recordset.Update
        cmdNoSave.Tag = Str(Val(cmdNoSave.Tag) + 1)
        frmPopUp.Visible = True
        cmdNoSave.Enabled = True
    Else
        Call MsgBox("This supplier already exists in the database, details will not be stored", , "Warning")
        Reset
    End If
Else
    frmInCorrect.Visible = True
End If
End Sub

Private Sub cmdBack_Click()
'exits the screen and shows the Supplier Maintenance Menu
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdNoSave_Click()
'Deletes any suppliers that were entered since the screen was last invoked
Dim counter As Integer
Dim Answer As Integer
If cmdNoSave.Tag > 0 Then
    Answer = MsgBox("Are you sure you want to delete all suppliers you have just entered?", vbYesNo, "Warning")
    If Answer = vbYes Then
        counter = 0
        datAddSupplier.RecordSource = "Select * from Supplier Order By Supplier_ID"
        datAddSupplier.Refresh
        datAddSupplier.Recordset.MoveLast
        While counter < Val(cmdNoSave.Tag)
            datAddSupplier.Recordset.Delete
            counter = counter + 1
            datAddSupplier.Recordset.MovePrevious
        Wend
    End If
End If
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdExit_Click()
'exits the screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub Form_Load()
'sets the data control to access the supplier table in the database
datAddSupplier.DatabaseName = strThePath
datAddSupplier.RecordSource = "Select * From Supplier Order By Supplier_ID"
datAddSupplier.Refresh
cmdNoSave.Tag = "0"
End Sub

Private Sub txtAdd1_Validate(Cancel As Boolean)
'makes sure there was data entered in the textbox, "txtAdd1"
If Len(txtAdd1.Text) = 0 Then
    Cancel = True
    If MsgBox("Please enter address 1", vbOKOnly, "Warning") = vbOK Then
        txtAdd1.SetFocus
    End If
Else
    Cancel = False
End If
End Sub

Private Sub txtAdd2_Validate(Cancel As Boolean)
'makes sure there was data entered in the textbox, "txtAdd2"
If Len(txtAdd2.Text) = 0 Then
    Cancel = True
    If MsgBox("Please enter address 2", vbOKOnly, "Warning") = vbOK Then
        txtAdd2.SetFocus
    End If
Else
    Cancel = False
End If
End Sub

Private Sub txtAdd3_Validate(Cancel As Boolean)
'makes sure there was data entered in the textbox, "txtAdd3"
If Len(txtAdd3.Text) = 0 Then
    Cancel = True
    If MsgBox("Please enter address 3", vbOKOnly, "Warning") = vbOK Then
        txtAdd3.SetFocus
    End If
Else
    Cancel = False
End If
End Sub

Private Sub txtSuppName_Validate(Cancel As Boolean)
'makes sure there was data entered in the textbox, "txtSuppName"
If Len(txtSuppName.Text) = 0 Then
    Cancel = True
    If MsgBox("Please enter supplier's name", vbOKOnly, "Warning") = vbOK Then
        txtSuppName.SetFocus
    End If
Else
    Cancel = False
End If
End Sub

Private Sub txtTelNum_Validate(Cancel As Boolean)
'ensures the user enters valid data into this textbox
Dim counter, CountUp As Integer
CountUp = 0
If Len(txtTelNum.Text) = 0 Then
    Cancel = True
    MsgBox "Please enter Telephone Number", vbOKOnly, "Warning"
    txtTelNum.SetFocus
Else
    For counter = 1 To Len(txtTelNum.Text)
        If IsNumeric(Mid(txtTelNum, counter, 1)) Then
                CountUp = CountUp + 1
            End If
        Next
        If (CountUp = (Len(txtTelNum.Text) - 1)) Or (CountUp = Len(txtTelNum.Text)) Then
            Cancel = False
        Else
            Cancel = True
            Call MsgBox("This is not a valid phone number", vbOKOnly, "Warning")
            txtTelNum.SetFocus
        End If
    Cancel = False
End If
End Sub

Public Sub Reset()
'Resets all the forms textboxes to nothing
txtSuppName.Text = ""
txtAdd1.Text = ""
txtAdd2.Text = ""
txtAdd3.Text = ""
txtTelNum.Text = ""
txtMobNum.Text = ""
txtEmail.Text = ""
End Sub

Private Sub txtMobNum_Validate(Cancel As Boolean)
'ensures the user enters valid data into this textbox
Dim counter, CountUp As Integer
CountUp = 0
If (Len(txtMobNum.Text) > 0) Then
    For counter = 1 To Len(txtMobNum.Text)
        If IsNumeric(Mid(txtMobNum, counter, 1)) Then
                CountUp = CountUp + 1
            End If
        Next
        If (CountUp = (Len(txtMobNum.Text) - 1)) Or (CountUp = Len(txtMobNum.Text)) Then
            Cancel = False
        Else
            Cancel = True
            Call MsgBox("This is not a valid mobile number", vbOKOnly, "Warning")
            txtMobNum.SetFocus
        End If
    Cancel = False
End If
End Sub

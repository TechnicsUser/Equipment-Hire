VERSION 5.00
Begin VB.Form frmAmViewSupp 
   Caption         =   "Amend/View Supplier"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSupplier 
      CausesValidation=   0   'False
      DataSource      =   "Data1"
      Height          =   645
      Left            =   5160
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search By:"
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
      Height          =   1920
      Left            =   3840
      TabIndex        =   28
      Top             =   1320
      Width           =   7215
      Begin VB.TextBox txtSuppName 
         DataSource      =   "Data1"
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
         Left            =   3570
         TabIndex        =   0
         Top             =   1320
         Width           =   2565
      End
      Begin VB.OptionButton OptName 
         Caption         =   "Supplier Name"
         CausesValidation=   0   'False
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
         Height          =   555
         Left            =   2160
         TabIndex        =   30
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optNumber 
         Caption         =   "Supplier Number"
         CausesValidation=   0   'False
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
         Height          =   555
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblNamesupp 
         AutoSize        =   -1  'True
         Caption         =   "Enter Suppliers Name:"
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
         TabIndex        =   33
         Top             =   1320
         Width           =   2280
      End
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         Caption         =   "Enter Suppliers Number:"
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
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   10200
      TabIndex        =   12
      Top             =   7530
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSuppNameAm 
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
      Left            =   5850
      TabIndex        =   4
      Top             =   5190
      Width           =   2565
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Changes"
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
      Left            =   10200
      TabIndex        =   11
      Top             =   6750
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data TheControl 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4005
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMainMenu 
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
      Left            =   12600
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      Left            =   10080
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "To Supplier Maintenance Menu"
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton cmdAmend 
      Caption         =   "Amend Details"
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
      Left            =   10200
      TabIndex        =   3
      Top             =   5970
      Width           =   2175
   End
   Begin VB.TextBox txtAdd1 
      DataSource      =   "Data1"
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
      Left            =   5850
      TabIndex        =   5
      Top             =   5760
      Width           =   2565
   End
   Begin VB.TextBox txtAdd2 
      DataSource      =   "Data1"
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
      Left            =   5850
      TabIndex        =   6
      Top             =   6240
      Width           =   2565
   End
   Begin VB.TextBox txtAdd3 
      DataSource      =   "Data1"
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
      Left            =   5850
      TabIndex        =   7
      Top             =   6720
      Width           =   2565
   End
   Begin VB.TextBox txtTelNum 
      DataSource      =   "Data1"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   7320
      Width           =   2565
   End
   Begin VB.TextBox txtEmail 
      DataSource      =   "Data1"
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
      Left            =   5850
      TabIndex        =   10
      Top             =   8520
      Width           =   2565
   End
   Begin VB.TextBox txtMobNum 
      DataSource      =   "Data1"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   7920
      Width           =   2565
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "View Details"
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
      Left            =   10200
      TabIndex        =   2
      Top             =   5190
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose a supplier:"
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
      TabIndex        =   34
      Top             =   3480
      Width           =   1860
   End
   Begin VB.Label lblAst5 
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
      Left            =   5610
      TabIndex        =   27
      Top             =   7320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
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
      Left            =   4485
      TabIndex        =   26
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
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
      Left            =   4485
      TabIndex        =   25
      Top             =   6240
      Width           =   1035
   End
   Begin VB.Label lblAst2 
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
      Left            =   5610
      TabIndex        =   24
      Top             =   5760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAst1 
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
      Left            =   5610
      TabIndex        =   23
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAst3 
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
      Left            =   5610
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAst4 
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
      Left            =   5610
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblAsterisk 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(Required fields marked with an asterisk)"
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
      Left            =   5250
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Supplier Name:"
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
      Left            =   3960
      TabIndex        =   19
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Amend/View Supplier"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   360
      Width           =   6825
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Left            =   4485
      TabIndex        =   15
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tele Number:"
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
      Left            =   4125
      TabIndex        =   14
      Top             =   7320
      Width           =   1395
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
      Left            =   4875
      TabIndex        =   13
      Top             =   8505
      Width           =   645
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
      Left            =   4185
      TabIndex        =   1
      Top             =   7920
      Width           =   1335
   End
End
Attribute VB_Name = "frmAmViewSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen will allows the user amend or simply view an existing supplier's details
'The user enters the supplier's id, name or browes to select a supplier.
'When found, the rest of the supplier's details are displayed in view mode. The user can then click
'the "Amend" button and amend the details. After clicking "Save Changes", and after confirming to save,
'the changes will be saved. If the user clicks "Cancel", all amendments are cancelled and the previous details
'are restored

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'MyArray    :   Variant (An array which holds the previous details of the supplier
'MyString   :   String  (holds all the previous details)
'strSQL     :   String  (holds an SQL statement)
'MyString   :   String  (Stores the supplier name)
'Answer     :   Boolean (indicates whether to save amendments)
'TheId      :   Integer (Holds a supplier id)
'counter    :   Integer (Control variable in a loop)
'CountUp    :   Integer (indicates whether the value in a textbox is numeric or not)

'Objects used -
'TheControl :   Data Control    (Used to access the supplier table in the database)


Option Explicit
Private Sub cmdAmend_Click()
'Sets the screen to amend mode, saves previous details
cmdAmend.Tag = TheControl.Recordset("Supplier Name") + "," + _
    TheControl.Recordset("Address 1") + "," + TheControl.Recordset("Address 2") + "," + _
    TheControl.Recordset("Address 3") + "," + TheControl.Recordset("Phone No") + "," + _
    TheControl.Recordset("Mobile No") + "," + TheControl.Recordset("E-mail")
txtSuppName.Enabled = False
optNumber.Enabled = False
optName.Enabled = False
lstSupplier.Enabled = False
cmdSave.Visible = True
Cmdcancel.Visible = True
lblAsterisk.Visible = True
lblAst1.Visible = True
lblAst2.Visible = True
lblAst3.Visible = True
lblAst4.Visible = True
lblAst5.Visible = True
txtSuppNameAm.Text = TheControl.Recordset("Supplier Name")
txtSuppNameAm.Enabled = True
txtAdd1.Enabled = True
txtAdd2.Enabled = True
txtAdd3.Enabled = True
txtTelNum.Enabled = True
txtMobNum.Enabled = True
txtEmail.Enabled = True
cmdBack.Enabled = False
cmdMainMenu.Enabled = False
cmdDisplay.Enabled = False
End Sub
Private Sub DisplayDetails()
'displays supplier details
txtSuppNameAm.Text = TheControl.Recordset("Supplier Name")
txtAdd1.Text = TheControl.Recordset("Address 1")
txtAdd2.Text = TheControl.Recordset("Address 2")
txtAdd3.Text = TheControl.Recordset("Address 3")
txtTelNum.Text = TheControl.Recordset("Phone No")
txtMobNum.Text = TheControl.Recordset("Mobile No")
txtEmail.Text = TheControl.Recordset("E-mail")
If optNumber.Value = True Then
    txtSuppName.Text = Str(TheControl.Recordset("Supplier_ID"))
Else
    txtSuppName.Text = TheControl.Recordset("Supplier Name")
End If
cmdAmend.Enabled = True
End Sub
Private Sub Reset1()
'Resets the screen to view mode
txtSuppName.Visible = True
lblNamesupp.Visible = True
optNumber.Enabled = True
optName.Enabled = True
lstSupplier.Enabled = True
txtSuppName.Enabled = True
cmdSave.Visible = False
Cmdcancel.Visible = False
cmdBack.Enabled = True
cmdMainMenu.Enabled = True
'cmdDisplay.Enabled = True
lblAsterisk.Visible = False
lblAst1.Visible = False
lblAst2.Visible = False
lblAst3.Visible = False
lblAst4.Visible = False
lblAst5.Visible = False
txtAdd1.Enabled = False
txtAdd2.Enabled = False
txtAdd3.Enabled = False
txtSuppNameAm.Enabled = False
txtTelNum.Enabled = False
txtMobNum.Enabled = False
txtEmail.Enabled = False
If optName.Value = True Then
    lblNamesupp.Visible = True
    lblNumber.Visible = False
Else
    lblNamesupp.Visible = False
    lblNumber.Visible = True
End If
End Sub

Private Sub cmdBack_Click()
'unloads this screen and shows the Supplier Maintenance Menu
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
'cancels any changes to the suppliers details and restores the original details
Dim MyArray
Dim MyString As String
MyString = cmdAmend.Tag
MyArray = Split(MyString, ",")
txtSuppNameAm.Text = MyArray(0)
txtAdd1.Text = MyArray(1)
txtAdd2.Text = MyArray(2)
txtAdd3.Text = MyArray(3)
txtTelNum.Text = MyArray(4)
txtMobNum.Text = MyArray(5)
txtEmail.Text = MyArray(6)
Call Reset1
End Sub

Private Sub cmdDisplay_Click()
'finds the first supplier which matches the criteria entered by the user, calls a procedure which displays the details
Dim strSQL, MyString As String
MyString = txtSuppName.Text
If optName.Value = True Then
    strSQL = "Select * From Supplier " & _
              "Where Deletion = False " & _
              "And [Supplier Name] Like '" & MyString & "'"
    txtSuppName.Text = ""
Else
    strSQL = "Select * From Supplier " & _
          "Where Deletion = False " & _
          "And Supplier_ID = " & MyString & ""
End If
TheControl.RecordSource = strSQL
TheControl.Refresh
If optName.Value = True Then
    TheControl.Recordset.FindFirst "[Supplier Name] = '" & MyString & "'"
Else
    TheControl.Recordset.FindFirst "[Supplier_ID] = " & MyString & ""
End If
If TheControl.Recordset.NoMatch Then
    frmNotFound.Visible = True
Else
    Call DisplayDetails
End If
End Sub

Private Sub cmdMainMenu_Click()
'unloads this screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdSave_Click()
'saves any amendments to the suppliers details
Dim Answer As Integer
cmdSave.Tag = TheControl.Recordset("Supplier_ID")
If (Len(txtSuppNameAm.Text) > 0) And (Len(txtAdd1.Text) > 0) And (Len(txtAdd2.Text) > 0) And (Len(txtAdd3.Text) > 0) And (Len(txtTelNum.Text) > 0) Then
    Answer = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Warning")
    If Answer = vbYes Then
        TheControl.Recordset.Edit
        TheControl.Recordset("Supplier Name") = txtSuppNameAm.Text
        TheControl.Recordset("Address 1") = txtAdd1.Text
        TheControl.Recordset("Address 2") = txtAdd2.Text
        TheControl.Recordset("Address 3") = txtAdd3.Text
        TheControl.Recordset("Phone No") = txtTelNum.Text
        TheControl.Recordset("Mobile No") = txtMobNum.Text
        TheControl.Recordset("E-mail") = txtEmail.Text
        TheControl.Recordset.Update
        Call MsgBox("The changes have been saved", , "Save Successful")
        If optName.Value = True Then
            txtSuppName.Text = TheControl.Recordset("Supplier Name")
        Else
            txtSuppName.Text = TheControl.Recordset("Supplier_ID")
        End If
        TheControl.RecordSource = "Select*from Supplier where [Deletion] = False"
        TheControl.Refresh
        lstSupplier.Clear
        While Not TheControl.Recordset.EOF
            lstSupplier.AddItem (TheControl.Recordset("Supplier Name") & " " & TheControl.Recordset("Address 1") & " " & TheControl.Recordset("Address 2") & " " & TheControl.Recordset("Phone No"))
            lstSupplier.ItemData(lstSupplier.NewIndex) = TheControl.Recordset("Supplier_ID")
            TheControl.Recordset.MoveNext
        Wend
        Call Reset1
    End If
Else
    Call MsgBox("One or more required fields left empty, details will not be stored", , "Warning")
End If
TheControl.Recordset.FindFirst "Supplier_ID = " & cmdSave.Tag & ""
End Sub

Private Sub Form_Activate()
'displays every supplier and some of their details in a listbox
txtSuppName.SetFocus
While Not TheControl.Recordset.EOF
    lstSupplier.AddItem (TheControl.Recordset("Supplier Name") & " " & TheControl.Recordset("Address 1") & " " & TheControl.Recordset("Address 2") & " " & TheControl.Recordset("Phone No"))
    lstSupplier.ItemData(lstSupplier.NewIndex) = TheControl.Recordset("Supplier_ID")
    TheControl.Recordset.MoveNext
Wend
TheControl.Refresh
End Sub

Private Sub Form_Load()
'sets the data control to access the details of the supplier table in the database
TheControl.DatabaseName = strThePath
TheControl.RecordSource = "Select*from Supplier where [Deletion] = False"
TheControl.Refresh
End Sub

Private Sub lstSupplier_Click()
'displays the details of whatever supplier that was chosen by the user from the listbox
Dim TheId As Integer
TheId = lstSupplier.ItemData(lstSupplier.ListIndex)
TheControl.RecordSource = "Select*from Supplier where [Deletion] = False;"
TheControl.Refresh
TheControl.Recordset.FindFirst "Supplier_ID = " & TheId & ""
txtSuppNameAm.Visible = True
lblName.Visible = True
cmdDisplay.Enabled = False
txtSuppName.Text = ""
Call DisplayDetails
End Sub

Private Sub OptName_Click()
'shows the label "lblNamesupp"
If optName.Value = True Then
    lblNumber.Visible = False
    lblNamesupp.Visible = True
    txtSuppName.SetFocus
End If
End Sub

Private Sub optNumber_Click()
'shows the label "lblNumber"
If optNumber.Value = True Then
    lblNumber.Visible = True
    lblNamesupp.Visible = False
    txtSuppName.SetFocus
End If
End Sub


Private Sub txtAdd1_GotFocus()
'highlights whatever text is in this textbox
txtAdd1.SelStart = 0
txtAdd1.SelLength = Len(txtAdd1)
End Sub

Private Sub txtAdd2_GotFocus()
'highlights whatever text is in this textbox
txtAdd2.SelStart = 0
txtAdd2.SelLength = Len(txtAdd2)
End Sub

Private Sub txtAdd3_GotFocus()
'highlights whatever text is in this textbox
txtAdd3.SelStart = 0
txtAdd3.SelLength = Len(txtAdd3)
End Sub

Private Sub txtEmail_GotFocus()
'highlights whatever text is in this textbox
txtEmail.SelStart = 0
txtEmail.SelLength = Len(txtEmail)
End Sub

Private Sub txtMobNum_GotFocus()
'highlights whatever text is in this textbox
txtMobNum.SelStart = 0
txtMobNum.SelLength = Len(txtMobNum)
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

Private Sub txtSuppName_GotFocus()
'empties all the supplier details out of the textboxes
txtSuppNameAm.Text = ""
txtSuppName.Text = ""
txtAdd1.Text = ""
txtAdd2.Text = ""
txtAdd3.Text = ""
txtEmail.Text = ""
txtMobNum.Text = ""
txtTelNum.Text = ""
cmdAmend.Enabled = False
cmdDisplay.Enabled = True
End Sub

Private Sub txtSuppName_Validate(Cancel As Boolean)
'makes sure that valid data is entered into this textbox
Dim counter, CountUp As Integer
CountUp = 0
If Len(txtSuppName.Text) = 0 Then
    Cancel = True
    If optName.Value = True Then
        Call MsgBox("Please enter supplier name", vbOKOnly, "Warning")
    Else
        Call MsgBox("Please enter supplier number", vbOKOnly, "Warning")
    End If
    txtSuppName.SetFocus
Else
    If optNumber = True Then
        For counter = 1 To Len(txtSuppName.Text)
            If IsNumeric(Mid(txtSuppName, counter, 1)) Then
                CountUp = CountUp + 1
            End If
        Next
        If Not (CountUp = Len(txtSuppName.Text)) Then
            Cancel = True
            Call MsgBox("Supplier ID must be numeric", vbOKOnly, "Warning")
            txtSuppName.Text = ""
        Else
            Cancel = False
        End If
    Else
        Cancel = False
    End If
End If
End Sub

Private Sub txtSuppNameAm_GotFocus()
'highlights whatever text is in this textbox
txtSuppNameAm.SelStart = 0
txtSuppNameAm.SelLength = Len(txtSuppNameAm)
End Sub

Private Sub txtTelNum_GotFocus()
'highlights whatever text is in this textbox
txtTelNum.SelStart = 0
txtTelNum.SelLength = Len(txtTelNum)
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

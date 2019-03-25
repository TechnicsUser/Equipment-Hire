VERSION 5.00
Begin VB.Form frmDelStaff 
   Caption         =   "Delete Staff "
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
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
      Left            =   9120
      TabIndex        =   24
      ToolTipText     =   "Exit screen without saving any deletions"
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Data datCount 
      Caption         =   "datCount"
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
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   1680
      Left            =   4560
      TabIndex        =   19
      Top             =   1560
      Width           =   5895
      Begin VB.TextBox txtEnteredName 
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
         Left            =   2760
         TabIndex        =   0
         ToolTipText     =   "Enter either employee name or employee number"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton optNumber 
         Caption         =   "Employee Number"
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
         Left            =   3000
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptName 
         Caption         =   "Employee Name"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         Caption         =   "Enter employee number"
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
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Enter employee name:"
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
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   2250
      End
   End
   Begin VB.ListBox lstStaff 
      CausesValidation=   0   'False
      DataSource      =   "datStaff"
      Height          =   450
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Click to view details"
      Top             =   3960
      Visible         =   0   'False
      Width           =   3975
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
      Height          =   405
      Left            =   6540
      TabIndex        =   16
      Top             =   8085
      Width           =   2415
   End
   Begin VB.Data datStaff 
      Caption         =   "datStaff"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewEmp 
      Caption         =   "Display Employee"
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
      TabIndex        =   1
      Top             =   4920
      Width           =   2175
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
      Height          =   405
      Left            =   6555
      TabIndex        =   9
      Top             =   6330
      Width           =   2415
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
      Height          =   405
      Left            =   6555
      TabIndex        =   8
      Top             =   5850
      Width           =   2415
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
      Left            =   11640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
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
      Left            =   6600
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Exit screen and save all deletions"
      Top             =   9720
      Width           =   2175
   End
   Begin VB.TextBox txtMobNum 
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
      Left            =   6555
      TabIndex        =   11
      Top             =   7530
      Width           =   2415
   End
   Begin VB.TextBox txtTelNum 
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
      Left            =   6555
      TabIndex        =   10
      Top             =   6930
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelEmp 
      Caption         =   "Delete Employee"
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
      Left            =   10560
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
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
      Height          =   405
      Left            =   6555
      TabIndex        =   7
      Top             =   5370
      Width           =   2415
   End
   Begin VB.TextBox txtEmpName 
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
      Left            =   6555
      TabIndex        =   6
      Top             =   4770
      Width           =   2415
   End
   Begin VB.Label lblMoreThan1 
      AutoSize        =   -1  'True
      Caption         =   "Multiple matches, please choose one:"
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
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Delete Staff"
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
      Left            =   5655
      TabIndex        =   18
      Top             =   480
      Width           =   3585
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email Address:"
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
      Left            =   4815
      TabIndex        =   17
      Top             =   8130
      Width           =   1530
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Mob. Number:"
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
      Left            =   4860
      TabIndex        =   15
      Top             =   7530
      Width           =   1485
   End
   Begin VB.Label Label4 
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
      Left            =   4890
      TabIndex        =   14
      Top             =   6930
      Width           =   1455
   End
   Begin VB.Label txtAdd1 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
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
      Left            =   5445
      TabIndex        =   13
      Top             =   5370
      Width           =   900
   End
   Begin VB.Label Label2 
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
      Left            =   5670
      TabIndex        =   12
      Top             =   4770
      Width           =   675
   End
End
Attribute VB_Name = "frmDelStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen allows the user to delete an existing staff member's details from the database
'The user enters the staff member's ID or name, the system then searches the Staff table for that number
'or name. When found, the rest of the staff member's details are displayed. The user then has the option
'of deleting the record from the database by clicking the "Delete Employee" button, repeat the process
'for another staff member or exit the screen without saving any deletions by clicking the "Cancel" button.

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'Answer     :   Integer (indicates whether or not to delete the employee)
'MyArray    :   Variant (An array which holds the ID's of the staff that was deleted)
'MyString   :   String  (holds all the id's of the staff that was deleted)
'strSQL     :   String  (holds an SQL statement)
'Answer     :   Boolean (indicates whether to save deletions)
'TheId      :   Integer (Holds an employee id)
'index      :   Integer (a control variable in a loop)
'MyString   :   String  (holds all the name of the employee)
'counter    :   Integer (Control variable in a loop)
'CountUp    :   Integer (indicates whether the value in a textbox is numeric or not

'Objects Used -
'datCount   :   Data Control    (Used to access the employee table in the database)
'datStaff   :   Data Control    (Used to access the employee table in the database)


Option Explicit

Private Sub cmdBack_Click()
'unloads this form and shows the Staff Maintenance Menu
frmStaffMenu.Show
Unload Me
End Sub

Public Sub DisplayDetails()
'displays the employees details
cmdDelEmp.Enabled = True
txtEmpName.Text = datStaff.Recordset("Employee Name")
txtAddress1.Text = datStaff.Recordset("Address 1")
txtAddress2.Text = datStaff.Recordset("Address 2")
txtAddress3.Text = datStaff.Recordset("Address 3")
txtTelNum.Text = datStaff.Recordset("Phone No")
txtMobNum.Text = datStaff.Recordset("Mobile No")
txtEmail.Text = datStaff.Recordset("E-mail")
cmdViewEmp.Enabled = False
End Sub

Private Sub Reset1()
'resets the screen, empties the textboxes
txtEmpName.Text = ""
txtAddress1.Text = ""
txtAddress2.Text = ""
txtAddress3.Text = ""
txtTelNum.Text = ""
txtMobNum.Text = ""
txtEnteredName.Text = ""
txtEmail.Text = ""
lstStaff.Visible = False
lblMoreThan1.Visible = False
End Sub

Private Sub cmdDelEmp_Click()
'deletes the employee from the database (marks for deletion)
Dim Answer As Integer
Answer = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Warning")
If Answer = vbYes Then
    datStaff.Recordset.Edit
    datStaff.Recordset("Deletion") = True
    cmdDelEmp.Tag = cmdDelEmp.Tag + Str(datStaff.Recordset("Employee_ID")) + ","
    frmDelStaff.Tag = Str(Val(frmDelStaff.Tag) + 1)
    datStaff.Recordset.Update
    Call MsgBox("The employee has been deleted", vbExclamation, "Deletion Sucessful")
    txtEnteredName.SetFocus
End If
End Sub

Private Sub cmdMainMenu_Click()
'unloads this form and shows the Main Menu
frmMainMenu.Visible = True
Unload Me
End Sub

Private Sub cmdNoSave_Click()
'Cancels all deletions made by the user since the screen was accessed
Dim TheArray
Dim MyString As String
Dim TheId, index, Answer As Integer
If Val(frmDelStaff.Tag) > 0 Then
    Answer = MsgBox("Are you sure you want to cancel all deletions that you made?", vbYesNo + vbQuestion, "Warning")
    If Answer = vbYes Then
        MyString = cmdDelEmp.Tag
        TheArray = Split(MyString, ",")
        datStaff.RecordSource = "Employee"
        datStaff.Refresh
        index = 0
        While index < Val(frmDelStaff.Tag)
            TheId = TheArray(index)
            datStaff.Recordset.FindFirst "Employee_ID = " & TheId & ""
            datStaff.Recordset.Edit
            datStaff.Recordset("Deletion") = False
            datStaff.Recordset.Update
            index = index + 1
        Wend
        Call MsgBox("The deletions that you made were not saved", vbOKOnly, "Deletions Cancelled")
    Else
        Call MsgBox("The changes that you made will be saved", vbOKOnly, "Deletions Saved")
    End If
End If
frmStaffMenu.Show
Unload Me
End Sub

Private Sub cmdViewEmp_Click()
'searches for the employee and when found, calls a procedure which displays the employees details
Dim MyString, strSQL As String
MyString = txtEnteredName.Text
If optName.Value = True Then
    strSQL = "Select * From Employee " & _
              "Where Deletion = False " & _
              "And [Employee Name] = '" & MyString & "'"
Else
    strSQL = "Select * From Employee " & _
          "Where Deletion = False " & _
          "And Employee_ID = " & MyString & ""
End If
datStaff.DatabaseName = strThePath
datStaff.RecordSource = strSQL
datStaff.Refresh
datCount.RecordSource = "SELECT Count(*) AS TheCount From Employee Where Deletion = False And [Employee Name] = '" & MyString & "'"
datCount.Refresh
If datCount.Recordset.Fields("TheCount") > 1 Then
    While Not datStaff.Recordset.EOF
        lstStaff.AddItem (datStaff.Recordset("Employee Name") & " " & datStaff.Recordset("Address 1"))
        lstStaff.ItemData(lstStaff.NewIndex) = datStaff.Recordset("Employee_ID")
        datStaff.Recordset.MoveNext
    Wend
    lblMoreThan1.Visible = True
    lstStaff.Visible = True
Else
    If optName.Value = True Then
        datStaff.Recordset.FindFirst "[Employee Name] = '" & MyString & "'"
    Else
        datStaff.Recordset.FindFirst "Employee_ID = " & MyString & ""
    End If
    If datStaff.Recordset.NoMatch Then
        txtEnteredName.SetFocus
        txtEnteredName.Text = ""
        If optName.Value = True Then
            Call MsgBox("Employee not found, please make sure you have the name spelled correctly and that the employee exists", vbOKOnly, "Search Failed")
        Else
            Call MsgBox("Employee not found, please make sure that the employee exists", vbOKOnly, "Search Failed")
        End If
    Else
        Call DisplayDetails
    End If
End If
datStaff.Refresh
End Sub

Private Sub Form_Activate()
'puts focus in a textbox called "txtEnteredName"
txtEnteredName.SetFocus
End Sub

Private Sub Form_Load()
'sets the data control to access the details of the employee table in the database
datCount.DatabaseName = strThePath
datCount.RecordSource = "Employee"
datCount.Refresh
cmdDelEmp.Tag = ""
frmDelStaff.Tag = "0"
End Sub

Private Sub lstStaff_Click()
'finds the employee and then calls a procedure to display its details
Dim TheId As Integer
TheId = lstStaff.ItemData(lstStaff.ListIndex)
datStaff.Recordset.FindFirst "Employee_ID = " & TheId & ""
If optNumber.Value = True Then
    txtEnteredName.Text = Str(TheId)
End If
Call DisplayDetails
End Sub

Private Sub OptName_Click()
If optName.Value = True Then
    lblNumber.Visible = False
    lblName.Visible = True
    txtEnteredName.SetFocus
End If
End Sub

Private Sub optNumber_Click()
If optNumber.Value = True Then
    lblNumber.Visible = True
    lblName.Visible = False
    txtEnteredName.SetFocus
End If
End Sub

Private Sub txtEnteredName_GotFocus()
cmdDelEmp.Enabled = False
cmdViewEmp.Enabled = True
Reset1
End Sub

Private Sub txtEnteredName_Validate(Cancel As Boolean)
'ensures valid data is entered in this textbox
Dim CountUp, counter As Integer
CountUp = 0
If Len(txtEnteredName.Text) = 0 Then
    If optName.Value = True Then
        Cancel = True
        If MsgBox("Please enter employee name", vbOKOnly, "Warning") = vbOK Then
            txtEnteredName.SetFocus
        End If
    Else
        Cancel = True
        If MsgBox("Please enter employee number", vbOKOnly, "Warning") = vbOK Then
            txtEnteredName.SetFocus
        End If
    End If
Else
    If optName.Value = True Then
        If IsNumeric(txtEnteredName.Text) Then
            Cancel = True
            txtEnteredName.SetFocus
            txtEnteredName.Text = ""
            Call MsgBox("Please enter a valid name", vbOKOnly, "Warning")
        Else
            Cancel = False
        End If
    Else
        For counter = 1 To Len(txtEnteredName.Text)
            If IsNumeric(Mid(txtEnteredName, counter, 1)) Then  'checks if the id entered is numeric
                CountUp = CountUp + 1
            End If
        Next
        If (CountUp = Len(txtEnteredName.Text)) Then
            Cancel = False
        Else
            Cancel = True
            Call MsgBox("Employee id must be numeric", vbOKOnly, "Warning")
            txtEnteredName.SetFocus
            txtEnteredName.Text = ""
        End If
    End If
End If
End Sub

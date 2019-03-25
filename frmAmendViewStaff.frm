VERSION 5.00
Begin VB.Form frmAmendViewStaff 
   Caption         =   "Amend/View Staff"
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
   Begin VB.CommandButton cmdAmend 
      Caption         =   "&Amend"
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
      Left            =   6840
      TabIndex        =   13
      ToolTipText     =   "Click to Amend"
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox txtName 
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
      Height          =   405
      Left            =   5160
      TabIndex        =   6
      Top             =   5040
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
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1935
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
      Height          =   3015
      Left            =   4560
      TabIndex        =   24
      Top             =   1440
      Width           =   7815
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
         Left            =   5880
         TabIndex        =   5
         ToolTipText     =   "Click to display previous match"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "N&ext"
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
         TabIndex        =   4
         ToolTipText     =   "Click to display next match"
         Top             =   1320
         Visible         =   0   'False
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
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Click to display staff details"
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
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Select if you wish to search by name"
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
         Left            =   3240
         TabIndex        =   1
         ToolTipText     =   "Select if you wish to search by ID"
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
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Enter staff name or number"
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
         Left            =   960
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Data datAmendView 
      Caption         =   "Employee"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   11160
      TabIndex        =   11
      Top             =   7040
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
      Left            =   11160
      TabIndex        =   12
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtPhone 
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
      Height          =   405
      Left            =   11160
      TabIndex        =   10
      Top             =   6040
      Width           =   1935
   End
   Begin VB.TextBox txtAdd3 
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
      Height          =   405
      Left            =   5160
      TabIndex        =   9
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtAdd2 
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
      Height          =   405
      Left            =   5160
      TabIndex        =   8
      Top             =   7040
      Width           =   1935
   End
   Begin VB.TextBox txtAdd1 
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
      Height          =   405
      Left            =   5160
      TabIndex        =   7
      Top             =   6040
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   10440
      TabIndex        =   15
      ToolTipText     =   "Click to ignore all amendments"
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
      TabIndex        =   14
      ToolTipText     =   "Click to exit screen and Save amendments"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdMM 
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
      Left            =   12960
      TabIndex        =   16
      ToolTipText     =   "Go to main menu"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label Label15 
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
      Left            =   480
      TabIndex        =   34
      Top             =   3840
      Width           =   3705
   End
   Begin VB.Label Label14 
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
      Left            =   3480
      TabIndex        =   33
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label13 
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
      Left            =   4080
      TabIndex        =   32
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label4 
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
      Left            =   9960
      TabIndex        =   31
      Top             =   6120
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
      Left            =   3480
      TabIndex        =   30
      Top             =   7080
      Width           =   135
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
      Left            =   3480
      TabIndex        =   29
      Top             =   8160
      Width           =   135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Employee Name"
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
      Left            =   2400
      TabIndex        =   28
      Top             =   5160
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID"
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
      TabIndex        =   27
      Top             =   5040
      Width           =   1320
   End
   Begin VB.Label Label10 
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
      Left            =   8400
      TabIndex        =   23
      Top             =   7080
      Width           =   1590
   End
   Begin VB.Label Label9 
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
      Left            =   8400
      TabIndex        =   22
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label Label8 
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
      Left            =   8400
      TabIndex        =   21
      Top             =   6120
      Width           =   1485
   End
   Begin VB.Label Label7 
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
      Left            =   2400
      TabIndex        =   20
      Top             =   8100
      Width           =   1005
   End
   Begin VB.Label Label6 
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
      Left            =   2400
      TabIndex        =   19
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label Label5 
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
      Left            =   2400
      TabIndex        =   18
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Amend/View Staff"
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
      Left            =   4935
      TabIndex        =   17
      Top             =   240
      Width           =   5685
   End
End
Attribute VB_Name = "frmAmendViewStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to make changes to an employees details

'It allows the user to search for the employee and show their details.
'It then lets the user change some of the fields.
'When user clicks "Amend" and confirms whether to Amend, the fields are updated
'using the data control "datAmendView". The original details are stored in their
'respective text boxes tag property. This method implies parallel arrays
'If the user clicks cancel at the bottom of the screen the system ignores the amendments and restores
'the staffs details to their original values that are stored in the text boxes tag.

'Author Fergal Purcell
'Date   06/03/2002

'Variables used
'strSql     :   String; This variable holds the SQL statement
'strID      :   String; Holds the Primary key of a customer
'intTab     :   Integer; Stores the location of the tab character
'intCounter :   Integer; Control variable in for loop

'Objects used
'datAmendView   :   Data Control; Used to display details of the staff

Option Explicit

Private Sub cmdAMend_Click()
If MsgBox("Are you sure you wish to amend this employee's details?", vbQuestion + vbYesNo, "Confirm deletion") = vbYes Then
    With datAmendView.Recordset
        cmdAMend.Tag = cmdAMend.Tag & txtID.Text & Chr(9)               'Store the ID's of every employee thats amended
        txtName.Tag = txtName.Tag & .Fields("Employee Name") & Chr(9)   'Store the original details of the employee
        txtAdd1.Tag = txtAdd1.Tag & .Fields("Address 1") & Chr(9)
        txtAdd2.Tag = txtAdd2.Tag & .Fields("Address 2") & Chr(9)
        txtAdd3.Tag = txtAdd3.Tag & .Fields("Address 3") & Chr(9)
        txtPhone.Tag = txtPhone.Tag & .Fields("Phone No") & Chr(9)
        txtMob.Tag = txtMob.Tag & .Fields("Mobile No") & Chr(9)
        txtMail.Tag = txtMail.Tag & .Fields("E-Mail") & Chr(9)
        .Edit
        .Fields("Employee Name") = txtName.Text                         'Amend the fields
        .Fields("Address 1") = txtAdd1.Text
        .Fields("Address 2") = txtAdd2.Text
        .Fields("Address 3") = txtAdd3.Text
        .Fields("Phone No") = txtPhone.Text
        .Fields("Mobile No") = txtMob.Text
        .Fields("E-Mail") = txtMail.Text
        .Update
    End With
Else
    AssignFields                    'Change the text boxes
End If
End Sub

Private Sub cmdCancel_Click()
If cmdAMend.Tag <> "" Then              'Only display this message if the user has made amendments
    If MsgBox("Do you want to save changes?", vbQuestion + vbYesNo, "Save changes") = vbNo Then
        CancelAmends
    End If
End If
frmStaffMenu.Show
Unload Me
End Sub

Private Sub cmdMM_Click()
If cmdAMend.Tag <> "" Then              'Only display this message if the user has made amendments
    If MsgBox("Do you wish to save amendments?", vbQuestion + vbYesNo, "Save Amendments") = vbNo Then
        CancelAmends
    End If
End If
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdNext_Click()
datAmendView.Recordset.MoveNext             'Move to the next match
If datAmendView.Recordset.EOF Then
    datAmendView.Recordset.MoveFirst
End If
AssignFields
End Sub

Private Sub cmdOk_Click()
frmStaffMenu.Show
Unload Me
End Sub

Private Sub cmdPrevious_Click()
datAmendView.Recordset.MovePrevious         'Move to the previous match
If datAmendView.Recordset.BOF Then
    datAmendView.Recordset.MoveLast
End If
AssignFields
End Sub

Private Sub cmdShow_Click()
Dim strSQL As String
cmdNext.Visible = False
cmdPrevious.Visible = False
datAmendView.DatabaseName = strThePath
datAmendView.RecordSource = "Employee"
datAmendView.Refresh
If optName = True Then                  'Search for matching name
    strSQL = "Select Count([Employee.Employee Name]) AS CountOfEmp From Employee Where Employee.[Employee Name] Like '" & txtSearch.Text & "*'" & _
                "And Employee.Deletion = False"
    datAmendView.RecordSource = strSQL
    datAmendView.Refresh
    If datAmendView.Recordset.Fields("CountOfEmp") = 0 Or txtSearch = "" Then
        MsgBox "There is no match.", vbExclamation, "No Match"
        txtSearch.SetFocus
        txtSearch.Text = ""
    ElseIf datAmendView.Recordset.Fields("CountOfEmp") = 1 Then
        strSQL = "Select * From Employee Where Employee.[Employee Name] Like '" & txtSearch.Text & "*'" & _
                    "And Employee.Deletion = False"
        datAmendView.RecordSource = strSQL
        datAmendView.Refresh
        AssignFields
    Else
        cmdNext.Visible = True              'Make the movenext and moveprevious buttons visible
        cmdPrevious.Visible = True
        strSQL = "Select * From Employee Where Employee.[Employee Name] Like '" & txtSearch.Text & "*'" & _
                    "And Employee.Deletion = False"
        datAmendView.RecordSource = strSQL
        datAmendView.Refresh
        AssignFields
    End If
Else                                                    'Search for primary keys
    If CheckNum = True Or txtSearch.Text = "" Then
        MsgBox "Invalid data.", vbExclamation, "Invalid data"
        txtSearch.SetFocus
        txtSearch.Text = ""
    Else
        datAmendView.Recordset.FindFirst "Employee_ID = " & txtSearch.Text
        If datAmendView.Recordset.NoMatch = True Or datAmendView.Recordset.Fields("Deletion") = True Then
            MsgBox "There is no match.", vbExclamation, "No Match"
            txtSearch.SetFocus
            txtSearch.Text = ""
        Else
            AssignFields
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
txtSearch.SetFocus
End Sub

Private Sub OptID_Click()
lblSearch.Caption = "Enter ID"              'Change the labels
txtSearch.Text = ""
txtSearch.SetFocus
End Sub

Private Sub optName_Click()
lblSearch.Caption = "Enter Name"
txtSearch.Text = ""
txtSearch.SetFocus
End Sub

Private Sub txtAdd1_Validate(Cancel As Boolean)
If Len(txtAdd1.Text) = 0 Then
    MsgBox "This field cannot be empty.", vbExclamation, "Empty field"
Cancel = True
End If
End Sub

Private Sub txtAdd2_Validate(Cancel As Boolean)
If Len(txtAdd2.Text) = 0 Then
    MsgBox "This field cannot be empty.", vbExclamation, "Empty field"
Cancel = True
End If
End Sub

Private Sub txtAdd3_Validate(Cancel As Boolean)
If Len(txtAdd3.Text) = 0 Then
    MsgBox "This field cannot be empty.", vbExclamation, "Empty field"
Cancel = True
End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
If Len(txtName.Text) = 0 Then
    MsgBox "This field cannot be empty.", vbExclamation, "Empty field"
    Cancel = True
End If
End Sub

Private Sub AssignFields()
'This procedure assigns all the text boxes there values

txtName.Text = datAmendView.Recordset.Fields("[Employee Name]")
txtID.Text = datAmendView.Recordset.Fields("Employee_ID")
txtAdd1.Text = datAmendView.Recordset.Fields("[Address 1]")
txtAdd2.Text = datAmendView.Recordset.Fields("[Address 2]")
txtAdd3.Text = datAmendView.Recordset.Fields("[Address 3]")
txtPhone.Text = datAmendView.Recordset.Fields("[Phone No]")
txtMob.Text = "" + datAmendView.Recordset.Fields("[Mobile No]")
txtMail.Text = "" + datAmendView.Recordset.Fields("E-Mail")
cmdAMend.Enabled = True
txtName.CausesValidation = True
txtAdd1.CausesValidation = True
txtAdd2.CausesValidation = True
txtAdd3.CausesValidation = True
txtPhone.CausesValidation = True
End Sub

Private Sub CancelAmends()
Dim strID As String, intTab As Integer
Do Until cmdAMend.Tag = ""                         'A loop that restores the original values to the database
    intTab = InStr(cmdAMend.Tag, Chr(9))           'Find the location of the tab because it marks the end of one primary key
    strID = Mid(cmdAMend.Tag, 1, intTab - 1)    'Find the primary key
    With datAmendView
        .DatabaseName = strThePath
        .RecordSource = "Employee"
        .Refresh
        .Recordset.FindFirst "Employee_ID = " & strID
        .Recordset.Edit
        
        intTab = InStr(txtName.Tag, Chr(9))                    'Find the position of the first tag
        .Recordset.Fields("Employee Name") = Mid(txtName.Tag, 1, intTab - 1)        'Restore fields to original settings
        txtName.Tag = Mid(txtName.Tag, intTab + 1, Len(txtName.Tag) - intTab)       'Delete field from tag
        
        intTab = InStr(txtAdd1.Tag, Chr(9))
        .Recordset.Fields("Address 1") = Mid(txtAdd1.Tag, 1, intTab - 1)
        txtAdd1.Tag = Mid(txtAdd1.Tag, intTab + 1, Len(txtAdd1.Tag) - intTab)
        
        intTab = InStr(txtAdd2.Tag, Chr(9))
        .Recordset.Fields("Address 2") = Mid(txtAdd2.Tag, 1, intTab - 1)
        txtAdd2.Tag = Mid(txtAdd2.Tag, intTab + 1, Len(txtAdd2.Tag) - intTab)
        
        intTab = InStr(txtAdd3.Tag, Chr(9))
        .Recordset.Fields("Address 3") = Mid(txtAdd3.Tag, 1, intTab - 1)
        txtAdd3.Tag = Mid(txtAdd3.Tag, intTab + 1, Len(txtAdd3.Tag) - intTab)
        
        intTab = InStr(txtPhone.Tag, Chr(9))
        .Recordset.Fields("Phone No") = Mid(txtPhone.Tag, 1, intTab - 1)
        txtPhone.Tag = Mid(txtPhone.Tag, intTab + 1, Len(txtPhone.Tag) - intTab)
        
        intTab = InStr(txtMob.Tag, Chr(9))
        .Recordset.Fields("Mobile No") = Mid(txtMob.Tag, 1, intTab - 1)
        txtMob.Tag = Mid(txtMob.Tag, intTab + 1, Len(txtMob.Tag) - intTab)
        
        intTab = InStr(txtMail.Tag, Chr(9))
        .Recordset.Fields("E-Mail") = Mid(txtMail.Tag, 1, intTab - 1)
        txtMail.Tag = Mid(txtMail.Tag, intTab + 1, Len(txtMail.Tag) - intTab)
        
        .Recordset.Update
    End With
    intTab = InStr(cmdAMend.Tag, Chr(9))                'Remove the first primary key
    cmdAMend.Tag = Mid(cmdAMend.Tag, intTab + 1, Len(cmdAMend.Tag) - intTab)
Loop
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
If Len(txtPhone.Text) = 0 Then
    MsgBox "This field cannot be empty.", vbExclamation, "Empty field"
    Cancel = True
End If
End Sub

Private Function CheckNum() As Boolean
'Bug in IsNumeric says that 6d6 is numeric
Dim intCounter As Integer
    For intCounter = 1 To Len(txtSearch.Text)
        If Not IsNumeric(Mid(txtSearch, intCounter, 1)) Then        'Check each character
            CheckNum = True
        End If
    Next
End Function

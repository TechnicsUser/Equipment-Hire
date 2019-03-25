VERSION 5.00
Begin VB.Form frmClockIn 
   Caption         =   "Clock In"
   ClientHeight    =   9255
   ClientLeft      =   2310
   ClientTop       =   1695
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   10380
   WindowState     =   2  'Maximized
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
      Left            =   11040
      TabIndex        =   4
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdServices 
      Caption         =   "Services"
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
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Hours Worked and clock in times for the day"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Data datCount 
      Caption         =   "Counter"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data datTime 
      Caption         =   "Time"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      ToolTipText     =   "Click to clock in"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtEmpId 
      Height          =   285
      Left            =   7920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Clock In Staff"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   4335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   8040
      X2              =   8040
      Y1              =   2880
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      X1              =   8040
      X2              =   8880
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   2535
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please enter Staff ID:"
      Height          =   195
      Left            =   6360
      TabIndex        =   1
      Top             =   5880
      Width           =   1515
   End
End
Attribute VB_Name = "frmClockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen allows the user to clock in or out, it also allows access to other services such as
'Calculating Hours Worked and Clock In Times for the day.
'The user enters his or her employee ID and clicks "OK"

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'TheLastId  :   Integer (Holds the last primary Clock-In ID in the Clock-In table)
'ID         :   Integer (holds the employee id)
'strSQL     :   String  (Holds SQL statements)
'Counter    :   Integer (Control variable in a loop)
'CountUp    :   Integer (indicates whether the value entered was numeric)


'Objects Used -
'datTime    :   Data Control    (Used to access the Employee and Clock-In tables in the database)
'datCount   :   Data Control    (Used to access the Clock-In table in the database)

Option Explicit

Private Sub cmdMainMenu_Click()
'unloads this screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdOk_Click()
'finds out whether the employee wants to clock in or out and then clocks the employee
'in or out
Dim TheLastId, ID As Integer
Dim strSQL As String
ID = Val(txtEmpId.Text)
strSQL = "Select * From Employee " & _
          "Where Deletion = False " & _
          "And Employee_ID = " & ID & ""
datTime.RecordSource = strSQL
datTime.Refresh
If (datTime.Recordset.RecordCount > 0) Then
    datTime.RecordSource = "Clock-In"
    datTime.Refresh
    If datCount.Recordset.RecordCount = 0 Then
        TheLastId = -1
    Else
        datCount.Recordset.MoveLast
        TheLastId = datCount.Recordset("ClockIn_ID")
    End If
    datCount.RecordSource = "SELECT Count(*) AS TheCount From [Clock-In] Where Employee_ID = " & ID & ""
    datCount.Refresh
    If datCount.Recordset.Fields("TheCount") > 0 Then
        datTime.RecordSource = "SELECT * From [Clock-In] Where Employee_ID = " & ID & ""
        datTime.Refresh
        datTime.Recordset.FindFirst ("[Time Out] Is Null")
        If datTime.Recordset.NoMatch Then
            PutInTable (TheLastId)
        Else
            datTime.Recordset.Edit
            datTime.Recordset("Time Out") = FormatDateTime(Now, vbShortTime)
            datTime.Recordset.Update
            datCount.RecordSource = "Clock-In"
            datCount.Refresh
        End If
    Else
        PutInTable (TheLastId)
    End If
Else
    MsgBox "This is not a valid id, please try again", vbExclamation
End If
txtEmpId.SetFocus
End Sub

Private Sub cmdServices_Click()
'hides this screen and shows the "Services" screen
frmClockIn.Hide
frmServices.Show
End Sub

Private Sub Form_Activate()
'puts the focus in the textbox "txtEmpId"
txtEmpId.SetFocus
End Sub

Private Sub Form_Load()
'sets the data control, "datCount" to access the details of the Clock-In table in the database
'sets the data control, "datTime" to access the details of the Clock-In table in the database
datCount.DatabaseName = strThePath
datCount.RecordSource = "Clock-In"
datCount.Refresh
datTime.DatabaseName = strThePath
datTime.RecordSource = "Clock-In"
datTime.Refresh
End Sub

Public Sub PutInTable(LastId As Integer)
'Enter the details into the Database
datTime.Recordset.AddNew
datTime.Recordset("Employee_ID") = txtEmpId.Text
datTime.Recordset("Time In") = FormatDateTime(Now, vbShortTime)
datTime.Recordset("Time Out") = Null
datTime.Recordset("Date") = FormatDateTime(Now, vbShortDate)
datTime.Recordset("ClockIn_ID") = LastId + 1
datTime.Recordset.Update
txtEmpId.SetFocus
End Sub

Private Sub txtEmpId_GotFocus()
'makes the screen ready for another clock in
txtEmpId.Text = ""
cmdOk.Enabled = True
datCount.RecordSource = "Clock-In"
datCount.Refresh
End Sub

Private Sub txtEmpId_Validate(Cancel As Boolean)
'ensures the user enters a valid ID
Dim counter, CountUp As Integer
If Len(txtEmpId.Text) = 0 Then
    Cancel = True
    If MsgBox("Please enter employee number", vbOKOnly, "Warning") = vbOK Then
        txtEmpId.SetFocus
    End If
Else
    For counter = 1 To Len(txtEmpId.Text)
        If IsNumeric(Mid(txtEmpId, counter, 1)) Then
            CountUp = CountUp + 1
        End If
    Next
    If Not (CountUp = Len(txtEmpId.Text)) Then
        Cancel = True
        Call MsgBox("Supplier ID must be numeric", vbOKOnly, "Warning")
        txtEmpId.Text = ""
    Else
        Cancel = False
    End If
End If
End Sub

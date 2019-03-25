VERSION 5.00
Begin VB.Form frmServices 
   Caption         =   "Services"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtClockInTime 
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
      Left            =   5115
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame frSupply 
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
      Height          =   4215
      Left            =   8880
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtTheTime 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "The time that will be entered"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cmbMinutes 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmServices.frx":0000
         Left            =   2280
         List            =   "frmServices.frx":0010
         TabIndex        =   6
         Text            =   "Please choose minute:"
         ToolTipText     =   "The minute field of the clock out time"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.ComboBox cmbHours 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmServices.frx":0024
         Left            =   120
         List            =   "frmServices.frx":0026
         TabIndex        =   5
         Text            =   "Please choose hour:"
         ToolTipText     =   "The hour field of the clock out time"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdSupplyTime 
         Caption         =   "Supply Clock-Out Time"
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
         Height          =   615
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Supply clock out time for current employee"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Click to supply the clock out time"
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "Please enter a time:"
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
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdMainMenu 
      Cancel          =   -1  'True
      Caption         =   "Main Menu"
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
      Left            =   12480
      TabIndex        =   9
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Data datTime 
      Caption         =   "Time"
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
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data datCount 
      Caption         =   "Counter"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClockIn 
      Caption         =   "Back"
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
      TabIndex        =   8
      ToolTipText     =   "To Clock In screen"
      Top             =   10080
      Width           =   2175
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Cheak all employees are out"
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
      Height          =   615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click to make sure that all employees are clocked out"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.ListBox lstStaff 
      CausesValidation=   0   'False
      Height          =   645
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   5115
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
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
      Left            =   5115
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   5115
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
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
      Left            =   5115
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   5115
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
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
      Left            =   5115
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdGetDiff 
      Caption         =   "Calculate Hours Worked"
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
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "For the week"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.ListBox lstDisplay 
      CausesValidation=   0   'False
      Height          =   1815
      Left            =   7440
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.ListBox lstDisplayTimes 
      CausesValidation=   0   'False
      Height          =   2790
      Left            =   5160
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.CommandButton cmdDisplayTimes 
      Caption         =   "Display Todays Clock-In Times"
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
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblHours 
      AutoSize        =   -1  'True
      Caption         =   "Hours worked this week:"
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
      Left            =   8760
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label lblClock 
      AutoSize        =   -1  'True
      Caption         =   "Clock-In Time:"
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
      Left            =   3375
      TabIndex        =   29
      Top             =   6840
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Click to display details:"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Services"
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
      Left            =   6300
      TabIndex        =   26
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblMobNum 
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
      Left            =   3420
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblTelNum 
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
      Left            =   3450
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAdd 
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
      Left            =   4005
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblName 
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
      Left            =   4230
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblTimes 
      AutoSize        =   -1  'True
      Caption         =   "Clock In Times for Today"
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
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   2565
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen allows the user to check that all employees are clocked out.Once all emplyees are clocked out,
'the user can then calculate the hours worked for the week for each employee. The user can also get a
'listing of clock in times for the day, the system automatically marks those employees
'who arrived late and those who left early.

'The user must first find out if all employees have clocked out. A listing of those
'who haven't clocked out, is displayed. The user clicks on their ID number to view
'their details and then has the option of supplying a clock out time for them. To do
'this, the user must first choose the hours and then the minutes from two Comboboxes
'and then click "OK".
'Once all employees are clocked out, hours worked for the week can be calculated by
'clicking "Calulate Hours Worked", and also see a list of clock in times for the day
'can be displayed by clicking "Display Todays Clock-In Times"

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'InTime         :   String  (Holds the last time the employee clocked in)
'EnteredTime    :   String  (Holds the clock out time for the employee)
'Counter        :   Integer (Control Variable in a loop)
'TheId          :   Integer (holds the employee's id)
'TheDiff        :   Integer (holds the hours worked for each employee)
'Diff           :   Currency(Holds the total amount of minutes worked for the week for each employee)
'strSQL         :   String  (Holds SQL statements)
'TheFirstHour   :   String  (Holds the Clock In time)
'TheSecHour     :   String  (Holds the Clock Out time)
'AllTheId       :   String  (Holds the Id's of all the employees that clocked in today)
'Status         :   String  (Indicates whether an employee arrived late)
'Left           :   String  (Indicates whether an employee left early)
'Temp           :   String  (Stores the variable "Left" temporarairly)
'TheArray       :   Variant (An array which holds the Id's from all the employees that clocked in today)
'TheDate        :   Date    (Holds todays date)
'ClockID        :   Integer (Holds the primary Id,"ClockIn_ID" in the Clock-In table)
'index          :   Integer (control variable in a loop)

'Objects Used -
'datTime        :   Data Control    (Used to access the Clock-In table in the database)
'datCount       :   Data Control    (Used to access the Employee and Clock-In tables in the database)

Dim EnteredTime, InTime As String
Option Explicit

Private Sub cmdClockIn_Click()
'Unloads this screen and shows the clock-in screen
Unload Me
frmClockIn.Show
End Sub

Private Sub cmdMainMenu_Click()
'Unloads this screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub cmdSupplyTime_Click()
'initializes the "cmbHours" combobox
Dim counter As Integer
counter = InTime
cmbHours.Enabled = True
While counter <= 23
    cmbHours.AddItem (Str(counter))
    counter = counter + 1
Wend
cmbHours.AddItem ("00")
lblTime.Enabled = True
End Sub

Private Sub Form_Load()
'sets both data controls to access the Clock-In table in the database
datCount.DatabaseName = strThePath
datCount.RecordSource = "Clock-In"
datCount.Refresh
datTime.DatabaseName = strThePath
datTime.RecordSource = "Clock-In"
datTime.Refresh
End Sub

Private Sub lstStaff_Click()
'finds an employee with a particular ID and then calls a procedure which displays his or her details
Dim TheId As Integer
datTime.RecordSource = "Clock-In"
datTime.Refresh
TheId = lstStaff.ItemData(lstStaff.ListIndex)
datTime.Recordset.FindLast "Employee_ID = " & TheId & ""
txtClockInTime.Text = datTime.Recordset("Time In")
InTime = Val(datTime.Recordset("Time In"))
TheId = lstStaff.ItemData(lstStaff.ListIndex)
datTime.RecordSource = "Employee"
datTime.Refresh
datTime.Recordset.FindFirst "Employee_ID = " & TheId & ""
Call DisplayDetails
End Sub

Private Sub cmdOk_Click()
'enters the clock out time for the employee, resets the screen
txtTheTime.Text = EnteredTime
Dim TheId As Integer
If IsDate(EnteredTime) Then
    TheId = lstStaff.ItemData(lstStaff.ListIndex)
    datCount.Recordset.FindFirst "Employee_ID = " & TheId & ""
    datCount.Recordset.Edit
    datCount.Recordset("Time Out") = EnteredTime
    datCount.Recordset.Update
    cmdOk.Enabled = False
    cmdSupplyTime.Enabled = False
    cmbHours.Text = "Please choose Hour"
    cmbHours.Enabled = False
    cmbMinutes.Text = "Please choose minute"
    cmbMinutes.Enabled = False
    lblTime.Enabled = False
    EnteredTime = 0
    txtTheTime.Text = ""
    lstStaff.RemoveItem (lstStaff.ListIndex)
    txtEmpName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddress3.Text = ""
    txtTelNum.Text = ""
    txtMobNum.Text = ""
    txtClockInTime.Text = ""
    If lstStaff.ListCount = 0 Then
        lstStaff.Visible = False
        lblDisplay.Visible = False
        cmdDisplayTimes.Enabled = True
        cmdGetDiff.Enabled = True
    End If
    Reset1
Else
    MsgBox "Please enter a valid time in the format hh:mm", vbExclamation
    cmbHours.SetFocus
End If
End Sub

Private Sub cmdGetDiff_Click()
'calulates the hours worked for each employee and displays it on the screen
Dim TheId, TheDiff As Integer
Dim Diff As Currency
Dim strSQL, TheFirstHour, TheSecHour As String
datCount.RecordSource = "Employee"
datCount.Refresh
datTime.RecordSource = "Clock-In"
datTime.Refresh
lblHours.Visible = True
lstDisplay.Visible = True
lstDisplay.Clear
lstDisplay.AddItem ("Employee ID" & Chr(9) & "Employee Name" & Chr(9) & "Hours Worked")
While Not datCount.Recordset.EOF
    TheId = datCount.Recordset("Employee_ID")
    strSQL = "Select * From [Clock-In] where Employee_ID = " & TheId & ""
    datTime.RecordSource = strSQL
    datTime.Refresh
    Diff = 0
    While Not datTime.Recordset.EOF
        TheFirstHour = datTime.Recordset("Time In")
        TheSecHour = datTime.Recordset("Time Out")
        Diff = Diff + DateDiff("n", TheFirstHour, TheSecHour)  'gets the total amount of minutes worked for the week
        datTime.Recordset.MoveNext
    Wend
    TheDiff = Round(Diff / 60)                                 'gets the hours worked for the week
    datCount.Recordset.Edit
    datCount.Recordset("HoursWorked") = TheDiff
    datCount.Recordset.Update
    lstDisplay.AddItem (datCount.Recordset("Employee_ID") & Chr(9) & Chr(9) & datCount.Recordset("Employee Name") & Chr(9) & Chr(9) & datCount.Recordset("HoursWorked"))
    datCount.Recordset.MoveNext
Wend
lstDisplay.Visible = True
End Sub

Private Sub cmdDisplayTimes_Click()
'Displays the clock in times for today
Dim strSQL, AllTheId, Status, Left, temp As String
Dim TheArray
Dim TheDate As Date
Dim TheId, ClockID, index As Integer
AllTheId = ""
cmdDisplayTimes.Tag = "0"
TheDate = FormatDateTime(Now, vbShortDate)
strSQL = "SELECT [Clock-In].[ClockIn_ID], [Clock-In].[Employee_ID], [Employee].[Employee Name], [Clock-In].[Time In]," & _
            "[Clock-In].[Time Out], [Clock-In].[Date]" & _
            "FROM [Clock-In] INNER JOIN Employee ON [Clock-In].[Employee_ID]=[Employee].[Employee_ID]" & _
            "Where ((([Clock-In].[Date]) = #" & TheDate & "#" & "))" & _
            "ORDER BY [Clock-In].[ClockIn_ID],[Clock-In].[Employee_ID];"    'gets everyone that clocked in today
datTime.RecordSource = strSQL
datTime.Refresh
While Not datTime.Recordset.EOF
    If InStr(AllTheId, Str(datTime.Recordset("Employee_ID"))) = False Then
        AllTheId = AllTheId + Str(datTime.Recordset("Employee_ID")) + ","    'gets the ID of each employee that clocked in today
        cmdDisplayTimes.Tag = Str(Val(cmdDisplayTimes.Tag) + 1)              'stores the number of employees that clocked in today
    End If
    datTime.Recordset.MoveNext
Wend
datTime.Refresh
TheArray = Split(AllTheId, ",")
index = 0
datCount.RecordSource = strSQL
datCount.Refresh
If datCount.Recordset.RecordCount = 0 Then
    Call MsgBox("No employees have clocked in today", vbOKOnly, "No Clock-Ins")
Else
    lblTimes.Visible = True
    lstDisplayTimes.Visible = True
    lstDisplayTimes.Clear
    lstDisplayTimes.AddItem ("Employee ID" & Chr(9) & "Employee Name" & Chr(9) & Chr(9) & "Clock In" & Chr(9) & Chr(9) & "Time Out" & Chr(9) & Chr(9) & "Arrived" & Chr(9) & Chr(9) & "Left")
    While index < Val(cmdDisplayTimes.Tag)
        TheId = TheArray(index)
        datCount.Recordset.FindFirst "Employee_ID = " & TheId & ""
        datTime.RecordSource = "Select Count(*) as Counter From [Clock-In]" & _
                                " Where [Employee_ID] = " & TheId & "" & _
                                " And Date = #" & TheDate & "#"
        datTime.Refresh
        ClockID = datCount.Recordset("ClockIn_ID")
        If (Val(datCount.Recordset("Time In")) >= Val("10:00")) Then
            Status = "Late"
        Else
            Status = "On Time"
        End If
        Left = ""
        datCount.Recordset.FindLast "Employee_ID = " & TheId & ""
        If (Val(datCount.Recordset("Time Out")) <= Val("15:30")) Then
            Left = "Early"
        Else
            Left = "On Time"
        End If
        If datTime.Recordset("Counter") = 1 Then
            lstDisplayTimes.AddItem (datCount.Recordset("Employee_ID") & Chr(9) & Chr(9) & datCount.Recordset("Employee Name") & Chr(9) & Chr(9) & datCount.Recordset("Time In") & Chr(9) & Chr(9) & datCount.Recordset("Time Out") & Chr(9) & Chr(9) & Status & Chr(9) & Chr(9) & Left)
        Else
            datCount.Recordset.FindFirst "Employee_ID = " & TheId & ""
            temp = Left
            Left = ""
            lstDisplayTimes.AddItem (datCount.Recordset("Employee_ID") & Chr(9) & Chr(9) & datCount.Recordset("Employee Name") & Chr(9) & Chr(9) & datCount.Recordset("Time In") & Chr(9) & Chr(9) & datCount.Recordset("Time Out") & Chr(9) & Chr(9) & Status & Chr(9) & Chr(9) & Left)
            datCount.Recordset.FindLast "Employee_ID = " & TheId & ""
            Left = temp
            Status = ""
            lstDisplayTimes.AddItem (datCount.Recordset("Employee_ID") & Chr(9) & Chr(9) & datCount.Recordset("Employee Name") & Chr(9) & Chr(9) & datCount.Recordset("Time In") & Chr(9) & Chr(9) & datCount.Recordset("Time Out") & Chr(9) & Chr(9) & Status & Chr(9) & Chr(9) & Left)
        End If
        index = index + 1
        datCount.Refresh
    Wend
End If
End Sub

Private Sub cmdCheck_Click()
'checks to see has all employees clocked out
lstStaff.Clear
datCount.RecordSource = "Select * From [Clock-In] Where [Time Out] Is Null"
datCount.Refresh
If datCount.Recordset.RecordCount = 0 Then
    txtEmpName.Text = ""
    txtAddress1.Text = ""
    txtAddress2.Text = ""
    txtAddress3.Text = ""
    txtTelNum.Text = ""
    txtMobNum.Text = ""
    Call MsgBox("There are no employees clocked in", vbOKOnly, "Warning")
    cmdGetDiff.Enabled = True
    cmdDisplayTimes.Enabled = True
Else
    lblDisplay.Visible = True
    lstStaff.Visible = True
    datCount.Recordset.MoveFirst
    While Not datCount.Recordset.EOF
        lstStaff.AddItem (datCount.Recordset("Employee_ID"))
        lstStaff.ItemData(lstStaff.NewIndex) = datCount.Recordset("Employee_ID")
        datCount.Recordset.MoveNext
    Wend
End If
End Sub

Public Sub DisplayDetails()
'Displays an employee's details
txtEmpName.Text = datTime.Recordset("Employee Name")
txtAddress1.Text = datTime.Recordset("Address 1")
txtAddress2.Text = datTime.Recordset("Address 2")
txtAddress3.Text = datTime.Recordset("Address 3")
txtTelNum.Text = datTime.Recordset("Phone No")
txtMobNum.Text = datTime.Recordset("Mobile No")
txtEmpName.Visible = True
txtAddress1.Visible = True
txtAddress2.Visible = True
txtAddress3.Visible = True
txtTelNum.Visible = True
txtMobNum.Visible = True
txtClockInTime.Visible = True
cmdSupplyTime.Enabled = True
lblName.Visible = True
lblAdd.Visible = True
lblTelNum.Visible = True
lblMobNum.Visible = True
lblClock.Visible = True
frSupply.Visible = True
End Sub

Private Sub cmbHours_Click()
'Gets the Hours field of the clock out time
txtTheTime.Text = ""
EnteredTime = cmbHours.Text
If cmbHours.Text = "00" Then
    txtTheTime.Text = "00:00"
    cmdOk.Enabled = True
    EnteredTime = "00:00"
Else
    cmbMinutes.Enabled = True
End If
End Sub

Private Sub cmbMinutes_Click()
'Gets the Minutes field of the clock out time
EnteredTime = EnteredTime + ":" + cmbMinutes.Text
txtTheTime.Text = EnteredTime
cmdOk.Enabled = True
cmbMinutes.Enabled = False
End Sub

Private Sub Reset1()
'Resets the screen
txtEmpName.Visible = False
txtAddress1.Visible = False
txtAddress2.Visible = False
txtAddress3.Visible = False
txtTelNum.Visible = False
txtMobNum.Visible = False
txtClockInTime.Visible = False
lblName.Visible = False
lblAdd.Visible = False
lblTelNum.Visible = False
lblMobNum.Visible = False
lblClock.Visible = False
frSupply.Visible = False
End Sub

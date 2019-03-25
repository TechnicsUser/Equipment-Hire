VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCustReport 
   Caption         =   "Customer Reports"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Please Choose Report:"
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
      Height          =   2280
      Left            =   5520
      TabIndex        =   13
      Top             =   1920
      Width           =   2535
      Begin VB.OptionButton optTown 
         Caption         =   "From a specific Town"
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
         Height          =   435
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "All customers matching the credit status from a particular town"
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optAllCust 
         Caption         =   "All Customers"
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
         Height          =   435
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "All customers matching the credit status"
         Top             =   480
         Width           =   1815
      End
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
      TabIndex        =   12
      ToolTipText     =   "To Report Menu"
      Top             =   9480
      Width           =   2175
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
      TabIndex        =   11
      Top             =   9480
      Width           =   2175
   End
   Begin VB.Frame OrderBy 
      Caption         =   "Order Report By:"
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
      Height          =   2280
      Left            =   9240
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
      Begin VB.OptionButton optCustNum 
         Caption         =   "Number"
         Enabled         =   0   'False
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
         Left            =   480
         TabIndex        =   9
         ToolTipText     =   "Show report in order of Customer Number"
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optSurName 
         Caption         =   "Surname"
         Enabled         =   0   'False
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
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Show report in order of Customer Surname"
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.TextBox txtTown 
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      ToolTipText     =   "Enter name of town"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid gridReport 
      Bindings        =   "CustReport.frx":0000
      CausesValidation=   0   'False
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColor       =   8388608
      SelectionMode   =   2
      AllowUserResizing=   3
      FormatString    =   $"CustReport.frx":0018
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Choices 
      Caption         =   "Credit Status:"
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
      Height          =   2280
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   2535
      Begin VB.OptionButton optAll 
         Caption         =   "All"
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
         Height          =   435
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Customers with any credit status"
         Top             =   1800
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optGood 
         Caption         =   "Good"
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
         Height          =   435
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Customers with a credit status of ""Good"""
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optBlack 
         Caption         =   "Blacklisted"
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
         Height          =   435
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Customers with a credit status of ""Blacklisted"""
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
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
         Height          =   435
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Customers with a credit status of ""Normal"""
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Data datReport 
      Caption         =   "datReport"
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
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblTown 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Town:"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Customer Report"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "frmCustReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This screen allows the user to produce reports on customer details.
'The user can select a Credit Status and then get a list of all customers matching it by clicking "All customers",
'or all customers from a particular town with the credit status by clicking "From a Specific town". The report can then
'be sorted by Customer Surname by clicking "SurName" or, it can be sorted by Customer Number by clicking "Number".

'Author: Sean O'Brien
'Date: 20-03-2002

'Variables Used -
'strSQL     :   String      (Holds SQL statements"
'MyString   :   String      (Holds credit status)
'TheTown    :   String      (holds the town name)
'TheName    :   String      (holds customer's full name)
'FName      :   String      (holds customer's first name)
'LName      :   String      (holds customer's last name)
'dbs        :   Database    (opens the database to create a temporary table called "Temp")
'Rec        :   Recordset   (Used to access a temporary table called "Temp", in the database)
'MyArray    :   Variant     (An array which holds the name the Surname and first name of the customer)
'counter    :   Integer     (indicates the amount of spaces in the name of the customer)

'Objects Used -
'datReport  :   Data Control    (used to access the Customer and Temp tables in the database)


Option Explicit
Private Sub cmdBack_Click()
'unloads this screen and shows the Reports Menu
frmReportsMenu.Show
Unload Me
End Sub

Public Sub Sort()
'displays all customers matching the criteria entered by the user
Dim strSQL, MyString, TheTown As String
If optGood.Value Then
    MyString = "Good"
    If optCustNum.Value Then
        If optTown.Value Then
            strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                        "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                        "Order By CustNum"
        Else
            strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                        "Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & " Order By CustNum"
        End If
    Else
        If optTown.Value Then
            strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                        "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                        "Order By LName"
        Else
            strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                        "Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & " Order By LName"
        End If
    End If
Else
    If optNormal.Value Then
        MyString = "Normal"
        If optCustNum.Value Then
            If optTown.Value Then
                strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                        "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                        "Order By CustNum"
            Else
                strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                        "Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & " Order By CustNum"
            End If
        Else
            If optTown.Value Then
                strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                        "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                        "Order By LName"
            Else
                strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                        "Add3 as [Address 3], Status From Temp " & _
                        "Where Status = '" & MyString & "'" & " Order By LName"
            End If
        End If
    Else
        If optBlack.Value = True Then
            MyString = "BlackListed"
            If optCustNum.Value Then
                If optTown.Value Then
                    strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                            "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                            "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                            "Order By CustNum"
                Else
                    strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                            "Add3 as [Address 3], Status From Temp " & _
                            "Where Status = '" & MyString & "'" & " Order By CustNum"
                End If
            Else
                If optTown.Value = True Then
                    strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                            "Add2 as [Address 2], Add3 as [Address 3], Status From Temp " & _
                            "Where Status = '" & MyString & "'" & "And Add1 = '" & txtTown.Text & "'" & _
                            "Order By LName"
                Else
                    strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1], Add2 as [Address 2]," & _
                                "Add3 as [Address 3], Status From Temp " & _
                                "Where Status = '" & MyString & "'" & " Order By LName"
                End If
            End If
        Else
            If optAll.Value = True Then
                If optTown.Value Then
                    If optCustNum.Value = True Then
                        strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                                    "Add2 as [Address 2], Add3 as [Address 3], Status From Temp Where Add1 = '" & _
                                    txtTown.Text & "'" & "Order By CustNum"
                    Else
                        strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                                    "Add2 as [Address 2], Add3 as [Address 3], Status From Temp Where Add1 = '" & _
                                    txtTown.Text & "'" & "Order By LName"
                    End If
                Else
                    If optCustNum.Value = True Then
                        strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                                    "Add2 as [Address 2], Add3 as [Address 3], Status From Temp Order By CustNum"
                    Else
                        strSQL = "Select CustNum as [Customer Number], FName as Name, Add1 as [Address 1]," & _
                                    "Add2 as [Address 2], Add3 as [Address 3], Status From Temp Order By LName"
                    End If
                End If
            End If
        End If
    End If
End If
datReport.DatabaseName = strThePath
datReport.RecordSource = strSQL
datReport.Refresh
If datReport.Recordset.RecordCount = 0 Then
    Call MsgBox("Search is complete, there was no results to display", vbOKOnly, "No Matches")
Else
    gridReport.Visible = True
End If
End Sub

Private Sub cmdViewRep_Click()
Sort
End Sub

Private Sub cmdMainMenu_Click()
'unloads this screen and shows the Main Menu
frmMainMenu.Show
Unload Me
End Sub

Private Sub Form_Activate()
OrderByNames
End Sub

Private Sub Form_Load()
'sets the data control, "datReport" to access the details of the Customer table in the database
datReport.DatabaseName = strThePath
datReport.RecordSource = "Select * from Customer where [Deletion] = False"
datReport.Refresh
gridReport.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Deletes the temporary table that was created
datReport.RecordSource = "Customer"
datReport.Refresh
Dim dbs As Database
Set dbs = OpenDatabase(strThePath)
dbs.Execute ("Drop Table Temp")
End Sub

Private Sub optAll_Click()
Reset
optTown.Value = False
optAllCust.Value = False
End Sub

Private Sub optAllCust_Click()
Reset
optTown.Value = False
optSurName.Enabled = True
optCustNum.Enabled = True
End Sub

Private Sub optBlack_Click()
Reset
optTown.Value = False
optAllCust.Value = False
End Sub

Private Sub optCustNum_Click()
Sort
End Sub

Private Sub optGood_Click()
Reset
optTown.Value = False
optAllCust.Value = False
End Sub

Private Sub optNormal_Click()
Reset
optTown.Value = False
optAllCust.Value = False
End Sub

Private Sub optSurName_Click()
Sort
End Sub

Public Sub OrderByNames()
'creates a temporary table called "Temp",for the purpose of ordering the report by SurName
'(The name of the Customer is all in the one field)
Dim strSQL, TheName, FName, LName As String
Dim dbs As Database
Dim Rec As Recordset
Dim MyArray
Dim counter As Integer
counter = 0
Set dbs = OpenDatabase(strThePath)
dbs.Execute ("CREATE TABLE Temp") _
    & ("( CustNum NUMBER, FName TEXT, LName TEXT,") _
    & ("Add1 TEXT, Add2 TEXT, Add3 TEXT, Status TEXT)")     'creates a temporary table called Temp
Set Rec = dbs.OpenRecordset("Temp", dbOpenDynaset)
strSQL = "Select Cust_ID, Name,[Address 1], [Address 2], [Address 3], Status From Customer " & _
         "Where Deletion = False "
datReport.RecordSource = strSQL
datReport.Refresh
While Not datReport.Recordset.EOF
    TheName = datReport.Recordset("Name")
    MyArray = Split(TheName, " ")               'breaks up the Name of the customer
    FName = MyArray(0)                          'first name goes into "FName"
    If Not (Len(FName) = Len(TheName)) Then     'this is true if there is only one word in the name
        LName = MyArray(1)
        counter = Len(Replace(TheName, " ", " " & "*")) - Len(TheName) 'finds the number of spaces in the Name
        If counter > 1 Then
            LName = LName + " " + MyArray(2)
        End If
        FName = FName & " " & LName
    Else
        LName = ""
    End If
    Rec.AddNew
    Rec("CustNum") = datReport.Recordset("Cust_ID")
    Rec("FName") = FName
    Rec("LName") = LName
    Rec("Add1") = datReport.Recordset("Address 1")
    Rec("Add2") = datReport.Recordset("Address 2")
    Rec("Add3") = datReport.Recordset("Address 3")
    Rec("Status") = datReport.Recordset("Status")
    Rec.Update
    datReport.Recordset.MoveNext
Wend
End Sub

Private Sub optTown_Click()
'allows the user to click the option buttons "optCustNum" and "optSurName"
optCustNum.Enabled = True
optSurName.Enabled = True
optSurName.Value = False
optCustNum.Value = False
gridReport.Visible = False
lblTown.Visible = True
txtTown.Visible = True
txtTown.SetFocus
End Sub

Private Sub txtTown_GotFocus()
txtTown.Text = ""
optCustNum.Value = False
optSurName.Value = False
gridReport.Visible = False
End Sub

Private Sub txtTown_Validate(Cancel As Boolean)
'ensures the user enters a valid town name
If Len(txtTown.Text) > 0 Then
    If Not IsNumeric(txtTown.Text) Then
        Cancel = False
    Else
        Cancel = True
        Call MsgBox("Please enter a valid town name", vbOKOnly, "Warning")
        txtTown.SetFocus
    End If
Else
    Cancel = True
    Call MsgBox("Please enter a town", vbOKOnly, "Warning")
    txtTown.SetFocus
End If
End Sub

Public Sub Reset()
'resets the screen
optSurName.Value = False
optCustNum.Value = False
optSurName.Enabled = False
optCustNum.Enabled = False
gridReport.Visible = False
lblTown.Visible = False
txtTown.Visible = False
End Sub

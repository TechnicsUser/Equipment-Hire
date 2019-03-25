VERSION 5.00
Begin VB.Form frmDeleteaSupplier 
   Caption         =   "Delete a Supplier"
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
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
      Left            =   3600
      TabIndex        =   24
      Top             =   1440
      Width           =   8415
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
         Left            =   4140
         TabIndex        =   4
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
         Left            =   6120
         TabIndex        =   5
         ToolTipText     =   "Click to display next match"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
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
         Left            =   6120
         TabIndex        =   6
         ToolTipText     =   "Click to browse for a Supplier"
         Top             =   1200
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
         ToolTipText     =   "Click to show details of Supplier"
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
         TabIndex        =   1
         ToolTipText     =   "Select to search by name"
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
         TabIndex        =   0
         ToolTipText     =   "Select to search by ID"
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
   Begin VB.Data datSupplier 
      Caption         =   "Supplier"
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   7200
      TabIndex        =   7
      ToolTipText     =   "Click to delete supplier"
      Top             =   8880
      Width           =   1575
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
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
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
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtPhone 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6040
      Width           =   1935
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
      Height          =   405
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1935
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
      Height          =   405
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7040
      Width           =   1935
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
      Height          =   405
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6040
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   10440
      TabIndex        =   10
      ToolTipText     =   "Click to exit screen and ignore deletions"
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
      TabIndex        =   8
      ToolTipText     =   "Click to exit screen "
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton cmdMM 
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
      Left            =   12960
      TabIndex        =   11
      ToolTipText     =   "Go to Main Menu"
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Left            =   2400
      TabIndex        =   29
      Top             =   5160
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Supplier ID"
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
      TabIndex        =   28
      Top             =   5040
      Width           =   1155
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
      Caption         =   "Delete Supplier"
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
      Left            =   5415
      TabIndex        =   17
      Top             =   240
      Width           =   4725
   End
End
Attribute VB_Name = "frmDeleteaSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to allow the user to mark a supplier for deletion.

'The user enters either the suppliers name or primary key, or they can browse for the
'supplier.
'When the user clicks "Show" the system finds the matching criteria in the supplier
'table and prints the suppliers details into text boxes.
'When user then clicks on "Delete", the supplier is marked for deletion and his
'primary key is concatenated into the tag property of cmdDelete followed by a space.
'If the user clicks "Cancel" the system undeletes the supplier by getting the primary
'key from the cmdDelete tag.

'Author Fergal Purcell
'Date   06/03/2002

'Variables used
'strSql     :   String ; This variable stores the Sql statement.
'intSpace   :   Integer; Holds the position of the space as it marks the end of one primary key.
'intCounter :   Integer; Control variable in for loop

'Objects used
'datSupplier    :   Data Control; This data control is used to query the "Supplier" table and mark it for deletion

Option Explicit
Private Function CheckNum() As Boolean
'Bug in IsNumeric says that 6d6 is numeric
Dim intCounter As Integer
    For intCounter = 1 To Len(txtSearch.Text)
        If Not IsNumeric(Mid(txtSearch, intCounter, 1)) Then        'Check each character
            CheckNum = True
        End If
    Next
End Function
Private Sub cmdCancel_Click()
If cmdDelete.Tag <> "" Then         'Display this message only if the user has deleted suppliers
    If MsgBox("Do you want to save deletions?", vbQuestion + vbYesNo, "Save deletions") = vbNo Then
        CancelDel
    End If
End If
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim strSQL As String
If MsgBox("Are you sure you wish to delete this supplier?", vbQuestion + vbYesNo, "Confirm deletion") = vbYes Then
    datSupplier.DatabaseName = strThePath
    strSQL = "SELECT Count(Order.Supplier_ID) As CountOfSupp " & _
              "From [Order] " & _
                "WHERE (((Order.Supplier_ID)= " & txtID.Text & ") AND ((Order.Delivered)<>True)); "
    datSupplier.RecordSource = strSQL
    datSupplier.Refresh
    If datSupplier.Recordset.Fields("CountOfSupp") > 0 Then         'if CountOfSupp is greater than zero then the supplier has delivered orders
        If MsgBox("This supplier has an undelivered order, are you sure you wish to delete?", vbExclamation + vbYesNo, "Undelivered order") = vbYes Then
           DeleteSup
           cmdDelete.Enabled = False
        Else
            strSQL = "Select * From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = False"
            datSupplier.RecordSource = strSQL
            datSupplier.Refresh
        End If
    Else
        DeleteSup
        cmdDelete.Enabled = False
    End If
End If
                          'Disable the delete button
End Sub

Private Sub cmdMM_Click()
If cmdDelete.Tag <> "" Then             'Display this message only if the user has deleted suppliers
    If MsgBox("Do you wish to save deletions?", vbQuestion + vbYesNo, "Save Deletions") = vbNo Then
        CancelDel
    End If
End If
Unload Me
frmMainMenu.Show
End Sub

Private Sub cmdNext_Click()
datSupplier.Recordset.MoveNext              'Move to the next match
If datSupplier.Recordset.EOF Then
    datSupplier.Recordset.MoveFirst
End If
AssignFields
End Sub

Private Sub cmdOk_Click()
frmSupplierMaintenanceMenu.Show
Unload Me
End Sub

Private Sub cmdPrevious_Click()
datSupplier.Recordset.MovePrevious         'Move to the previous match
If datSupplier.Recordset.BOF Then
    datSupplier.Recordset.MoveLast
End If
AssignFields
End Sub

Private Sub cmdShow_Click()
Dim strSQL As String
cmdNext.Visible = False
cmdPrevious.Visible = False
datSupplier.DatabaseName = strThePath
datSupplier.RecordSource = "Supplier"
datSupplier.Refresh
If optName = True Then                  'Search for matching name
    strSQL = "Select Count([Supplier.Supplier Name]) AS CountOfSup From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = false"
    datSupplier.RecordSource = strSQL
    datSupplier.Refresh
    If datSupplier.Recordset.Fields("CountOfSup") = 0 Or txtSearch = "" Then
        MsgBox "There is no match.", vbExclamation, "No match"
        txtSearch.SetFocus
        txtSearch.Text = ""
    ElseIf datSupplier.Recordset.Fields("CountOfSup") = 1 Then
        strSQL = "Select * From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = False"
        datSupplier.RecordSource = strSQL
        datSupplier.Refresh
        AssignFields
    Else
        cmdNext.Visible = True                  'Display these buttons if more than one match is found
        cmdPrevious.Visible = True
        strSQL = "Select * From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = False"
        datSupplier.RecordSource = strSQL
        datSupplier.Refresh
        AssignFields
    End If
Else                                                                'Search for primary keys
    If CheckNum = True Or txtSearch.Text = "" Then          'Function CheckNum returns True if the ID entered is not numeric
        MsgBox "Invalid data.", vbExclamation, "Invalid data"
        txtSearch.SetFocus
        txtSearch.Text = ""
    Else
        datSupplier.Recordset.FindFirst "Supplier_ID = " & txtSearch.Text
        If datSupplier.Recordset.NoMatch = True Or datSupplier.Recordset.Fields("Deletion") = True Then
            MsgBox "There is no match.", vbExclamation, "No match"
            txtSearch.SetFocus
            txtSearch.Text = ""
        Else
            AssignFields
        End If
    End If
End If
End Sub

Private Sub cmdBrowse_Click()
optID.Value = True              'Set the search type to ID search since browse returns the primary key
optName.Value = False
lblSearch.Caption = "Enter ID"              'Change the labels
frmBrowse.Tag = "Supplier"
frmBrowse.Show

End Sub

Private Sub Form_Activate()
txtSearch.SetFocus
End Sub

Private Sub OptID_Click()
lblSearch.Caption = "Enter ID"              'Change the labels
txtSearch.SetFocus
txtSearch.Text = ""
End Sub

Private Sub optName_Click()
lblSearch.Caption = "Enter Name"
txtSearch.Text = ""
txtSearch.SetFocus
End Sub

Private Sub AssignFields()
'Assign all the text fields
txtName.Text = datSupplier.Recordset.Fields("[Supplier Name]")
txtID.Text = datSupplier.Recordset.Fields("Supplier_ID")
txtAdd1.Text = datSupplier.Recordset.Fields("[Address 1]")
txtAdd2.Text = datSupplier.Recordset.Fields("[Address 2]")
txtAdd3.Text = datSupplier.Recordset.Fields("[Address 3]")
txtPhone.Text = datSupplier.Recordset.Fields("[Phone No]")
txtMob.Text = "" + datSupplier.Recordset.Fields("[Mobile No]")
txtMail.Text = "" + datSupplier.Recordset.Fields("E-Mail")
cmdDelete.Enabled = True
End Sub

Private Sub AssignBlank()
'Assign all the text fields blank
txtSearch.Text = ""
txtName.Text = ""
txtID.Text = ""
txtAdd1.Text = ""
txtAdd2.Text = ""
txtAdd3.Text = ""
txtPhone.Text = ""
txtMob.Text = ""
txtMail.Text = ""
End Sub

Private Sub CancelDel()
'This procedure returns the suppliers deletions fields to false
Dim intSpace As Integer
intSpace = 1
Do Until cmdDelete.Tag = ""                     'Do until all suppliers ID's have been processed
    intSpace = InStr(cmdDelete.Tag, " ")           'Find the location of the space because it marks the end of one primary key
    txtID.Text = Mid(cmdDelete.Tag, 1, intSpace - 1)    'Find the primary key
    datSupplier.DatabaseName = strThePath
    datSupplier.RecordSource = "Supplier"
    datSupplier.Refresh
    datSupplier.Recordset.FindFirst "Supplier_ID = " & txtID.Text   'Find matching key field
    datSupplier.Recordset.Edit
    datSupplier.Recordset.Fields("Deletion") = False                'Dont delete supplier
    datSupplier.Recordset.Update
    If Len(cmdDelete.Tag) <= 2 Then                     'No more supplier numbers
        cmdDelete.Tag = ""
    Else
        cmdDelete.Tag = Mid(cmdDelete.Tag, intSpace + 1, (Len(cmdDelete.Tag) - (intSpace - 1))) 'Delete the first supplier number everytime
    End If
Loop
End Sub

Private Sub DeleteSup()
Dim strSQL As String, strSearch As String
'This procedure marks a supplier for deletion
cmdDelete.Tag = cmdDelete.Tag + txtID.Text + " "      'Add the ID of the supplier to be deleted to the tag of OK
strSearch = txtSearch.Text
datSupplier.RecordSource = "Supplier"
datSupplier.Refresh
datSupplier.Recordset.FindFirst "Supplier_ID = " & txtID.Text
datSupplier.Recordset.Edit
datSupplier.Recordset.Fields("Deletion") = True         'Mark the supplier for deletion
datSupplier.Recordset.Update
'Refresh the sql statement
strSQL = "Select Count([Supplier.Supplier Name]) AS CountOfSup From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = false"
datSupplier.RecordSource = strSQL
datSupplier.Refresh
If datSupplier.Recordset.Fields("CountOfSup") = 0 Then          'Hide buttons if all matches have been deleted
    cmdNext.Visible = False
    cmdPrevious.Visible = False
End If
strSQL = "Select * From Supplier Where Supplier.[Supplier Name] Like '" & txtSearch.Text & "*'" & _
                "And Supplier.Deletion = False"
datSupplier.RecordSource = strSQL
datSupplier.Refresh
AssignBlank
txtSearch.Text = strSearch
End Sub

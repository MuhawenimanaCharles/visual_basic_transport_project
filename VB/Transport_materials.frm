VERSION 5.00
Begin VB.Form Transport_materials 
   BackColor       =   &H80000015&
   Caption         =   "Form1Form1"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Text            =   "  PLEASE THE BELLOW ID LABEL IS ONLY USED TO UPDATE AND DELETE THE DATA ONLY"
      Top             =   3840
      Width           =   7215
   End
   Begin VB.PictureBox PictureBox1 
      BackColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   7920
      ScaleHeight     =   4515
      ScaleWidth      =   5955
      TabIndex        =   15
      Top             =   960
      Width           =   6015
   End
   Begin VB.TextBox txtTransactionID 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtRoundTrip 
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   3120
      Width           =   4095
   End
   Begin VB.TextBox txtDriverName 
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txtVehicleNumber 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton read 
      BackColor       =   &H0000FFFF&
      Caption         =   "read"
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton delete 
      Caption         =   "delete record"
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton update 
      Caption         =   "update record"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton insert 
      Caption         =   "insert record"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "      ID"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "  ROUND"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "  DRIVER NAME"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "  VEHICLE NUMBER"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "       please insert transport materials records in the following blank space"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "                 WELCOME TO TRANSPORT MATERIALS PAGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Transport_materials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub delete_Click(Index As Integer)

    ' Declare a variable for the transaction ID
    Dim transaction_id As String

    ' Retrieve the transaction ID from the textbox
    transaction_id = txtTransactionID.Text

    ' Validate that the transaction ID field is not empty
    If txtTransactionID.Text = "" Then
        MsgBox "Please enter the Transaction ID to delete the record.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Confirm deletion with the user
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete this transaction?", vbQuestion + vbYesNo, "Confirm Deletion")
    If response = vbNo Then Exit Sub

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for deleting the record
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM TransportMaterials WHERE ID = ?"
        ' Append parameter
        .Parameters.Append .CreateParameter("ID", 200, 1, Len(transaction_id), transaction_id) ' 200 = adVarChar
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transaction deleted successfully!", vbInformation, "Deletion Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub


Private Sub insert_Click(Index As Integer)
    ' Declare variables for input fields
    Dim vehicle_number As String, driver_name As String, round_trip As String

    ' Retrieve input values from textboxes
    vehicle_number = Trim(txtVehicleNumber.Text)
    driver_name = Trim(txtDriverName.Text)
    round_trip = Trim(txtRoundTrip.Text)

    ' Validate required fields
    If vehicle_number = "" Or driver_name = "" Or round_trip = "" Then
        MsgBox "Please fill all fields before adding the transport material.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for inserting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO [Transportmaterials] ([Vehicle Number], [Driver Name], [Round]) VALUES (?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("Vehicle Number", 200, 1, Len(vehicle_number), vehicle_number) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("Driver Name", 200, 1, Len(driver_name), driver_name)
        .Parameters.Append .CreateParameter("Round", 200, 1, Len(round_trip), round_trip)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transport material added successfully!", vbInformation, "Registration Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    On Error Resume Next ' Avoid further errors in error handler
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtVehicleNumber.Text = ""
    txtDriverName.Text = ""
    txtRoundTrip.Text = ""
End Sub

Private Sub ExitBtn_Click()
    ' Close the form properly
    Unload Me
End Sub

End Sub

Private Sub Command2_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to EXIT?", vbYesNo + vbQuestion, "EXIT Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        login.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "EXIT Cancelled"
    End If
End Sub

Private Sub read_Click(Index As Integer)

    Dim conn As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False;"
    conn.Open

    ' Query to retrieve data
    Dim sql As String
    sql = "SELECT ID, [Vehicle Number], [Driver Name], Round FROM TransportMaterials"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "ID" & vbTab & "Vehicle Number" & vbTab & "Driver Name" & vbTab & "Round"
    PictureBox1.Print header
    PictureBox1.Line (0, 30)-(PictureBox1.ScaleWidth, 30), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("ID") & vbTab & rs("Vehicle Number") & vbTab & rs("Driver Name") & vbTab & rs("Round")
        PictureBox1.Print line
        y = y + 15
        rs.MoveNext
    Loop

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


Private Sub txtRound_Change()

End Sub

Private Sub update_Click(Index As Integer)


    ' Declare variables for the input fields
    Dim vehicle_number As String, driver_name As String, round As String, transaction_id As String

    ' Retrieve input values from textboxes
    transaction_id = txtTransactionID.Text
    vehicle_number = txtVehicleNumber.Text
    driver_name = txtDriverName.Text
    round = txtRoundTrip.Text

    ' Validate required fields
    If txtTransactionID.Text = "" Or txtVehicleNumber.Text = "" Or txtDriverName.Text = "" Or txtRoundTrip.Text = "" Then
        MsgBox "Please fill all fields before updating the transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection and command objects
    Dim conn As Object
    Dim cmd As Object
    On Error GoTo ErrorHandler ' Add error handling

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for updating data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "UPDATE Transportmaterials SET [Vehicle Number] = ?, [Driver Name] = ?, [Round] = ? WHERE ID = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("vehicle_number", 200, 1, Len(vehicle_number), vehicle_number) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("driver_name", 200, 1, Len(driver_name), driver_name)
        .Parameters.Append .CreateParameter("round", 200, 1, Len(round), round)
        .Parameters.Append .CreateParameter("transaction_id", 200, 1, Len(transaction_id), transaction_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "updated successfully!", vbInformation, "Update Complete"

    ' Clear input fields
    ClearInputFields

    ' Clean up
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    ' Handle errors
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing
End Sub


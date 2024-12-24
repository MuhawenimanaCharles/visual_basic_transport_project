VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form transation 
   BackColor       =   &H80000015&
   Caption         =   "Form1Form1"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14070
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00008000&
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Text            =   "  PLEASE THE BELLOW ID LABEL IS ONLY USED TO UPDATE AND DELETE THE DATA ONLY"
      Top             =   5760
      Width           =   7815
   End
   Begin VB.PictureBox PictureBox1 
      BackColor       =   &H0000FFFF&
      Height          =   5415
      Left            =   10560
      ScaleHeight     =   5355
      ScaleWidth      =   6675
      TabIndex        =   19
      Top             =   960
      Width           =   6735
   End
   Begin VB.TextBox txtTransactionID 
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   6480
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2040
      Top             =   7200
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Horizon_ticket­_transaction"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtModeOfPayment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox txtDepartureTime 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox txtBookingPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox txtPhoneNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox txtFullName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton read 
      Caption         =   "read all transaction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   10
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton delete 
      Caption         =   "delete record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   8160
      TabIndex        =   9
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton update 
      Caption         =   "update record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton insert 
      Caption         =   "insert record"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   8160
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "             ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "  MODE OF PAYMENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "  DEPARTURE TIME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "  BOOKING PERIOD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "  PHONE NUMBER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "   FULL NAMES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "      Put the following customer credententials in the system"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "       HORIZON BUS TICKETING MANAGEMENT TABLE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "transation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboModeOfPayment_Change()

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

    ' Prepare the SQL command for deleting the transaction
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM Horizontransaction WHERE ID = ?"
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
    Dim full_name As String, phone_number As String
    Dim booking_period As String, departure_time As String, mode_of_payment As String

    ' Retrieve input values from textboxes and combo box
    full_name = Trim(txtFullName.Text)
    phone_number = Trim(txtPhoneNumber.Text)
    booking_period = Trim(txtBookingPeriod.Text)
    departure_time = Trim(txtDepartureTime.Text)
    mode_of_payment = Trim(txtModeOfPayment.Text)

    ' Validate required fields
    If full_name = "" Or phone_number = "" Or booking_period = "" Or departure_time = "" Or mode_of_payment = "" Then
        MsgBox "Please fill all fields before adding the ticket transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate phone number (basic validation for numeric input)
    If Not IsNumeric(phone_number) Then
        MsgBox "Please enter a valid phone number.", vbExclamation, "Invalid Phone Number"
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
        .CommandText = "INSERT INTO [Horizontransaction] ([Full Names], [Phone number], [Booking Period], [Departure Time], [Mode of payment]) " & _
                       "VALUES (?, ?, ?, ?, ?)"
        ' Append parameters
        .Parameters.Append .CreateParameter("Full Names", 200, 1, Len(full_name), full_name) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("Phone number", 200, 1, Len(phone_number), phone_number)
        .Parameters.Append .CreateParameter("Booking Period", 200, 1, Len(booking_period), booking_period)
        .Parameters.Append .CreateParameter("Departure Time", 200, 1, Len(departure_time), departure_time)
        .Parameters.Append .CreateParameter("Mode of payment", 200, 1, Len(mode_of_payment), mode_of_payment)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Ticket transaction added successfully!", vbInformation, "Registration Complete"

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

Private Sub clear_Click()
    ' Clear all input fields
    ClearInputFields

    ' Notify user of cancellation
    MsgBox "Input fields have been cleared.", vbInformation, "Canceled"
End Sub

' Subroutine to clear all input fields
Private Sub ClearInputFields()
    txtFullName.Text = ""
    txtPhoneNumber.Text = ""
    txtBookingPeriod.Text = ""
    txtDepartureTime.Text = ""
    ComboModeOfPayment.Text = ""
End Sub

Private Sub ExitBtn_Click()
    ' Close the form properly
    Unload Me
End Sub
Private Sub DeleteBtn_Click()

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
        .CommandText = "DELETE FROM HorizonTransaction WHERE ID = ?"
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
    sql = "SELECT ID, [Full Names], [Phone Number], [Booking Period], [Departure Time], [Mode of Payment] FROM Horizontransaction"

    ' Execute the query
    Set rs = conn.Execute(sql)

    ' Clear the PictureBox
    PictureBox1.Cls
    PictureBox1.Font.Size = 10
    PictureBox1.Font.Name = "Arial"

    ' Display column headers
    Dim header As String
    header = "ID" & vbTab & "Full Names" & vbTab & "Phone Number" & vbTab & "Booking Period" & vbTab & "Departure Time" & vbTab & "Mode of Payment"
    PictureBox1.Print header
    PictureBox1.Line (0, 40)-(PictureBox1.ScaleWidth, 40), vbBlack

    ' Display data rows
    Dim y As Single
    y = 45
    Do While Not rs.EOF
        Dim line As String
        line = rs("ID") & vbTab & rs("Full Names") & vbTab & rs("Phone Number") & vbTab & rs("Booking Period") & vbTab & rs("Departure Time") & vbTab & rs("Mode of Payment")
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




Private Sub update_Click(Index As Integer) ' Update HorizonTransaction
    
    ' Declare variables for the input fields
    Dim full_names As String, phone_number As String, booking_period As String
    Dim departure_time As Date, mode_of_payment As String, transaction_id As String

    ' Retrieve input values from textboxes
    transaction_id = txtTransactionID.Text
    full_names = txtFullName.Text
    phone_number = txtPhoneNumber.Text
    booking_period = txtBookingPeriod.Text
    departure_time = CDate(txtDepartureTime.Text) ' Convert to Date
    mode_of_payment = txtModeOfPayment.Text

    ' Validate required fields
    If txtTransactionID.Text = "" Or txtFullName.Text = "" Or txtPhoneNumber.Text = "" Or txtBookingPeriod.Text = "" Or txtDepartureTime.Text = "" Or txtModeOfPayment.Text = "" Then
        MsgBox "Please fill all fields before updating the transaction.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Validate date field
    If Not IsDate(txtDepartureTime.Text) Then
        MsgBox "Please enter a valid departure time.", vbExclamation, "Invalid Date"
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
        .CommandText = "UPDATE Horizontransaction SET [Full Names] = ?, [Phone Number] = ?, [Booking Period] = ?, [Departure Time] = ?, [Mode of Payment] = ? WHERE ID = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("full_names", 200, 1, Len(full_names), full_names) ' 200 = adVarChar
        .Parameters.Append .CreateParameter("phone_number", 200, 1, Len(phone_number), phone_number)
        .Parameters.Append .CreateParameter("booking_period", 200, 1, Len(booking_period), booking_period)
        .Parameters.Append .CreateParameter("departure_time", 7, 1, , departure_time) ' 7 = adDate
        .Parameters.Append .CreateParameter("mode_of_payment", 200, 1, Len(mode_of_payment), mode_of_payment)
        .Parameters.Append .CreateParameter("transaction_id", 200, 1, Len(transaction_id), transaction_id)
        ' Execute the command
        .Execute
    End With

    ' Notify user of success
    MsgBox "Transaction updated successfully!", vbInformation, "Update Complete"

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


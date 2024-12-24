VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H80000015&
   Caption         =   "Form2"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4650
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1920
      TabIndex        =   4
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TRANSPORT MATERIALS"
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   5400
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "HORIZON TRANSACTION BUS TICKETS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   2
      Top             =   3960
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "     PRESS ON THE BUTTON FOR WHERE YOU WANT TO REACH"
      Height          =   975
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "               WELCOME TO HORIZON BUS TICKETING MNAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   12255
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EXIT_Click()

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

Private Sub Command1_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to reach on transaction page?", vbYesNo + vbQuestion, "EXIT Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        transation.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "EXIT Cancelled"
    End If
End Sub

Private Sub Command2_Click()

    ' Ask the user if they want to log out
    Dim response As VbMsgBoxResult
    response = MsgBox("Do you really want to reach on transport material page form?", vbYesNo + vbQuestion, "EXIT Confirmation")

    ' Check the user's response
    If response = vbYes Then
        ' Show the Welcome form if the user clicked Yes
        Transport_materials.Show
        ' Optionally, hide or close the current form
        Me.Hide ' or Me.Close if you want to completely close the current form
    Else
        ' If the user clicked No, do nothing or display a message
        MsgBox "You are still logged in.", vbInformation, "EXIT Cancelled"
    End If
End Sub
Private Sub InsertBtn_Click()
    ' Declare variables for input fields
    Dim full_name As String, phone_number As String
    Dim booking_period As String, departure_time As String, mode_of_payment As String

    ' Retrieve input values from textboxes and combo box
    full_name = Trim(txtFullName.Text)
    phone_number = Trim(txtPhoneNumber.Text)
    booking_period = Trim(txtBookingPeriod.Text)
    departure_time = Trim(txtDepartureTime.Text)
    mode_of_payment = Trim(ComboModeOfPayment.Text)

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
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USER\Desktop\DATABASE\Horizon_ticket_transaction.mdb;Persist Security Info=False;"
    conn.Open

    ' Prepare the SQL command for inserting data
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO Horizon_ticket_transaction ([Full Names], [Phone number], [Booking Period], [Departure Time], [Mode of payment]) " & _
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

Private Sub Label1_Click()

End Sub

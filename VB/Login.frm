VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "Admin"
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
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   315
      Left            =   3960
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CommandButton login 
      Caption         =   "LOGIN"
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox passwrd 
      Height          =   405
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox uname 
      Height          =   405
      Left            =   3480
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "      PASSWORD"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "      USERNAME"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "                                                  SIGN UP"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub login_Click()
    ' Declare variables for the input fields
    Dim Username As String, Password As String

    ' Retrieve input values from textboxes
    Username = uname.Text
    Password = passwrd.Text

    ' Validate if both fields are filled
    If uname.Text = "" Or passwrd.Text = "" Then
        MsgBox "Please enter both Username and Password.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Database connection string
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    ' Initialize database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=A:\VB\HORIZON_BUS_TICKETING_MS_DB.mdb;Persist Security Info=False;"
    conn.Open

    ' Create the SQL command to check for username and password
    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = conn
        .CommandText = "SELECT * FROM [Admin] WHERE Username = ? AND [Password] = ?"
        ' Append parameters
        .Parameters.Append .CreateParameter("username", 200, 1, 50, Username)
        .Parameters.Append .CreateParameter("password", 200, 1, 50, Password)
        
        ' Execute the command and get the recordset
        Set rs = .Execute
    End With

    ' Check if any records were returned (i.e., the user exists with the provided credentials)
    If Not rs.EOF Then
        MsgBox "Login successful!", vbInformation, "Login"
        Menu.Show
        Me.Hide
        
        ' You can redirect to another form or open the main application window here
        ' For example: OpenMainForm
    Else
        MsgBox "Invalid username or password. Please try again.", vbCritical, "Login Failed"
    End If

    ' Clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during login: " & Err.Description, vbCritical, "Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub

Private Sub CancelBtn_Click()
    ' Clear the username and password fields
    uname.Text = ""
    passwrd.Text = ""

    ' Optionally, you can reset the focus to the username field
    uname.SetFocus
End Sub




VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000015&
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOG IN THE SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   6120
      Width           =   9855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Caption         =   $"Index.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   3720
      Width           =   9975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "                                   WELCOME TO HORIZON BUS TICKET TRANSACTION SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   13695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
login.Show
End Sub


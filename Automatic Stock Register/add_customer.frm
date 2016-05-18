VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add customer"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_customer.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2460
      Left            =   720
      TabIndex        =   5
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Add"
      Height          =   495
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Available customers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer name"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1650
      TabIndex        =   0
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\automated_stock_register.mdb;Persist Security Info=false;"
    conn.Open
    sql = "INSERT INTO customer_list VALUES('4','" & Text1.Text & "')"
    rs.Open (sql), conn, adOpenDynamic, adLockOptimistic
    rs.Close
    conn.Close
End Sub


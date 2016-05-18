VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add supplier"
   ClientHeight    =   6540
   ClientLeft      =   5250
   ClientTop       =   1890
   ClientWidth     =   6465
   Icon            =   "add_supplier.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6465
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   600
      TabIndex        =   0
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter supplier"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Available suppliers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
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
    sql = "INSERT INTO supplier_list VALUES('4','" & Text1.Text & "')"
    rs.Open (sql), conn, adOpenDynamic, adLockOptimistic
    rs.Close
    conn.Close
End Sub


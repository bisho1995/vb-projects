VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Add category"
   ClientHeight    =   5100
   ClientLeft      =   4965
   ClientTop       =   2895
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_category.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1740
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Available Categories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter category"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add category"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Index As Integer
Private Sub Command1_Click()
    sql = "INSERT INTO category_list VALUES('" & Index + 1 & "','" & Text1.Text & "')"
    rs.Open (sql), conn, adOpenDynamic, adLockOptimistic
    conn.Close
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\automated_stock_register.mdb;Persist Security Info=false;"
    conn.Open
    rs.Open ("SELECT * FROM category_list"), conn, adOpenDynamic, adLockReadOnly
    Index = rs.RecordCount + 1
    x = MsgBox(rs.RecordCount)
    rs.Close
    'conn.Close
End Sub

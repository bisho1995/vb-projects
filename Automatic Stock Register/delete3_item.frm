VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H8000000E&
   Caption         =   "Delete Item"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   ScaleHeight     =   2295
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter the product Id to delete"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3120
   End
End
Attribute VB_Name = "Form14"
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
    sql = "DELETE FROM items WHERE product_id='" & Text1.Text & "'"
    rs.Open (sql), conn, adOpenDynamic, adLockReadOnly
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conn.Close
    'rs.Close
End Sub

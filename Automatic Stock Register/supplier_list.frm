VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form12"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form12"
   ScaleHeight     =   6285
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3900
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier List"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\automated_stock_register.mdb;Persist Security Info=false;"
    conn.Open
    sql = "SELECT * FROM supplier_list"
    rs.Open (sql), conn, adOpenDynamic, adLockOptimistic
    Dim i As Integer
    x = MsgBox(rs.RecordCount)
    'For i = 0 To rs.GetRows Step 1
    '    List1.AddItem (rs.Fields("category"))
    'Next i
    
    rs.Close
    conn.Close
End Sub



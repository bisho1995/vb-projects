VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form10"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   ScaleHeight     =   6630
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4140
      ItemData        =   "category_list.frx":0000
      Left            =   360
      List            =   "category_list.frx":0002
      TabIndex        =   1
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category List"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form10"
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
    sql = "SELECT * FROM category_list"
    rs.Open (sql), conn, adOpenDynamic, adLockOptimistic
    Dim i As Integer
    x = MsgBox(rs.RecordCount)
    'For i = 0 To rs.GetRows Step 1
    '    List1.AddItem (rs.Fields("category"))
    'Next i
    
    rs.Close
    conn.Close
End Sub

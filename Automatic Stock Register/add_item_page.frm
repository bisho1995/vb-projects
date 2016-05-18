VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Item Add"
   ClientHeight    =   7965
   ClientLeft      =   4530
   ClientTop       =   2610
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "add_item_page.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   4980
   Begin VB.ComboBox Combo4 
      Height          =   360
      ItemData        =   "add_item_page.frx":054A
      Left            =   2040
      List            =   "add_item_page.frx":054C
      TabIndex        =   17
      Top             =   4560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "add_item_page.frx":054E
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   360
      Left            =   3600
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   2880
      TabIndex        =   6
      Top             =   3840
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item cost per item"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Items purchased"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of purchase"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manufacturer"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter index no or product id"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\automated_stock_register.mdb;Persist Security Info=false;"
    conn.Open
    Dim productDate As String
    productDate = Combo1.Text & "-" & Combo2.Text & "-" & Combo3.Text
    'writing it in a date format then it will be added to rthe datavbase
    'x = MsgBox(productDate)
    sql = "INSERT INTO items VALUES('1','" & Text1.Text & "','" & Text5.Text & "','" & Text2.Text & "','" & productDate & "','" & Combo4.Text & "','" & Text3.Text & "','" & Text4.Text & "')"
    rs.Open (sql), conn, adOpenDynamic, adLockReadOnly
         
     conn.Close
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    For i = 1 To 31 Step 1
        Combo1.AddItem (i)
    Next i
    For i = 1 To 12 Step 1
        Combo2.AddItem (i)
    Next i
    For i = Year(Now) To 1921 Step -1
        Combo3.AddItem (i)
    Next i
    
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
End Sub



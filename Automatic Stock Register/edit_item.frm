VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Edit Item"
   ClientHeight    =   8430
   ClientLeft      =   4095
   ClientTop       =   1170
   ClientWidth     =   7200
   Icon            =   "edit_item.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   7200
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2760
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "edit_item.frx":054A
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "edit_item.frx":0552
      Left            =   3240
      List            =   "edit_item.frx":0554
      TabIndex        =   0
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Item"
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
      Left            =   3000
      TabIndex        =   18
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter index no or product id"
      Height          =   195
      Left            =   840
      TabIndex        =   17
      Top             =   960
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manufacturer"
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   3120
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of purchase"
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   3960
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Items purchased"
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   5280
      Width           =   1950
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item cost per item"
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   6240
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
      Height          =   195
      Left            =   840
      TabIndex        =   12
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
      Height          =   195
      Left            =   1080
      TabIndex        =   11
      Top             =   4680
      Width           =   630
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
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

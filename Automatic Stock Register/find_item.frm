VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Fine Item"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "find_item.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "find_item.frx":054A
      Left            =   3720
      List            =   "find_item.frx":054C
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   3480
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter as many details as possible"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   360
      Width           =   4665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item cost per item"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Items purchased"
      Height          =   195
      Left            =   1080
      TabIndex        =   12
      Top             =   4920
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of purchase"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manufacturer"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter index no or product id"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   1080
      Width           =   1950
   End
End
Attribute VB_Name = "Form9"
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

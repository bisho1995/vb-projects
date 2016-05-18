VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automated Stock Register"
   ClientHeight    =   8895
   ClientLeft      =   3075
   ClientTop       =   1455
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "automatic_stock_register.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14865
   Begin VB.ListBox List7 
      Height          =   6300
      Left            =   12480
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ListBox List6 
      Height          =   5820
      Left            =   10440
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   6060
      Left            =   8520
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ListBox List4 
      Height          =   5820
      Left            =   6240
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ListBox List3 
      Height          =   6060
      Left            =   3960
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   6060
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Timer splashScreen 
      Interval        =   1000
      Left            =   5520
      Top             =   8520
   End
   Begin VB.PictureBox ProgressBar1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   5025
      TabIndex        =   2
      Top             =   8520
      Width           =   5055
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   14805
      TabIndex        =   1
      Top             =   8400
      Width           =   14865
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shortcuts"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Picture         =   "automatic_stock_register.frx":054A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   1080
         Picture         =   "automatic_stock_register.frx":5479
         Stretch         =   -1  'True
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Cost"
      Height          =   240
      Left            =   12720
      TabIndex        =   16
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of items"
      Height          =   240
      Left            =   10320
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
      Height          =   240
      Left            =   8760
      TabIndex        =   14
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date of Purchase"
      Height          =   240
      Left            =   6240
      TabIndex        =   13
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manufacturer"
      Height          =   240
      Left            =   3960
      TabIndex        =   12
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
      Height          =   240
      Left            =   1920
      TabIndex        =   11
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Id"
      Height          =   240
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   900
   End
   Begin VB.Menu inventory 
      Caption         =   "&Inventory"
      Index           =   0
      Begin VB.Menu add_item 
         Caption         =   "Add Item"
         Shortcut        =   ^A
      End
      Begin VB.Menu edit_item 
         Caption         =   "Edit Item"
         Shortcut        =   ^E
      End
      Begin VB.Menu delete_item 
         Caption         =   "Delete Item"
         Shortcut        =   ^D
      End
      Begin VB.Menu find_item 
         Caption         =   "Find Item"
         Shortcut        =   ^F
      End
      Begin VB.Menu exit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu category 
      Caption         =   "&Category"
      Index           =   1
      Begin VB.Menu new_category 
         Caption         =   "New Category"
      End
      Begin VB.Menu category_list 
         Caption         =   "Open categorie list"
      End
   End
   Begin VB.Menu location 
      Caption         =   "&Location"
      Index           =   2
      Begin VB.Menu new_location 
         Caption         =   "New Location"
      End
      Begin VB.Menu location_list 
         Caption         =   "Open locations list .."
      End
   End
   Begin VB.Menu suppliers 
      Caption         =   "Suppliers"
      Index           =   3
      Begin VB.Menu new_supplier 
         Caption         =   "New supplier"
      End
      Begin VB.Menu open_supliers_list 
         Caption         =   "Open suppliers list"
      End
   End
   Begin VB.Menu customers 
      Caption         =   "C&ustomers"
      Index           =   4
      Begin VB.Menu add_customer 
         Caption         =   "Add customer"
      End
      Begin VB.Menu open_customers_list 
         Caption         =   "Open customers list .."
      End
   End
   Begin VB.Menu report 
      Caption         =   "&Report"
      Index           =   5
   End
   Begin VB.Menu about 
      Caption         =   "About"
      Index           =   6
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click(Index As Integer)
    x = MsgBox("Automated Stock Register created by Bisvarup Mukherjee roll 54 and Gaurav Ganguly roll 67 of IEM Saltlake", vbInformation, "About")
End Sub

Private Sub add_customer_Click()
    Form7.Show
End Sub

Private Sub add_item_Click()
    Form2.Show
End Sub

Private Sub category_list_Click()
    Form10.Show
End Sub

Private Sub delete_item_Click()
    Form14.Show
End Sub

Private Sub edit_item_Click()
    Form8.Show
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub find_item_Click()
    Form9.Show
End Sub

Private Sub Form_Load()
    Form3.Show
    Form3.SetFocus
    Form1.Hide
End Sub

Private Sub Image1_Click()
    Form2.Show
End Sub

Private Sub Image2_Click()
    Form8.Show
End Sub

Private Sub location_list_Click()
    Form11.Show
End Sub

Private Sub new_category_Click()
    Form4.Show
End Sub

Private Sub new_location_Click()
    Form5.Show
End Sub

Private Sub new_supplier_Click()
    Form6.Show
End Sub

Private Sub open_customers_list_Click()
    Form13.Show
End Sub

Private Sub open_supliers_list_Click()
    Form12.Show
End Sub

Private Sub splashScreen_Timer()
    Form3.Hide
    Form1.Show
    splashScreen.Enabled = False
    
End Sub

VERSION 5.00
Begin VB.Form frmChooseColor 
   Caption         =   "Choose A Color"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit and Update Form"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
   Begin VB.ListBox lstColor 
      Height          =   2205
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmChooseColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Color As Integer

Private Sub cmdExit_Click()
    frmBezierCurves.Picture1.BackColor = QBColor(Color + 1)
    frmChooseColor.Visible = False
End Sub
Private Sub Form_Load()
    initializeColorListBox
End Sub
Private Sub initializeColorListBox()
    'lstColor.AddItem("Black")
    lstColor.
    lstColor.AddItem ("Blue")
    lstColor.AddItem ("Green")
    lstColor.AddItem ("Cyan")
    lstColor.AddItem ("Red")
    lstColor.AddItem ("Magenta")
    lstColor.AddItem ("Yellow")
    lstColor.AddItem ("White")
    lstColor.AddItem ("Gray")
    lstColor.AddItem ("Light Blue")
    lstColor.AddItem ("Light Green")
    lstColor.AddItem ("Light Cyan")
    lstColor.AddItem ("Light Red")
    lstColor.AddItem ("Light Magenta")
    lstColor.AddItem ("Light Yellow")
    lstColor.AddItem ("Bright White")
End Sub
Private Sub lstColor_ItemCheck(Item As Integer)
    Color = Item
End Sub

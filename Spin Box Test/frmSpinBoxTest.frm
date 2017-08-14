VERSION 5.00
Begin VB.Form frmSpinBoxTest 
   Caption         =   "Spin Box Test"
   ClientHeight    =   4668
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   4668
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   972
   End
   Begin VB.ListBox lstSpinBoxBad 
      Height          =   240
      ItemData        =   "frmSpinBoxTest.frx":0000
      Left            =   1080
      List            =   "frmSpinBoxTest.frx":0007
      TabIndex        =   1
      Top             =   1800
      Width           =   1452
   End
   Begin VB.ListBox lstSpinBox 
      Height          =   240
      ItemData        =   "frmSpinBoxTest.frx":0017
      Left            =   1080
      List            =   "frmSpinBoxTest.frx":0019
      TabIndex        =   0
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label lblBadSpinBox 
      Caption         =   "Bad Spin Box - Does not select the correct item when you click on the up/down arrows."
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3612
   End
   Begin VB.Label lblGoodSpinBox 
      Caption         =   "Good Spin Box - Selects the correct item when you click on the up/down arrows."
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3612
   End
End
Attribute VB_Name = "frmSpinBoxTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtEnvironments(5) As String
Private Sub Form_Load()
    txtEnvironments(0) = "Test Version"
    txtEnvironments(1) = "Systems Test"
    txtEnvironments(2) = "Customer Test"
    txtEnvironments(3) = "Integrated Acceptance"
    txtEnvironments(4) = "Model Office"
    For i = 0 To 4 Step 1
        lstSpinBox.AddItem (txtEnvironments(i))
    Next i
    lstSpinBox.Selected(0) = True
    lstSpinBoxBad.RemoveItem (0)
    For i = 0 To 1 Step 1
        lstSpinBoxBad.AddItem (txtEnvironments(i))
    Next i
    lstSpinBoxBad.Selected(0) = True
End Sub
Private Sub lstSpinBox_Scroll()
    lstSpinBox.ListIndex = (lstSpinBox.ListIndex + 1) Mod (lstSpinBox.ListCount)
    'For i = 0 To lstSpinBox.ListCount - 1 Step 1
    '    If (lstSpinBox.Selected(i)) = True Then
    '        If (lstSpinBox.ListIndex) = (i - 1) Mod lstSpinBox.ListCount Then
    '            lstSpinBox.Selected(i) = False
    '            lstSpinBox.Selected((i - 1) Mod lstSpinBox.ListCount) = True
    '        Else
    '            lstSpinBox.Selected(i) = False
    '            lstSpinBox.Selected((i + 1) Mod lstSpinBox.ListCount) = True
    '
    '        End If
    '        Exit For
    '    End If
    'Next i
End Sub


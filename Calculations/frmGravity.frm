VERSION 5.00
Begin VB.Form frmGravity 
   Caption         =   "Gravity"
   ClientHeight    =   5550
   ClientLeft      =   4020
   ClientTop       =   1440
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate!"
      Height          =   615
      Left            =   1680
      TabIndex        =   15
      Top             =   4680
      Width           =   975
   End
   Begin VB.ComboBox cboUnitsM2 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cboUnitsR 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox cboUnitsResult 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cboUnitsM 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtR 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtM2 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtM 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   240
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label lblSphere2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblSphere1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   13
      Top             =   4200
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblForce 
      BackStyle       =   0  'Transparent
      Caption         =   "F = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "r:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblM2 
      Caption         =   "m:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblM 
      Caption         =   "M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   255
   End
   Begin VB.OLE oleGravityFormula 
      BackStyle       =   0  'Transparent
      Class           =   "Equation.3"
      Height          =   1095
      Left            =   1200
      OleObjectBlob   =   "frmGravity.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmGravity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim kilogrammass1 As Double
    Dim kilogrammass2 As Double
    Dim meterradius As Double
    
    Dim convm1 As frmConvUnits
    Set convm1 = New frmConvUnits
    kilogrammass1 = convm1.convert_units("mass", cboUnitsM.Text, "kilogram", CDbl(txtM.Text))
    Unload convm1
    
    Dim convm2 As frmConvUnits
    Set convm2 = New frmConvUnits
    kilogrammass2 = convm2.convert_units("mass", cboUnitsM2.Text, "kilogram", CDbl(txtM2.Text))
    Unload convm2
    
    Dim convr As frmConvUnits
    Set convr = New frmConvUnits
    meterradius = convr.convert_units("length", cboUnitsR.Text, "meter", CDbl(txtR.Text))
    Unload convr
    
    Dim G As Double
    G = 0.0000000000667
    Dim Result As Double
    If kilogrammass1 = "" Or kilogrammass2 = "" Or meterradius = "" Then
    Else
        If meterradius = 0 Then
        Else
            Result = G * kilogrammass1 * kilogrammass2 / (meterradius) ^ 2
        End If
    End If
    
    Dim convresult As frmConvUnits
    Set convresult = New frmConvUnits
    txtResult.Text = convresult.convert_units("force", "newton", cboUnitsResult.Text, Result)
    Unload convresult
End Sub

Private Sub Form_Load()
    'Initialize the two mass units combo boxes
    cboUnitsM.AddItem ("atomic mass unit")
    cboUnitsM.AddItem ("carat")
    cboUnitsM.AddItem ("grain")
    cboUnitsM.AddItem ("gram")
    cboUnitsM.AddItem ("kilogram")
    cboUnitsM.AddItem ("metric ton")
    cboUnitsM.AddItem ("MeV/c^2")
    cboUnitsM.AddItem ("ounce")
    cboUnitsM.AddItem ("pound")
    cboUnitsM.AddItem ("short ton")
    cboUnitsM.AddItem ("slug")
    cboUnitsM.AddItem ("tonne")
    cboUnitsM2.AddItem ("atomic mass unit")
    cboUnitsM2.AddItem ("carat")
    cboUnitsM2.AddItem ("grain")
    cboUnitsM2.AddItem ("gram")
    cboUnitsM2.AddItem ("kilogram")
    cboUnitsM2.AddItem ("metric ton")
    cboUnitsM2.AddItem ("MeV/c^2")
    cboUnitsM2.AddItem ("ounce")
    cboUnitsM2.AddItem ("pound")
    cboUnitsM2.AddItem ("short ton")
    cboUnitsM2.AddItem ("slug")
    cboUnitsM2.AddItem ("tonne")
    'Initialize the radius units combo box
    cboUnitsR.AddItem ("angstrom")
    cboUnitsR.AddItem ("astronomical unit")
    cboUnitsR.AddItem ("centimeter")
    cboUnitsR.AddItem ("fermi")
    cboUnitsR.AddItem ("foot")
    cboUnitsR.AddItem ("inch")
    cboUnitsR.AddItem ("kilometer")
    cboUnitsR.AddItem ("light-year")
    cboUnitsR.AddItem ("meter")
    cboUnitsR.AddItem ("micron")
    cboUnitsR.AddItem ("nautical mile")
    cboUnitsR.AddItem ("statute mile")
    cboUnitsR.AddItem ("parsec")
    cboUnitsR.AddItem ("yard")
    'Initialize the result units combo box
    cboUnitsResult.AddItem ("dyne")
    cboUnitsResult.AddItem ("kilogram force")
    cboUnitsResult.AddItem ("kilopond")
    cboUnitsResult.AddItem ("newton")
    cboUnitsResult.AddItem ("pound-force")
    cboUnitsResult.AddItem ("short ton-force")
End Sub

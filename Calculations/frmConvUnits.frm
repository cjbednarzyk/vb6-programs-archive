VERSION 5.00
Begin VB.Form frmConvUnits 
   Caption         =   "Convert Units"
   ClientHeight    =   3795
   ClientLeft      =   2865
   ClientTop       =   1845
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate!"
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtUnits2 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtUnits1 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboUnits 
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
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox cboUnits 
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
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox cboType 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Units:"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblUnits2 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblUnits1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert From:"
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
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Menu mnuPrefix 
      Caption         =   "&Prefix List"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmConvUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboType_Change()
    initializeUnitsComboBoxes
End Sub
Private Sub cboType_Click()
    cboType_Change
End Sub
Private Sub cmdCalculate_Click()
    Dim const1 As String
    Dim const2 As String
    frmConvUnits.getConstValue (0)
    frmConvUnits.getConstValue (1)
    const1 = cboUnits(0).Tag
    const2 = cboUnits(1).Tag
    If txtUnits1.Text = "" Then
    Else
        txtUnits2.Text = Format(CDbl(txtUnits1.Text) * CDbl(const2) / CDbl(const1), "Scientific")
    End If
End Sub
Private Sub Form_Load()
    initializeTypeComboBox
    frmPrefixList.Visible = True
End Sub
Private Sub initializeTypeComboBox()
    cboType.AddItem ("acceleration")
    cboType.AddItem ("angle")
    cboType.AddItem ("area")
    cboType.AddItem ("capacitance")
    cboType.AddItem ("density")
    cboType.AddItem ("electric charge")
    cboType.AddItem ("electric current")
    cboType.AddItem ("electric potential")
    cboType.AddItem ("electric field")
    cboType.AddItem ("electric resistance")
    cboType.AddItem ("electric resistivity")
    cboType.AddItem ("energy")
    cboType.AddItem ("force")
    cboType.AddItem ("inductance")
    cboType.AddItem ("length")
    cboType.AddItem ("magnetic field")
    cboType.AddItem ("magnetic flux")
    cboType.AddItem ("mass")
    cboType.AddItem ("power")
    cboType.AddItem ("pressure")
    cboType.AddItem ("radioactivity")
    cboType.AddItem ("speed")
    cboType.AddItem ("time")
    cboType.AddItem ("volume")
End Sub
Private Sub initializeUnitsComboBoxes()
    Dim I As Integer
    Select Case cboType.Text
        Case "acceleration"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("cm/s^2")
                cboUnits(I).AddItem ("ft/s^2")
                cboUnits(I).AddItem ("G")
                cboUnits(I).AddItem ("meter/s^2")
            Next
        Case "angle"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("degree")
                cboUnits(I).AddItem ("minute")
                cboUnits(I).AddItem ("radian")
                cboUnits(I).AddItem ("revolution")
                cboUnits(I).AddItem ("second")
            Next
        Case "area"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("barn")
                cboUnits(I).AddItem ("square centimeter")
                cboUnits(I).AddItem ("square foot")
                cboUnits(I).AddItem ("square inch")
                cboUnits(I).AddItem ("square kilometer")
                cboUnits(I).AddItem ("square meter")
                cboUnits(I).AddItem ("square statute mile")
                cboUnits(I).AddItem ("square yard")
            Next
        Case "capacitance"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abfarad (emu)")
                cboUnits(I).AddItem ("centimeter")
                cboUnits(I).AddItem ("farad")
                cboUnits(I).AddItem ("statfarad (esu)")
            Next
        Case "density"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("gram/cm^3")
                cboUnits(I).AddItem ("kg/cm^3")
                cboUnits(I).AddItem ("lb/ft^3")
                cboUnits(I).AddItem ("lb/gal")
                cboUnits(I).AddItem ("short ton/yd^3")
                cboUnits(I).AddItem ("slug/ft^3")
            Next
        Case "electric charge"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abcoulomb (emu)")
                cboUnits(I).AddItem ("coulomb")
                cboUnits(I).AddItem ("statcoulomb (esu)")
            Next
        Case "electric current"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abampere (emu)")
                cboUnits(I).AddItem ("ampere")
                cboUnits(I).AddItem ("statampere (esu)")
            Next
        Case "electric field"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abvolt/cm")
                cboUnits(I).AddItem ("statvolt/cm")
                cboUnits(I).AddItem ("volt/m")
            Next
        Case "electric potential"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abvolt (emu)")
                cboUnits(I).AddItem ("statvolt (esu)")
                cboUnits(I).AddItem ("volt")
            Next
        Case "electric resistance"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abohm")
                cboUnits(I).AddItem ("ohm")
                cboUnits(I).AddItem ("sec/cm")
                cboUnits(I).AddItem ("statohm")
            Next
        Case "electric resistivity"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abohm-cm")
                cboUnits(I).AddItem ("ohm-m")
                cboUnits(I).AddItem ("statohm-cm")
            Next
        Case "energy"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("Btu")
                cboUnits(I).AddItem ("calorie")
                cboUnits(I).AddItem ("erg")
                cboUnits(I).AddItem ("eV")
                cboUnits(I).AddItem ("foot-pound")
                cboUnits(I).AddItem ("joule")
                cboUnits(I).AddItem ("kilocalorie")
                cboUnits(I).AddItem ("kilowatt-hour")
            Next
        Case "force"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("dyne")
                cboUnits(I).AddItem ("kilogram force")
                cboUnits(I).AddItem ("kilopond")
                cboUnits(I).AddItem ("newton")
                cboUnits(I).AddItem ("pound-force")
                cboUnits(I).AddItem ("short ton-force")
            Next
        Case "inductance"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("abhenry (emu)")
                cboUnits(I).AddItem ("henry")
                cboUnits(I).AddItem ("sec^2/cm")
                cboUnits(I).AddItem ("stathenry (esu)")
            Next
        Case "length"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("angstrom")
                cboUnits(I).AddItem ("astronomical unit")
                cboUnits(I).AddItem ("centimeter")
                cboUnits(I).AddItem ("fermi")
                cboUnits(I).AddItem ("foot")
                cboUnits(I).AddItem ("inch")
                cboUnits(I).AddItem ("kilometer")
                cboUnits(I).AddItem ("light-year")
                cboUnits(I).AddItem ("meter")
                cboUnits(I).AddItem ("micron")
                cboUnits(I).AddItem ("nautical mile")
                cboUnits(I).AddItem ("statute mile")
                cboUnits(I).AddItem ("parsec")
                cboUnits(I).AddItem ("yard")
            Next
        Case "magnetic field"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("gauss")
                cboUnits(I).AddItem ("tesla")
                cboUnits(I).AddItem ("Weber/m^2")
            Next
        Case "magnetic flux"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("maxwell")
                cboUnits(I).AddItem ("Weber")
            Next
        Case "mass"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("atomic mass unit")
                cboUnits(I).AddItem ("carat")
                cboUnits(I).AddItem ("grain")
                cboUnits(I).AddItem ("gram")
                cboUnits(I).AddItem ("kilogram")
                cboUnits(I).AddItem ("metric ton")
                cboUnits(I).AddItem ("MeV/c^2")
                cboUnits(I).AddItem ("ounce")
                cboUnits(I).AddItem ("pound")
                cboUnits(I).AddItem ("short ton")
                cboUnits(I).AddItem ("slug")
                cboUnits(I).AddItem ("tonne")
            Next
        Case "power"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("Btu/hour")
                cboUnits(I).AddItem ("calorie/sec")
                cboUnits(I).AddItem ("erg/sec")
                cboUnits(I).AddItem ("ft*lb/sec")
                cboUnits(I).AddItem ("horsepower")
                cboUnits(I).AddItem ("kilowatt")
                cboUnits(I).AddItem ("watt")
            Next
        Case "pressure"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("atmosphere")
                cboUnits(I).AddItem ("bar")
                cboUnits(I).AddItem ("centimeter of mercury")
                cboUnits(I).AddItem ("dyne/cm^2")
                cboUnits(I).AddItem ("inch of mercury")
                cboUnits(I).AddItem ("kilopond/cm^2")
                cboUnits(I).AddItem ("millimeter of mercury")
                cboUnits(I).AddItem ("newton/m^2")
                cboUnits(I).AddItem ("pascal")
                cboUnits(I).AddItem ("pound/in^2")
                cboUnits(I).AddItem ("torr")
            Next
        Case "radioactivity"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("becquerel")
                cboUnits(I).AddItem ("Curie")
                cboUnits(I).AddItem ("decays/sec")
            Next
        Case "speed"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("c")
                cboUnits(I).AddItem ("cm/sec")
                cboUnits(I).AddItem ("ft/sec")
                cboUnits(I).AddItem ("km/hour")
                cboUnits(I).AddItem ("knot")
                cboUnits(I).AddItem ("meter/sec")
                cboUnits(I).AddItem ("nautical mile/hour")
            Next
        Case "time"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("day")
                cboUnits(I).AddItem ("hour")
                cboUnits(I).AddItem ("minute")
                cboUnits(I).AddItem ("second")
                cboUnits(I).AddItem ("sidereal day")
                cboUnits(I).AddItem ("year")
            Next
        Case "volume"
            For I = 0 To 1
                cboUnits(I).Clear
                cboUnits(I).AddItem ("cm^3")
                cboUnits(I).AddItem ("ft^3")
                cboUnits(I).AddItem ("gallon")
                cboUnits(I).AddItem ("in^3")
                cboUnits(I).AddItem ("liter")
                cboUnits(I).AddItem ("m^3")
                cboUnits(I).AddItem ("yd^3")
            Next
        Case Else
            For I = 0 To 1
                cboUnits(I).Clear
            Next
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPrefixList
End Sub

Private Sub mnuClose_Click()
    Unload frmConvUnits
End Sub
Private Sub mnuPrefix_Click()
    If frmPrefixList.Visible = True Then
        Unload frmPrefixList
    Else
        frmPrefixList.Visible = True
    End If
End Sub
Public Sub getConstValue(I As Integer)
    Dim X As Double
    Select Case cboType.Text
        Case "acceleration"
            Select Case cboUnits(I).Text
                Case "cm/s^2"
                    X = 100
                Case "ft/s^2"
                    X = 3.281
                Case "G"
                    X = 0.102
                Case "meter/s^2"
                    X = 1
            End Select
        Case "angle"
            Select Case cboUnits(I).Text
                Case "degree"
                    X = 1
                Case "minute"
                    X = 60
                Case "radian"
                    X = 0.01745
                Case "revolution"
                    X = 1 / 360
                Case "second"
                    X = 3600
            End Select
        Case "area"
            Select Case cboUnits(I).Text
                Case "barn"
                    X = 1E+28
                Case "square centimeter"
                    X = 10000#
                Case "square foot"
                    X = 10.76
                Case "square inch"
                    X = 1550#
                Case "square kilometer"
                    X = 0.000001
                Case "square meter"
                    X = 1
                Case "square statute mile"
                    X = 0.0000003861
                Case "square yard"
                    X = 1.196
            End Select
        Case "capacitance"
            Select Case cboUnits(I).Text
                Case "abfarad (emu)"
                    X = 0.000000001
                Case "centimeter"
                    X = 900000000000#
                Case "farad"
                    X = 1
                Case "statfarad (esu)"
                    X = 898800000000#
            End Select
        Case "density"
            Select Case cboUnits(I).Text
                Case "gram/cm^3"
                    X = 0.001
                Case "kg/cm^3"
                    X = 1
                Case "lb/ft^3"
                    X = 0.06243
                Case "lb/gal"
                    X = 0.008345
                Case "short ton/yd^3"
                    X = 0.0008428
                Case "slug/ft^3"
                    X = 0.00194
            End Select
        Case "electric charge"
            Select Case cboUnits(I).Text
                Case "abcoulomb (emu)"
                    X = 0.1
                Case "coulomb"
                    X = 1
                Case "statcoulomb (esu)"
                    X = 2998000000#
            End Select
        Case "electric current"
            Select Case cboUnits(I).Text
                Case "abampere (emu)"
                    X = 0.1
                Case "ampere"
                    X = 1
                Case "statampere (esu)"
                    X = 2998000000#
            End Select
        Case "electric field"
            Select Case cboUnits(I).Text
                Case "abvolt/cm"
                    X = 1000000#
                Case "statvolt/cm"
                    X = 0.00003336
                Case "volt/m"
                    X = 1
            End Select
        Case "electric potential"
            Select Case cboUnits(I).Text
                Case "abvolt (emu)"
                    X = 100000000#
                Case "statvolt (esu)"
                    X = 0.003336
                Case "volt"
                    X = 1
            End Select
        Case "electric resistance"
            Select Case cboUnits(I).Text
                Case "abohm"
                    X = 1000000000#
                Case "ohm"
                    X = 1
                Case "sec/cm"
                    X = 0.00000000001 / 9
                Case "statohm"
                    X = 0.000000000001113
            End Select
        Case "electric resistivity"
            Select Case cboUnits(I).Text
                Case "abohm-cm"
                    X = 100000000000#
                Case "ohm-m"
                    X = 1
                Case "statohm-cm"
                    X = 0.0000000001113
            End Select
        Case "energy"
            Select Case cboUnits(I).Text
                Case "Btu"
                    X = 0.0009478
                Case "calorie"
                    X = 0.2388
                Case "erg"
                    X = 10000000#
                Case "eV"
                    X = 6.242E+18
                Case "foot-pound"
                    X = 0.7376
                Case "joule"
                    X = 1
                Case "kilocalorie"
                    X = 0.0002388
                Case "kilowatt-hour"
                    X = 0.0000002778
            End Select
        Case "force"
            Select Case cboUnits(I).Text
                Case "dyne"
                    X = 100000#
                Case "kilogram force"
                    X = 0.102
                Case "kilopond"
                    X = 0.102
                Case "newton"
                    X = 1
                Case "pound-force"
                    X = 0.2248
                Case "short ton-force"
                    X = 0.0001124
            End Select
        Case "inductance"
            Select Case cboUnits(I).Text
                Case "abhenry (emu)"
                    X = 1000000000#
                Case "henry"
                    X = 1
                Case "sec^2/cm"
                    X = 0.00000000001 / 9
                Case "stathenry (esu)"
                    X = 0.000000000001113
            End Select
        Case "length"
            Select Case cboUnits(I).Text
                Case "angstrom"
                    X = 10000000000#
                Case "astronomical unit"
                    X = 0.000000000006685
                Case "centimeter"
                    X = 100
                Case "fermi"
                    X = 1E+15
                Case "foot"
                    X = 3.281
                Case "inch"
                    X = 39.37
                Case "kilometer"
                    X = 0.001
                Case "light-year"
                    X = 1.057E-16
                Case "meter"
                    X = 1
                Case "micron"
                    X = 1000000#
                Case "nautical mile"
                    X = 0.00054
                Case "statute mile"
                    X = 0.0006214
                Case "parsec"
                    X = 3.241E-17
                Case "yard"
                    X = 1.094
            End Select
        Case "magnetic field"
            Select Case cboUnits(I).Text
                Case "gauss"
                    X = 10000#
                Case "tesla"
                    X = 1
                Case "Weber/m^2"
                    X = 1
            End Select
        Case "magnetic flux"
            Select Case cboUnits(I).Text
                Case "maxwell"
                    X = 100000000#
                Case "Weber"
                    X = 1
            End Select
        Case "mass"
            Select Case cboUnits(I).Text
                Case "atomic mass unit"
                    X = 6.024E+26
                Case "carat"
                    X = 5000
                Case "grain"
                    X = 15430#
                Case "gram"
                    X = 1000
                Case "kilogram"
                    X = 1
                Case "metric ton"
                    X = 0.001
                Case "MeV/c^2"
                    X = 6.024E+26 * 931.5
                Case "ounce"
                    X = 35.27
                Case "pound"
                    X = 2.205
                Case "short ton"
                    X = 0.001102
                Case "slug"
                    X = 0.06852
                Case "tonne"
                    X = 0.001
            End Select
        Case "power"
            Select Case cboUnits(I).Text
                Case "Btu/hour"
                    X = 3.412
                Case "calorie/sec"
                    X = 0.2388
                Case "erg/sec"
                    X = 10000000#
                Case "ft*lb/sec"
                    X = 0.7376
                Case "horsepower"
                    X = 0.001341
                Case "kilowatt"
                    X = 0.001
                Case "watt"
                    X = 1
            End Select
        Case "pressure"
            Select Case cboUnits(I).Text
                Case "atmosphere"
                    X = 0.000009869
                Case "bar"
                    X = 0.00001
                Case "centimeter of mercury"
                    X = 0.0007501
                Case "dyne/cm^2"
                    X = 10
                Case "inch of mercury"
                    X = 0.0002953
                Case "kilopond/cm^2"
                    X = 0.102
                Case "millimeter of mercury"
                    X = 0.007501
                Case "newton/m^2"
                    X = 1
                Case "pascal"
                    X = 1
                Case "pound/in^2"
                    X = 0.000145
                Case "torr"
                    X = 0.007501
            End Select
        Case "radioactivity"
            Select Case cboUnits(I).Text
                Case "becquerel"
                    X = 1
                Case "Curie"
                    X = 1 / 37000000000#
                Case "decays/sec"
                    X = 1
            End Select
        Case "speed"
            Select Case cboUnits(I).Text
                Case "c"
                    X = 1 / 299800000#
                Case "cm/sec"
                    X = 100
                Case "ft/sec"
                    X = 3.281
                Case "km/hour"
                    X = 3.6
                Case "knot"
                    X = 1.944
                Case "meter/sec"
                    X = 1
                Case "nautical mile/hour"
                    X = 1 / 0.5144
            End Select
        Case "time"
            Select Case cboUnits(I).Text
                Case "day"
                    X = 0.00001157
                Case "hour"
                    X = 1 / 3600
                Case "minute"
                    X = 1 / 60
                Case "second"
                    X = 1
                Case "sidereal day"
                    X = 0.00001161
                Case "year"
                    X = 0.00000003169
            End Select
        Case "volume"
            Select Case cboUnits(I).Text
                Case "cm^3"
                    X = 1000000#
                Case "ft^3"
                    X = 35.31
                Case "gallon"
                    X = 264.2
                Case "in^3"
                    X = 61020#
                Case "liter"
                    X = 1000#
                Case "m^3"
                    X = 1
                Case "yd^3"
                    X = 1.308
            End Select
    End Select
    cboUnits(I).Tag = X
End Sub
Public Function convert_units(UnitType As String, ConvertFrom As String, ConvertTo As String, ConvertNumber As Double) As Double
    cboType.Text = UnitType
    cboUnits(0).Text = ConvertFrom
    cboUnits(1).Text = ConvertTo
    txtUnits1.Text = ConvertNumber
    cmdCalculate.Value = True
    convert_units = CDbl(txtUnits.Text)
End Function

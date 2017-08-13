VERSION 5.00
Begin VB.Form frmPrefixList 
   Caption         =   "List of Acceptable Prefixes"
   ClientHeight    =   3930
   ClientLeft      =   8235
   ClientTop       =   1695
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   3465
   Begin VB.ListBox lstPrefixes 
      BackColor       =   &H80000004&
      Height          =   3180
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.ListBox lstPrefixes 
      BackColor       =   &H80000004&
      Height          =   3180
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.ListBox lstPrefixes 
      BackColor       =   &H80000004&
      Height          =   3180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblSymbol 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Symbol"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblPrefix 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prefix"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblFactor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Powers of Ten"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrefixList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    initializeListBoxes
    lstPrefixes(0).Visible = True
End Sub
Private Sub initializeListBoxes()
    lstPrefixes(0).AddItem ("18")
    lstPrefixes(0).AddItem ("15")
    lstPrefixes(0).AddItem ("12")
    lstPrefixes(0).AddItem ("9")
    lstPrefixes(0).AddItem ("6")
    lstPrefixes(0).AddItem ("3")
    lstPrefixes(0).AddItem ("2")
    lstPrefixes(0).AddItem ("1")
    lstPrefixes(0).AddItem ("-1")
    lstPrefixes(0).AddItem ("-2")
    lstPrefixes(0).AddItem ("-3")
    lstPrefixes(0).AddItem ("-6")
    lstPrefixes(0).AddItem ("-9")
    lstPrefixes(0).AddItem ("-12")
    lstPrefixes(0).AddItem ("-15")
    lstPrefixes(0).AddItem ("-18")
    lstPrefixes(1).AddItem ("exa")
    lstPrefixes(1).AddItem ("peta")
    lstPrefixes(1).AddItem ("tera")
    lstPrefixes(1).AddItem ("giga")
    lstPrefixes(1).AddItem ("mega")
    lstPrefixes(1).AddItem ("kilo")
    lstPrefixes(1).AddItem ("hecto")
    lstPrefixes(1).AddItem ("deka")
    lstPrefixes(1).AddItem ("deci")
    lstPrefixes(1).AddItem ("centi")
    lstPrefixes(1).AddItem ("milli")
    lstPrefixes(1).AddItem ("micro")
    lstPrefixes(1).AddItem ("nano")
    lstPrefixes(1).AddItem ("pico")
    lstPrefixes(1).AddItem ("femto")
    lstPrefixes(1).AddItem ("atto")
    lstPrefixes(2).AddItem ("E")
    lstPrefixes(2).AddItem ("P")
    lstPrefixes(2).AddItem ("T")
    lstPrefixes(2).AddItem ("G")
    lstPrefixes(2).AddItem ("M")
    lstPrefixes(2).AddItem ("k")
    lstPrefixes(2).AddItem ("h")
    lstPrefixes(2).AddItem ("da")
    lstPrefixes(2).AddItem ("d")
    lstPrefixes(2).AddItem ("c")
    lstPrefixes(2).AddItem ("m")
    lstPrefixes(2).AddItem ("mu symbol")
    lstPrefixes(2).AddItem ("n")
    lstPrefixes(2).AddItem ("p")
    lstPrefixes(2).AddItem ("f")
    lstPrefixes(2).AddItem ("a")
End Sub


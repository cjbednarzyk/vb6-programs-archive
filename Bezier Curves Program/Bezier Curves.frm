VERSION 5.00
Begin VB.Form frmBezierCurves 
   Caption         =   "Draw Bezier Curves"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hscNumPoints 
      Height          =   375
      Left            =   2160
      Max             =   100
      Min             =   1
      TabIndex        =   9
      Top             =   6120
      Value           =   100
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      Height          =   3975
      Left            =   1080
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   3915
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
      Begin VB.Label lblPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "P3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "P2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   7
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "P0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "P1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.OLE olePoint1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Paint.Picture"
         DragMode        =   1  'Automatic
         Height          =   75
         Index           =   1
         Left            =   480
         OleObjectBlob   =   "Bezier Curves.frx":0000
         OLETypeAllowed  =   1  'Embedded
         TabIndex        =   4
         Top             =   1800
         Width           =   75
      End
      Begin VB.OLE olePoint1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Paint.Picture"
         DragMode        =   1  'Automatic
         Height          =   75
         Index           =   3
         Left            =   2760
         OleObjectBlob   =   "Bezier Curves.frx":36018
         OLETypeAllowed  =   1  'Embedded
         TabIndex        =   3
         Top             =   1080
         Width           =   75
      End
      Begin VB.OLE olePoint1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Paint.Picture"
         DragMode        =   1  'Automatic
         Height          =   75
         Index           =   2
         Left            =   3360
         OleObjectBlob   =   "Bezier Curves.frx":6C030
         OLETypeAllowed  =   1  'Embedded
         TabIndex        =   2
         Top             =   2160
         Width           =   75
      End
      Begin VB.OLE olePoint1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   0  'None
         Class           =   "Paint.Picture"
         DragMode        =   1  'Automatic
         Height          =   75
         Index           =   0
         Left            =   840
         OleObjectBlob   =   "Bezier Curves.frx":A2048
         OLETypeAllowed  =   1  'Embedded
         TabIndex        =   1
         Top             =   960
         Width           =   75
      End
   End
   Begin VB.Label lblShowNumPoints 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label lblNumPoints 
      Caption         =   "Number of Points in Curve"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Menu mnuVis 
      Caption         =   "&Visibility Properties"
      Begin VB.Menu mnuTanVis 
         Caption         =   "Make &Tangent Lines Invisible"
      End
      Begin VB.Menu mnuCtlPts 
         Caption         =   "Make &Control Points Invisible"
      End
      Begin VB.Menu mnuLabVis 
         Caption         =   "Make &Labels Invisible"
      End
   End
   Begin VB.Menu mnuEnd 
      Caption         =   "&Exit Program"
   End
End
Attribute VB_Name = "frmBezierCurves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tanVis As Boolean
Private Sub Form_Load()
    tanVis = True
    Call Picture1_DragDrop(olePoint1(0), olePoint1(0).Left, olePoint1(0).Top)
End Sub
Private Sub hscNumPoints_Change()
    lblShowNumPoints.Caption = CStr(hscNumPoints.Value)
    Call Picture1_DragDrop(olePoint1(0), olePoint1(0).Left, olePoint1(0).Top)
End Sub
Private Sub mnuEnd_Click()
    End
End Sub
Private Sub mnuLabVis_Click()
    If mnuLabVis.Caption = "Make &Labels Visible" Then
        mnuLabVis.Caption = "Make &Labels Invisible"
        For i = 0 To 3 Step 1
            lblPoint(i).Visible = True
        Next
    Else
        mnuLabVis.Caption = "Make &Labels Visible"
        For i = 0 To 3 Step 1
            lblPoint(i).Visible = False
        Next
    End If
End Sub
Private Sub mnuCtlPts_Click()
    If mnuCtlPts.Caption = "Make &Control Points Invisible" Then
        mnuCtlPts.Caption = "Make &Control Points Visible"
        For i = 0 To 3 Step 1
            olePoint1(i).Visible = False
        Next
    Else
        mnuCtlPts.Caption = "Make &Control Points Invisible"
        For i = 0 To 3 Step 1
            olePoint1(i).Visible = True
        Next
    End If
End Sub
Private Sub mnuTanVis_Click()
    If mnuTanVis.Caption = "Make &Tangent Lines Invisible" Then
        mnuTanVis.Caption = "Make &Tangent Lines Visible"
        tanVis = False
    Else
        mnuTanVis.Caption = "Make &Tangent Lines Invisible"
        tanVis = True
    End If
    Call Picture1_DragDrop(olePoint1(0), olePoint1(0).Left, olePoint1(0).Top)
End Sub
Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    Picture1.Cls
    Source.Left = X
    Source.Top = Y
    X0 = olePoint1(0).Left
    X1 = olePoint1(1).Left
    X2 = olePoint1(2).Left
    X3 = olePoint1(3).Left
    Y0 = olePoint1(0).Top
    Y1 = olePoint1(1).Top
    Y2 = olePoint1(2).Top
    Y3 = olePoint1(3).Top
    Separation = hscNumPoints.Value
    Call DrawBSpline(X0, X1, X2, X3, Y0, Y1, Y2, Y3, Separation)
    Call ctlPointMove(Source)
    If tanVis = True Then
        Call tanDraw(X0, X1, X2, X3, Y0, Y1, Y2, Y3)
    End If
End Sub
Sub DrawBSpline(X0, X1, X2, X3, Y0, Y1, Y2, Y3, Separation)
    For t = 0 To 1 Step (1 / Separation)
        Q1 = (1 - t) ^ 3 * X0 + 3 * t * (1 - t) ^ 2 * X1 + 3 * t ^ 2 * (1 - t) * X2 + t ^ 3 * X3
        Q2 = (1 - t) ^ 3 * Y0 + 3 * t * (1 - t) ^ 2 * Y1 + 3 * t ^ 2 * (1 - t) * Y2 + t ^ 3 * Y3
        u = t + (1 / Separation)
        Q3 = (1 - u) ^ 3 * X0 + 3 * u * (1 - u) ^ 2 * X1 + 3 * u ^ 2 * (1 - u) * X2 + u ^ 3 * X3
        Q4 = (1 - u) ^ 3 * Y0 + 3 * u * (1 - u) ^ 2 * Y1 + 3 * u ^ 2 * (1 - u) * Y2 + u ^ 3 * Y3
        Red = 256 * t
        Blue = 256 * (1 - t)
        Green = 256 * (Abs(0.5 - t))
        ErrorValue = 10000000
        If u < 1 + 1 / ErrorValue * Separation Then
            Picture1.Line (Q1, Q2)-(Q3, Q4), RGB(Red, Blue, Green)
        End If
    Next
End Sub
Sub tanDraw(X0, X1, X2, X3, Y0, Y1, Y2, Y3)
        Picture1.Line (X0, Y0)-(X1, Y1), QBColor(1)
        Picture1.Line (X2, Y2)-(X3, Y3), QBColor(1)
End Sub
Sub ctlPointMove(Source)
    lblPoint(Source.Index).Left = olePoint1(Source.Index).Left + 100
    lblPoint(Source.Index).Top = olePoint1(Source.Index).Top + 100
End Sub

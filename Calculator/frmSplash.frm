VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7245
   ClientLeft      =   600
   ClientTop       =   360
   ClientWidth     =   9690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   9345
      Begin VB.Timer tmrSplash 
         Interval        =   1000
         Left            =   4440
         Top             =   600
      End
      Begin VB.Image imgLogo 
         Height          =   2430
         Left            =   3720
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   6420
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   3
         Top             =   6630
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   " Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   6600
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8250
         TabIndex        =   5
         Top             =   6060
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7860
         TabIndex        =   6
         Top             =   5700
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3360
         TabIndex        =   8
         Top             =   1860
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3120
         TabIndex        =   7
         Top             =   1425
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    lblCompany.Caption = "CJB Corp."
    lblCompanyProduct.Caption = "Another CJB Corp. Product..."
    lblCopyright.Caption = "Copyright " + Chr(169) + "1999 by Christopher J. Bednarzyk"
    lblLicenseTo.Caption = "Licensed to:  Susan Bednarzyk"
    lblPlatform.Caption = "Microsoft Windows 98"
    lblProductName.Caption = "Handy-Dandy Calculator"
    lblVersion.Caption = "Calculator Version 1.0"
    lblWarning.Caption = " Warning:  Use of this calculator may be too outrageous for you!"
End Sub
Private Sub tmrSplash_Timer()
    tmrSplash.Enabled = False
    Me.Visible = False
    frmCalculator.Visible = True
End Sub
Private Sub Form_LostFocus()
    Me.Visible = False
End Sub
Private Sub Form_Click()
    Me.Visible = False
End Sub
Private Sub Frame1_Click()
    frmSplash.Visible = False
End Sub
Private Sub imgLogo_Click()
    frmSplash.Visible = False
End Sub

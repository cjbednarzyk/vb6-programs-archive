VERSION 5.00
Begin VB.Form frmListOps 
   BorderStyle     =   0  'None
   Caption         =   "List of Operations"
   ClientHeight    =   6330
   ClientLeft      =   585
   ClientTop       =   1095
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   350
      Left            =   1250
      TabIndex        =   1
      Top             =   5880
      Width           =   1000
   End
   Begin VB.ListBox lstOps 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      ItemData        =   "frmListOps.frx":0000
      Left            =   150
      List            =   "frmListOps.frx":0002
      TabIndex        =   0
      ToolTipText     =   "List of Previous Operations"
      Top             =   120
      Width           =   3200
   End
End
Attribute VB_Name = "frmListOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    frmListOps.Visible = False
End Sub

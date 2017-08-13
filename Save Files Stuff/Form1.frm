VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbSaveData 
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2566
      _Version        =   327680
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox rtbFile 
      Height          =   4455
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":00C9
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      DialogTitle     =   "Save File As..."
      InitDir         =   "c:\mydocu~1"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save As..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Open "c:\mydocu~1\hi" For Input As #1
  strSaveData = ""
  Do While Not EOF(1) ' Loop until end of file.
    strSaveData = strSaveData + CStr(Input(1, #1))
  Loop
  Close #1
  rtbFile.Text = strSaveData
End Sub

Private Sub mnuSave_Click()
    CommonDialog1.ShowSave
    rtbFile.SaveFile CommonDialog1.filename
End Sub

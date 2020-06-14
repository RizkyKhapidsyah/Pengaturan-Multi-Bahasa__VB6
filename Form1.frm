VERSION 5.00
Begin VB.Form form1 
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdopen 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton opt2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton opt1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbljdl 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- MENULIS BAHASA=1/0 KE FILE DATA.INI
Private Sub cmdopen_Click()
  If Me.opt1.Value = True Then
     WriteINI App.Path & "\Data.ini", "data", "bahasa", 1
  Else
     WriteINI App.Path & "\Data.ini", "data", "bahasa", 0
  End If
  
  Form2.Show
  Unload Me
End Sub

'-- PENERAPAN BAHASA & POSISI OPTION
Private Sub Form_Load()
  Bahasa Me
  Me.opt1.Value = indo
End Sub

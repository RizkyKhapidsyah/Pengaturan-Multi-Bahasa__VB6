VERSION 5.00
Begin VB.Form form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdclose 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lbljdl 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblalamat 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblnama 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- MENERAPKAN BAHASA
Private Sub Form_Load()
  Bahasa Me
End Sub

'-- MENUTUP FORM
Private Sub cmdclose_Click()
  Form1.Show
  Unload Me
End Sub



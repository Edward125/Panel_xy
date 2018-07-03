VERSION 5.00
Begin VB.Form frm_Help 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "請單擊換頁"
   ClientHeight    =   8250
   ClientLeft      =   14145
   ClientTop       =   8445
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   8265
      Index           =   4
      Left            =   0
      Picture         =   "frm_Help.frx":0000
      Top             =   0
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   8265
      Index           =   3
      Left            =   0
      Picture         =   "frm_Help.frx":128806
      Top             =   0
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   8265
      Index           =   2
      Left            =   0
      Picture         =   "frm_Help.frx":25100C
      Top             =   0
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   8265
      Index           =   1
      Left            =   0
      Picture         =   "frm_Help.frx":379812
      Top             =   0
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   8265
      Index           =   0
      Left            =   0
      Picture         =   "frm_Help.frx":4A2018
      Top             =   0
      Width           =   11010
   End
   Begin VB.Image Image1 
      Height          =   8235
      Index           =   5
      Left            =   0
      Picture         =   "frm_Help.frx":5CA81E
      Top             =   0
      Width           =   10965
   End
End
Attribute VB_Name = "frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
For i = 1 To 5
 Image1(i).Visible = False
Next
 
End Sub

Private Sub Image1_Click(Index As Integer)

If Index = 0 Then
 Image1(1).Visible = True
 Image1(0).Visible = False
End If
If Index = 1 Then
 Image1(2).Visible = True
 Image1(1).Visible = False
End If
If Index = 2 Then
 Image1(3).Visible = True
 Image1(2).Visible = False
End If
If Index = 3 Then
 Image1(4).Visible = True
 Image1(3).Visible = False
End If
If Index = 4 Then
 Image1(5).Visible = True
 Image1(4).Visible = False
End If
If Index = 5 Then
form1.Show                '把form1扔到前排
Unload Me                 '關掉自己
End If
End Sub

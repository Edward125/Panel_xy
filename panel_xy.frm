VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Rain's first exe---Panel XY Calculator "
   ClientHeight    =   5370
   ClientLeft      =   2925
   ClientTop       =   2685
   ClientWidth     =   13395
   Icon            =   "panel_xy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   13395
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   13180
      Begin VB.CommandButton Clear_SB_XY 
         Caption         =   "�M���p�O�y��"
         Height          =   495
         Left            =   10620
         TabIndex        =   39
         Top             =   4560
         Width           =   1000
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   9240
         TabIndex        =   37
         Top             =   3240
         Width           =   3735
         Begin VB.Label Label_help 
            Caption         =   "�����ڬݭ�z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   2520
            TabIndex        =   40
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "�`�N����r"
            ForeColor       =   &H000000FF&
            Height          =   915
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   10
         Top             =   4560
         Width           =   1000
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   9015
         Begin VB.TextBox txt_SB_B_Cons_Y 
            Height          =   375
            Left            =   7560
            MaxLength       =   7
            TabIndex        =   7
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txt_SB_B_Cons_X 
            Height          =   375
            Left            =   5400
            MaxLength       =   7
            TabIndex        =   6
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txt_MB_A_Cons_Y 
            Height          =   375
            Left            =   7560
            MaxLength       =   7
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txt_MB_A_Cons_X 
            Height          =   375
            Left            =   5400
            MaxLength       =   7
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image Image2 
            Height          =   630
            Left            =   240
            Picture         =   "panel_xy.frx":058A
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label22 
            Caption         =   "X3="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label16 
            Caption         =   "Y4="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   32
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "X4="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   31
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "S/B B point in Board Cons"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   30
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label15 
            Caption         =   "Y3="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6840
            TabIndex        =   29
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "MB A point in Board Cons"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1320
            TabIndex        =   28
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   9015
         Begin VB.TextBox txt_MB_A_Tebo_X 
            Height          =   375
            Left            =   5400
            MaxLength       =   8
            TabIndex        =   0
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txt_MB_A_Tebo_Y 
            Height          =   375
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   1
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txt_SB_B_Tebo_X 
            Height          =   375
            Left            =   5400
            MaxLength       =   8
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_SB_B_Tebo_Y 
            Height          =   375
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   3
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   750
            Left            =   240
            Picture         =   "panel_xy.frx":1C1C
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "MB A point in Tebo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1320
            TabIndex        =   26
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "X1="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   4800
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Y1="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6840
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "S/B B point in Tebo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   23
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label9 
            Caption         =   "X2="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   22
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Y2="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   21
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "&GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   8
         Top             =   4560
         Width           =   1000
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   9015
         Begin VB.OptionButton Option1 
            Caption         =   "90��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2800
            TabIndex        =   19
            Top             =   269
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "180��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4880
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "270��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6600
            TabIndex        =   17
            Top             =   270
            Width           =   1095
         End
         Begin VB.CheckBox Same_board_panel 
            Caption         =   "�ۦP�O�l���O"
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
            TabIndex        =   9
            Top             =   242
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   3000
         Index           =   0
         Left            =   9240
         Picture         =   "panel_xy.frx":3C66
         ScaleHeight     =   2940
         ScaleWidth      =   3735
         TabIndex        =   15
         Top             =   240
         Width           =   3800
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   3000
         Index           =   1
         Left            =   9240
         Picture         =   "panel_xy.frx":20A84
         ScaleHeight     =   2940
         ScaleWidth      =   3735
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3800
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   3000
         Index           =   2
         Left            =   9240
         Picture         =   "panel_xy.frx":3E0C6
         ScaleHeight     =   2940
         ScaleWidth      =   3735
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3800
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000007&
         Height          =   3000
         Index           =   3
         Left            =   9240
         Picture         =   "panel_xy.frx":5A988
         ScaleHeight     =   2940
         ScaleWidth      =   3735
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   3800
      End
      Begin VB.Label ld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4800
         TabIndex        =   36
         Top             =   4680
         Width           =   60
      End
      Begin VB.Label ly 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3840
         TabIndex        =   35
         Top             =   4680
         Width           =   60
      End
      Begin VB.Label lx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2760
         TabIndex        =   34
         Top             =   4680
         Width           =   60
      End
      Begin VB.Image Image3 
         Height          =   1350
         Left            =   120
         Picture         =   "panel_xy.frx":73BFA
         Top             =   3720
         Width           =   8985
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Del_Dian(strText As String)
Dim strTmp As String          '�w�qstrTmp���r�Ŧ�
Dim FindOk As Boolean         '�w�qFindOk��Boolean�ܶq,Boolean�ܶq�u�ରTrue �άO False
Dim c() As String             '�w�qc()���r�Ŧ�
FindOk = False
ReDim c(Len(strText))         '�w�qc()���ʺA�Ʋ�,�w�q�ʺA�Ʋիe�������n���w�q�@��.��p�o�̪�Dim c() As String

For i = 1 To Len(strText)     'i�q1��strText������
   
       c(i) = Mid(strText, i, 1)  'Mid(string, start, [length]),�o�̬O��^�r�Ŧꪺi��}�l��1��.
   
   If FindOk = False Then         '�p�G�٨S�����L�I,�N...
        If c(i) = "." Then           '�A�p�G���F�I,�N...
           FindOk = True             '��FindOk�ȳ]���u,�N����F�Ĥ@���I
        End If
       Else                       '�p�GFindOk =True,�����e�����`�����w�g���L�I�N...
        If c(i) = "." Then           '���I��������
           c(i) = ""
        End If
   End If
Next i

  For i = 1 To Len(strText)
     If c(i) <> "" Then
         strTmp = strTmp & c(i)          ' ��c(1),c(2),c(3)...�s���@����strTmp�̭��h
     End If
  Next i


  
Del_Dian = Trim(strTmp)                  '�h��strTmp�e�᪺�Ů��ȵ�Del_Dian,��ƦW�Y����^��.
End Function





Function Del_JianHao(strText As String)
Dim strTmp As String
Dim c() As String
ReDim c(Len(strText))
For i = 1 To Len(strText)
  If i > 1 Then                         'i=1�ɤ��i�J�`��
     c(i) = Mid(strText, i, 1)
        If c(i) = "-" Then
           c(i) = ""
        End If
     Else
      c(i) = Mid(strText, i, 1)
   End If
Next i
  
  For i = 1 To Len(strText)
     If c(i) <> "" Then
         strTmp = strTmp & c(i)
     End If
  Next i
  
Del_JianHao = strTmp
End Function




Private Sub Label17_Click()

End Sub

Private Sub Clear_SB_XY_Click()
txt_SB_B_Tebo_X = ""    '�M�Ťp�O��XY�y�Э�
txt_SB_B_Tebo_Y = ""
txt_SB_B_Cons_X = ""
txt_SB_B_Cons_Y = ""
txt_SB_B_Tebo_X.SetFocus   '��J�I���p�OX��

End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Label4.Caption = "�`�N:1.��JTebo�y�нп�J���,�p�⾹�����|�۰ʭ��H10." & Chr(13) & Chr(10) & "          2.�ۦP�O�l���O�p�⪺�y���ٯʥF����,�Цۦ��ˬd."
End Sub

Private Sub Label_help_DblClick()
frm_Help.Show

End Sub

Private Sub Label4_DblClick()
frm_Help.Show
End Sub

Private Sub Option1_Click()
  If Option1.Value = True Then
        Picture1(1).Visible = True
        Picture1(0).Visible = False
        Picture1(2).Visible = False
        Picture1(3).Visible = False
        Else
        Picture1(1).Visible = False
  End If
End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
        Picture1(2).Visible = True
        Picture1(0).Visible = False
        Picture1(1).Visible = False
        Picture1(3).Visible = False
        Else
        Picture1(2).Visible = False
  End If
End Sub

Private Sub Option3_Click()
  If Option3.Value = True Then
        Picture1(3).Visible = True
        Picture1(0).Visible = False
        Picture1(1).Visible = False
        Picture1(2).Visible = False
        Else
        Picture1(3).Visible = False
  End If
End Sub

Private Sub Same_board_panel_Click()

If Same_board_panel.Value = 0 Then  '�T�{�����ī�~��﨤��
 Option1.Enabled = False
 Option2.Enabled = False
 Option3.Enabled = False
 End If
If Same_board_panel.Value = 1 Then  '�T�{�����ī�~��﨤��
 Option1.Enabled = True
 Option2.Enabled = True
 Option3.Enabled = True
 
    If Option1.Value = True Then     '�T�{�����ī���ܪ�����Ϥ�1
     Picture1(1).Visible = True
     Picture1(0).Visible = False
     Picture1(2).Visible = False
     Picture1(3).Visible = False
   End If
   
    If Option2.Value = True Then     '�T�{�����ī���ܪ�����Ϥ�2
     Picture1(2).Visible = True
     Picture1(0).Visible = False
     Picture1(1).Visible = False
     Picture1(3).Visible = False
   End If
   
    If Option3.Value = True Then     '�T�{�����ī���ܪ�����Ϥ�3
     Picture1(3).Visible = True
     Picture1(0).Visible = False
     Picture1(1).Visible = False
     Picture1(2).Visible = False
   End If
 
 
End If
 
If Same_board_panel.Value = 1 Then   '�T�{���ۦP�O�l���O�hS/B��cons �y�Ф��ݭn�A��J
   txt_SB_B_Cons_X.Enabled = False
   txt_SB_B_Cons_Y.Enabled = False
   
    'Picture1(0).Visible = False
Else
   txt_SB_B_Cons_X.Enabled = True
   txt_SB_B_Cons_Y.Enabled = True
    Picture1(0).Visible = True
    Picture1(1).Visible = False
    Picture1(2).Visible = False
    Picture1(3).Visible = False
    
End If

End Sub

Private Sub cmdGO_Click()
Dim TEBO_A_X As Double
Dim TEBO_A_Y As Double
Dim TEBO_B_X As Double
Dim TEBO_B_Y As Double
'
Dim CONS_A_X As Long
Dim CONS_A_Y  As Long
Dim CONS_B_X As Long
Dim CONS_B_Y As Long
'
Dim consX As Long
Dim consY As Long
Dim Du As Integer

Du = 0
TEBO_A_X = Val(txt_MB_A_Tebo_X.Text) * 10
TEBO_A_Y = Val(txt_MB_A_Tebo_Y.Text) * 10
TEBO_B_X = Val(txt_SB_B_Tebo_X.Text) * 10
TEBO_B_Y = Val(txt_SB_B_Tebo_Y.Text) * 10

CONS_A_X = Val(txt_MB_A_Cons_X.Text)
CONS_A_Y = Val(txt_MB_A_Cons_Y.Text)
CONS_B_X = Val(txt_SB_B_Cons_X.Text)
CONS_B_Y = Val(txt_SB_B_Cons_Y.Text)

If Same_board_panel.Value = 1 Then

If Option1.Value = True Then     '90�� �h CONS_B_X=-CONS_A_Y,CONS_B_Y=CONS_A_X
       Du = 90
        
       CONS_B_Y = CONS_A_X       '��j�O��X�y�н�ȵ�B point��Y�y��

   
      If CONS_A_Y > 0 Then       '��j�O��Y�y�Ш��ۤϼƫ��ȵ�B point��X�y��
 
     CONS_B_X = Val("-" & CONS_A_Y)
     
      Else
     
     CONS_B_X = Abs(CONS_A_Y)    '���M���
     
     
      End If

      txt_SB_B_Cons_X.Text = CONS_B_X         '���ƻs���p�O��cons �y�Ц������
      txt_SB_B_Cons_Y.Text = CONS_B_Y

End If

If Option2.Value = True Then '180 �h CONS_B_X=-CONS_A_X,CONS_B_Y=-CONS_A_Y
      Du = 180
      
      If CONS_A_X > 0 Then       '��j�O��X�y�Ш��ۤϼƫ��ȵ�B point��X�y��
 
     CONS_B_X = Val("-" & CONS_A_X)
     
      Else
     
     CONS_B_X = Abs(CONS_A_X)
     
      End If
      
      If CONS_A_Y > 0 Then       '��j�O��Y�y�Ш��ۤϼƫ��ȵ�B point��Y�y��
 
     CONS_B_Y = Val("-" & CONS_A_Y)
     
      Else
     
     CONS_B_Y = Abs(CONS_A_Y)    '���M���
     
      End If
      txt_SB_B_Cons_X.Text = CONS_B_X         '���ƻs���p�O��cons �y�Ц������
      txt_SB_B_Cons_Y.Text = CONS_B_Y

End If

If Option3.Value = True Then '270  �h CONS_B_X=CONS_A_Y,CONS_B_Y=-CONS_A_X
       Du = 270
         
       CONS_B_X = CONS_A_Y       '��j�O��Y�y�н�ȵ�B point��X�y��
       
      If CONS_A_X > 0 Then       '��j�O��X�y�Ш��ۤϼƫ��ȵ�B point��Y�y��
 
     CONS_B_Y = Val("-" & CONS_A_X)
     
      Else
     
     CONS_B_Y = Abs(CONS_A_X)    '���M���
     
      End If
      txt_SB_B_Cons_X.Text = CONS_B_X         '���ƻs���p�O��cons �y�Ц������
      txt_SB_B_Cons_Y.Text = CONS_B_Y
End If

End If

consX = TEBO_B_X - TEBO_A_X - CONS_B_X + CONS_A_X

consY = TEBO_B_Y - TEBO_A_Y - CONS_B_Y + CONS_A_Y

'~~~~board cons panel �w�q�س̦h�u���J7��,�]�A�t��.
If consX > 9999999 Or consX < -999999 Or consY > 9999999 Or consY < -999999 Then

MsgBox ("Board Cons Panel�y�г̦h�u���J7��,���G�y�жW�X�d��!"), vbCritical

'txt_MB_A_Tebo_X = ""         '�M�ųo�ǭ�
'txt_MB_A_Tebo_Y = ""
'txt_SB_B_Tebo_X = ""
'txt_SB_B_Tebo_Y = ""
'txt_MB_A_Cons_X = ""
'txt_MB_A_Cons_Y = ""
'txt_SB_B_Cons_X = ""
'txt_SB_B_Cons_Y = ""
'
'consX = 0
'
'consY = 0



End If
'~~~board cons panel �w�q�س̦h�u���J7��,�]�A�t��~~~end
lx.Caption = consX
ld.Caption = Du
ly.Caption = consY



End Sub

Private Sub txt_MB_A_Cons_X_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Or KeyAscii = 46 Or KeyAscii = 47 Then '�u���J�Ʀr(48-57)�M�t��(45)
   KeyAscii = 0
    txt_MB_A_Cons_X.Text = KeyAscii
 End If
End Sub

Private Sub txt_MB_A_Cons_X_LostFocus()                  '�btxt_MB_A_Cons_X�إ��h�J�I���ɭ�....
txt_MB_A_Cons_X.Text = Del_JianHao(txt_MB_A_Cons_X.Text) '�ե� "�R���"��function
txt_MB_A_Cons_X.Text = Del_Dian(txt_MB_A_Cons_X.Text)    '�ե� "�R���p���I"��function
End Sub

Private Sub txt_MB_A_Cons_Y_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Or KeyAscii = 46 Or KeyAscii = 47 Then '�u���J�Ʀr(48-57)�M�t��(45)
   KeyAscii = 0
    txt_MB_A_Cons_Y.Text = KeyAscii
 End If
End Sub

Private Sub txt_MB_A_Cons_Y_LostFocus()                  '�btxt_MB_A_Cons_Y�إ��h�J�I���ɭ�....
txt_MB_A_Cons_Y.Text = Del_JianHao(txt_MB_A_Cons_Y.Text) '�ե� "�R���"��function
txt_MB_A_Cons_Y.Text = Del_Dian(txt_MB_A_Cons_Y.Text) '�ե� "�R���p���I"��function
End Sub

Private Sub txt_MB_A_Tebo_X_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then                 '�u���J�p���I(46)�M�Ʀr(48-57)�M�t��(45).47�O����,�����J
   KeyAscii = 0
    txt_MB_A_Tebo_X.Text = KeyAscii
 Exit Sub
 End If
 
 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Then
   KeyAscii = 0
    txt_MB_A_Tebo_X.Text = KeyAscii
 End If
End Sub

Private Sub txt_MB_A_Tebo_X_LostFocus()                  '�btxt_MB_A_Tebo_X�إ��h�J�I���ɭ�....
txt_MB_A_Tebo_X.Text = Del_JianHao(txt_MB_A_Tebo_X.Text) '�ե� "�R���"��function
txt_MB_A_Tebo_X.Text = Del_Dian(txt_MB_A_Tebo_X.Text)    '�ե� "�R���p���I"��function

End Sub

Private Sub txt_MB_A_Tebo_Y_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then                 '�u���J�p���I(46)�M�Ʀr(48-57)�M�t��(45).47�O����,�����J
   KeyAscii = 0
    txt_MB_A_Tebo_Y.Text = KeyAscii
 Exit Sub
 End If
 
 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Then
   KeyAscii = 0
    txt_MB_A_Tebo_Y.Text = KeyAscii
 End If
End Sub

Private Sub txt_MB_A_Tebo_Y_LostFocus()                  '�btxt_MB_A_Tebo_Y�إ��h�J�I���ɭ�....
txt_MB_A_Tebo_Y.Text = Del_JianHao(txt_MB_A_Tebo_Y.Text) '�ե� "�R���"��function
txt_MB_A_Tebo_Y.Text = Del_Dian(txt_MB_A_Tebo_Y.Text)    '�ե� "�R���"��function
End Sub

Private Sub txt_SB_B_Cons_X_KeyPress(KeyAscii As Integer)

 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Or KeyAscii = 46 Or KeyAscii = 47 Then '�u���J�Ʀr(48-57)�M�t��(45)
   KeyAscii = 0
    txt_SB_B_Cons_X.Text = KeyAscii
 End If
End Sub

Private Sub txt_SB_B_Cons_X_LostFocus()                  '�btxt_SB_B_Cons_X�إ��h�J�I���ɭ�....
txt_SB_B_Cons_X.Text = Del_JianHao(txt_SB_B_Cons_X.Text) '�ե� "�R���"��function
txt_SB_B_Cons_X.Text = Del_Dian(txt_SB_B_Cons_X.Text)    '�ե� "�R���p���I"��function
End Sub

Private Sub txt_SB_B_Cons_Y_KeyPress(KeyAscii As Integer)
 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 If KeyAscii < 45 Or KeyAscii > 57 Or KeyAscii = 46 Or KeyAscii = 47 Then '�u���J�Ʀr(48-57)�M�t��(45)
   KeyAscii = 0
    txt_SB_B_Cons_Y.Text = KeyAscii
 End If
End Sub

Private Sub txt_SB_B_Cons_Y_LostFocus()                  '�btxt_SB_B_Cons_Y�إ��h�J�I���ɭ�....
txt_SB_B_Cons_Y.Text = Del_JianHao(txt_SB_B_Cons_Y.Text) '�ե� "�R���"��function
txt_SB_B_Cons_Y.Text = Del_Dian(txt_SB_B_Cons_Y.Text)    '�ե� "�R���"��function
End Sub

Private Sub txt_SB_B_Tebo_X_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then                 '�u���J�p���I(46)�M�Ʀr(48-57)�M�t��(45).47�O����,�����J
   KeyAscii = 0
    txt_SB_B_Tebo_X.Text = KeyAscii
 Exit Sub
 End If

  If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Then
   KeyAscii = 0
    txt_SB_B_Tebo_X.Text = KeyAscii
 End If
End Sub

Private Sub txt_SB_B_Tebo_X_LostFocus()                  '�btxt_SB_A_Tebo_X�إ��h�J�I���ɭ�....
txt_SB_B_Tebo_X.Text = Del_JianHao(txt_SB_B_Tebo_X.Text) '�ե� "�R���"��function
txt_SB_B_Tebo_X.Text = Del_Dian(txt_SB_B_Tebo_X.Text) '�ե� "�R���"��function
End Sub

Private Sub txt_SB_B_Tebo_Y_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then                 '�u���J�p���I(46)�M�Ʀr(48-57)�M�t��(45).47�O����,�����J
   KeyAscii = 0
    txt_SB_B_Tebo_Y.Text = KeyAscii
 Exit Sub
 End If

 If KeyAscii = 8 Then        '�h����(8)�i�H��J.
 Exit Sub
 End If
 
 If KeyAscii < 45 Or KeyAscii > 57 Then
   KeyAscii = 0
    txt_SB_B_Tebo_Y.Text = KeyAscii
 End If
 
 
End Sub

Private Sub txt_SB_B_Tebo_Y_LostFocus()            '�btxt_SB_A_Tebo_X�إ��h�J�I���ɭ�....
txt_SB_B_Tebo_Y.Text = Del_JianHao(txt_SB_B_Tebo_Y.Text) '�ե� "�R���"��function
txt_SB_B_Tebo_Y.Text = Del_Dian(txt_SB_B_Tebo_Y.Text) '�ե� "�R���"��function
End Sub

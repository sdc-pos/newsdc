VERSION 5.00
Begin VB.Form PI000501 
   Caption         =   "���ޔ��㏈��"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   12315
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   8295
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   5355
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   5355
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   2625
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   12
      Top             =   3600
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   2625
      MaxLength       =   11
      TabIndex        =   11
      Top             =   3240
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   7
      Left            =   2625
      MaxLength       =   8
      TabIndex        =   10
      Top             =   2880
      Width           =   1485
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   3360
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   9
      Top             =   2280
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2280
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   3360
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   7
      Top             =   1920
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1920
      Width           =   750
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   5145
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2625
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4515
      MaxLength       =   7
      TabIndex        =   1
      Top             =   240
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2625
      MaxLength       =   5
      TabIndex        =   2
      Top             =   840
      Width           =   750
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   3360
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   840
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1890
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10440
      TabIndex        =   27
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8760
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7920
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W������"
      Height          =   255
      Index           =   10
      Left            =   6930
      TabIndex        =   39
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�W������"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   38
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌ɐ���"
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   37
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���z"
      Height          =   255
      Index           =   7
      Left            =   1890
      TabIndex        =   36
      Top             =   3720
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�P��"
      Height          =   255
      Index           =   5
      Left            =   1890
      TabIndex        =   35
      Top             =   3360
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   4
      Left            =   1890
      TabIndex        =   34
      Top             =   3000
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�̔��敪"
      Height          =   255
      Index           =   3
      Left            =   1470
      TabIndex        =   33
      Top             =   2400
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���x�P��"
      Height          =   255
      Index           =   2
      Left            =   1470
      TabIndex        =   32
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   0
      Left            =   1470
      TabIndex        =   31
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�����N��"
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   30
      Top             =   360
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���Ӑ�"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   29
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����N����"
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   28
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "PI000501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private POS_UMU     As Boolean              'POS���т̗L��
    
Private YOIN        As String * 2           'POS���і��̏o�ɗv��
Private TANTO       As String * 5           'POS���і��̒S���Һ���

    
Dim WS_NO           As String * 3
    
    
'�e�L�X�g�p�Y��
Private Const ptxURIAGE_DT% = 0             '����N����
Private Const ptxKEIJYO_YM% = 1             '�v�㌎

Private Const ptxTOKUI_CODE% = 2            '���Ӑ溰��

Private Const ptxHIN_GAI% = 3               '�i��
Private Const ptxHIN_NAME% = 4              '�i��

Private Const ptxG_SYUSHI% = 5              '���x�P��
Private Const ptxG_HANBAI_KBN% = 6          '�̔��敪

Private Const ptxURIAGE_QTY% = 7            '���㐔��
Private Const ptxTANKA% = 8                 '�P��
Private Const ptxKINGAKU% = 9               '���z

Private Const ptxZAIKO_QTY% = 10            '�݌Ɏc
Private Const ptxG_ST_URITAN% = 11          '�W���e������
Private Const ptxG_ST_SHITAN% = 12          '�W���e������

Private Const ptxZEI_KIN% = 13              '�����


'�R���{�p�Y��
Private Const pcmbTOKUI% = 0                '���Ӑ�
Private Const pcmbG_SYUSHI% = 1             '���x�P��
Private Const pcmbG_HANBAI_KBN% = 2         '�̔��P��
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI000501.MousePointer = vbHourglass

    Call Ctrl_Lock(PI000501)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI000501)


    PI000501.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
    
Dim wkDate      As String * 10
    
Dim ST_Sumi_Qty As Long
Dim ST_Mi_Qty   As Long
    
Dim ZEI         As Long
Dim wkKINGAKU   As Long
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        
        
        Case ptxURIAGE_DT       '����N����
            
            If Not IsDate(Text1(ptxURIAGE_DT).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxURIAGE_DT).Text = Format(CDate(Text1(ptxURIAGE_DT).Text), "YYYY/MM/DD")
            End If
        
        Case ptxKEIJYO_YM       '�����N��
            
            
            wkDate = Text1(ptxKEIJYO_YM).Text & "/01"
            
            If Not IsDate(wkDate) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                
                wkDate = Format(CDate(Text1(ptxKEIJYO_YM).Text), "YYYY/MM/DD")
                
                Text1(ptxKEIJYO_YM).Text = Mid(wkDate, 1, 7)
            End If
        
        Case ptxTOKUI_CODE   '���Ӑ�
            
           Combo1(pcmbTOKUI).ListIndex = -1
           For i = 0 To Combo1(pcmbTOKUI).ListCount - 1
               If Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).List(i), 5)) Then
                   Combo1(pcmbTOKUI).ListIndex = i
                   Exit For
               End If
           
           Next i
    
           If i > Combo1(pcmbTOKUI).ListCount - 1 Then
               MsgBox "���͂������ڂ̓G���[�ł��B"
               Text1(Mode).SetFocus
               Exit Function
           End If
        
        Case ptxHIN_GAI         '�i��
    
                    
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI And _
                StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI And _
                Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = Trim(Text1(ptxHIN_GAI).Text) Then
    
            Else
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        Text1(ptxHIN_NAME).Text = ""
                        Text1(ptxZAIKO_QTY).Text = ""
                        Text1(ptxG_ST_URITAN).Text = ""
                        Text1(ptxG_ST_SHITAN).Text = ""
                        
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(Mode).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                
                End Select
                
                            
                
                Text1(ptxHIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                
                '���x�P��
                Text1(ptxG_SYUSHI).Text = StrConv(ITEMREC.G_SYUSHI, vbUnicode)
                Combo1(pcmbG_SYUSHI).ListIndex = -1
                For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                    If Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
                        Combo1(pcmbG_SYUSHI).ListIndex = i
                        Exit For
                    End If
                
                Next i
                '�̔��敪
                Text1(ptxG_HANBAI_KBN).Text = StrConv(ITEMREC.G_HANBAI_KBN, vbUnicode)
                Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
                For i = 0 To Combo1(pcmbG_HANBAI_KBN).ListCount - 1
                    If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
                        Combo1(pcmbG_HANBAI_KBN).ListIndex = i
                        Exit For
                    End If
                
                Next i
                
                
                If Not POS_UMU Then
                '�o�n�r�����ŕW���I�Ԗ��ݒ�͏o�ɕs��2006.04.26
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                        Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then

                        MsgBox "�W���I�Ԃ��ݒ肳��Ă��܂���B"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                End If
                
                
                
                
                If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                           StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                           StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                
                End If
                            
                            
                                        

                
                Text1(ptxZAIKO_QTY).Text = Format(SUMI_QTY + MI_QTY, "#,##0")
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Text1(ptxG_ST_URITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#,##0.00")
                Else
                    Text1(ptxG_ST_URITAN).Text = ""
                End If
                
                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                    Text1(ptxTANKA).Text = Format(CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)), "#0.00")
                Else
                    Text1(ptxTANKA).Text = ""
                End If
                
                If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                    Text1(ptxG_ST_SHITAN).Text = Format(CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)), "#,##0.00")
                Else
                    Text1(ptxG_ST_SHITAN).Text = ""
                End If
            End If
           
            
            
                    
        
        
        
        Case ptxG_SYUSHI        '���x�P��
            
            Combo1(pcmbG_SYUSHI).ListIndex = -1
            For i = 0 To Combo1(pcmbG_SYUSHI).ListCount - 1
                If Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).List(i), 3)) Then
                    Combo1(pcmbG_SYUSHI).ListIndex = i
                    Exit For
                End If
               
            Next i
        
            If i > Combo1(pcmbG_SYUSHI).ListCount - 1 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            End If
        
        Case ptxG_HANBAI_KBN    '�̔��敪
            
            Combo1(pcmbG_HANBAI_KBN).ListIndex = -1
            For i = 0 To Combo1(pcmbG_HANBAI_KBN).ListCount - 1
                If Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).List(i), 3), 2)) Then
                    Combo1(pcmbG_HANBAI_KBN).ListIndex = i
                    Exit For
                End If
           
           Next i
    
           If i > Combo1(pcmbG_HANBAI_KBN).ListCount - 1 Then
               MsgBox "���͂������ڂ̓G���[�ł��B"
               Text1(Mode).SetFocus
               Exit Function
           End If
        
        
        
        Case ptxURIAGE_QTY       '���㐔��
    
            If Not IsNumeric(Text1(ptxURIAGE_QTY).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                If CLng(Text1(ptxURIAGE_QTY).Text) = 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
                
                Text1(ptxURIAGE_QTY).Text = Format(CLng(Text1(ptxURIAGE_QTY).Text), "#0")
            
                
                If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
                    
                    If CLng(Text1(ptxURIAGE_QTY).Text) <= 0 Then
                    Else
                        If CLng(Text1(ptxURIAGE_QTY).Text) > CLng(Text1(ptxZAIKO_QTY).Text) Then
                            MsgBox "���͂������ڂ̓G���[�ł��B�i���݌ɐ��s���j"
                            Text1(Mode).SetFocus
                            Exit Function
                        End If
                    
                    
                    
                    
                        If Not POS_UMU Then
                        '�o�n�r�����ŕW���I�ԍ݌ɂōă`�F�b�N2006.04.26
                            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" And _
                                Trim(StrConv(ITEMREC.ST_RETU, vbUnicode)) = "" And _
                                Trim(StrConv(ITEMREC.ST_REN, vbUnicode)) = "" And _
                                Trim(StrConv(ITEMREC.ST_DAN, vbUnicode)) = "" Then
        
                                MsgBox "�W���I�Ԃ��ݒ肳��Ă��܂���B"
                                Text1(Mode).SetFocus
                                Exit Function
        
                            End If
                        
                        
                            If Zaiko_Syukei_Proc(ST_Sumi_Qty, ST_Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                       StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                       StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                                       StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                       StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                       StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                       StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                                Exit Function
                            
                            End If
                            
                            If CLng(Text1(ptxURIAGE_QTY).Text) > ST_Sumi_Qty + ST_Mi_Qty Then
                                MsgBox "���͂������ڂ̓G���[�ł��B�i�W���I�ԍ݌ɐ��s���j"
                                Text1(Mode).SetFocus
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        
                    
                    
                    
                    
                    
                    End If
                
                
                
                
                End If
            
            
                            
            
            
            
            
                If IsNumeric(Text1(ptxTANKA).Text) Then
                    
                    If Text1(ptxKINGAKU).Text = "" Then
                        Text1(ptxKINGAKU).Text = Format(CLng(CDbl(Text1(ptxTANKA).Text) * _
                                                    CLng(Text1(ptxURIAGE_QTY).Text)), "#,##0")
                        
                        
                    End If
'-----------------------
                
                
                
                
                
                
                
                
                
                
                
                Else
                    Text1(ptxKINGAKU).Text = "0"
                End If
            End If
    
    
        Case ptxTANKA           '�P��
    
            If Not IsNumeric(Text1(ptxTANKA).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            Else
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
            
            End If
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                  ���ޔ����ް��X�V
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer




    Update_Proc = True
                                        
    Call Input_Lock
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If

                                        '�Ǘ��t�@�C����莑�ޔ���ԍ��̊l��
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            
                If P_KANRI_MAKE_Proc() Then
                    GoTo Abort_Tran
                End If
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = True
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                GoTo Abort_Tran
        
        End Select
    
    
    Loop
    
    '�����ް����{�P
    If CLng(StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)) = 99999 Then
        Call UniCode_Conv(P_KANRIREC.URIAGE_NO, "00001")
    Else
        Call UniCode_Conv(P_KANRIREC.URIAGE_NO, Format(CLng(StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)) + 1, "00000"))
    End If
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�Ǘ��}�X�^")
                    End If
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�Ǘ��}�X�^")
                GoTo Abort_Tran
        End Select
    Loop
    
    '---------------------------------------------------    '���ޔ���f�[�^�X�V
    Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_NO, StrConv(P_KANRIREC.URIAGE_NO, vbUnicode))       '���㇂
    Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_DT, Format(Text1(ptxURIAGE_DT).Text, "YYYYMMDD"))   '�����
    Call UniCode_Conv(P_SHURIAGE_REC.KEIJYO_YM, Mid(Text1(ptxKEIJYO_YM), 1, 4) & Mid(Text1(ptxKEIJYO_YM), 6, 2))  '�v��N��
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxTOKUI_CODE).Text)
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            '���o�^�͈�ʈ����i�����ɂ͂��Ȃ��͂��j
            Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, P_TORI_GENERAL)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Function
        
    End Select

    
    
    
    Call UniCode_Conv(P_SHURIAGE_REC.TORI_KBN, StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode))      '�����敪
    Call UniCode_Conv(P_SHURIAGE_REC.TOKUI_CODE, Text1(ptxTOKUI_CODE).Text)                     '���Ӑ溰��
    Call UniCode_Conv(P_SHURIAGE_REC.JGYOBU, SHIZAI)                                            '���ƕ�
    Call UniCode_Conv(P_SHURIAGE_REC.NAIGAI, NAIGAI_NAI)                                        '�����O
    Call UniCode_Conv(P_SHURIAGE_REC.HIN_GAI, Text1(ptxHIN_GAI).Text)                           '�i��
    Call UniCode_Conv(P_SHURIAGE_REC.G_SYUSHI, Text1(ptxG_SYUSHI).Text)                         '���x�P��
    Call UniCode_Conv(P_SHURIAGE_REC.G_HANBAI_KBN, Text1(ptxG_HANBAI_KBN).Text)                 '�̔��敪
                                                                                                '����
    
    If CDbl(Text1(ptxURIAGE_QTY).Text) < 0 Then
        Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_QTY, Format(CDbl(Text1(ptxURIAGE_QTY).Text), "0000000.00"))
    Else
        Call UniCode_Conv(P_SHURIAGE_REC.URIAGE_QTY, Format(CDbl(Text1(ptxURIAGE_QTY).Text), "00000000.00"))
    End If
                                                                                                '�P��
    Call UniCode_Conv(P_SHURIAGE_REC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), "00000000.00"))
                                                                                                '���z
    
    If CLng(Text1(ptxKINGAKU).Text) < 0 Then
        Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(CLng(Text1(ptxKINGAKU).Text), "00000000"))
    Else
        Call UniCode_Conv(P_SHURIAGE_REC.KINGAKU, Format(CLng(Text1(ptxKINGAKU).Text), "000000000"))
    End If
    
    
    
    Call UniCode_Conv(P_SHURIAGE_REC.SEIKU_F, P_SEIKYU_NON)                       '�����׸�
    
    Call UniCode_Conv(P_SHURIAGE_REC.FILLER, "")
    
                                                                                    '�X�V����
    Call UniCode_Conv(P_SHURIAGE_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpInsert, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHURIAGE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���ޔ����ް�")
                GoTo Abort_Tran
        End Select
    
    Loop
    

    If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = P_ZAIKO_F_ON Then
        If Not POS_UMU Then
            'POS���тȂ��́A�W���I�Ԃ��݌Ɉ������Ƃ�
        
            If CLng(Text1(ptxURIAGE_QTY).Text) > 0 Then
                sts = Syuko_Update_Proc(SHIZAI, _
                                        NAIGAI_NAI, _
                                        Text1(ptxHIN_GAI).Text, _
                                        "", _
                                        (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)), _
                                        YOIN, _
                                        0, _
                                        CLng(Text1(ptxURIAGE_QTY).Text), _
                                        0, _
                                        WS_NO, _
                                        TANTO)
        
            End If
            Select Case sts
                Case False
                Case Else
                    Update_Proc = sts
                    GoTo Abort_Tran
            End Select
        
        
        
                    
        
        
        
        End If
    End If

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    
    Call Input_UnLock
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    
    Call Input_UnLock

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbTOKUI          '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_SYUSHI       '���x�P��
            Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbG_HANBAI_KBN   '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
        Case pcmbTOKUI          '���Ӑ�
            Text1(ptxTOKUI_CODE).Text = Trim(Right(Combo1(pcmbTOKUI).Text, 5))
        Case pcmbG_SYUSHI       '���x�P��
            Text1(ptxG_SYUSHI).Text = Trim(Right(Combo1(pcmbG_SYUSHI).Text, 3))
        Case pcmbG_HANBAI_KBN   '�̔��敪
            Text1(ptxG_HANBAI_KBN).Text = Trim(Left(Right(Combo1(pcmbG_HANBAI_KBN).Text, 3), 2))
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans         As Integer
Dim i           As Integer


Dim sts         As Integer

    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            For i = ptxURIAGE_DT To ptxG_ST_SHITAN
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
                
                Call Init_Proc
            
            End If
            
            Text1(ptxURIAGE_DT).SetFocus
        
        Case P_CMD_DEL                      '�폜
    
        Case P_CMD_DSP                      '����/�\��
        
        
        Case P_CMD_OUT                      '�ް��o��
        
        Case P_CMD_PRT                      '���
            
        Case P_CMD_End                      '�I��
    
            Unload Me
    
    End Select

End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

Dim sBuffer As String * 255
Dim com     As String


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                'POS���їL���̎�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "POS_UMU", "P_SYS", c) Then
        POS_UMU = False
    Else
        If RTrim(c) = "0" Then
            POS_UMU = False
        Else
            POS_UMU = True
        End If
    End If
                                
    If Not POS_UMU Then
                                'POS���і����A�o�ɗv��
        If GetIni(StrConv(App.EXEName, vbUpperCase), "YOIN", "P_SYS", c) Then
            Beep
            MsgBox "�o�ɗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        YOIN = Trim(c)
    
                                'POS���і����A�S���Һ���
    
        If GetIni(StrConv(App.EXEName, vbUpperCase), "TANTO", "P_SYS", c) Then
            TANTO = ""
        End If
        TANTO = Trim(c)
    
    
    End If
                                
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�ړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '������n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�v���n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ޔ����ް��n�o�d�m
    If P_SHURIAGE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����Ͻ��n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    '�Ǘ��}�X�^�̓ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)

    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            If P_KANRI_MAKE_Proc() Then
                Unload Me
            End If
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^")
            Unload Me
    End Select
        
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc

    
    '���Ӑ�
    If Ukeharai_Set_Proc(pcmbTOKUI) Then
        Unload Me
    End If
    
    '���x�P�ʂ̃Z�b�g
    If Code_Set_Proc(pcmbG_SYUSHI, P_KBN03_CD, 0) Then
        Unload Me
    End If
    
    '�̔��敪�̃Z�b�g
    If Code_Set_Proc(pcmbG_HANBAI_KBN, P_KBN02_CD, 0) Then
        Unload Me
    End If
    
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
    
    '��ʏ����ݒ�
    Call Init_Proc

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '���ޔ����ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_SHURIAGE_POS, P_SHURIAGE_REC, Len(P_SHURIAGE_REC), K0_P_SHURIAGE, Len(K0_P_SHURIAGE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޔ����ް�")
        End If
    End If
                                            '�݌��ް��b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI000501 = Nothing

    End
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub
Private Sub Init_Proc()
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i       As Integer
Dim sts     As Integer


    
    
    
    For i = ptxURIAGE_DT To ptxG_ST_SHITAN
        Text1(i).Text = ""
    Next i
    '���ぁ����
    Text1(ptxURIAGE_DT).Text = Format(Now, "YYYY/MM/DD")
    '�v�㌎
    Text1(ptxKEIJYO_YM).Text = Mid(Format(Now, "YYYY/MM/DD"), 1, 7)


    For i = pcmbTOKUI To pcmbG_HANBAI_KBN
        
        Combo1(i).ListIndex = -1
    
    Next i




    Call UniCode_Conv(ITEMREC.JGYOBU, "")
    Call UniCode_Conv(ITEMREC.NAIGAI, "")
    Call UniCode_Conv(ITEMREC.HIN_GAI, "")

End Sub
Private Function Ukeharai_Set_Proc(Index As Integer) As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(Index).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        
        End Select

        
        
        Combo1(Index).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function
Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function



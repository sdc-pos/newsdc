VERSION 5.00
Begin VB.Form F1030701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�o�ח\��o�^"
   ClientHeight    =   6015
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   13455
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
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13455
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   15
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   51
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   840
      MaxLength       =   12
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   2640
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   1320
      Width           =   972
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   3600
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   17
      Top             =   4080
      Width           =   4950
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   8640
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   5040
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2280
      Width           =   852
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   9840
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   960
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
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
      Left            =   10200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   9360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   8520
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   7680
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   5
      Left            =   4920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   3000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   2160
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   1320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
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
      Left            =   480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�����́F�������ԁj"
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   50
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�����́F�������ԁj"
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   49
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IDNo"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   48
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   47
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   46
      Top             =   5400
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   17
      Left            =   1560
      TabIndex        =   45
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   16
      Left            =   8400
      TabIndex        =   44
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   7680
      TabIndex        =   43
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   6960
      TabIndex        =   42
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�z�X�g�I��"
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   41
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�z�X�g�q��"
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   40
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��j"
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   39
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\�Z�P�ʁi���j"
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   38
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[��"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   37
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   36
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   35
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   34
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ח\�萔"
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   33
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����敪"
      Height          =   255
      Index           =   14
      Left            =   960
      TabIndex        =   32
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   31
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   30
      Top             =   1080
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const pcmbC_Kbn% = 0
Private Const pcmbNAIGAI% = 1
Private Const pcmbMUKE_CODE% = 2

Private Const ptxMAX% = 15

Private Const ptxID_No% = 0
Private Const ptxCode% = 1
Private Const ptxName% = 2
Private Const ptxS_Qty% = 3
Private Const ptxYY% = 4
Private Const ptxMM% = 5
Private Const ptxDD% = 6
Private Const ptxNo% = 7
Private Const ptxMoto% = 8
Private Const ptxSaki% = 9
Private Const ptxSoko% = 10
Private Const ptxS_No% = 11
Private Const ptxRetu% = 12
Private Const ptxRen% = 13
Private Const ptxDan% = 14
Private Const ptxMUKE_CODE% = 15         '������i�R�[�h���͗p�j
                                   
'Private Const LAST_UPDATE_DAY$ = "[F103070]2018.04.21 09:00"
'Private Const LAST_UPDATE_DAY$ = "[F103070]2018.04.27 16:45"
Private Const LAST_UPDATE_DAY$ = "[F103070]2020.04.14 14:00 �X�V��O�� �\���c����C��"
                                   
Private DEF_CYU_KBN As String * 1       '2009.04.14
Private OSAKA_MODE As String * 1        '2010.03.23
                                   
                                   
                                   
                                   '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Optional Start_Pos As Integer = 0)
Dim i As Integer
    
    For i = Start_Pos To ptxMAX
        Text(i).Text = ""
    Next i
    
End Sub
                                    '�i�ڃ}�X�^���e���ڂ�\������
Private Function Item_Dsp() As Integer

Dim sts As Integer


    Item_Dsp = True
                                                '�����O�`�F�b�N
                                                '�܂��O���i�Ԃœǂݍ���
    
    Text(ptxCode).Text = StrConv(Text(ptxCode).Text, vbUpperCase)
    
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxCode))
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxName) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text(ptxSoko) = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
            Text(ptxS_No) = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Text(ptxRetu) = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Text(ptxRen) = StrConv(ITEMREC.ST_REN, vbUnicode)
            Text(ptxDan) = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '�����i�Ԃœǂݍ���
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxCode).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    Text(ptxCode).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxName).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Text(ptxSoko).Text = StrConv(ITEMREC.BIKOU_SOKO, vbUnicode)
                    Text(ptxS_No).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Text(ptxRetu).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Text(ptxRen).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Text(ptxDan).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
                Case BtErrKeyNotFound
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Item_Dsp = SYS_ERR
                    Exit Function
            End Select
                
                
                
        Case Else
                
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Item_Dsp = SYS_ERR
            Exit Function
        
    End Select
    
    Item_Dsp = False
    
End Function

'                                       ���͍��ڂ̃G���[�`�F�b�N
Private Function Err_Chk() As Integer

Dim yn          As Integer
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim Flg         As Boolean
Dim Qty         As Long
Dim W_CYU_KBN   As String

    Err_Chk = True
                                        '�i�ڃ`�F�b�N
    
    If Trim(Text(ptxCode).Text) = "" Then   '2018.04.20
        Beep                                '2018.04.20
        MsgBox "�i�Ԃ͕K�{���͂ł��B "      '2018.04.20
        Text(ptxS_Qty).SetFocus             '2018.04.20
        Exit Function                       '2018.04.20
    End If                                  '2018.04.20
    
    
    
    
    sts = Item_Dsp
    Select Case sts
        Case False
        Case True
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� "
            Text(ptxCode).SetFocus
            Exit Function
        Case Else
            Err_Chk = SYS_ERR
            Exit Function
    End Select
                                        '�o�ח\�萔�ʃ`�F�b�N
    If Trim(Text(ptxS_Qty).Text) = "" Then
        Text(ptxS_Qty) = "0"
    End If
        
    If Not IsNumeric(Trim(Text(ptxS_Qty).Text)) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł�� "
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
    
    Text(ptxS_Qty).Text = Format(CLng(Text(ptxS_Qty).Text), "#0")
    If CLng(Text(ptxS_Qty).Text) <= 0 Then
        Beep
        MsgBox "�o�ח\�萔�����͂���Ă��܂���"
        Text(ptxS_Qty).SetFocus
        Exit Function
    End If
                                        '�`�[���t
    For i = ptxYY To ptxDD
        If Trim(Text(i)) = "" Then
            Text(i).Text = "0"
        End If
        
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� "
            Text(i).SetFocus
            Exit Function
        Else
            RetBuf = Format(CLng(Text(i).Text), "0000")
            Text(i).Text = Right(RetBuf, Text(i).MaxLength)
        End If
    Next i
        
    If Not IsDate(Text(ptxYY) & "/" & Text(ptxMM) & "/" & Text(ptxDD)) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł�� "
        Text(ptxYY).SetFocus
        Err_Chk = True
        Exit Function
    End If
                
    If Not IsNumeric(Trim(Text(ptxNo))) Then
'        Beep
'        MsgBox "���͂������ڂ̓G���[�ł��B"
'        Text(ptxNo).SetFocus
'        Err_Chk = True
'        Exit Function
    Else
        Text(ptxNo) = Format(CLng(Text(ptxNo)), "000000")
    End If
                                    '�u�O�v�ȊO�����͂��ꂽ��o�^�ς݃`�F�b�N
                                '�h�c��
    If Len(Text(ptxID_No).Text) = 0 Then
    Else
                                                '�������ԈȊO�����͂��ꂽ��o�^�ς݃`�F�b�N
        If Not IsNumeric(Text(ptxID_No).Text) Then
'            Beep                               '�p�����G���[�ɂ��Ȃ�
'            MsgBox "���͂������ڂ̓G���[�ł��B"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxID_No).Text = Format(CDbl(Text(ptxID_No).Text), "00000000")
        End If
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'        Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))2004.04.08
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Text(ptxID_No).Text)
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i�o�ח\��o�^�ς݁j"
                Text(ptxID_No).SetFocus
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
                                '�`�[��
    If Len(Text(ptxNo).Text) = 0 Then
    Else
        If Not IsNumeric(Text(ptxNo).Text) Then
'            Beep                               '�p�����G���[�ɂ��Ȃ�
'            MsgBox "���͂������ڂ̓G���[�ł��B"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxNo).Text = Format(CLng(Text(ptxNo).Text), "000000")
        End If
    
    End If
                                                '������`�F�b�N
    Call UniCode_Conv(K0_MTS.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
                                                '������R�[�h
    Call UniCode_Conv(K0_MTS.SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B�i������j"
            If MTS_Set_Proc() Then
                Err_Chk = SYS_ERR
                Exit Function
            End If
            Combo(pcmbMUKE_CODE).SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "������}�X�^")
            Err_Chk = SYS_ERR
            Exit Function
    End Select
    
    Err_Chk = False
    
    
End Function
                                            '�o�ח\��̒ǉ�
Private Function Update_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer

Dim ID_NO   As String * 12
Dim DEN_NO  As String * 6
    
    Update_Proc = True
                                            '�o�ח\��ҏW
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                  '�g�p�q�@�h�c
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                  '�g�p���v���O����
    Call UniCode_Conv(Y_SYUREC.KAN_KBN, "0")                                '�����敪
    Call UniCode_Conv(Y_SYUREC.DT_SYU, "R")                                 '�f�[�^���
    Call UniCode_Conv(Y_SYUREC.JGYOBU, Last_JGYOBU)                         '���ƕ��敪
    
                                                                            '�����敪
    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))
    Call UniCode_Conv(Y_SYUREC.CYU_KBN, Right(Combo(pcmbC_Kbn).Text, 1))
    
                                                                            '�h�c��
    If Len(Trim(Text(ptxID_No).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Text(ptxID_No).Text)
        Call UniCode_Conv(Y_SYUREC.ID_NO, Text(ptxID_No).Text)
    Else
        
        
        If OSAKA_MODE = "1" Then
        
            sts = Den_No_Set_Proc(31, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Trim(ID_NO) & "01")
            Call UniCode_Conv(Y_SYUREC.ID_NO, Trim(ID_NO) & "01")
        Else
        
            sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
            Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
        End If
    End If
                                                                            
    Call UniCode_Conv(Y_SYUREC.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))    '�����O
                                                                    
    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, Text(ptxCode).Text)              '�i�ڔԍ�
    Call UniCode_Conv(Y_SYUREC.HIN_NO, Text(ptxCode).Text)                  '�i�ڔԍ�
                                                                            '���Ӑ�R�[�h
    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
                                                                            '������R�[�h
    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    Call UniCode_Conv(Y_SYUREC.SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
                                                                            '�o�ד�
    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, (Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text))
                
    

    
    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")                                  '���Ə�
    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")                                '�f�[�^�敪
    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")                                '����敪
                                                                            
    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                                                                            
                                                                            '�`�[��
    If Len(Trim(Text(ptxNo).Text)) <> 0 Then
        Call UniCode_Conv(Y_SYUREC.DEN_NO, Text(ptxNo).Text)
    Else
        
        
        If OSAKA_MODE = "1" Then
            sts = Den_No_Set_Proc(32, Last_JGYOBU, ID_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.DEN_NO, Trim(ID_NO))
        Else
        
            sts = Den_No_Set_Proc(20, Last_JGYOBU, DEN_NO)
            If sts Then
                Update_Proc = SYS_ERR
                Exit Function
            End If
            Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
        End If
    End If
                                                                            '�o�ɐ���
    Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
        
    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")                             '�o�Ɏ��x
    
    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
    
    
    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                                 '�I�[�_�[�ԍ�
    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                                 '�A�C�e���ԍ�
    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                                                                            '���Ӑ於��
    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
                                                                            '�����敪����
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, Left(Combo(pcmbC_Kbn).Text, Len(Combo(pcmbC_Kbn).Text) - 1))
                                                                            '�i��
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        
    
    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
        
        
    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
        
        
        
        
        
        
        
                                                                            '�z�X�g�I��
    Call UniCode_Conv(Y_SYUREC.HTANABAN, Text(ptxS_No).Text & Text(ptxRetu).Text & Text(ptxRen).Text & Text(ptxDan).Text)
    
    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")                               '�������t
    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")                                 '�������t
    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")                              '���i���t
    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")                                 '������敪
    
    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")                      '���ѐ���
    
                                                                            '�X�V����
    Call UniCode_Conv(Y_SYUREC.INS_NOW, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                                                            '2006.07.20
    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")              '2006.07.20
    
    
    
    Call UniCode_Conv(Y_SYUREC.FILLER, "")

    Do
        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case BtErrDuplicates
                If Len(Trim(Text(ptxID_No).Text)) = 0 Then
                                            '�������ԃf�[�^�d���͍Ĕ��s
                    sts = Den_No_Set_Proc(21, Last_JGYOBU, ID_NO)
                    If sts Then
                        Update_Proc = SYS_ERR
                        Exit Function
                    End If
    
                    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                
                Else
                    Call File_Error(sts, BtOpInsert, "�o�ח\��f�[�^")
                    Update_Proc = SYS_ERR
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�o�ח\��f�[�^")
                Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
    
    
    If OSAKA_MODE = "1" Then
    
    
    
        '-------------------------------------------    �o�ח\��(νĲҰ��)
        'ID_NO
        Call UniCode_Conv(Y_SYU_HREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        '��
        Call UniCode_Conv(Y_SYU_HREC.SYUKA_NO, "")
        '�o�ד�
        Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, Text(ptxYY).Text & _
                                                Text(ptxMM).Text & _
                                                Text(ptxDD).Text)
        '����於
        Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, "")
        '����`
        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "0")
        '�`�[�ԍ�
        Call UniCode_Conv(Y_SYU_HREC.DEN_NO, Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 7))
        '�ǔ�
        Call UniCode_Conv(Y_SYU_HREC.SEQ_NO, Right(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)), 1))
        '�i��
        Call UniCode_Conv(Y_SYU_HREC.HIN_NO, Text(ptxCode).Text)
        '����
        Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(Text(ptxS_Qty).Text), "0000000"))
        '������
        Call UniCode_Conv(Y_SYU_HREC.ODER_NO, "")
        '���Ӑ�
        Call UniCode_Conv(Y_SYU_HREC.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        '���Ӑ於
        Call UniCode_Conv(Y_SYU_HREC.MUKE_NAME, StrConv(MTSREC.MUKE_NAME, vbUnicode))
        '���l
        Call UniCode_Conv(Y_SYU_HREC.BIKOU, "")
        '�^����Ж�
        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "")
        '�捞�ݓ����i���͓����j
        Call UniCode_Conv(Y_SYU_HREC.INS_NOW, Format(Now, "YYYYMMDDHHMMSS"))
        '�o�����و�������i���͓����j
        Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, "")
        
        
        '�ް�������
        Call UniCode_Conv(Y_SYU_HREC.DATA_CNT, "00001")
        '�����
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
        '���i����
        Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
        '���i�S����
        Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
        '����
        Call UniCode_Conv(Y_SYU_HREC.xKUTI_SU, "")
        '���������׸�
        Call UniCode_Conv(Y_SYU_HREC.KYOSEI_END, "")
        '��ݾ��׸�
        Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "")
        '���͔��l
        Call UniCode_Conv(Y_SYU_HREC.INPUT_BIKOU, "")
        '���͔��l
        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, "09")
        '���͔��l
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "")
        '���ƕ�
        Call UniCode_Conv(Y_SYU_HREC.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
        '�����O
        Call UniCode_Conv(Y_SYU_HREC.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
        '�o�ɕ\��
        Call UniCode_Conv(Y_SYU_HREC.SYU_NO, "")
        '�o�Ɏ��ѐ�
        Call UniCode_Conv(Y_SYU_HREC.J_SURYO, "")
        '�W�񑗂��CD
        Call UniCode_Conv(Y_SYU_HREC.COL_OKURISAKI_CD, "")
        '�����CD
        Call UniCode_Conv(Y_SYU_HREC.OKURISAKI_CD, "")
        '�Z��
        Call UniCode_Conv(Y_SYU_HREC.JYUSHO, "")
        '�d�b�ԍ�
        Call UniCode_Conv(Y_SYU_HREC.TEL_NO, "")
        '�X�֔ԍ�
        Call UniCode_Conv(Y_SYU_HREC.YUBIN_NO, "")
        '�d��
        Call UniCode_Conv(Y_SYU_HREC.JURYO, "")
        '�ː�
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU, "")
        '����󇂁@�}��
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, "")
        '����敪
        Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, "")
        '����(�P��)
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, "")
        '�ː�(�P��)
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, "")
        '����󇂁@�}��
        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ_TO, "")
        '�ː�(�P��:�C���s��)    2010.11.01
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN_SAV, "")
        '�ː��v�Z�l(����P��)   2010.11.01
        Call UniCode_Conv(Y_SYU_HREC.SAI_SU_CALC, "")
        '�����v�Z�l(����P��)   2010.11.9
        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_CALC, "")
        '���Ǉ��@�@�@���Ǘ���(��)   2011.04.30
        Call UniCode_Conv(Y_SYU_HREC.SEK_KEN_NO, "")
        '�i�Ǉ��@�@�@���Ǘ���(��)   2011.04.30
        Call UniCode_Conv(Y_SYU_HREC.SEK_HIN_NO, "")
        '�����ް��ƍ��S��       2011.05.02
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, "")
        '�����ް��ƍ�����       2011.05.02
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, "")
        '���i���с@�o��     2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.CNT_BARA_SU, "")
        '���i���с@��       2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.CNT_HAKO_SU, "")
        '�O�����萔         2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.GAISO_IRI_QTY, "")
        '�i�ԓǍ��݉�     2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.Y_HIN_CHK_CNT, "")
        '�i�ԓǍ��ݍς݉� 2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.J_HIN_CHK_CNT, "")
        '���i���i��         2012.10.02
        Call UniCode_Conv(Y_SYU_HREC.KEN_HINBAN, "")
        '���X�R�[�h         2017.02.08
        Call UniCode_Conv(Y_SYU_HREC.TYAKUTEN, "")
        'FILLER
        Call UniCode_Conv(Y_SYU_HREC.FILLER, "")
        '�ǉ��@�S����       2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.INS_TANTO, "F103070")
        '�ǉ��@����         2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
        '�X�V�@�S����       2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, "")
        '�X�V�@����         2011.05.06
        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, "")
        
        
        
        Do
            sts = BTRV(BtOpInsert, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                Case BtErrDuplicates
                    '�ް��̖�������������̂ł��̂܂܏������f
                    ans = MsgBox("�`�[ID���A�d�����Ă��܂��B�X�V�������~���܂�", vbOK, "�m�F����")
                    
                    Call File_Error(sts, BtOpInsert, "�o�ח\��(ν��ް�)�f�[�^")
                    Update_Proc = SYS_CANCEL
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpInsert, "�o�ח\��(ν��ް�)�f�[�^")
                    Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    End If
    
    If SYUKA_LOG_ON Then
        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
    End If

    Beep
    MsgBox "�`�[��:" & StrConv(Y_SYUREC.DEN_NO, vbUnicode) & " ID:" & StrConv(Y_SYUREC.ID_NO, vbUnicode)

    Update_Proc = False
End Function


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
                                            '���͍��ڂ̃N���A�[�Ƃ݂Ȃ�
    Select Case Index
        Case pcmbC_Kbn
            Combo(pcmbNAIGAI).SetFocus
        Case pcmbNAIGAI
            Text(ptxCode).SetFocus
            
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE) = Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))
    End Select

End Sub

Private Sub Combo_LostFocus(Index As Integer)

    Select Case Index
        Case pcmbMUKE_CODE
            Text(ptxMUKE_CODE) = Trim(Right(Combo(pcmbMUKE_CODE).Text, 16))
    
    
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0                      '�X�V
                                    '�G���[�`�F�b�N
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                sts = Update_Proc()
                Select Case sts
                    Case False, True
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
            'Call Clear_Field()
            Text(ptxCode) = ""
            Text(ptxNo) = ""
            Text(ptxName) = ""      '2020/04/14 �i����
            Text(ptxS_Qty) = ""     '2020/04/14 �o�ח\�萔��
            Text(ptxS_No) = ""      '2020/04/14 �q�ɋ�
            Text(ptxRetu) = ""      '2020/04/14 ���
            Text(ptxRen) = ""       '2020/04/14 �A��
            Text(ptxDan) = ""       '2020/04/14 �i��
            Text(ptxCode).SetFocus
            Call Text_GotFocus(ptxCode)
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

                                '���ƕ���荞��
    If JGYOB_TB_Set() Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030701.Caption = "�o�ח\��o�^�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)



                                '��̫�Ē����敪��荞�� 2009.04.14
    If GetIni(App.EXEName, "DEF_CYU_KBN", "SYS", c) Then
        DEF_CYU_KBN = CYU_KBN_TUK
    Else
        DEF_CYU_KBN = Trim(c)
    End If
                                '���H 2010.03.23
    If GetIni(App.EXEName, "OSAKA_MODE", "SYS", c) Then
        OSAKA_MODE = "0"
    Else
        OSAKA_MODE = Trim(c)
    End If





                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�t�@�C���n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                '�o�ח\��f�[�^�t�@�C���n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '��ʏ����ݒ�
    Combo(pcmbC_Kbn).Clear
    Combo(pcmbC_Kbn).AddItem CYU_KBN_1$ & Space(5) & CYU_KBN_TUK$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_2$ & Space(5) & CYU_KBN_SPO$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_3$ & Space(5) & CYU_KBN_HJU$
    Combo(pcmbC_Kbn).AddItem CYU_KBN_E$ & Space(5) & CYU_KBN_BOU$
    '��97.08.06 �u�����i�ً}�j�v�̗\��͑��݂��Ȃ�
'    Combo(pcmbC_Kbn).AddItem CYU_KBN_T$ & Space(5) & CYU_KBN_KIN$
    'Combo(pcmbC_Kbn).Text = CYU_KBN_1$
    '��2001.03.28 �u�����v�̗\����o�^�Ƃ����I
'    Combo(pcmbC_Kbn).AddItem CYU_KBN_4$ & Space(5) & CYU_KBN_TOK$
    
    
    '2009.04.14
    Combo(pcmbC_Kbn).ListIndex = 0
    For i = 0 To Combo(pcmbC_Kbn).ListCount - 1
        If DEF_CYU_KBN = Right(Combo(pcmbC_Kbn).List(i), 1) Then
            Combo(pcmbC_Kbn).ListIndex = i
            Exit For
        End If
    Next i
    
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & Space(5) & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & Space(5) & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
                        '������ݒ�
    If MTS_Set_Proc() Then
        Unload Me
    End If
    
    Call Clear_Field
    Text(ptxYY) = Mid(Date, 1, 4)
    Text(ptxMM) = Mid(Date, 6, 2)
    Text(ptxDD) = Mid(Date, 9, 2)  '2020/04/14 �`�[���t��{�����t�����\���ɕύX
    
    Text(ptxID_No).SetFocus

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
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^")
        End If
    End If
                                            '�o�ח\��f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^�t�@�C��")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030701 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1030701.Caption = "�o�ח\��o�^�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf  As String
Dim i       As Integer
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Select Case Index
        Case ptxCode                '�i�ڃR�[�h
            
            
            
            If Trim(Text(ptxCode).Text) = "" Then   '2018.04.20
                Beep                                '2018.04.20
                MsgBox "�i�Ԃ͕K�{���͂ł��B "      '2018.04.20
                Text(ptxCode).SetFocus              '2018.04.20
                Exit Sub                            '2018.04.20
            End If                                  '2018.04.20
            
            
            
            
            sts = Item_Dsp()
            Select Case sts
                Case False
                Case True
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                    Text(Index).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
        Case ptxMUKE_CODE         '������i�R�[�h���͗p�j
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '������
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            
    End Select
    
    For i = Index + 1 To ptxMAX
        If Text(i).Visible And Text(i).Enabled And Not Text(i).Locked Then
            Text(i).SetFocus
            Exit Sub
        End If
    Next i
    Combo(pcmbMUKE_CODE).SetFocus
    
End Sub
Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String


    MTS_Set_Proc = True
    
    com = BtOpGetFirst
    
    Combo(pcmbMUKE_CODE).Clear
    
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������}�X�^")
                MTS_Set_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        Combo(pcmbMUKE_CODE).AddItem Edit
    
    
        com = BtOpGetNext
    Loop




    If Combo(pcmbMUKE_CODE).ListCount = 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If


    MTS_Set_Proc = False
End Function


Private Sub Text_LostFocus(Index As Integer)
    
    If Index = ptxCode Then
        Text(ptxCode).Text = StrConv(Text(ptxCode).Text, vbUpperCase)
    End If
End Sub

VERSION 5.00
Begin VB.Form CONV20060515F 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^�R���o�[�g����"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   ControlBox      =   0   'False
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
   ScaleHeight     =   6495
   ScaleWidth      =   11070
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8760
      MaxLength       =   2
      TabIndex        =   24
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   20
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�V�ړ����폜"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   3720
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ݾ�"
      Height          =   555
      Index           =   2
      Left            =   7980
      TabIndex        =   18
      Top             =   5820
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ްĊJ�n"
      Height          =   555
      Index           =   1
      Left            =   5160
      TabIndex        =   17
      Top             =   5820
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�S�I��"
      Height          =   1095
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   3540
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�o�ח\��@�@�@�@��"
      Height          =   375
      Index           =   6
      Left            =   1380
      TabIndex        =   15
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���ח\��@�@�@�@��"
      Height          =   375
      Index           =   5
      Left            =   1380
      TabIndex        =   14
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���i���w�}(�q)�@��"
      Height          =   375
      Index           =   4
      Left            =   1380
      TabIndex        =   13
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��Ǝ��у��O�@�@��"
      Height          =   375
      Index           =   3
      Left            =   1380
      TabIndex        =   12
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�݌Ɉړ����@�@�@��"
      Height          =   375
      Index           =   2
      Left            =   1380
      TabIndex        =   11
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���ԃ}�X�^�@�@�@��"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�폜�ςݏo�ח\�聁"
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   9
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Height          =   375
      Left            =   7320
      TabIndex        =   26
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "���ȑO��"
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   8520
      TabIndex        =   23
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�N"
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   22
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3900
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   3900
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3900
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3900
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3900
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3900
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3900
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "�f�[�^�R���o�[�g(ID_NO 8��12��)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "CONV20060515F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Convert_Proc() As Integer
Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim count           As Long

Dim DISP_INTERVAL   As Long

Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

Dim lngWk           As Double

    Convert_Proc = True

'---------------------------------------------  �폜�ςݏo�ח\��̃R���o�[�g
Convert_P0:
    If Check1(0).Value <> 1 Then GoTo Convert_P1

    MsgLab(1) = "�폜�ςݏo�ח\��R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), K0_O_DEL_SYU, Len(K0_O_DEL_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�폜�ςݏo�ח\��")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(DEL_SYUREC.WEL_ID, StrConv(O_DEL_SYUREC.WEL_ID, vbUnicode))   '�g�p�q�@ID
        Call UniCode_Conv(DEL_SYUREC.PRG_ID, StrConv(O_DEL_SYUREC.PRG_ID, vbUnicode))   '�g�p����۸���
        Call UniCode_Conv(DEL_SYUREC.KAN_KBN, StrConv(O_DEL_SYUREC.KAN_KBN, vbUnicode)) '�����敪
        Call UniCode_Conv(DEL_SYUREC.DT_SYU, StrConv(O_DEL_SYUREC.DT_SYU, vbUnicode))   '�ް����
        Call UniCode_Conv(DEL_SYUREC.JGYOBU, StrConv(O_DEL_SYUREC.JGYOBU, vbUnicode))   '���ƕ��敪
        Call UniCode_Conv(DEL_SYUREC.KEY_CYU_KBN, _
                                        StrConv(O_DEL_SYUREC.KEY_CYU_KBN, vbUnicode))   '�����敪
        Call UniCode_Conv(DEL_SYUREC.KEY_ID_NO, _
                                        StrConv(O_DEL_SYUREC.KEY_ID_NO, vbUnicode))     'ID-NO(8����12��)
        Call UniCode_Conv(DEL_SYUREC.NAIGAI, StrConv(O_DEL_SYUREC.NAIGAI, vbUnicode))   '�����O
        Call UniCode_Conv(DEL_SYUREC.KEY_HIN_NO, _
                                        StrConv(O_DEL_SYUREC.KEY_HIN_NO, vbUnicode))    '�i�ڔԍ�
        Call UniCode_Conv(DEL_SYUREC.KEY_MUKE_CODE, _
                                    StrConv(O_DEL_SYUREC.KEY_MUKE_CODE, vbUnicode))     '���Ӑ�R�[�h
        Call UniCode_Conv(DEL_SYUREC.KEY_SS_CODE, _
                                    StrConv(O_DEL_SYUREC.KEY_SS_CODE, vbUnicode))       '������R�[�h
        Call UniCode_Conv(DEL_SYUREC.KEY_SYUKA_YMD, _
                                    StrConv(O_DEL_SYUREC.KEY_SYUKA_YMD, vbUnicode))     '�o�ד��t
        Call UniCode_Conv(DEL_SYUREC.JGYOBA, StrConv(O_DEL_SYUREC.JGYOBA, vbUnicode))   '���Ə�
        Call UniCode_Conv(DEL_SYUREC.DATA_KBN, _
                                        StrConv(O_DEL_SYUREC.DATA_KBN, vbUnicode))      '�f�[�^�敪
        Call UniCode_Conv(DEL_SYUREC.TORI_KBN, _
                                        StrConv(O_DEL_SYUREC.TORI_KBN, vbUnicode))      '����敪
        Call UniCode_Conv(DEL_SYUREC.ID_NO, StrConv(O_DEL_SYUREC.ID_NO, vbUnicode))     'ID-NO(8����12��)

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.KAIKEI_JGYOBA, "")                         '��v�p���Ə꺰��
        Call UniCode_Conv(DEL_SYUREC.SHISAN_JGYOBA, "")                         '���Y�Ǘ��p���Ə꺰��
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.HIN_NO, StrConv(O_DEL_SYUREC.HIN_NO, vbUnicode))   '�i�ڔԍ�
        Call UniCode_Conv(DEL_SYUREC.DEN_NO, StrConv(O_DEL_SYUREC.DEN_NO, vbUnicode))   '�`�[�ԍ�
        Call UniCode_Conv(DEL_SYUREC.SURYO, StrConv(O_DEL_SYUREC.SURYO, vbUnicode))     '�o�ɐ���
        Call UniCode_Conv(DEL_SYUREC.MUKE_CODE, _
                                        StrConv(O_DEL_SYUREC.MUKE_CODE, vbUnicode))     '���Ӑ�R�[�h
        Call UniCode_Conv(DEL_SYUREC.SYUKO_SYUSI, _
                                        StrConv(O_DEL_SYUREC.SYUKO_SYUSI, vbUnicode))   '�݌Ɏ��x

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.SHISAN_SYUSI, "")                          '���Y�Ǘ��p�݌Ɏ��x����
        Call UniCode_Conv(DEL_SYUREC.HOJYO_SYUSI, "")                           '�⏕�݌Ɏ��x����
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.SYUKA_YMD, _
                                        StrConv(O_DEL_SYUREC.SYUKA_YMD, vbUnicode))     '�o�ד��t

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.TANKA, _
                                    StrConv(O_DEL_SYUREC.TANKA, vbUnicode))     '���ےP��
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.ODER_NO, _
                                        StrConv(O_DEL_SYUREC.ODER_NO, vbUnicode))       '�I�[�_�[�ԍ�
        Call UniCode_Conv(DEL_SYUREC.ITEM_NO, StrConv(O_DEL_SYUREC.ITEM_NO, vbUnicode)) '�A�C�e���ԍ�

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.ODER_NO_R, "")                             '�����Ǘ��ԍ�����
        Call UniCode_Conv(DEL_SYUREC.KOSO_KEITAI, "")                           '���`�Ժ���
        Call UniCode_Conv(DEL_SYUREC.SYUKO_YMD, "")                             '�o�ח\���
        Call UniCode_Conv(DEL_SYUREC.TANABAN1, "")                              '۹����1
        Call UniCode_Conv(DEL_SYUREC.TANABAN2, "")                              '۹����2
        Call UniCode_Conv(DEL_SYUREC.TANABAN3, "")                              '۹����3
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.MUKE_NAME, _
                                        StrConv(O_DEL_SYUREC.MUKE_NAME, vbUnicode))     '���Ӑ於��
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN, StrConv(O_DEL_SYUREC.CYU_KBN, vbUnicode)) '�����敪
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN_NAME, _
                                        StrConv(O_DEL_SYUREC.CYU_KBN_NAME, vbUnicode))  '�����敪����

'''        Call UniCode_Conv(DEL_SYUREC.EXPORT_KBN, _
'''                                        StrConv(O_DEL_SYUREC.EXPORT_KBN, vbUnicode))    '�A�o�o�׌����敪
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_KBN, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '�����ٔ��s�敪
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_UNIT, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_ISSUE_UNIT, vbUnicode))  '�����ٔ��s�P�ʐ�
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_TANKA_KBN, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '�����ْP���\���敪
'''        Call UniCode_Conv(DEL_SYUREC.TANKA, StrConv(O_DEL_SYUREC.TANKA, vbUnicode))     '�P��
'''        Call UniCode_Conv(DEL_SYUREC.KINGAKU, StrConv(O_DEL_SYUREC.KINGAKU, vbUnicode)) '���z
'''        Call UniCode_Conv(DEL_SYUREC.BIKOU2, StrConv(O_DEL_SYUREC.BIKOU2, vbUnicode))   '���l�Q
'''        Call UniCode_Conv(DEL_SYUREC.REBATE_KBN, _
'''                                        StrConv(O_DEL_SYUREC.REBATE_KBN, vbUnicode))    '���x�[�g�敪
'''        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, _
'''                                        StrConv(O_DEL_SYUREC.CHOHA_KBN, vbUnicode))     '���[�敪
'''        Call UniCode_Conv(DEL_SYUREC.ATAISA_KBN, _
'''                                        StrConv(O_DEL_SYUREC.ATAISA_KBN, vbUnicode))    '�l���敪
'''        Call UniCode_Conv(DEL_SYUREC.REP_KISHU, _
'''                                        StrConv(O_DEL_SYUREC.REP_KISHU, vbUnicode))     '��\�@��
'''        Call UniCode_Conv(DEL_SYUREC.NS_KANRI_NO, _
'''                                        StrConv(O_DEL_SYUREC.NS_KANRI_NO, vbUnicode))   '�m�r�Ǘ��ԍ�
'''        Call UniCode_Conv(DEL_SYUREC.MTS_HIN_CODE, _
'''                                        StrConv(O_DEL_SYUREC.MTS_HIN_CODE, vbUnicode))  '�l�s�r���i�R�[�h
'''        Call UniCode_Conv(DEL_SYUREC.BIKOU1, StrConv(O_DEL_SYUREC.BIKOU1, vbUnicode))   '���l�P
'''        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, _
'''                                        StrConv(O_DEL_SYUREC.CHOKU_KBN, vbUnicode))     '�����敪
'''        Call UniCode_Conv(DEL_SYUREC.REBATE_RATE, _
'''                                        StrConv(O_DEL_SYUREC.REBATE_RATE, vbUnicode))   '���x�[�g��
'''        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, _
'''                                        StrConv(O_DEL_SYUREC.HIN_NAME, vbUnicode))      '�i��
'''        Call UniCode_Conv(DEL_SYUREC.JGYOBA_GAI, _
'''                                        StrConv(O_DEL_SYUREC.JGYOBA_GAI, vbUnicode))    '�ΊO���Ə�
'''        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, _
'''                                        StrConv(O_DEL_SYUREC.KISHU_CODE, vbUnicode))    '�@��R�[�h
'''        Call UniCode_Conv(DEL_SYUREC.SS_CODE, StrConv(O_DEL_SYUREC.SS_CODE, vbUnicode)) '������R�[�h


'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.ORIGIN1, "")                              '���Y��1
        Call UniCode_Conv(DEL_SYUREC.ORIGIN2, "")                              '���Y��2
        Call UniCode_Conv(DEL_SYUREC.BIKOU2, _
                                        StrConv(O_DEL_SYUREC.BIKOU2, vbUnicode))    '���l2
        Call UniCode_Conv(DEL_SYUREC.HAN_KBN, "")                                   '�̔��敪
        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, _
                                        StrConv(O_DEL_SYUREC.CHOKU_KBN, vbUnicode)) '�����w���敪
        Call UniCode_Conv(DEL_SYUREC.UNIT_ID_NO, "")                                   '�ƯďC���Ǘ��ԍ�
        Call UniCode_Conv(DEL_SYUREC.ZAIKO_HIKIATE, "")                               '�݌Ɉ�������
        Call UniCode_Conv(DEL_SYUREC.GOKON_KANRI_NO, "")                                  '�����Ǘ��ԍ�
        Call UniCode_Conv(DEL_SYUREC.JYUCHU_ZAN, "")                                '�󒍎c����
        Call UniCode_Conv(DEL_SYUREC.KYOKYU_KBN, "")                                '�����敪
        Call UniCode_Conv(DEL_SYUREC.SHOHIN_SYUSI, "")                             '���i���[�i�݌Ɏ��x����
        Call UniCode_Conv(DEL_SYUREC.S_SHISAN_SYUSI, "")                            '���i���[�i���Y�Ǘ����x����
        Call UniCode_Conv(DEL_SYUREC.S_HOJYO_SYUSI, "")                             '���i���[�i�⏕���x����
        Call UniCode_Conv(DEL_SYUREC.BIKOU1, _
                                        StrConv(O_DEL_SYUREC.BIKOU1, vbUnicode))    '���l1
        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, _
                                        StrConv(O_DEL_SYUREC.CHOHA_KBN, vbUnicode)) '���[�敪
        Call UniCode_Conv(DEL_SYUREC.JYU_HIN_NO, "")                                '��t�i�ڔԍ�
        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, _
                                        StrConv(O_DEL_SYUREC.HIN_NAME, vbUnicode))  '�i��
        Call UniCode_Conv(DEL_SYUREC.HIN_CHANGE_KBN, "")                          '�i�ڔԍ��ύX�敪
        Call UniCode_Conv(DEL_SYUREC.MODULE_EXCHANGE, "")                                'Ӽޭ�ٌ����敪
        Call UniCode_Conv(DEL_SYUREC.ZAIKO_SYUSI, "")                           '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
        Call UniCode_Conv(DEL_SYUREC.ZAN_SHISAN_SYUSI, "")                          '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
        Call UniCode_Conv(DEL_SYUREC.ZAN_HOJYO_SYUSI, "")                           '�c�݌ɂ܂Ƃߕ⏕���x����
        Call UniCode_Conv(DEL_SYUREC.NOUKI_YMD, "")                                     '�w��[��
        Call UniCode_Conv(DEL_SYUREC.SERVICE_KANRI_NO, "")                                '���޽��ЊǗ��ԍ�
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, _
                                    StrConv(O_DEL_SYUREC.KISHU_CODE, vbUnicode))    '�@��i�ں���
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, "")                                 '����敔�i�敪
        Call UniCode_Conv(DEL_SYUREC.SS_CODE, _
                                    StrConv(O_DEL_SYUREC.SS_CODE, vbUnicode))       '��������溰��
        Call UniCode_Conv(DEL_SYUREC.KEPIN_KAIJYO, "")                              '���i�����敪
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.HIN_NAI, StrConv(O_DEL_SYUREC.HIN_NAI, vbUnicode)) '�i�ԁi�����j
        Call UniCode_Conv(DEL_SYUREC.HTANABAN, _
                                        StrConv(O_DEL_SYUREC.HTANABAN, vbUnicode))      '�z�X�g�I��
        Call UniCode_Conv(DEL_SYUREC.PRINT_YMD, _
                                        StrConv(O_DEL_SYUREC.PRINT_YMD, vbUnicode))     '�o�ɕ\������t
        Call UniCode_Conv(DEL_SYUREC.KAN_YMD, StrConv(O_DEL_SYUREC.KAN_YMD, vbUnicode)) '�������t
        Call UniCode_Conv(DEL_SYUREC.KENPIN_YMD, _
                                        StrConv(O_DEL_SYUREC.KENPIN_YMD, vbUnicode))    '���i���t
        Call UniCode_Conv(DEL_SYUREC.TOK_KBN, StrConv(O_DEL_SYUREC.TOK_KBN, vbUnicode)) '������敪
        Call UniCode_Conv(DEL_SYUREC.JITU_SURYO, _
                                        StrConv(O_DEL_SYUREC.JITU_SURYO, vbUnicode))    '�o�Ɏ��ѐ���
        Call UniCode_Conv(DEL_SYUREC.INS_NOW, StrConv(O_DEL_SYUREC.INS_NOW, vbUnicode)) '�捞�ݓ���
        Call UniCode_Conv(DEL_SYUREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<DEL_SYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�폜�ςݏo�ח\��")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(0).Caption = Format(count, "#0")

'---------------------------------------------  ���ԃ}�X�^�̃R���o�[�g
Convert_P1:
    If Check1(1).Value <> 1 Then GoTo Convert_P2

    MsgLab(1) = "���ԃ}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), K0_O_HATUBAN, Len(K0_O_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j���ԃ}�X�^")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(HATUBANREC.JGYOBU, StrConv(O_HATUBANREC.JGYOBU, vbUnicode))   '���ƕ��敪
        Call UniCode_Conv(HATUBANREC.NYK_KBN, StrConv(O_HATUBANREC.NYK_KBN, vbUnicode)) '���ד`�[���敪
        Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, _
                                        StrConv(O_HATUBANREC.NYK_DEN_NO, vbUnicode))    '�����ד`�[��
        Call UniCode_Conv(HATUBANREC.SYK_KBN, StrConv(O_HATUBANREC.SYK_KBN, vbUnicode)) '�o�ד`�[���敪
        Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, _
                                        StrConv(O_HATUBANREC.SYK_DEN_NO, vbUnicode))    '���o�ד`�[��
        Call UniCode_Conv(HATUBANREC.NYK_ID_KBN, _
                                        StrConv(O_HATUBANREC.NYK_ID_KBN, vbUnicode))    '����ID���敪
        lngWk = Val(StrConv(O_HATUBANREC.NYK_ID_NO, vbUnicode))
        Call UniCode_Conv(HATUBANREC.NYK_ID_NO, Format(lngWk, String(11, "0")))         '������ID��(8����11��)
        Call UniCode_Conv(HATUBANREC.SYK_ID_KBN, _
                                        StrConv(O_HATUBANREC.SYK_ID_KBN, vbUnicode))    '�o��ID���敪
        lngWk = Val(StrConv(O_HATUBANREC.SYK_ID_NO, vbUnicode))
        Call UniCode_Conv(HATUBANREC.SYK_ID_NO, Format(lngWk, String(11, "0")))         '���o��ID��(7����11��)
        Call UniCode_Conv(HATUBANREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���ԃ}�X�^")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop


    Cnt(1).Caption = Format(count, "#0")

'---------------------------------------------  �݌Ɉړ����̃R���o�[�g
Convert_P2:
    If Check1(2).Value <> 1 Then GoTo Convert_P3

    MsgLab(1) = "�݌Ɉړ����R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_IDO_POS, O_IDOREC, Len(O_IDOREC), K0_O_IDO, Len(K0_O_IDO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�݌Ɉړ���")
                Exit Function
        End Select


        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(O_IDOREC.JITU_DT, vbUnicode))         '���ѓ��t
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(O_IDOREC.JITU_TM, vbUnicode))         '���ю���
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(O_IDOREC.JGYOBU, vbUnicode))           '���ƕ��敪
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(O_IDOREC.NAIGAI, vbUnicode))           '�����O
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(O_IDOREC.HIN_GAI, vbUnicode))         '�i�ځi�O���j
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(O_IDOREC.RIRK_ID, vbUnicode))         '�������
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, _
                                        StrConv(O_IDOREC.SUMI_JITU_QTY, vbUnicode))     '���ѐ���(���i���ς�)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, StrConv(O_IDOREC.MI_JITU_QTY, vbUnicode)) '���ѐ���(���ѐ���(�����i))
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(O_IDOREC.FROM_SOKO, vbUnicode))     'From �q�ɇ�
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(O_IDOREC.FROM_RETU, vbUnicode))     'From ��
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(O_IDOREC.FROM_REN, vbUnicode))       'From �A
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(O_IDOREC.FROM_DAN, vbUnicode))       'From �i
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(O_IDOREC.TO_SOKO, vbUnicode))         'TO �q�ɇ�
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(O_IDOREC.TO_RETU, vbUnicode))         'TO ��
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(O_IDOREC.TO_REN, vbUnicode))           'TO �A
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(O_IDOREC.TO_DAN, vbUnicode))           'TO �i
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(O_IDOREC.DEN_DT, vbUnicode))           '�`�[���t
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(O_IDOREC.DEN_NO, vbUnicode))           '�`�[��
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(O_IDOREC.PRG_ID, vbUnicode))           '�o�͌��v���O����
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(O_IDOREC.HIN_NAI, vbUnicode))         '�i�ԁi�����j
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(O_IDOREC.NYUKA_DT, vbUnicode))       '���ד��t
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(O_IDOREC.NYUKO_DT, vbUnicode))       '���ɓ��t
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(O_IDOREC.WEL_ID, vbUnicode))           '�Ώے[����
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(O_IDOREC.RIRK_NAME, vbUnicode))     '������ʖ���
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(O_IDOREC.HIN_NAME, vbUnicode))       '�i��
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                    StrConv(O_IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode))    '�i�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, _
                                    StrConv(O_IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))      '�i�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.SUMI_FROM_TANA_Zaiko_Qty, vbUnicode))  'FROM�I�ʕi�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.SUMI_TO_TANA_Zaiko_Qty, vbUnicode))    'TO�I�ʕi�ڕʍ݌ɐ��i���i���ς݁j
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.MI_FROM_TANA_Zaiko_Qty, vbUnicode))    'FROM�I�ʕi�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.MI_TO_TANA_Zaiko_Qty, vbUnicode))      'TO�I�ʕi�ڕʍ݌ɐ��i�����i�j
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(O_IDOREC.TOKU_MARK, vbUnicode))     '������}�[�N
        Call UniCode_Conv(IDOREC.MEMO, StrConv(O_IDOREC.MEMO, vbUnicode))               '����
        Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(O_IDOREC.TANTO_CODE, vbUnicode))                                       '�S���҃R�[�h
        Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(O_IDOREC.TANTO_NAME, vbUnicode))                                        '�S���Җ���
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(O_IDOREC.MUKE_CODE, vbUnicode))     '���Ӑ�R�[�h
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(O_IDOREC.MUKE_DNAME, vbUnicode))    '���Ӑ於��
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(O_IDOREC.SS_CODE, vbUnicode))                                           '������R�[�h
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(O_IDOREC.SS_NAME, vbUnicode))                                           '�����於��
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(O_IDOREC.MUKE_DNAME, vbUnicode))   '���Ӑ旪��
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(O_IDOREC.MUKE_CHG_CD, vbUnicode)) '������Ǒւ��R�[�h
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(O_IDOREC.SUM_KBN, vbUnicode))         '�W�v�敪
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(O_IDOREC.ID_NO, vbUnicode))             'ID-NO(8����12��)
        Call UniCode_Conv(IDOREC.Ins_DateTime, _
                                        StrConv(O_IDOREC.Ins_DateTime, vbUnicode))      '�}������
        Call UniCode_Conv(IDOREC.SHIIRE_CODE, StrConv(O_IDOREC.SHIIRE_CODE, vbUnicode)) '�d���溰��
        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, _
                                        StrConv(O_IDOREC.SHIIRE_TANKA, vbUnicode))      '�d���P��(9(8)V99)
        Call UniCode_Conv(IDOREC.KEIJYO_YM, StrConv(O_IDOREC.KEIJYO_YM, vbUnicode))     '�v��N��(YYYYMM)
        Call UniCode_Conv(IDOREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ɉړ���")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(2).Caption = Format(count, "#0")

'---------------------------------------------  ��Ǝ��у��O�̃R���o�[�g
Convert_P3:
    If Check1(3).Value <> 1 Then GoTo Convert_P4

    MsgLab(1) = "��Ǝ��у��O�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_P_SAGYO_LOG_POS, O_P_SAGYO_LOG_REC, Len(O_P_SAGYO_LOG_REC), K0_O_P_SAGYO_LOG, Len(K0_O_P_SAGYO_LOG), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j��Ǝ��у��O")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_DT, _
                                StrConv(O_P_SAGYO_LOG_REC.JITU_DT, vbUnicode))      '���ѓ��t
        Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_TM, _
                                StrConv(O_P_SAGYO_LOG_REC.JITU_TM, vbUnicode))      '���ю���
        Call UniCode_Conv(P_SAGYO_LOG_REC.TANTO_CODE, _
                                StrConv(O_P_SAGYO_LOG_REC.TANTO_CODE, vbUnicode))   '�S���҃R�[�h
        Call UniCode_Conv(P_SAGYO_LOG_REC.WEL_ID, _
                                StrConv(O_P_SAGYO_LOG_REC.WEL_ID, vbUnicode))       '�Ώے[����
        Call UniCode_Conv(P_SAGYO_LOG_REC.JGYOBU, _
                                StrConv(O_P_SAGYO_LOG_REC.JGYOBU, vbUnicode))       '���ƕ��敪
        Call UniCode_Conv(P_SAGYO_LOG_REC.NAIGAI, _
                                StrConv(O_P_SAGYO_LOG_REC.NAIGAI, vbUnicode))       '�����O
        Call UniCode_Conv(P_SAGYO_LOG_REC.MENU_NO, _
                                StrConv(O_P_SAGYO_LOG_REC.MENU_NO, vbUnicode))      '���j���[�O���[�v��
        Call UniCode_Conv(P_SAGYO_LOG_REC.RIRK_ID, _
                                StrConv(O_P_SAGYO_LOG_REC.RIRK_ID, vbUnicode))      '�������
        Call UniCode_Conv(P_SAGYO_LOG_REC.ID_NO, _
                                StrConv(O_P_SAGYO_LOG_REC.ID_NO, vbUnicode))        'ID-NO
        Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, _
                                StrConv(O_P_SAGYO_LOG_REC.HIN_GAI, vbUnicode))      '�i�ԁi�O���j
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, _
                            StrConv(O_P_SAGYO_LOG_REC.SUMI_JITU_QTY, vbUnicode))    '���ѐ���(���i���ς�)
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, _
                            StrConv(O_P_SAGYO_LOG_REC.MI_JITU_QTY, vbUnicode))      '���ѐ���(�����i)
        Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, _
                            StrConv(O_P_SAGYO_LOG_REC.MUKE_CODE, vbUnicode))        '���Ӑ�R�[�h
        Call UniCode_Conv(P_SAGYO_LOG_REC.SS_CODE, _
                            StrConv(O_P_SAGYO_LOG_REC.SS_CODE, vbUnicode))          '������R�[�h
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_SOKO, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_SOKO, vbUnicode))        'From �q�ɇ�
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_RETU, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_RETU, vbUnicode))        '   �@��
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_REN, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_REN, vbUnicode))         '   �@�A
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_DAN, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_DAN, vbUnicode))         '   �@�i
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_SOKO, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_SOKO, vbUnicode))          '�s�n �q�ɇ�
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_RETU, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_RETU, vbUnicode))          '   �@��
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_REN, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_REN, vbUnicode))           '   �@�A
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_DAN, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_DAN, vbUnicode))           '   �@�i
        Call UniCode_Conv(P_SAGYO_LOG_REC.PRG_ID, _
                            StrConv(O_P_SAGYO_LOG_REC.PRG_ID, vbUnicode))           '�o�͌��v���O����
        Call UniCode_Conv(P_SAGYO_LOG_REC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SAGYO_LOG.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "��Ǝ��у��O")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(3).Caption = Format(count, "#0")

'---------------------------------------------  ���i���w�}�i�q�j�̃R���o�[�g
Convert_P4:
    If Check1(4).Value <> 1 Then GoTo Convert_P5

    MsgLab(1) = "���i���w�}�i�q�j�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(4).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), K0_O_P_SSHIJI_K, Len(K0_O_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j���i���w�}�i�q�j")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(4).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If


        Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_NO, _
                                    StrConv(O_P_SSHIJI_K_REC.SHIJI_NO, vbUnicode))      '�w�}�[��
        Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, _
                                    StrConv(O_P_SSHIJI_K_REC.DATA_KBN, vbUnicode))      '�ް��敪
        Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, _
                                    StrConv(O_P_SSHIJI_K_REC.SEQNO, vbUnicode))         '�ǔ�
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode))   '�q ���
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))     '�q ���ƕ�
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))     '�q �����O
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))    '�q �i��
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_QTY, vbUnicode))        '�q ����(999V99)
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode))  '�w����(9(8)V99)
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))      '�q ���l
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_ID_NO, vbUnicode))      '�q ID_NO
        Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, _
                                    StrConv(O_P_SSHIJI_K_REC.CALCEL_F, vbUnicode))      '��ݾ�F
        Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, _
                                StrConv(O_P_SSHIJI_K_REC.CANCEL_DATETIME, vbUnicode))   '��ݾٓ���
        Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
        Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, _
                                    StrConv(O_P_SSHIJI_K_REC.UPD_DATETIME, vbUnicode))  '�X�V ����

        Do
            sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���i���w�}�i�q�j")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(4).Caption = Format(count, "#0")

'    GoTo Convert_End

'---------------------------------------------  ���ח\��̃R���o�[�g
Convert_P5:
    If Check1(5).Value <> 1 Then GoTo Convert_P6

    MsgLab(1) = "���ח\��R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(5).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j���ח\��")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(5).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(Y_NYUREC.KAN_KBN, StrConv(O_Y_NYUREC.KAN_KBN, vbUnicode))     '�����敪
        Call UniCode_Conv(Y_NYUREC.DT_SYU, StrConv(O_Y_NYUREC.DT_SYU, vbUnicode))       '�f�[�^���
        Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(O_Y_NYUREC.JGYOBU, vbUnicode))       '���ƕ��敪
        Call UniCode_Conv(Y_NYUREC.NAIGAI, StrConv(O_Y_NYUREC.NAIGAI, vbUnicode))       '�����O
        Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(O_Y_NYUREC.TEXT_NO, vbUnicode))     '�e�L�X�g��
        Call UniCode_Conv(Y_NYUREC.JGYOBA, StrConv(O_Y_NYUREC.JGYOBA, vbUnicode))       '���Ə�
        Call UniCode_Conv(Y_NYUREC.DATA_KBN, StrConv(O_Y_NYUREC.DATA_KBN, vbUnicode))   '�f�[�^�敪
        Call UniCode_Conv(Y_NYUREC.TORI_KBN, StrConv(O_Y_NYUREC.TORI_KBN, vbUnicode))   '����敪
        Call UniCode_Conv(Y_NYUREC.ID_NO, StrConv(O_Y_NYUREC.ID_NO, vbUnicode))         'ID-NO(8����12��)
        Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(O_Y_NYUREC.HIN_NO, vbUnicode))       '�i�ڔԍ�
        Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(O_Y_NYUREC.DEN_NO, vbUnicode))       '�`�[�ԍ�
        Call UniCode_Conv(Y_NYUREC.SURYO, StrConv(O_Y_NYUREC.SURYO, vbUnicode))         '�o�ɐ���
        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, StrConv(O_Y_NYUREC.MUKE_CODE, vbUnicode)) '�o�ɐ�
        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, _
                                        StrConv(O_Y_NYUREC.SYUKO_SYUSI, vbUnicode))     '�o�Ɏ��x
        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(O_Y_NYUREC.SYUKO_YMD, vbUnicode)) '�o�ɓ��t
        Call UniCode_Conv(Y_NYUREC.TANKA, StrConv(O_Y_NYUREC.TANKA, vbUnicode))         '�P��
        Call UniCode_Conv(Y_NYUREC.ODER_NO, StrConv(O_Y_NYUREC.ODER_NO, vbUnicode))     '�I�[�_�[�ԍ�
        Call UniCode_Conv(Y_NYUREC.ITEM_NO, StrConv(O_Y_NYUREC.ITEM_NO, vbUnicode))     '�A�C�e���ԍ�
        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, StrConv(O_Y_NYUREC.ODER_R_NO, vbUnicode)) '�I�[�_�[����
        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, _
                                        StrConv(O_Y_NYUREC.KOSO_KEITAI, vbUnicode))     '���`��
        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(O_Y_NYUREC.SYUKA_YMD, vbUnicode)) '�o�ד�
        Call UniCode_Conv(Y_NYUREC.TANABAN1, StrConv(O_Y_NYUREC.TANABAN1, vbUnicode))   '�I�ԂP
        Call UniCode_Conv(Y_NYUREC.TANABAN2, StrConv(O_Y_NYUREC.TANABAN2, vbUnicode))   '�I�ԂQ
        Call UniCode_Conv(Y_NYUREC.TANABAN3, StrConv(O_Y_NYUREC.TANABAN3, vbUnicode))   '�I�ԂR
        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, StrConv(O_Y_NYUREC.MUKE_NAME, vbUnicode)) '�o�ɐ於��
        Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(O_Y_NYUREC.CYU_KBN, vbUnicode))     '�����敪
        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, _
                                        StrConv(O_Y_NYUREC.CYU_KBN_NAME, vbUnicode))    '�����敪����
        Call UniCode_Conv(Y_NYUREC.ORIGIN1, StrConv(O_Y_NYUREC.ORIGIN1, vbUnicode))     '���Y���P
        Call UniCode_Conv(Y_NYUREC.ORIGIN2, StrConv(O_Y_NYUREC.ORIGIN2, vbUnicode))     '���Y���Q
        Call UniCode_Conv(Y_NYUREC.BIKOU2, StrConv(O_Y_NYUREC.BIKOU2, vbUnicode))       '���l�Q
        Call UniCode_Conv(Y_NYUREC.HAN_KBN, StrConv(O_Y_NYUREC.HAN_KBN, vbUnicode))     '�̔��敪
        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, StrConv(O_Y_NYUREC.CHOKU_KBN, vbUnicode)) '�����敪
        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, _
                                        StrConv(O_Y_NYUREC.UNIT_ID_NO, vbUnicode))      '�ƯďC��ID-NO
        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, _
                                        StrConv(O_Y_NYUREC.ZAIKO_HIKIATE, vbUnicode))   '�݌Ɉ�������
        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, _
                                        StrConv(O_Y_NYUREC.GOKON_KANRI_NO, vbUnicode))  '�����Ǘ��ԍ�
        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, StrConv(O_Y_NYUREC.JUCHU_ZAN, vbUnicode)) '�󒍎c����
        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, _
                                        StrConv(O_Y_NYUREC.KYOKYU_KBN, vbUnicode))      '�����敪
        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, _
                                        StrConv(O_Y_NYUREC.SHOHIN_SYUSI, vbUnicode))    '���i���[������x
        Call UniCode_Conv(Y_NYUREC.BIKOU1, StrConv(O_Y_NYUREC.BIKOU1, vbUnicode))       '���l�P
        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, StrConv(O_Y_NYUREC.CHOHA_KBN, vbUnicode)) '���[�敪
        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, _
                                        StrConv(O_Y_NYUREC.JYU_HIN_NO, vbUnicode))      '�󒍕i�ڔԍ�
        Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(O_Y_NYUREC.HIN_NAME, vbUnicode))   '�i��
        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, _
                                        StrConv(O_Y_NYUREC.HIN_CHANGE_KBN, vbUnicode))  '�i�ԕύX�敪
        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, _
                                        StrConv(O_Y_NYUREC.MODULE_EXCHANGE, vbUnicode)) '���W���[�������敪
        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, _
                                        StrConv(O_Y_NYUREC.ZAIKO_SYUSI, vbUnicode))     '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, StrConv(O_Y_NYUREC.NOUKI_YMD, vbUnicode)) '�w��[��
        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, _
                                    StrConv(O_Y_NYUREC.SERVICE_KANRI_NO, vbUnicode))    '�T�[�r�X��ЊǗ��ԍ�
        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, StrConv(O_Y_NYUREC.KI_HIN_NO, vbUnicode)) '�@��i�ڃR�[�h
        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, _
                                        StrConv(O_Y_NYUREC.ENVIRONMENT_KBN, vbUnicode)) '���K�i���i�敪
        Call UniCode_Conv(Y_NYUREC.KAN_DT, StrConv(O_Y_NYUREC.KAN_DT, vbUnicode))       '�������t
        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, _
                                        StrConv(O_Y_NYUREC.BEF_NYU_QTY, vbUnicode))     '��s���א�
        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, _
                                        StrConv(O_Y_NYUREC.YOSAN_FROM, vbUnicode))      '�\�Z�P�ʁi���j
        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(O_Y_NYUREC.YOSAN_TO, vbUnicode))   '�\�Z�P�ʁi��j
        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(O_Y_NYUREC.HTANABAN, vbUnicode))   '�W���I��
        Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(O_Y_NYUREC.HIN_NAI, vbUnicode))     '�i�ԁi�����j
        Call UniCode_Conv(Y_NYUREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "���ח\��")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(5).Caption = Format(count, "#0")

'    GoTo Convert_End

'---------------------------------------------  �o�ח\��̃R���o�[�g
Convert_P6:
    If Check1(6).Value <> 1 Then GoTo Convert_End

    MsgLab(1) = "�o�ח\��f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(6).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i���j�o�ח\��f�[�^")
                Exit Function
        End Select


        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(6).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

                                                                '�g�p�q�@ID
        Call UniCode_Conv(Y_SYUREC.WEL_ID, StrConv(O_Y_SYUREC.WEL_ID, vbUnicode))
                                                                '�g�p���v���O����
        Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(O_Y_SYUREC.PRG_ID, vbUnicode))
                                                                '�����敪
        Call UniCode_Conv(Y_SYUREC.KAN_KBN, StrConv(O_Y_SYUREC.KAN_KBN, vbUnicode))
                                                                '�f�[�^���
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(O_Y_SYUREC.DT_SYU, vbUnicode))
                                                                '���ƕ��敪
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(O_Y_SYUREC.JGYOBU, vbUnicode))
                                                                '�����敪�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(O_Y_SYUREC.KEY_CYU_KBN, vbUnicode))
                                                                '�`�[�h�c�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(O_Y_SYUREC.KEY_ID_NO, vbUnicode))
                                                                '�����O
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(O_Y_SYUREC.NAIGAI, vbUnicode))
                                                                '�i�ڔԍ��i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(O_Y_SYUREC.KEY_HIN_NO, vbUnicode))
                                                                '���Ӑ�R�[�h�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(O_Y_SYUREC.KEY_MUKE_CODE, vbUnicode))
                                                                '������R�[�h�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(O_Y_SYUREC.KEY_SS_CODE, vbUnicode))
                                                                '�o�ד��t�i�j�d�x�j
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(O_Y_SYUREC.KEY_SYUKA_YMD, vbUnicode))
                                                                '���Ə�
        Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(O_Y_SYUREC.JGYOBA, vbUnicode))
                                                                '�f�[�^�敪
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(O_Y_SYUREC.DATA_KBN, vbUnicode))
                                                                '����敪
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(O_Y_SYUREC.TORI_KBN, vbUnicode))
                                                                '�h�c��
        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(O_Y_SYUREC.ID_NO, vbUnicode))

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")           '��v�p���Ə꺰��
        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")           '���Y�Ǘ��p���Ə꺰��
'------------------------------------------------------------------------------

                                                                '�i�ڔԍ�
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(O_Y_SYUREC.HIN_NO, vbUnicode))
                                                                '�`�[�ԍ�
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(O_Y_SYUREC.DEN_NO, vbUnicode))
                                                                '�o�א���
        Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(O_Y_SYUREC.SURYO, vbUnicode))
                                                                '���Ӑ�R�[�h
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(O_Y_SYUREC.MUKE_CODE, vbUnicode))
                                                                '�݌Ɏ��x
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(O_Y_SYUREC.SYUKO_SYUSI, vbUnicode))

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")            '���Y�Ǘ��p�݌Ɏ��x����
        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")             '�⏕�݌Ɏ��x����
'------------------------------------------------------------------------------

                                                                '�o�ד��t
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(O_Y_SYUREC.SYUKA_YMD, vbUnicode))

'--- �ǉ����� ------------------------------------------------------------------
                                                                '���ےP��
        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(O_Y_SYUREC.TANKA, vbUnicode))
'------------------------------------------------------------------------------

                                                                '�I�[�_�[�ԍ�
        Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(O_Y_SYUREC.ODER_NO, vbUnicode))
                                                                '�A�C�e���ԍ�
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(O_Y_SYUREC.ITEM_NO, vbUnicode))

'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")               '�����Ǘ��ԍ�����
        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")             '���`�Ժ���
        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, "")               '�o�ח\���
        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")                '۹����1
        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")                '۹����2
        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")                '۹����3
'------------------------------------------------------------------------------

                                                                '���Ӑ於��
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(O_Y_SYUREC.MUKE_NAME, vbUnicode))
                                                                '�����敪
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(O_Y_SYUREC.CYU_KBN, vbUnicode))
                                                                '�����敪����
        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(O_Y_SYUREC.CYU_KBN_NAME, vbUnicode))

'''                                                                '�A�o�o�׌����敪
'''        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, StrConv(O_Y_SYUREC.EXPORT_KBN, vbUnicode))
'''                                                                '�����x�����s�敪
'''        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, StrConv(O_Y_SYUREC.LABEL_ISSUE_KBN, vbUnicode))
'''                                                                '�����x�����s�P�ʐ�
'''        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, StrConv(O_Y_SYUREC.LABEL_ISSUE_UNIT, vbUnicode))
'''                                                                '�����x���P���\���敪
'''        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, StrConv(O_Y_SYUREC.LABEL_TANKA_KBN, vbUnicode))
'''                                                                '�P��
'''        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(O_Y_SYUREC.TANKA, vbUnicode))
'''                                                                '���z
'''        Call UniCode_Conv(Y_SYUREC.KINGAKU, StrConv(O_Y_SYUREC.KINGAKU, vbUnicode))
'''                                                                '���l�Q
'''        Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(O_Y_SYUREC.BIKOU2, vbUnicode))
'''                                                                '���x�[�g�敪
'''        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, StrConv(O_Y_SYUREC.REBATE_KBN, vbUnicode))
'''                                                                '���[�敪
'''        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(O_Y_SYUREC.CHOHA_KBN, vbUnicode))
'''                                                                '�l���敪
'''        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, StrConv(O_Y_SYUREC.ATAISA_KBN, vbUnicode))
'''                                                                '��\�@��
'''        Call UniCode_Conv(Y_SYUREC.REP_KISHU, StrConv(O_Y_SYUREC.REP_KISHU, vbUnicode))
'''                                                                '�m�r�Ǘ��ԍ�
'''        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, StrConv(O_Y_SYUREC.NS_KANRI_NO, vbUnicode))
'''                                                                '�l�s�r���i�R�[�h
'''        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, StrConv(O_Y_SYUREC.MTS_HIN_CODE, vbUnicode))
'''                                                                '���l�P
'''        Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(O_Y_SYUREC.BIKOU1, vbUnicode))
'''                                                                '�����敪
'''        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(O_Y_SYUREC.CHOKU_KBN, vbUnicode))
'''                                                                '���x�[�g��
'''        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, StrConv(O_Y_SYUREC.REBATE_RATE, vbUnicode))
'''                                                                '�i��
'''        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(O_Y_SYUREC.HIN_NAME, vbUnicode))
'''                                                                '�ΊO���Ə�
'''        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, StrConv(O_Y_SYUREC.JGYOBA_GAI, vbUnicode))
'''                                                                '�@��R�[�h
'''        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(O_Y_SYUREC.KISHU_CODE, vbUnicode))
'''                                                                '������R�[�h
'''        Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(O_Y_SYUREC.SS_CODE, vbUnicode))


'--- �ǉ����� ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")                 '���Y��1
        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")                 '���Y��2
        Call UniCode_Conv(Y_SYUREC.BIKOU2, _
                    StrConv(O_Y_SYUREC.BIKOU2, vbUnicode))      '���l2
        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")                 '�̔��敪
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, _
                    StrConv(O_Y_SYUREC.CHOKU_KBN, vbUnicode))   '�����w���敪
        Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")              '�ƯďC���Ǘ��ԍ�
        Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")           '�݌Ɉ�������
        Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")          '�����Ǘ��ԍ�
        Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")              '�󒍎c����
        Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")              '�����敪
        Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")            '���i���[�i�݌Ɏ��x����
        Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")          '���i���[�i���Y�Ǘ����x����
        Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")           '���i���[�i�⏕���x����
        Call UniCode_Conv(Y_SYUREC.BIKOU1, _
                    StrConv(O_Y_SYUREC.BIKOU1, vbUnicode))      '���l1
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, _
                    StrConv(O_Y_SYUREC.CHOHA_KBN, vbUnicode))   '���[�敪
        Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")              '��t�i�ڔԍ�
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, _
                    StrConv(O_Y_SYUREC.HIN_NAME, vbUnicode))    '�i��
        Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")          '�i�ڔԍ��ύX�敪
        Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")         'Ӽޭ�ٌ����敪
        Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")             '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
        Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")        '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
        Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")         '�c�݌ɂ܂Ƃߕ⏕���x����
        Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")               '�w��[��
        Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")        '���޽��ЊǗ��ԍ�
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, _
                StrConv(O_Y_SYUREC.KISHU_CODE, vbUnicode))      '�@��i�ں���
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")              '����敔�i�敪
        Call UniCode_Conv(Y_SYUREC.SS_CODE, _
                    StrConv(O_Y_SYUREC.SS_CODE, vbUnicode))     '��������溰��
        Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")            '���i�����敪
'------------------------------------------------------------------------------

                                                                '�i�ԁi�����j
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(O_Y_SYUREC.HIN_NAI, vbUnicode))
                                                                '�z�X�g�I��
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(O_Y_SYUREC.HTANABAN, vbUnicode))
                                                                '�o�ɕ\������t
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, StrConv(O_Y_SYUREC.PRINT_YMD, vbUnicode))
                                                                '�������t
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(O_Y_SYUREC.KAN_YMD, vbUnicode))
                                                                '���i���t
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(O_Y_SYUREC.KENPIN_YMD, vbUnicode))
                                                                '������敪
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(O_Y_SYUREC.TOK_KBN, vbUnicode))
                                                                '���яo�ɐ�
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(O_Y_SYUREC.JITU_SURYO, vbUnicode))
                                                                '�捞�ݓ���
        Call UniCode_Conv(Y_SYUREC.INS_NOW, StrConv(O_Y_SYUREC.INS_NOW, vbUnicode))
                                                                        
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
                Case Else
                    Cnt(6).Caption = Format(count, "#0")
                    DoEvents
                    Call File_Error(sts, BtOpInsert, "�o�ח\��")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(6).Caption = Format(count, "#0")


'---------------------------------------------  �I��
Convert_End:
    
    Convert_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
Dim ans     As Integer
Dim i       As Integer

    Select Case Index

        Case 0      '�S�I��
            For i = 0 To 6
                Check1(i).Value = 1
            Next i

        Case 1      '���ްĊJ�n
            Beep
            ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                Command1(0).Enabled = False
                Command1(1).Enabled = False
                Command1(2).Enabled = False
                DoEvents

                If Convert_Proc() Then
                    Unload Me
                End If
            End If
            MsgBox "�I�����܂����B"
            Unload Me

        Case 2      '��ݾ�
            Unload Me

    End Select

End Sub

Private Sub Command2_Click()

Dim yn      As Integer


    If Not IsNumeric(Text1(0).Text) Or _
        Not IsNumeric(Text1(1).Text) Or _
        Not IsNumeric(Text1(2).Text) Then
        MsgBox "���t�װ"
        Exit Sub
    End If
    
    Text1(1).Text = Format(CInt(Text1(1).Text), "00")
    Text1(2).Text = Format(CInt(Text1(2).Text), "00")
    
    yn = MsgBox("�ړ����폜���s���܂��H", vbYesNo + vbDefaultButton2, "���ӁI�I")

    If yn = vbYes Then
        If IDO_DELETE_PROC() Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_DblClick()
    PrintForm
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

                                '�폜�ςݏo�ח\��n�o�d�m
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�폜�ςݏo�ח\��n�o�d�m
    If O_DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j���ԃ}�X�^�n�o�d�m
    If O_HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�݌Ɉړ����n�o�d�m
    If O_IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '��Ǝ��у��O�n�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j��Ǝ��у��O�n�o�d�m
    If O_P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w�}�i�q�j�n�o�d�m
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j���i���w�}�i�q�j�n�o�d�m
    If O_P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j���ח\��n�o�d�m
    If O_Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i���j�o�ח\��n�o�d�m
    If O_Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '�폜�ςݏo�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�폜�ςݏo�ח\��")
        End If
    End If
                                            '(��)�폜�ςݏo�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), K0_O_DEL_SYU, Len(K0_O_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i���j�폜�ςݏo�ח\��")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '(��)�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), K0_O_HATUBAN, Len(K0_O_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ƀf�[�^")
        End If
    End If
    
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '(��)�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, O_IDO_POS, O_IDOREC, Len(O_IDOREC), K0_O_IDO, Len(K0_O_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�݌Ɉړ���")
        End If
    End If
                                            '��Ǝ��у��O�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "��Ǝ��у��O")
        End If
    End If
                                            '(��)��Ǝ��у��O�b�k�n�r�d
    sts = BTRV(BtOpClose, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)��Ǝ��у��O")
        End If
    End If
                                            '���i���w�}�i�q�j�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�i�q�j")
        End If
    End If
                                            '(��)���i���w�}�i�q�j�b�k�n�r�d
    sts = BTRV(BtOpClose, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)���i���w�}�i�q�j")
        End If
    End If
                                            '���ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '(��)���ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�o�ח\��f�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '(��)�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(��)�o�ח\��f�[�^")
        End If
    End If


    sts = BTRV(BtOpReset, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20060515F = Nothing

    End
End Sub

Private Function IDO_DELETE_PROC() As Integer
    
Dim count           As Long
Dim DISP_INTERVAL   As Long
Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    
Dim DEL_COUNT           As Long
Dim DISP_DEL_INTERVAL   As Long
    
    IDO_DELETE_PROC = True
    
    MsgLab(1) = "�݌Ɉړ����폜�������I�I"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    DEL_COUNT = 0
    DISP_DEL_INTERVAL = 0
    Cnt(2).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If



        If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text1(0).Text & Text1(1).Text & Text1(2).Text) Then
        Else

    
            If Trim(StrConv(IDOREC.TANTO_CODE, vbUnicode)) = "" Then
                DEL_COUNT = DEL_COUNT + 1
                
                DISP_DEL_INTERVAL = DISP_DEL_INTERVAL + 1
                If DISP_DEL_INTERVAL = 100 Then
                    Label2.Caption = Format(DEL_COUNT, "#0")
                    DISP_DEL_INTERVAL = 0
                End If
                
                
                
        
        
                Do
                    sts = BTRV(BtOpDelete, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "�݌Ɉړ���")
                            Exit Function
                    End Select
                Loop
            End If
        End If
        com = BtOpGetNext

    Loop
    
    Cnt(2).Caption = Format(count, "#0")
    Label2.Caption = Format(DEL_COUNT, "#0")
    
    MsgLab(1) = ""
    Me.MousePointer = vbDefault
    IDO_DELETE_PROC = False

End Function

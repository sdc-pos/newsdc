VERSION 5.00
Begin VB.Form F1020101 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  '��°� ����޳
   Caption         =   "���o�ח\��f�[�^�捞�� '2019/12/13 ����DC ���xR8�Ή� "
   ClientHeight    =   4170
   ClientLeft      =   1920
   ClientTop       =   2280
   ClientWidth     =   8580
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
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ListBox LBox_Hin 
      Height          =   300
      Left            =   1560
      TabIndex        =   25
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   19
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "F1020101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String               'ܰ��ð��ݔԍ�

Private FileName    As String               '�e�L�X�g�t�@�C����
Private FileNo      As Integer              '�t�@�C����


Private ER_IN_FileName As String            '�װ�@�e�L�X�g�t�@�C����    2015.11.19
Private ER_IN_FileNo   As Integer           '�װ�@�t�@�C����            2015.11.19
Private ER_OUT_FileName As String           '�װ�@�e�L�X�g�t�@�C����    2015.11.19
Private ER_OUT_FileNo   As Integer          '�װ�@�t�@�C����            2015.11.19

Private TP_IN_FileName As String            '�װ�@�e�L�X�g�t�@�C����    2015.11.19
Private TP_IN_FileNo   As Integer           '�װ�@�t�@�C����            2015.11.19
Private TP_OUT_FileName As String           '�װ�@�e�L�X�g�t�@�C����    2015.11.19
Private TP_OUT_FileNo   As Integer          '�װ�@�t�@�C����            2015.11.19


Private KASO_NYUKA_SOKO      As String * 2  '���z���בq�ɔԍ�
Private KASO_SMODOSHI_SOKO   As String * 2  '���z�x���߂��q�ɔԍ�

Private Proc_F      As Integer              '�i�ԁ��݌ɗL���@����t���O
Private Last_Proc_F As Integer              '���������ް��폜�����@���s�L���t���O
                                            
Private Type YUKO_SOKO_TBL                  '�L��νđq�Ɏ�荞�݃e�[�u��
    HS_SOKO             As String * 3
    NAIGAI              As String * 1
End Type

Dim Soko_T()            As YUKO_SOKO_TBL  '�q�ɏ��

'-                                          2005.12.30
Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '�d������
    JGYOBU                  As String * 1   '���ƕ�
    NAIGAI                  As String * 1   '�����O
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean

'-                                          2005.12.30


Private HS_IN_SIJ   As String               '���Ƀf�[�^�t�@�C����
Private HS_OUT_SIJ  As String               '�o�Ƀf�[�^�t�@�C����
Private New_HS_OUT_SIJ  As String           '�Vڲ��ďo�Ƀf�[�^�t�@�C����2006.05.23


Private In_Cnt      As Integer              '�f�[�^�ǂݍ��݌���
Private Out_Cnt     As Integer              '�f�[�^�o�͌���

Private Const In_Mode% = 1                  '���׏���
Private Const Out_Mode% = 2                 '�o�׏���

                                            
Dim NormalFont As New StdFont               '����t�H���g

Private Const LMAX% = 46                    '�œ��ő�s��
Private Const MGN_L% = 1                    '���׈���J�n���ʒu�i�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j
Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j


Private Const NAI_CHANGE% = 1
Private Const GAI_CHANGE% = 2
Private Const NOT_GAI_CHANGE% = 3


Private ETC_MTS_NAI As String * 8             '���̑�������(����)
Private ETC_SS_NAI  As String * 8             '���̑�������(����)

Private ETC_MTS_GAI As String * 8             '���̑�������(�C�O)
Private ETC_SS_GAI  As String * 8             '���̑�������(�C�O)

Dim DUP_SYUKA_DATA  As String                 '�o�׃f�[�^�t���p�X


                                        
Dim MyCenter        As String

Private MENU_NO     As String * 2       '���у��O�o�͗p�ƭ���   2007.11.06

Dim Err_FLg         As Boolean          '2008.10.07

Dim TANA_SPACE      As Boolean          '2009.03.07

Dim KAMOKU_FURIKAE      As String * 2       '�ȖڐU�֗v�� 2009.06.26

'2010.07.20 ��
'Private Const GENSANKOKU_ON% = 1
'Private Const GENSANKOKU_OFF% = 0
'2010.07.20 ��


'���i���v��x�� 2011.07.07
Dim NOT_Hin_Name    As Variant          '���O�i��
Dim NOT_Hin_Name_F  As Boolean          '���O�i���L��
'���i���v��x�� 2011.07.07


Dim GOODS_F         As String * 1       '���i���L���@��̫�� 2012.12.20




Dim GENSAN_T()      As String * 1       '���Y���X�V�L�� 2016.12.28


'Private Const Last_Update_Day$ = "[F102010] 2019.03.06 11:55"
'Private Const Last_Update_Day$ = "[F102010] 2019.04.15 09:30"
Private Const Last_Update_Day$ = "[F102010] 2019.12.13 17:00 �G�A�R�� ���xR8�捞�Ή�"






Private Function Nyuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   �u���ח\��f�[�^�v�X�V����
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Skip_Flg    As Boolean
    
Dim WK_Y_QTY    As Long     '�o�א����[�N
Dim WK_Qty      As Long     '�O�؎c���[�N
Dim WK_E_QTY    As Long     '��s�o�א����[�N

Dim SUMI_QTY    As Long     '���i���ς݂Ƃ��ēo�^
Dim MI_QTY      As Long     '�����i�Ƃ��ēo�^



Dim WORK_SOKO   As String * 2
    
Dim sts         As Integer
Dim ans         As Integer
Dim Not_SHUSI   As Boolean
    
Dim wkText      As String
Dim Length      As Integer
    
    
Dim NAIGAI      As String * 1   '2007.06.15
    
    
Dim TEXT_NO     As String * 9           '÷�ć�
Dim JGYOBU_Code As String * 1           '���ƕ��敪
Dim CYOK_KBN    As String * 1           '�����敪
Dim DEN_DT      As String * 8           '�`�[���t
Dim IO_KBN      As String * 1           '���o�ɋ敪
Dim PM_KBN      As String * 1           '�ԍ��敪
Dim DEN_SYU     As String * 1           '�`�[���
Dim DEN_NO      As String * 6           '�`�[��
Dim CYU_KBN     As String * 1           '�����敪
'Dim HIN_GAI     As String * 13          '�i�ԁi�O���j  '2016.04.26
'Dim HIN_NAI     As String * 13          '�i�ԁi�����j  '2016.04.26
Dim HIN_GAI     As String * 20          '�i�ԁi�O���j   '2016.04.26
Dim HIN_NAI     As String * 20          '�i�ԁi�����j   '2016.04.26
Dim HIN_NAME    As String * 25          '�i��
Dim YOTEI_QTY   As String * 6           '����
Dim YOSAN_FROM  As String * 5           '�\�Z�P�ʁi���j
Dim YOSAN_TO    As String * 5           '�\�Z�P�ʁi��j
Dim HOST_SOKO   As String * 8           '�q�ɋ敪�iνāj
Dim HOST_TANA   As String * 8           '�I�ԁiνāj
Dim SYUK_CODE   As String * 5           '�x����^�o�א�
Dim SYUK_NAME   As String * 20          '�x����^�o�א於
Dim REC_END     As String * 1           'ں��ޏI�[ϰ�(@)
    
    
    
Dim GENSANKOKU          As String * 20  '���Y����
Dim GEN_GENSANKOKU      As String * 20  '�����\�����Y����
Dim SHIIRE_WORK_CENTER  As String * 8   '���ގd����ܰ�����
Dim KANKYO_KBN          As String * 3   '����ދ敪
Dim KANKYO_KBN_ST       As String * 8   '����ދ敪�K�p�J�n
Dim KANKYO_KBN_SURYO    As String * 10  '����ދ敪����
Dim ID_NO2              As String * 12  'ID_NO
Dim AITESAKI_CODE       As String * 16  '����溰��
Dim JYUCHU_YMD          As String * 8   '�󒍔N����
Dim SHITEI_NOUKI_YMD    As String * 8   '�w��[���N����
    
    
    
    
Dim GENSANKOKU_CHG_F    As Boolean      '���Y���ύXF
    
    
'�o�ח\�� �ҏW�O���� ################################################################# 2005/05/16 Add ��
Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
Dim INS_NOW         As String * 14
Dim wkStr           As String
    
Dim wkMUKE_CODE     As String
    
'2010.11.01
Dim DUP_FLG         As Boolean

'2011.01.19
Dim Loop_Cnt        As Integer


'2011.03.23
Dim MOTO_TEXT_NO    As String * 9

    
Dim Rec_LENG        As Long         '2016.04.19
    
    
Dim MAEGARI_FLG     As Boolean      '2018.11.16
    
    
    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'#################################################################################### 2005/05/16 Add ��
    
    Nyuka_Update_Proc = True


    Do Until EOF(FileNo)
        Line Input #FileNo, wkText
    
    
    
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And LenB(StrConv(wkText, vbFromUnicode)) <> 251 Then    '2016.04.26
        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And LenB(StrConv(wkText, vbFromUnicode)) <> 265 Then     '2016.04.26
            
'            Call NG_File_Make_Proc
             Err_FLg = True
           
    Call LOG_OUT(LOG_F, wkText)
            Exit Do
        End If
    
    
        Rec_LENG = LenB(StrConv(wkText, vbFromUnicode)) '2016.04.19
    
    
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
    
        MAEGARI_FLG = False     '2018.11.16
    
                                                                    '÷�ć�
        Length = 1
        TEXT_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TEXT_NO)), vbUnicode)
                                                                    '���ƕ��敪
        Length = Length + Len(TEXT_NO)
        JGYOBU_Code = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JGYOBU_Code)), vbUnicode)
                                                                    '�����敪
        Length = Length + Len(JGYOBU_Code)
        CYOK_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYOK_KBN)), vbUnicode)
                                                                    '�`�[���t
        Length = Length + Len(CYOK_KBN)
        DEN_DT = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_DT)), vbUnicode)
                                                                    '���o�ɋ敪
        Length = Length + Len(DEN_DT)
        IO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(IO_KBN)), vbUnicode)
                                                                    '�ԍ��敪
        Length = Length + Len(IO_KBN)
        PM_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(PM_KBN)), vbUnicode)
                                                                    '�`�[���
        Length = Length + Len(PM_KBN)
        DEN_SYU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_SYU)), vbUnicode)
                                                                    '�`�[��
        Length = Length + Len(DEN_SYU)
        DEN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_NO)), vbUnicode)
                                                                    '�����敪
        Length = Length + Len(DEN_NO)
        CYU_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYU_KBN)), vbUnicode)
                                                                    '�i�ԁi�O���j
        Length = Length + Len(CYU_KBN)
        HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI)), vbUnicode)
                                                                    '�i�ԁi�����j
        Length = Length + Len(HIN_GAI)
        HIN_NAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAI)), vbUnicode)
                                                                    '�i��
        Length = Length + Len(HIN_NAI)
        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)
                                                                    '����
        Length = Length + Len(HIN_NAME)
        YOTEI_QTY = Trim(StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOTEI_QTY)), vbUnicode))
                                                                    '�\�Z�P�ʁi���j
        Length = Length + Len(YOTEI_QTY)
        YOSAN_FROM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_FROM)), vbUnicode)
                                                                    '�\�Z�P�ʁi��j
        Length = Length + Len(YOSAN_FROM)
        YOSAN_TO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_TO)), vbUnicode)
                                                                    '�q�ɋ敪�iνāj
        Length = Length + Len(YOSAN_TO)
        HOST_SOKO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_SOKO)), vbUnicode)
                                                                    '�I�ԁiνāj
        Length = Length + Len(HOST_SOKO)
        HOST_TANA = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_TANA)), vbUnicode)
                                                                    '�x����^�o�א�
        Length = Length + Len(HOST_TANA)
        SYUK_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_CODE)), vbUnicode)
                                                                    '�x����^�o�א於
        Length = Length + Len(SYUK_CODE)
        SYUK_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_NAME)), vbUnicode)
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) = 251 Then     '2016.04.26
        If LenB(StrConv(wkText, vbFromUnicode)) = 265 Then      '2016.04.26
                                                                    '���Y��
            Length = Length + Len(SYUK_NAME)
            GENSANKOKU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(GENSANKOKU)), vbUnicode)
                                                                    '�����\�����Y����
            Length = Length + Len(GENSANKOKU)
            GEN_GENSANKOKU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(GEN_GENSANKOKU)), vbUnicode)
                                                                    '���ގd����ܰ�����
            Length = Length + Len(GEN_GENSANKOKU)
            SHIIRE_WORK_CENTER = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SHIIRE_WORK_CENTER)), vbUnicode)
                                                                    '����ދ敪
            Length = Length + Len(SHIIRE_WORK_CENTER)
            KANKYO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN)), vbUnicode)
                                                                    '����ދ敪�K�p�J�n
            Length = Length + Len(KANKYO_KBN)
            KANKYO_KBN_ST = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN_ST)), vbUnicode)
                                                                    '����ދ敪����
            Length = Length + Len(KANKYO_KBN_ST)
            KANKYO_KBN_SURYO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN_SURYO)), vbUnicode)
                                                                    'ID_NO
            Length = Length + Len(KANKYO_KBN_SURYO)
            ID_NO2 = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(ID_NO2)), vbUnicode)
                                                                    '�����
            Length = Length + Len(ID_NO2)
            AITESAKI_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(AITESAKI_CODE)), vbUnicode)
                                                                    '�󒍔N����
            Length = Length + Len(AITESAKI_CODE)
            JYUCHU_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JYUCHU_YMD)), vbUnicode)
                                                                    '�w��[���N����
            Length = Length + Len(JYUCHU_YMD)
            SHITEI_NOUKI_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SHITEI_NOUKI_YMD)), vbUnicode)
        
        
        
        Else
            
            GENSANKOKU = ""             '���Y����
            GEN_GENSANKOKU = ""         '�����\�����Y����
            SHIIRE_WORK_CENTER = ""     '���ގd����ܰ�����
            KANKYO_KBN = ""             '����ދ敪
            KANKYO_KBN_ST = ""          '����ދ敪�K�p�J�n
            KANKYO_KBN_SURYO = ""       '����ދ敪����
            ID_NO2 = ""                 'ID_NO
            AITESAKI_CODE = ""          '����溰��
            JYUCHU_YMD = ""             '�󒍔N����
            SHITEI_NOUKI_YMD = ""       '�w��[���N����
        End If
'        Length = 1
'        TEXT_NO = Mid(wkText, Length, Len(TEXT_NO))                 '÷�ć�
        
'        Length = Length + Len(TEXT_NO)
'        JGYOBU_Code = Mid(wkText, Length, Len(JGYOBU_Code))         '���ƕ��敪
    
'        Length = Length + Len(JGYOBU_Code)
'        CYOK_KBN = Mid(wkText, Length, Len(CYOK_KBN))               '�����敪
    
'        Length = Length + Len(CYOK_KBN)
'        DEN_DT = Mid(wkText, Length, Len(DEN_DT))                   '�`�[���t
    
'        Length = Length + Len(DEN_DT)
'        IO_KBN = Mid(wkText, Length, Len(IO_KBN))                   '���o�ɋ敪
    
'        Length = Length + Len(IO_KBN)
'        PM_KBN = Mid(wkText, Length, Len(PM_KBN))                   '�ԍ��敪
    
'        Length = Length + Len(PM_KBN)
'        DEN_SYU = Mid(wkText, Length, Len(DEN_SYU))                 '�`�[���
    
'        Length = Length + Len(DEN_SYU)
'        DEN_NO = Mid(wkText, Length, Len(DEN_NO))                   '�`�[��
    
'        Length = Length + Len(DEN_NO)
'        CYU_KBN = Mid(wkText, Length, Len(CYU_KBN))                 '�����敪
    
'        Length = Length + Len(CYU_KBN)
'        HIN_GAI = Mid(wkText, Length, Len(HIN_GAI))                 '�i�ԁi�O���j
    
'        Length = Length + Len(HIN_GAI)
'        HIN_NAI = Mid(wkText, Length, Len(HIN_NAI))                 '�i�ԁi�����j
    
'        Length = Length + Len(HIN_NAI)
'        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)             '�i��
    
'        Length = Length + Len(HIN_NAME)
'        YOTEI_QTY = Trim(Mid(wkText, Length, Len(YOTEI_QTY)))       '����
    
'        Length = Length + Len(YOTEI_QTY)
'        YOSAN_FROM = Mid(wkText, Length, Len(YOSAN_FROM))           '�\�Z�P�ʁi���j
    
'        Length = Length + Len(YOSAN_FROM)
'        YOSAN_TO = Mid(wkText, Length, Len(YOSAN_TO))               '�\�Z�P�ʁi��j
    
'        Length = Length + Len(YOSAN_TO)
'        HOST_SOKO = Mid(wkText, Length, Len(HOST_SOKO))             '�q�ɋ敪�iνāj
    
'        Length = Length + Len(HOST_SOKO)
'        HOST_TANA = Mid(wkText, Length, Len(HOST_TANA))             '�I�ԁiνāj
        
'        Length = Length + Len(HOST_TANA)
'        SYUK_CODE = Mid(wkText, Length, Len(SYUK_CODE))             '�x����^�o�א�
        
'        Length = Length + Len(SYUK_CODE)
'        SYUK_NAME = Mid(wkText, Length, Len(SYUK_NAME))             '�x����^�o�א於
    
    
    
    
    
    
        Skip_Flg = True
        Not_SHUSI = False
        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(HOST_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        If Skip_Flg Then
            Not_SHUSI = True
        End If
    
    
    
    
    
    
    
    
'-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
        '�ƍ��p���ח\��d���`�F�b�N
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
        
'2010.11.01
        DUP_FLG = False
'2010.11.01
        
        
        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
'2010.11.01
                DUP_FLG = True
'2010.11.01
            
            Case BtErrKeyNotFound
            Case Else
                'Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��", 0)                '2016.06.23
                Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��", 1, Y_GLICS_ID)     '2016.06.23
'                Exit Function      '2015.11.19
                GoTo Abort_Tran     '2015.11.19
        End Select



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' �������ɂ��f�[�^�m�F  2011.03.23
        MOTO_TEXT_NO = ""
        
        If DUP_FLG Then
            If StrConv(App.EXEName, vbUpperCase) = Trim(StrConv(Y_GLICSREC.MOTO_PROG_ID, vbUnicode)) Then
            Else
                If Trim(TEXT_NO) = Trim(StrConv(Y_GLICSREC.MOTO_TEXT_NO, vbUnicode)) Then
                Else
                    MOTO_TEXT_NO = TEXT_NO
                    Mid(TEXT_NO, 5, 1) = "A"
                    DUP_FLG = False
                End If
            
            End If
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' �������ɂ��f�[�^�m�F  2011.03.23

















        If Not_SHUSI Then
            NAIGAI = "1"
        Else
            NAIGAI = Soko_T(i, j).NAIGAI
        End If

'        If sts = BtErrKeyNotFound Then
        If Not DUP_FLG Then

'            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME) Then
                
                
                
    '2010.07.20 ��
            If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                If Item_Check_Proc(In_Mode, JGYOBU, "1", HIN_GAI, HIN_NAI, HIN_NAME, GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO) Then
'                    GoTo Abort_Tran
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
                End If
            End If
    '2010.07.20 ��
                

                
                
'''''''''''''''''''''2011.03.23 �����ǉ�
'            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME, _
'                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD) Then
            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
                                TEXT_NO, _
                                JGYOBU_Code, _
                                CYOK_KBN, _
                                DEN_DT, _
                                IO_KBN, _
                                PM_KBN, _
                                DEN_SYU, _
                                DEN_NO, _
                                CYU_KBN, _
                                HIN_GAI, _
                                HIN_NAI, _
                                HIN_NAME, _
                                YOTEI_QTY, _
                                YOSAN_FROM, _
                                YOSAN_TO, _
                                HOST_SOKO, _
                                HOST_TANA, _
                                SYUK_CODE, _
                                SYUK_NAME, _
                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD, MOTO_TEXT_NO) Then
'''''''''''''''''''''2011.03.23 �����ǉ�
'               Exit Function      '2015.11.19
                GoTo Abort_Tran     '2015.11.19
            End If

        End If



'-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
    
    
    
    
    
    
    
    
    
        If IO_KBN <> "1" Then
            
            If IO_KBN = "4" And Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" And Trim(HOST_SOKO) = "11B" Then
            Else
                Skip_Flg = True
            End If
        End If
    
    
        If PM_KBN = "-" Then
            Skip_Flg = True
        End If
    
        'NOPOS  2006.05.01
        If Trim(DEN_NO) = "NOPOS" Then
            Skip_Flg = True
        End If
    
        '�\�Z����36003���O  2006.07.15
        If Trim(YOSAN_FROM) = "36003" Then
            Skip_Flg = True
        End If
    
        '�\�Z����PP���O  2008.01.10
        If Left(YOSAN_FROM, 2) = "PP" Then
            
            If Trim(YOSAN_FROM) = "PPP4" And JGYOBU = SHOKUSEN Then     '2017.02.17
            Else                                                        '2017.02.17
                Skip_Flg = True
            End If                                                      '2017.02.17
        End If
    
    
    
    
        WORK_SOKO = KASO_NYUKA_SOKO
    
    
    
        Select Case JGYOBU
            
''            Case SENTAKU                        '����@
''
''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
''                    Skip_Flg = True
''                End If
''
''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
''                    Skip_Flg = True
''                End If
                            
            
            
            
            Case SOJIKI                         '�|���@
                
            
                If Left(YOSAN_FROM, 2) = "KM" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "KK" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "GG" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "SS" Then
                    Skip_Flg = True
                End If

                '2005.04.07 ���x�ǉ�
                If Left(YOSAN_FROM, 5) = "0090K" Then
                    Skip_Flg = True
                End If
                '2006.07.27 ���x�ǉ�
                If Left(YOSAN_FROM, 5) = "0092H" Then
                    Skip_Flg = True
                End If
                '2006.07.27 ���x�ǉ�
                If Left(YOSAN_FROM, 2) = "AA" Then
                    Skip_Flg = True
                End If
            
                '2009.08.28 ���x�ǉ�
                If Left(YOSAN_FROM, 2) = "ZZ" Then
                    Skip_Flg = True
                End If
            
            
            
            
                If Trim(YOSAN_FROM) <> "91H" Then
                    WORK_SOKO = KASO_SMODOSHI_SOKO
                End If
            
            
            
            Case DENKA, SUIHAN, SENTAKU, BLBU        '�d���A���сA����@�i�A�C�����j    BLBU�ǉ� 2012.04.06
            
            
                Select Case MyCenter
                    
                    Case "O"
                
                        If Left(YOSAN_FROM, 2) = "01" Then
                            Skip_Flg = True
                        End If
                    
                        If Left(YOSAN_FROM, 3) = "H33" Then    '2004.07.16
                            Skip_Flg = True
                        End If
                        If Left(YOSAN_FROM, 3) = "H22" Then    '2004.07.16
                            Skip_Flg = True
                        End If
        
                        If Left(YOSAN_FROM, 2) = "05" Then
                            Skip_Flg = True
                        End If
        
                        '2006.08.17
                        If Left(YOSAN_FROM, 2) = "08" Then
                            Skip_Flg = True
                        End If
                        
                        '2006.10.13 �d�������͗\�Z��="02"�̂ݑΏ�
                        If JGYOBU = DENKA Then
                            
                            '2008.01.07 "02"-->"0201"�ɕύX 2008.01.08 Left(Trim(YOSAN_FROM), 2) <> "02" �ɕύX
                            If Left(Trim(YOSAN_FROM), 2) <> "02" And _
                            Trim(YOSAN_FROM) <> "G11" And _
                            Trim(YOSAN_FROM) <> "G22" And _
                            Trim(YOSAN_FROM) <> "KA01" Then         '2012.08.31 KA01 �ǉ�
                                Skip_Flg = True
                            End If
                        End If
        
        
        
        
                        '2012.08.31
                        
        
        
        
                        '2006.11.22 ����/�A�C�����̏��O�����ǉ� 2012.04.06 BLBU�ǉ�
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If (Left(YOSAN_FROM, 2) = "P3" Or _
                                Left(YOSAN_FROM, 2) = "S3") Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        '2007.10.25 ����/�A�C�����̏��O�����ǉ� 2012.04.06 BLBU�ǉ�
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If Left(YOSAN_FROM, 2) = "RO" Then
                                Skip_Flg = True
                            End If
                        End If
        
                        '2007.12.06 ����/�A�C�����̏��O�����ǉ�  2012.04.06 BLBU�ǉ�
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If Left(YOSAN_FROM, 2) = "07" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2008.06.26 ���т̏��O�����ǉ� 2012.04.06 BLBU�ǉ�
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "04" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2008.10.14 ���т̏��O�����ǉ� 2012.04.06 BLBU�ǉ�
        
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "NC" Or Left(YOSAN_FROM, 2) = "99" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        '2016.06.17 ����,BLBU�̏��O�����ǉ�
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "RX" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2017.07.22 ����,BLBU�̏��O�����ǉ�
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "RZ" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        If Left(YOSAN_FROM, 3) = "G22" Then
                            WORK_SOKO = "80"
                        End If
        
                        If Left(YOSAN_FROM, 3) = "G11" Then
                            WORK_SOKO = "81"
                        End If
                    
                        '2006.04.29�p
                        If Left(YOSAN_FROM, 2) = "S1" And _
                            Left(YOSAN_TO, 2) = "S3" Then
                            WORK_SOKO = "87"
                        End If
                        '2006.05.01
                        If Trim(DEN_NO) = "POS87" Then
                            WORK_SOKO = "87"
                        End If
                
                                
                
                
                
                
                
                
                        '2008.06.26 ���т̑q�ɔԍ��̐ݒ�ǉ�  2012.04.06 BLBU�ǉ�
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "02" And Left(YOSAN_TO, 3) = "SDC" Then
                                WORK_SOKO = "90"
                            End If
                        End If
                
                
                
                
                        '2009.06.01 65�ԑq�ɏo�͒ǉ� 2012.04.06 BLBU�ǉ�
                        If (JGYOBU = SUIHAN Or JGYOBU = DENKA Or JGYOBU = BLBU) Then
                            If IO_KBN = "4" Then
                                If Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" Then
                                    
                                    If Trim(HOST_SOKO) = "11B" Then
                                        WORK_SOKO = "65"
                                    End If
                                End If
                            End If
                        End If
                        
                        
                
                
                
                
                
                
                
                
                    Case "F"
            
            
            
                        If Left(YOSAN_FROM, 2) = "P2" Then
                            Skip_Flg = True
                        End If

                        If Left(YOSAN_FROM, 2) = "U2" Then      '2008.01.11
                            Skip_Flg = True
                        End If


                        If Left(YOSAN_FROM, 3) <> "904" Then
                            If Left(YOSAN_FROM, 1) = "9" Then
                              Skip_Flg = True
                            End If
                        End If
                        
                        
                        
                        '�\�Z����PP�܈�̂�  2009.11.10
                        If Left(YOSAN_FROM, 2) = "PP" Then
                            
                            
                            If Not Not_SHUSI Then
                            
                                Skip_Flg = False
                            End If
                        End If
                        
                        If Left(YOSAN_FROM, 2) = "S1" And _
                            Left(YOSAN_TO, 2) = "S2" Then
                            WORK_SOKO = "88"
                        End If
            
                        '2006.05.01
                        If Trim(DEN_NO) = "POS88" Then
                            WORK_SOKO = "88"
                        End If
            
                End Select
             Case AIRCON                     '�G�A�R��
                '���O�q�ɂɁuCA�v��ǉ� 2006.07.27
                If Trim(HOST_SOKO) = "J4" Or _
                   Trim(HOST_SOKO) = "JG" Or _
                    Trim(HOST_SOKO) = "JW" Or _
                    Trim(HOST_SOKO) = "JV" Or _
                    Trim(HOST_SOKO) = "HY" Or _
                    Trim(HOST_SOKO) = "CA" Then
                    Skip_Flg = True
                End If
        
        
                If Left(YOSAN_FROM, 2) = "SH" Then
                    Skip_Flg = True
                End If
        
        
                If Left(YOSAN_FROM, 2) = "S1" Then
                    If Trim(HOST_SOKO) = "OS" Then
                      Skip_Flg = True
                    End If
                End If
        
                If Not Skip_Flg Then
                    'S2��ǉ� 2009.11.04
                    'SS��ǉ� 2010.03.08

'                    If Trim(HOST_SOKO) = "S8" Then
                    If Trim(HOST_SOKO) = "S8" Or Trim(HOST_SOKO) = "S2" Or Trim(HOST_SOKO) = "SS" Then
                        
                        
                        WORK_SOKO = "80"
                    Else
                        If CYU_KBN = "A" Then
                        Else
                            If CYU_KBN = "D" Then
                                WORK_SOKO = "70"
                            Else
                            End If
                        End If
                    End If
                End If
           
        
        
            Case OVEN           '�d�q�����W 2012.09.28
                
        
    '6�@1    ��     SDC    ��     90
    '6�@1    001    SDC    ��     70���ǉ� �����ɁE���j�b�g
    '6  1    0102   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0201   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0601   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0602   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0701   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0702   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0801   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0802   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    0899   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9101   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9102   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9301   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9601   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9602   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9901   SDC    ��     - ���ǉ� ���x�U�֕��͏��O
    '6  1    9902   SDC    ��     - ���ǉ� ���x�U�֕��͏��O


        
                WORK_SOKO = "90"
                Select Case Trim(YOSAN_FROM)
                
                    Case "001"
                        WORK_SOKO = "70"
                

                        MAEGARI_FLG = True      '2018.11.16

                
                
                
                
                    Case "WP555"                    '2017.05.16
                        WORK_SOKO = "WP"            '2017.05.16
                    
                    
                        MAEGARI_FLG = True      '2018.11.16
                    
                    
                    
                    Case "0102"
                        Skip_Flg = True
                    Case "0201"
                        Skip_Flg = True
                    Case "0601"
                        Skip_Flg = True
                    Case "0602"
                        Skip_Flg = True
                    Case "0701"
                        If Trim(HOST_SOKO) = "01" Then          '2019.01.11
                        Else                                    '2019.01.11
                            Skip_Flg = True
                        End If                                  '2019.01.11
                    Case "0702"
                        Skip_Flg = True
                    Case "0801"
                        Skip_Flg = True
                    Case "0802"
                        Skip_Flg = True
                    Case "0899"
                        Skip_Flg = True
                    Case "9101"
                        Skip_Flg = True
                    Case "9102"
                        Skip_Flg = True
                    Case "9301"
                        Skip_Flg = True
                    Case "9601"
                        Skip_Flg = True
                    Case "9602"
                        Skip_Flg = True
                    Case "9901"
                        Skip_Flg = True
                    Case "9902"
                        Skip_Flg = True
                
                
                
                    Case "ZA071"                                    '2018.12.07
                        
                        
                        If (Trim(HOST_SOKO) = "01" Or Trim(HOST_SOKO) = "02" Or Trim(HOST_SOKO) = "99") Then    '2018.12.11
                        Else                                                                                    '2018.12.11
                            If Trim(HOST_SOKO) <> "06" Then     '2018.12.07
                                Skip_Flg = True                 '2018.12.07
                            End If                              '2018.12.07
                        End If                                                                                  '2018.12.11
                
                End Select
        
        
                If Trim(HOST_SOKO) = "06" And Trim(YOSAN_FROM) = "ZA071" Then     '2018.12.07
                Else                                                        '2018.12.70
        
        
                    If Trim(YOSAN_TO) <> "SDC" Then
                        Skip_Flg = True
                    End If
        
                End If                                                      '2018.12.7
        
        
                If Trim(HOST_SOKO) = "06" Then                  '2018.12.12
                    If (Trim(YOSAN_FROM) <> "ZA071" And Trim(YOSAN_FROM) <> "WP555") Then           '2018.12.12,2019.02.07
                        Skip_Flg = True                         '2018.12.12
                    End If                                      '2018.12.12
                End If                                          '2018.12.12
        
        
                '>> 2019.03.06
                If Trim(YOSAN_FROM) = "WP555" Then
                    If Trim(HOST_SOKO) = "01" Or Trim(HOST_SOKO) = "02" Or Trim(HOST_SOKO) = "06" Or Trim(HOST_SOKO) = "93" Or Trim(HOST_SOKO) = "99" Then
                        If Trim(YOSAN_TO) = "SDC" Then
                        
Debug.Print
                        
                        Else
                            Skip_Flg = True
                        End If
                    Else
                        Skip_Flg = True
                    End If
                End If
                '>> 2019.03.06
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>2015.10.16     �H���ǉ�
            Case SHOKUSEN
                                                                                                                '"903" �ǉ� 2015.10.21
'                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" Then
                                                                                                                '"PPP4" �ǉ� 2017.02.17 "906"�@�ǉ��@2019.04.15
'                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" And Trim(YOSAN_FROM) <> "PPP4" Then
                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" And Trim(YOSAN_FROM) <> "PPP4" _
                    And Trim(YOSAN_FROM) <> "906" Then
                    Skip_Flg = True
                End If
                
                If Trim(YOSAN_FROM) = "904" Then
'                   If Trim(HOST_SOKO) <> "S4" Then                                '2017.08.04
                    If Trim(HOST_SOKO) <> "S4" And Trim(HOST_SOKO) <> "P4" Then     '2017.08.04
                        Skip_Flg = True
                    End If
                End If
                
                If Trim(YOSAN_FROM) = "903" Then            '2015.10.21
                    If Trim(HOST_SOKO) <> "P4" Then         '2015.10.21
                        Skip_Flg = True                     '2015.10.21
                    End If                                  '2015.10.21
                End If                                      '2015.10.21
                
                
                If Trim(YOSAN_FROM) = "906" Then            '2019.04.15
                    If Trim(HOST_SOKO) <> "P4" Then         '2019.04.15
                        Skip_Flg = True                     '2019.04.15
                    End If                                  '2019.04.15
                End If                                      '2019.04.15
                
                
                
                
                
                If Trim(YOSAN_FROM) = "PPP4" And Trim(HOST_SOKO) = "P4" Then            '2017.02.17
                    WORK_SOKO = "81"                                                    '2017.02.17
                End If                                                                  '2017.02.17
                
                
                If Trim(YOSAN_TO) <> "SDC" Then
                    Skip_Flg = True
                End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>2015.10.16     �H���ǉ�
                
        
        End Select
            
            
            
            
        
    
    
        If Not Skip_Flg Then
                                        
                                        
            
                
                                        '���ח\��d���`�F�b�N
            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
    
            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
                    Skip_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    'Call File_Error(sts, BtOpGetEqual, "���ח\��", 0)              '2016.06.23
                    Call File_Error(sts, BtOpGetEqual, "���ח\��", 1, Y_NYU_ID)     '2016.06.23
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
            End Select
        
        
        
        
        
            If Not Skip_Flg Then
                                                '�g�����U�N�V�����J�n
                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
                End If
                                            '�i�ڃ}�X�^�`�F�b�N
'                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
'                    GoTo Abort_Tran
'                End If
                                            
                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME, , , , KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO) Then
                    GoTo Abort_Tran
                End If
                                            
                                            
                '2012.12.20
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                End If
                '2012.12.20
                                            
                                            '���׃f�[�^�쐬
                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
        
        
                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
        
        
                Last_Proc_F = True              '���������ް��폜�����@���s�L��
        
        
                '���������ް��X�V
                
                
                If MAEGARI_FLG Then                                         '2018.11.16
                
                    
                    If WORK_SOKO = "70" And JGYOBU = OVEN Then
                    
                        'Call LOG_OUT(LOG_F, "HIN_GAI=" & HIN_GAI & " YOTEI_QTY=" & YOTEI_QTY)
                        If MAEGARI_PROC(JGYOBU_Code, HIN_GAI, YOTEI_QTY) Then   '2018.11.16
                            Unload Me                                           '2018.11.16
                        End If                                                  '2018.11.16
                
                    End If
                
                
                    WK_E_QTY = 0                                            '2018.11.16
                    
                Else                                                        '2018.11.16
                
                
                
                
                
                    If JGYOBU = OVEN And Trim(YOSAN_FROM) <> "4HHK" Then        '2019.02.08
                        WK_E_QTY = 0                                            '2019.02.08
                    Else                                                        '2019.02.08
                
                
                        Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
                        Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                        Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
            
                        WK_Y_QTY = CLng(YOTEI_QTY)
            
            
                        Loop_Cnt = 0
            
                        Do
                            sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                        WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                                
                                        Loop_Cnt = 0
                                        Do
                                        
                                            sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                                            Beep
        '                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
        '                                            If ans = vbCancel Then
        ''                                                Exit Function
        '                                                GoTo Abort_Tran
        '                                            End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                                
                                                Case Else
                                                    'Call File_Error(sts, BtOpUpdate, "���������ް�", 0)            '2016.06.23
                                                    Call File_Error(sts, BtOpUpdate, "���������ް�", 1, J_NYU_ID)   '2016.06.23
        '                                            Exit Function
                                                    GoTo Abort_Tran
                                            End Select
                                        
                                        Loop
                                        WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                    Else
                                        
                                        
                                        Loop_Cnt = 0
                                                                        
                                        Do
                                            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                                            Beep
        '                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
        '                                            If ans = vbCancel Then
        ''                                                Exit Function
        '                                                GoTo Abort_Tran
        '                                            End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                                
                                                
                                                Case Else
                                                    'Call File_Error(sts, BtOpDelete, "���������ް�", 0)            '2016.06.23
                                                    Call File_Error(sts, BtOpDelete, "���������ް�", 1, J_NYU_ID)   '2016.06.23
        '                                            Exit Function
                                                    GoTo Abort_Tran
                                            End Select
                                        Loop
                                        WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                                    End If
                            
                                    Exit Do
                                Case BtErrKeyNotFound
                                    WK_E_QTY = 0
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                            Beep
        '                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
        '                            If ans = vbCancel Then
        ''                                Exit Function
        '                                GoTo Abort_Tran
        '                           End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                
                                
                                Case Else
                                    'Call File_Error(sts, BtOpGetEqual, "���������ް�", 0)              '2016.06.23
                                    Call File_Error(sts, BtOpGetEqual, "���������ް�", 1, J_NYU_ID)     '2016.06.23
        '                            Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
                    End If                                          '2019.02.08
                End If
                                    
                                    
                                    '��s���א��i���׎��ѐ��j
                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
        
                                    '�\�Z�P�ʌ�
                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                    '�\�Z�P�ʐ�
                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                    '�W���I��
                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                                    'H�q�� 2006.10.17
                Call UniCode_Conv(Y_NYUREC.H_SOKO, HOST_SOKO)

                                    '���׃��X�g�o�̓t���O   2007.06.12
                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")








                '----------------   2010.07.08 ��
                
                
                If Trim(GENSANKOKU) = "" And Trim(GEN_GENSANKOKU) = "" And Trim(SHIIRE_WORK_CENTER) = "" Then
                
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))                    '���Y����
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))            '�����\�����Y����
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))    '���ގd����ܰ�����
                
                Else
                
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)                    '���Y����
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '�����\�����Y����
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '���ގd����ܰ�����
                End If
                
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                          '����ދ敪
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)                    '����ދ敪�K�p�J�n
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)              '����ދ敪����
                Call UniCode_Conv(Y_NYUREC.ID_NO2, ID_NO2)                                  'ID_NO
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)                    '����溰��
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                          '�󒍔N����
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)              '�w��[���N����
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "0")                             '���Ɋ֘Aؽďo��F
                    
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "8")                           '���ɊǗ�ؽďo��F
                If StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode) <> "" And Mid(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), 1, 1) > " " Then
                    
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                        
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                            Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                            Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
                            sts = BTRV(BtOpGetGreaterEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        
                            Select Case sts
                                Case BtNoErr
                                
                                    If Trim(HIN_GAI) = Trim(StrConv(GENSANREC.HIN_GAI, vbUnicode)) Then
                                        Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "0")                   '���ɊǗ�ؽďo��F
                                    End If
                                
                                
                                
                                Case BtErrEOF
                        
                                Case Else
                                  
                                    'Call File_Error(sts, BtOpGetGreaterEqual, "���Y��Ͻ�", 0)              '2016.06.23
                                    Call File_Error(sts, BtOpGetGreaterEqual, "���Y��Ͻ�", 1, GENSAN_ID)    '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                    End Select
                End If

                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "0")                       '��������ؽďo��F
                
                
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, (WORK_SOKO & _
                                                            "01" & "01" & "01"))        '���ɒI��
                                                                                        '�O�ؑ��E��
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, Format(WK_E_QTY, "00000000"))
                
                
                
                
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                
                
                
                '2011.03.23 �������v���O����
                Call UniCode_Conv(Y_NYUREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
                '2011.03.23 ���e�L�X�g��
                If Trim(MOTO_TEXT_NO) = "" Then
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, "")
                Else
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
                End If
                
                '----------------   2010.07.08 ��








                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                
                
                Loop_Cnt = 0
                            
                Do
                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                            If ans = vbCancel Then
''                                Exit Function
'                                GoTo Abort_Tran
'                            End If
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                GoTo Abort_Tran
                            End If
                        
                            DoEvents
                            Sleep (500)
                        
                        Case Else
                            'Call File_Error(sts, BtOpInsert, "���ח\��", 0)            '2016.06.23
                            Call File_Error(sts, BtOpInsert, "���ח\��", 1, Y_NYU_ID)   '2016.06.23
'                            Exit Function
                            GoTo Abort_Tran
                    End Select
                Loop
            
            
                '----------------   2010.07.08 ��
                '���Y���̃`�F�b�N���o�^
                If StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode) <> "" And Mid(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), 1, 1) > " " Then
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.28
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "2010")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                            
                            Loop_Cnt = 0
                            
                            Do
                                sts = BTRV(BtOpUpdate, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "���Y��Ͻ�", 0)               '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "���Y��Ͻ�", 1, GENSAN_ID)     '2016.06.23
                                        GoTo Abort_Tran
                                End Select
                            Loop
                            '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.28
                        
                        
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(GENSANREC.JGYOBU, JGYOBU)
                            Call UniCode_Conv(GENSANREC.NAIGAI, NAIGAI)
                            Call UniCode_Conv(GENSANREC.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                            Call UniCode_Conv(GENSANREC.FILLER, "")
                            Call UniCode_Conv(GENSANREC.INS_TANTO, "2010")
                            Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                            
                            Loop_Cnt = 0
                            
                            Do
                                sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                        If ans = vbCancel Then
''                                            Exit Function
'                                            GoTo Abort_Tran
'                                        End If
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "���Y��Ͻ�", 0)               '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "���Y��Ͻ�", 1, GENSAN_ID)     '2016.06.23
'                                        Exit Function
                                        GoTo Abort_Tran
                                End Select
                            Loop
                        Case Else
                            'Call File_Error(sts, BtOpGetEqual, "���Y��Ͻ�", 0)                 '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "���Y��Ͻ�", 1, GENSAN_ID)       '2016.06.23
'                            Exit Function
                            GoTo Abort_Tran
                    End Select
                End If
                '----------------   2010.07.08 ��
            
            
            
            
            
            
            
'------------ 2005.12.30
                Select Case JGYOBU
                    Case AIRCON, SENTAKU
                        Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                'Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)            '2016.06.23
                                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 1, SOKO_ID)    '2016.06.23
'                                Exit Function
                                GoTo Abort_Tran
                        End Select
        
                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
        
                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            MI_QTY = 0
                        Else
                        
                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                SUMI_QTY = 0
                            Else
                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                MI_QTY = 0
                            End If
                        End If
                        
'------------ 2005.12.30
                        
                    Case Else
                        
                        
                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            SUMI_QTY = 0
                        Else
                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            MI_QTY = 0
                        End If
                End Select
                
        
'                Wk_SOKO = KASO_NYUKA_SOKO
'                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'                    Wk_SOKO = KASO_SMODOSHI_SOKO
'
'                End If
        
        
        
        
        
                '���א��ō݌Ƀf�[�^�X�V�i�{�j
                If Nyuko_Update_Proc(JGYOBU, _
                                    Soko_T(i, j).NAIGAI, _
                                    HIN_GAI, _
                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                    (WORK_SOKO & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, MI_QTY, _
                                    WS_NO, WS_NO, 5, _
                                    DEN_DT & " �`��:" & DEN_NO, , , , MENU_NO, , , StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode), ID_NO2, YOSAN_FROM, YOSAN_TO) Then
'                    Exit Function
                    GoTo Abort_Tran
            
                End If
            
                '�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
                If WK_E_QTY <> 0 Then
                '�݌Ƀf�[�^LOCK
                    If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                        JGYOBU, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        WS_NO, , , 5) Then
'                        Exit Function
                        GoTo Abort_Tran
    
                    End If
        
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                        MI_QTY = WK_E_QTY
                    Else
                        SUMI_QTY = WK_E_QTY
                    End If
            
            
                    If Syuko_Update_Proc(JGYOBU, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        DEN_DT, _
                                        (WORK_SOKO & "01" & "01" & "01"), _
                                        YOIN_MAE_SOUSAI, _
                                        SUMI_QTY, MI_QTY, 0, _
                                        WS_NO, WS_NO, 5) Then
'                        Exit Function
                        GoTo Abort_Tran
        
                    End If
            
            
            
            
            
            
                End If
                
                
                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    GoTo Abort_Tran
                End If
                
                
                Out_Cnt = Out_Cnt + 1
                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                DoEvents
    
            End If
        
        
        
        
        
'�o�ח\��ϊ�################################################## 2005/05/16 Add ���ꕨ����
        Else
            
            If JGYOBU = AIRCON Then
            
                If Not_SHUSI And Trim(HOST_SOKO) <> "R8" Then     '2019/12/13 ����DC ���xR8�Ή�
                Else
                    If IO_KBN = "2" Then
                                
                        
                        wkMUKE_CODE = ""
                        
                        If Trim(HOST_SOKO) = "S8" Then
                            wkMUKE_CODE = "S8"
                        ElseIf Trim(HOST_SOKO) = "R8" Then '2019/12/13 ����DC ���xR8�Ή�
                            wkMUKE_CODE = "R8"             '2019/12/13 ����DC ���xR8�Ή�
                        Else
                            If Trim(HOST_SOKO) = "ST" Then              'ST�ǉ��@   2016.03.11
                                wkMUKE_CODE = "ST"                      '           2016.03.11
                            Else                                        '           2016.03.11
                                If Trim(HOST_SOKO) = "SH" Then
                                Else
                                    Select Case Trim(YOSAN_TO)
                                    
                                        Case "Z0014"
                                            wkMUKE_CODE = "LM"
                                        Case "B0070"
                                            If Trim(HOST_SOKO) = "S2" Then
                                                wkMUKE_CODE = "S2"
                                            Else
                                                wkMUKE_CODE = "AC"
                                            End If
                                        Case Else
                                             wkMUKE_CODE = "AC"
                                    End Select
                                End If                                  '           2016.03.11
                            End If
                        
                            If Trim(HOST_SOKO) <> "S2" Then             '��"S2" and ="B0070"    '2019.01.22
                                If Trim(YOSAN_TO) = "B0070" Then
                                    wkMUKE_CODE = "AC"
                                End If
                            End If
                        End If
            
            
                        If wkMUKE_CODE = "" Then
                        Else
'                            Skip_Flg = False
                                                        '���ח\��d���`�F�b�N
'                            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'                            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
'                            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
'
'                            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
'                                    Skip_Flg = True
'                                Case BtErrKeyNotFound
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "���ח\��")
'                                    Exit Function
'                            End Select
            
            
            
            
'
'                            Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
'                            Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
'                            Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
'
'                            sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
'                                    Skip_Flg = True
'                                Case BtErrKeyNotFound
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��")
'                                    Exit Function
'                            End Select
            
            
                            
                            If Not DUP_FLG Then
            
                                                                '�g�����U�N�V�����J�n
                                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                    Exit Function
                                End If
                                                                '�i�ڃ}�X�^�`�F�b�N
                                If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                                    GoTo Abort_Tran
                                End If
            
        '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
                                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'                                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI) '2019/12/13 ����DC ���xR8�Ή�
                                Call UniCode_Conv(Y_NYUREC.NAIGAI, "1")
                                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                        
                        
                                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                
                
                                Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
                
                
                
                                                    '���׃��X�g�o�̓t���O   2007.06.12
                                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
                
                
                
                
                
                
                
                
                
                                '----------------   2010.07.08 ��
                                Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)                      '���Y����
                                Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, GEN_GENSANKOKU)              '�����\�����Y����
                                Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)      '���ގd����ܰ�����
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                      '����ދ敪
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)                '����ދ敪�K�p�J�n
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)          '����ދ敪����
                                Call UniCode_Conv(Y_NYUREC.ID_NO2, ID_NO2)                              'ID_NO
                                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)                '����溰��
                                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                      '�󒍔N����
                                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)          '�w��[���N����
                                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "")                          '����ؽďo��F
                                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, "")                           '���ɒI��
                                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, "")                           '�O�ؑ��E��
                                
                                
                                
                                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                                
                                
                                
                                
                                
                                '----------------   2010.07.08 ��
                
                
                
                
                
                
                
                
                
                
                
                
                                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                                
'                                Do
'                                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                                    Select Case sts
'                                        Case BtNoErr
'                                            Exit Do
'                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
'                                        Case Else
'                                            Call File_Error(sts, BtOpInsert, "���ח\��")
'                                            Exit Function
'                                    End Select
'                                Loop
        '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
            
                                Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                                Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                                Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                                Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
            
                                If Rec_LENG = 138 Then                                      '2016.04.19
                                    If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
                                        GoTo Abort_Tran
                                    Else
                                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
                                    End If
                                Else                                                        '2016.04.19
                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)           '2016.04.19
                                End If                                                      '2016.04.19
            
            
                                Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                                Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
                                Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                                Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                                Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                                Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                                Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                                Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                                
                                Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                                
                                'If JGYOBU = AIRCON Then             '2008.02.01
                                '    If Left(DEN_NO, 1) = "0" Then
                                '        DEN_NO = Right(DEN_NO, Len(DEN_NO) - 1)
                                '    End If
                                'End If
                                
                                Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                                
                                
                                
                                
                                
                                wkStr = Format(Val(YOTEI_QTY), "0000000")
                                Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
                                Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                                Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                                
                                Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                                Call UniCode_Conv(Y_SYUREC.TANKA, "")
                                
                                
                                Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                                Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                                Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                                Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                                
                                
                                
                                Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)
    
                                Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
    
                                Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)
    
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                                Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                                Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                                Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                                Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)
    
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
                                Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
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
    
                                Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                                Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                                Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                                Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                                Call UniCode_Conv(Y_SYUREC.FILLER, "")
            
                                Loop_Cnt = 0
                                
                                Do
                                    sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
''                                                Exit Function
'                                                GoTo Abort_Tran
'                                            End If
                                        
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                GoTo Abort_Tran
                                            End If
                                        
                                            DoEvents
                                            Sleep (500)
                                        
                                        
                                        Case Else
                                            'Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)            '2016.06.23
                                            Call File_Error(sts, BtOpInsert, "�o�ח\��", 1, Y_SYU_ID)   '2016.06.23
'                                            Exit Function
                                            GoTo Abort_Tran
                                    End Select
                                Loop
            
                                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts <> BtNoErr Then
                                    GoTo Abort_Tran
                                End If
            
                                Out_Cnt = Out_Cnt + 1
                                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                                DoEvents
            
                                If SYUKA_LOG_ON Then
                                    Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
                                End If
            
'                                If Not Fast_Flg Then
'                                    Close #DUP_SYUKANo
'                                End If
                            End If
                        End If
                    End If
                End If
            End If
'#################################################################################### 2005/05/16 Add ��
        
'#################################################################################### 2008/02/22 Add ��
            If JGYOBU = SOJIKI Then
            
                If IO_KBN = "1" Then
                            
                    wkMUKE_CODE = ""
                            
                    If Trim(HOST_SOKO) = "SS" Then
                        wkMUKE_CODE = "00000000"
                    End If
                    If Trim(HOST_SOKO) = "ZZ" Then
                        wkMUKE_CODE = "88888888"
                    End If
        
        
                    If wkMUKE_CODE = "" Then
                    Else
 '                       Skip_Flg = False
                                                    '���ח\��d���`�F�b�N
 '                       Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
 '                       Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
 '                       Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
 '
 '                       sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
 '                       Select Case sts
 '                           Case BtNoErr
 '                               Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
 '                               Skip_Flg = True
 '                           Case BtErrKeyNotFound
 '                           Case Else
 '                               Call File_Error(sts, BtOpGetEqual, "���ח\��")
 '                               Exit Function
 '                       End Select
        
        
 '                       Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
 '                       Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
 '                       Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
 '
 '                       sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
 '                       Select Case sts
 '                           Case BtNoErr
 '                               Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
 '                               Skip_Flg = True
 '                           Case BtErrKeyNotFound
 '                           Case Else
 '                               Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��")
 '                               Exit Function
 '                       End Select
        
        
        
                        If Not DUP_FLG Then
        
                                                            '�g�����U�N�V�����J�n
                            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                Exit Function
                            End If
                                                            '�i�ڃ}�X�^�`�F�b�N
                            If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                                GoTo Abort_Tran
                            End If
        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
                            Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                            Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                            Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                            Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                    
                    
                            Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                            Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                            Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                            Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                            Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                            Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                            Call UniCode_Conv(Y_NYUREC.TANKA, "")
                            Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                            Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                            Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                            Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                            Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                            Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                            Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                            Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                            Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                            Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                            Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                            Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                            Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                            Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                            Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                            Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                            Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                            Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                            Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                            Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                            Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                            Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                            Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                            Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
            
            
                            Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
            
            
            
                                                '���׃��X�g�o�̓t���O   2007.06.12
                            Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
            
                            Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                            Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                            Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                            Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
            
                            Call UniCode_Conv(Y_NYUREC.FILLER, "")
                            
'                            Do
'                                sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                       Exit Do
'�@                                  Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                        If ans = vbCancel Then
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpInsert, "���ח\��")
'                                        Exit Function
'                                End Select
'                            Loop
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
        
                            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                            Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                            Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                            
                            If Trim(HOST_SOKO) = "SS" Then
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                            Else
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                            End If
        
                            
                            If Rec_LENG = 138 Then                                  '2016.04.19
                                If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
                                    GoTo Abort_Tran
                                Else
                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
                                End If
                            Else                                                    '2016.04.19
                                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)       '2016.04.19
                            End If                                                  '2016.04.19
        
                            Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                            Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
                            Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                            Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                            Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                            Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                            Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                            Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                            
                            Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                            'If JGYOBU = AIRCON Then             '2008.02.01
                            '    If Left(DEN_NO, 1) = "0" Then
                            '        DEN_NO = Right(DEN_NO, Len(DEN_NO) - 1)
                            '    End If
                            'End If
                            
                            Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                            
                            
                            
                            
                            
                            wkStr = Format(Val(YOTEI_QTY), "0000000")
                            Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                            Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
                            Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                            Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                            
                            Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                            Call UniCode_Conv(Y_SYUREC.TANKA, "")
                            
                            
                            Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                            Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                            Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                            Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                            
                            
                            
                            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                            Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                            Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                            Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                            Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                            If Trim(HOST_SOKO) = "SS" Then
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                            Else
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                            End If

                            
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                            Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                            Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                            Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                            Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

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
                            Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
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

                            Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                            Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                            Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                            Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
        
                            Loop_Cnt = 0
        
                            Do
                                sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                        If ans = vbCancel Then
''                                            Exit Function
'                                            GoTo Abort_Tran
'                                        End If
                                    
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)            '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "�o�ח\��", 1, Y_SYU_ID)   '2016.06.23
'                                        Exit Function
                                        GoTo Abort_Tran
                                End Select
                            Loop
        
                            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                GoTo Abort_Tran
                            End If
        
                            Out_Cnt = Out_Cnt + 1
                            lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                            DoEvents
        
                            If SYUKA_LOG_ON Then
                                Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
                            End If
        
'                            If Not Fast_Flg Then
'                                Close #DUP_SYUKANo
'                            End If
                        End If
                    End If
                End If
            End If
        
'#################################################################################### 2008/02/22 Add ��
        
        
        
        
        
        
        
        
        
        
        
'#################################################################################### 2010/07/21 Add ��
            If JGYOBU = DENKA Then
            
                If IO_KBN = "2" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "01KA" Then
                            
'                    Skip_Flg = False
                                                '���ח\��d���`�F�b�N
'                    Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'                    Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
'                    Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
'
'                    sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
'                            Skip_Flg = True
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "���ח\��")
'                            Exit Function
'                    End Select
    
'                    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
'                    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
'                    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
'
'                    sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
'                            Skip_Flg = True
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��")
'                            Exit Function
'                    End Select
    
    
                    If Not DUP_FLG Then
    
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '�i�ڃ}�X�^�`�F�b�N
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '���׃��X�g�o�̓t���O   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
'                        Do
'                            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Exit Do
'                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                    If ans = vbCancel Then
'                                        Exit Function
'                                    End If
'                                Case Else
'                                    Call File_Error(sts, BtOpInsert, "���ח\��")
'                                    Exit Function
'                            End Select
'                        Loop
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

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
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
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

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "�o�ח\��", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
                        End If
    
'                        If Not Fast_Flg Then
'                            Close #DUP_SYUKANo
'                        End If
                    End If
                End If
            End If
        
'#################################################################################### 2008/02/22 Add ��
        
        
'#################################################################################### 2018/09/19 Add ��
            If JGYOBU = OVEN Then
                If (IO_KBN = "2" And Trim(HOST_SOKO) = "01" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "0107") Or _
                    (IO_KBN = "2" And Trim(HOST_SOKO) = "06" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "0607") Or _
                    (IO_KBN = "2" And Trim(YOSAN_FROM) = "SDC" And Mid(AITESAKI_CODE, 3, 2) = "KA") Then
                            
    
    
                    If Not DUP_FLG Then
    
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '�i�ڃ}�X�^�`�F�b�N
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '���׃��X�g�o�̓t���O   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

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
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
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

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "�o�ח\��", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
                        End If
    
                    End If
                End If
            End If
        
'#################################################################################### 2018/09/19 Add ��
        
'#################################################################################### 2018/09/20 Add ��
            If JGYOBU = SHOKUSEN Then
                If IO_KBN = "2" And Trim(YOSAN_TO) = "904" Then
                            
    
    
                    If Not DUP_FLG Then
    
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '�i�ڃ}�X�^�`�F�b�N
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '���׃��X�g�o�̓t���O   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
    '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

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
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
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

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "�o�ח\��", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
                        End If
    
                    End If
                End If
            End If
        
'#################################################################################### 2018/09/19 Add ��
        
        
        
        
        
        End If
        
        
        
    
    Loop

    Nyuka_Update_Proc = False
    Exit Function

Abort_Tran:
    
'>>>>>  2015.11.19
    If Fast_Flg Then
        Open (FileName) For Output As DUP_SYUKANo
        Write #DUP_SYUKANo, , , "���Ɏ捞�ُ݈탊�X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
        Write #DUP_SYUKANo, "�G���[���e", "�`�[���t", "�`�[��", "�\�Z��", "�\�Z��", "νđq��", "�i��", "����", "TEXT_NO"      '2015.11.19
        Fast_Flg = False
    End If


    Write #DUP_SYUKANo, "���d����",
    Write #DUP_SYUKANo, DEN_DT,
    Write #DUP_SYUKANo, DEN_NO,
    Write #DUP_SYUKANo, YOSAN_FROM,
    Write #DUP_SYUKANo, YOSAN_TO,
    Write #DUP_SYUKANo, HOST_SOKO,

    Write #DUP_SYUKANo, HIN_GAI,
    Write #DUP_SYUKANo, YOTEI_QTY,
    Write #DUP_SYUKANo, TEXT_NO
'>>>>>  2015.11.19
    
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


    

End Function
    
Private Function Syuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   �u�o�ח\��f�[�^�v�X�V����
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim Skip_Flg    As Boolean
Dim sts         As Integer
    
Dim ans         As Integer

Dim c               As String * 128

Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
    
Dim INS_NOW         As String * 14

Dim wkSS            As String
Dim wkMUKE_CODE     As String
Dim wkCHOKU_KBN     As String * 1


Dim wkText          As String
Dim Length      As Integer


Dim JGYOBA              As String * 8       '���Ə�
Dim DATA_KBN            As String * 1       '�f�[�^�敪
Dim TORI_KBN            As String * 2       '����敪
Dim ID_NO               As String * 12      'ID-NO
Dim KAIKEI_JGYOBA       As String * 8       '��v�p���Ə꺰��
Dim SHISAN_JGYOBA       As String * 8       '���Y�Ǘ����Ə꺰��
Dim HIN_NO              As String * 20      '�i�ڔԍ�
Dim DEN_NO              As String * 10      '�`�[�ԍ�
Dim SURYO               As String * 7       '�o�ɐ���
Dim MUKE_CODE           As String * 8       '�o�ɐ�
Dim SYUKO_SYUSI         As String * 8       '�o�Ɏ��x
Dim SHISAN_SYUSI        As String * 8       '���Y�Ǘ��p�݌Ɏ��x����
Dim HOJYO_SYUSI         As String * 8       '�⏕�݌Ɏ��x����
Dim SYUKO_YMD           As String * 8       '�o�ɓ��t
Dim TANKA               As String * 10      '�P��
Dim ODER_NO             As String * 12      '�I�[�_�[�ԍ�
Dim ITEM_NO             As String * 5       '�A�C�e���ԍ�
Dim ODER_NO_R           As String * 5       '�I�[�_�[����
Dim KOSO_KEITAI         As String * 14      '���`��       10-->14 2011.10.31
Dim SYUKA_YMD           As String * 8       '�o�ד�
Dim TANABAN1            As String * 10      '�I�ԂP
Dim TANABAN2            As String * 10      '�I�ԂQ
Dim TANABAN3            As String * 10      '�I�ԂR
Dim MUKE_NAME           As String * 24      '�o�ɐ於��
Dim CYU_KBN             As String * 1       '�����敪
Dim CYU_KBN_NAME        As String * 40      '�����敪����
Dim ORIGIN1             As String * 10      '���Y���P
Dim ORIGIN2             As String * 10      '���Y���Q
Dim BIKOU2              As String * 40      '���l�Q
Dim HAN_KBN             As String * 1       '�̔��敪
Dim CHOKU_KBN           As String * 1       '�����敪
Dim UNIT_ID_NO          As String * 12      '�ƯďC��ID-NO
Dim ZAIKO_HIKIATE       As String * 3       '�݌Ɉ�������
Dim GOKON_KANRI_NO      As String * 8       '�����Ǘ��ԍ�
Dim JYUCHU_ZAN          As String * 7       '�󒍎c����
Dim KYOKYU_KBN          As String * 1       '�����敪
Dim SHOHIN_SYUSI        As String * 8       '���i���[������x
Dim S_SHISAN_SYUSI      As String * 8       '���i���[�i���Y�Ǘ����x����
Dim S_HOJYO_SYUSI       As String * 8       '���i���[�i�⏕���x����
Dim BIKOU1              As String * 40      '���l�P
Dim CHOHA_KBN           As String * 1       '���[�敪
Dim JYU_HIN_NO          As String * 40      '�󒍕i�ڔԍ�
Dim HIN_NAME            As String * 40      '�i��
Dim HIN_CHANGE_KBN      As String * 1       '�i�ԕύX�敪
Dim MODULE_EXCHANGE     As String * 1       '���W���[�������敪
Dim ZAIKO_SYUSI         As String * 8       '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
Dim ZAN_SHISAN_SYUSI    As String * 8       '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
Dim ZAN_HOJYO_SYUSI     As String * 8       '�c�݌ɂ܂Ƃߕ⏕���x����
Dim NOUKI_YMD           As String * 8       '�w��[��
Dim SERVICE_KANRI_NO    As String * 9       '�T�[�r�X��ЊǗ��ԍ�
Dim KISHU_CODE          As String * 3       '�@��i�ڃR�[�h
Dim ENVIRONMENT_KBN     As String * 1       '���K�i���i�敪
Dim SS_CODE             As String * 8       '������R�[�h
Dim KEPIN_KAIJYO        As String * 1       '���i�����敪


Dim wkSyukaRec      As wkSyukaRec_tag










Dim Upd_com             As Integer          '2008.02.23

Dim wkTemp              As String


Dim WK_Y_QTY            As Long             '2009.04.14
Dim WK_Qty              As Long             '2009.04.14
Dim WK_E_QTY            As Long             '2009.04.14

Dim WORK_SOKO           As String * 2       '2009.04.14

Dim SUMI_QTY            As Long             '2009.04.14
Dim MI_QTY              As Long             '2009.04.14

'2011.01.19
Dim GENSAN_CNT          As Integer
Dim com                 As Integer
Dim GENSANKOKU          As String * 20

Dim Loop_Cnt            As Integer

'2011.01.19



    Syuka_Update_Proc = True



    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)


    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")

    Do Until EOF(FileNo)
'        Line Input #FileNo, wkText
        Get #FileNo, , wkSyukaRec
        
        
        
'        If StrConv(wkSyukaRec.CRLF, vbUnicode) <> vbCrLf Then
'            Call NG_File_Make_Proc
'            Exit Do
'        End If
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
'        Length = 1
'        JGYOBA = Mid(wkText, Length, Len(JGYOBA))                   '���Ə�
        JGYOBA = StrConv(wkSyukaRec.JGYOBA, vbUnicode)
        
        
'        Length = Length + Len(JGYOBA)
'        DATA_KBN = Mid(wkText, Length, Len(DATA_KBN))               '�f�[�^�敪
        DATA_KBN = StrConv(wkSyukaRec.DATA_KBN, vbUnicode)
        
        
        
'        Length = Length + Len(DATA_KBN)
'        TORI_KBN = Mid(wkText, Length, Len(TORI_KBN))               '����敪
        TORI_KBN = StrConv(wkSyukaRec.TORI_KBN, vbUnicode)
    
'        Length = Length + Len(TORI_KBN)
'        ID_NO = Mid(wkText, Length, Len(ID_NO))                     'ID-NO
        ID_NO = StrConv(wkSyukaRec.ID_NO, vbUnicode)
    
'        Length = Length + Len(ID_NO)
'        KAIKEI_JGYOBA = Mid(wkText, Length, Len(KAIKEI_JGYOBA))     '��v�p���Ə꺰��
        KAIKEI_JGYOBA = StrConv(wkSyukaRec.KAIKEI_JGYOBA, vbUnicode)
    
'        Length = Length + Len(KAIKEI_JGYOBA)
'        SHISAN_JGYOBA = Mid(wkText, Length, Len(SHISAN_JGYOBA))     '���Y�Ǘ����Ə꺰��
        SHISAN_JGYOBA = StrConv(wkSyukaRec.SHISAN_JGYOBA, vbUnicode)
    
'        Length = Length + Len(SHISAN_JGYOBA)
'        HIN_NO = Mid(wkText, Length, Len(HIN_NO))                   '�i�ڔԍ�
        HIN_NO = StrConv(wkSyukaRec.HIN_NO, vbUnicode)
    
'        Length = Length + Len(HIN_NO)
'        DEN_NO = Mid(wkText, Length, Len(DEN_NO))                   '�`�[�ԍ�
        DEN_NO = StrConv(wkSyukaRec.DEN_NO, vbUnicode)
    
'        Length = Length + Len(DEN_NO)
'        SURYO = Mid(wkText, Length, Len(SURYO))                     '�o�ɐ���
        SURYO = StrConv(wkSyukaRec.SURYO, vbUnicode)
    
'        Length = Length + Len(SURYO)
'        MUKE_CODE = Mid(wkText, Length, Len(MUKE_CODE))             '�o�ɐ�
        MUKE_CODE = StrConv(wkSyukaRec.MUKE_CODE, vbUnicode)
    
'        Length = Length + Len(MUKE_CODE)
'        SYUKO_SYUSI = Mid(wkText, Length, Len(SYUKO_SYUSI))         '�o�Ɏ��x
        SYUKO_SYUSI = StrConv(wkSyukaRec.SYUKO_SYUSI, vbUnicode)
    
'        Length = Length + Len(SYUKO_SYUSI)
'        SHISAN_SYUSI = Mid(wkText, Length, Len(SHISAN_SYUSI))       '���Y�Ǘ��p�݌Ɏ��x����
        SHISAN_SYUSI = StrConv(wkSyukaRec.SHISAN_SYUSI, vbUnicode)
    
    
'        Length = Length + Len(SHISAN_SYUSI)
'        HOJYO_SYUSI = Mid(wkText, Length, Len(HOJYO_SYUSI))         '�⏕�݌Ɏ��x����
        HOJYO_SYUSI = StrConv(wkSyukaRec.HOJYO_SYUSI, vbUnicode)
        
'        Length = Length + Len(HOJYO_SYUSI)
'        SYUKO_YMD = Mid(wkText, Length, Len(SYUKO_YMD))             '�o�ɓ��t
        SYUKO_YMD = StrConv(wkSyukaRec.SYUKO_YMD, vbUnicode)
    
'        Length = Length + Len(SYUKO_YMD)
'        TANKA = Mid(wkText, Length, Len(TANKA))                     '�P��
        TANKA = StrConv(wkSyukaRec.TANKA, vbUnicode)
    
'        Length = Length + Len(TANKA)
'        ODER_NO = Mid(wkText, Length, Len(ODER_NO))                 '�I�[�_�[�ԍ�
        ODER_NO = StrConv(wkSyukaRec.ODER_NO, vbUnicode)
    
'        Length = Length + Len(ODER_NO)
'        ITEM_NO = Mid(wkText, Length, Len(ITEM_NO))                 '�A�C�e���ԍ�
        ITEM_NO = StrConv(wkSyukaRec.ITEM_NO, vbUnicode)
    
'        Length = Length + Len(ITEM_NO)
'        ODER_NO_R = Mid(wkText, Length, Len(ODER_NO_R))             '�I�[�_�[����
        ODER_NO_R = StrConv(wkSyukaRec.ODER_NO_R, vbUnicode)
    
'        Length = Length + Len(ODER_NO_R)
'        KOSO_KEITAI = Mid(wkText, Length, Len(KOSO_KEITAI))         '���`��
        KOSO_KEITAI = StrConv(wkSyukaRec.KOSO_KEITAI, vbUnicode)
    
'        Length = Length + Len(KOSO_KEITAI)
'        SYUKA_YMD = Mid(wkText, Length, Len(SYUKA_YMD))             '�o�ד�
        SYUKA_YMD = StrConv(wkSyukaRec.SYUKA_YMD, vbUnicode)
    
'        Length = Length + Len(SYUKA_YMD)
'        TANABAN1 = Mid(wkText, Length, Len(TANABAN1))               '�I�ԂP
        TANABAN1 = StrConv(wkSyukaRec.TANABAN1, vbUnicode)
    
'        Length = Length + Len(TANABAN1)
'        TANABAN2 = Mid(wkText, Length, Len(TANABAN2))               '�I�ԂQ
        TANABAN2 = StrConv(wkSyukaRec.TANABAN2, vbUnicode)
    
'        Length = Length + Len(TANABAN2)
'        TANABAN3 = Mid(wkText, Length, Len(TANABAN3))               '�I�ԂR
        TANABAN3 = StrConv(wkSyukaRec.TANABAN3, vbUnicode)
    
    
    
    
'        Length = Length + Len(TANABAN3)
'        MUKE_NAME = Mid(wkText, Length, Len(MUKE_NAME))             '�o�ɐ於��
        MUKE_NAME = StrConv(wkSyukaRec.MUKE_NAME, vbUnicode)
    
            
    
    
    
    
    
    
'        Length = Length + Len(MUKE_NAME)
'        CYU_KBN = Mid(wkText, Length, Len(CYU_KBN))                 '�����敪
        CYU_KBN = StrConv(wkSyukaRec.CYU_KBN, vbUnicode)
    
    
    
    
    
'        Length = Length + Len(CYU_KBN)
'        CYU_KBN_NAME = Mid(wkText, Length, Len(CYU_KBN_NAME))       '�����敪����
        CYU_KBN_NAME = StrConv(wkSyukaRec.CYU_KBN_NAME, vbUnicode)
        
        
        
'        Length = Length + Len(CYU_KBN_NAME)
'        ORIGIN1 = Mid(wkText, Length, Len(ORIGIN1))                 '���Y���P
        ORIGIN1 = StrConv(wkSyukaRec.ORIGIN1, vbUnicode)
    
'        Length = Length + Len(ORIGIN1)
'        ORIGIN2 = Mid(wkText, Length, Len(ORIGIN2))                 '���Y���Q
        ORIGIN2 = StrConv(wkSyukaRec.ORIGIN2, vbUnicode)
    
'        Length = Length + Len(ORIGIN2)
'        BIKOU2 = Mid(wkText, Length, Len(BIKOU2))                   '���l�Q
        BIKOU2 = StrConv(wkSyukaRec.BIKOU2, vbUnicode)
    
'        Length = Length + Len(BIKOU2)
'        HAN_KBN = Mid(wkText, Length, Len(HAN_KBN))                 '�̔��敪
        HAN_KBN = StrConv(wkSyukaRec.HAN_KBN, vbUnicode)
    
'        Length = Length + Len(HAN_KBN)
'        CHOKU_KBN = Mid(wkText, Length, Len(CHOKU_KBN))             '�����敪
        CHOKU_KBN = StrConv(wkSyukaRec.CHOKU_KBN, vbUnicode)
    
'        Length = Length + Len(CHOKU_KBN)
'        UNIT_ID_NO = Mid(wkText, Length, Len(UNIT_ID_NO))           '�ƯďC��ID-NO
        UNIT_ID_NO = StrConv(wkSyukaRec.UNIT_ID_NO, vbUnicode)
    
'        Length = Length + Len(UNIT_ID_NO)
'        ZAIKO_HIKIATE = Mid(wkText, Length, Len(ZAIKO_HIKIATE))     '�݌Ɉ�������
        ZAIKO_HIKIATE = StrConv(wkSyukaRec.ZAIKO_HIKIATE, vbUnicode)
    
'        Length = Length + Len(ZAIKO_HIKIATE)
'        GOKON_KANRI_NO = Mid(wkText, Length, Len(GOKON_KANRI_NO))   '�����Ǘ��ԍ�
        GOKON_KANRI_NO = StrConv(wkSyukaRec.GOKON_KANRI_NO, vbUnicode)
    
'        Length = Length + Len(GOKON_KANRI_NO)
'        JYUCHU_ZAN = Mid(wkText, Length, Len(JYUCHU_ZAN))           '�󒍎c����
        JYUCHU_ZAN = StrConv(wkSyukaRec.JYUCHU_ZAN, vbUnicode)
    
'        Length = Length + Len(JYUCHU_ZAN)
'        KYOKYU_KBN = Mid(wkText, Length, Len(KYOKYU_KBN))           '�����敪
        KYOKYU_KBN = StrConv(wkSyukaRec.KYOKYU_KBN, vbUnicode)
    
'        Length = Length + Len(KYOKYU_KBN)
'        SHOHIN_SYUSI = Mid(wkText, Length, Len(SHOHIN_SYUSI))       '���i���[������x
        SHOHIN_SYUSI = StrConv(wkSyukaRec.SHOHIN_SYUSI, vbUnicode)
    
'        Length = Length + Len(SHOHIN_SYUSI)
'        S_SHISAN_SYUSI = Mid(wkText, Length, Len(S_SHISAN_SYUSI))   '���i���[�i���Y�Ǘ����x����
        S_SHISAN_SYUSI = StrConv(wkSyukaRec.S_SHISAN_SYUSI, vbUnicode)
    
'        Length = Length + Len(S_SHISAN_SYUSI)
'        S_HOJYO_SYUSI = Mid(wkText, Length, Len(S_HOJYO_SYUSI))     '���i���[�i�⏕���x����
        S_HOJYO_SYUSI = StrConv(wkSyukaRec.S_HOJYO_SYUSI, vbUnicode)
    
'        Length = Length + Len(S_SHISAN_SYUSI)
'        BIKOU1 = Mid(wkText, Length, Len(BIKOU1))                   '���l�P
        BIKOU1 = StrConv(wkSyukaRec.BIKOU1, vbUnicode)
    
'        Length = Length + Len(BIKOU1)
'        CHOHA_KBN = Mid(wkText, Length, Len(CHOHA_KBN))             '���[�敪
        CHOHA_KBN = StrConv(wkSyukaRec.CHOHA_KBN, vbUnicode)
    
'        Length = Length + Len(CHOHA_KBN)
'        JYU_HIN_NO = Mid(wkText, Length, Len(JYU_HIN_NO))           '�󒍕i�ڔԍ�
        JYU_HIN_NO = StrConv(wkSyukaRec.JYU_HIN_NO, vbUnicode)
    
'        Length = Length + Len(JYU_HIN_NO)
'        HIN_NAME = Mid(wkText, Length, Len(HIN_NAME))               '�i��
        HIN_NAME = StrConv(wkSyukaRec.HIN_NAME, vbUnicode)
    
'        Length = Length + Len(HIN_NAME)
'        HIN_CHANGE_KBN = Mid(wkText, Length, Len(HIN_CHANGE_KBN))   '�i�ԕύX�敪
        HIN_CHANGE_KBN = StrConv(wkSyukaRec.HIN_CHANGE_KBN, vbUnicode)
    
'        Length = Length + Len(HIN_CHANGE_KBN)
'        MODULE_EXCHANGE = Mid(wkText, Length, Len(MODULE_EXCHANGE)) '���W���[�������敪
        MODULE_EXCHANGE = StrConv(wkSyukaRec.MODULE_EXCHANGE, vbUnicode)
    
'        Length = Length + Len(MODULE_EXCHANGE)
'        ZAIKO_SYUSI = Mid(wkText, Length, Len(ZAIKO_SYUSI))         '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
        ZAIKO_SYUSI = StrConv(wkSyukaRec.ZAIKO_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAIKO_SYUSI)
'        ZAN_SHISAN_SYUSI = Mid(wkText, Length, Len(ZAN_SHISAN_SYUSI))   '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
        ZAN_SHISAN_SYUSI = StrConv(wkSyukaRec.ZAN_SHISAN_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAN_SHISAN_SYUSI)
'        ZAN_HOJYO_SYUSI = Mid(wkText, Length, Len(ZAN_HOJYO_SYUSI)) '�c�݌ɂ܂Ƃߕ⏕���x����
        ZAN_HOJYO_SYUSI = StrConv(wkSyukaRec.ZAN_HOJYO_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAN_HOJYO_SYUSI)
'        NOUKI_YMD = Mid(wkText, Length, Len(NOUKI_YMD))             '�w��[��
        NOUKI_YMD = StrConv(wkSyukaRec.NOUKI_YMD, vbUnicode)
    
'        Length = Length + Len(NOUKI_YMD)
'        SERVICE_KANRI_NO = Mid(wkText, Length, Len(SERVICE_KANRI_NO))   '�T�[�r�X��ЊǗ��ԍ�
        SERVICE_KANRI_NO = StrConv(wkSyukaRec.SERVICE_KANRI_NO, vbUnicode)
    
'        Length = Length + Len(SERVICE_KANRI_NO)
'        KISHU_CODE = Mid(wkText, Length, Len(KISHU_CODE))           '�@��i�ڃR�[�h
        KISHU_CODE = StrConv(wkSyukaRec.KISHU_CODE, vbUnicode)
    
'        Length = Length + Len(KISHU_CODE)
'        ENVIRONMENT_KBN = Mid(wkText, Length, Len(ENVIRONMENT_KBN)) '���K�i���i�敪
        ENVIRONMENT_KBN = StrConv(wkSyukaRec.ENVIRONMENT_KBN, vbUnicode)
    
'        Length = Length + Len(ENVIRONMENT_KBN)
'        SS_CODE = Mid(wkText, Length, Len(SS_CODE))                 '������R�[�h
        SS_CODE = StrConv(wkSyukaRec.SS_CODE, vbUnicode)
    
'        Length = Length + Len(SS_CODE)
'        KEPIN_KAIJYO = Mid(wkText, Length, Len(KEPIN_KAIJYO))       '���i�����敪
        KEPIN_KAIJYO = StrConv(wkSyukaRec.KEPIN_KAIJYO, vbUnicode)
        
If ID_NO = "700092591973" Then
     Debug.Print ID_NO
End If
        
        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(SYUKO_SYUSI) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
'-------------------------- PPSC��荞�݂��
        If Trim(CYU_KBN) = "" Then
            wkCHOKU_KBN = ""
        Else
            wkCHOKU_KBN = "1"
        End If
                                                                                                
                                                        
        If Trim(CYU_KBN) = "" Then
            If Trim(MUKE_CODE) = "A1" Or _
                Trim(MUKE_CODE) = "A2" Or _
                Trim(MUKE_CODE) = "A3" Or _
                Trim(MUKE_CODE) = "A4" Or _
                Trim(MUKE_CODE) = "A5" Or _
                Trim(MUKE_CODE) = "A6" Or _
                Trim(MUKE_CODE) = "A7" Then
                CYU_KBN = "3"
            End If
        
            If MUKE_CODE = "22000440" Or _
                MUKE_CODE = "22000441" Or _
                MUKE_CODE = "22000442" Or _
                MUKE_CODE = "22000443" Or _
                MUKE_CODE = "22000444" Or _
                MUKE_CODE = "22000445" Or _
                MUKE_CODE = "22000446" Then
                CYU_KBN = "2"
            End If
        
        End If
                                                        
        If Trim(CYU_KBN) = "" Then
            CYU_KBN = "3"
        End If
    
'-------------------------- PPSC��荞�݂��
    
    
    
    
    
    
    
    
        '�u00036003�v�̑Ή�2006.06.03
    
        'If JGYOBU = AIRCON Then                        '�G�A�R���͏��O 2006.11.10
'        If JGYOBU = AIRCON Or JGYOBU = OVEN Then        '�G�A�R���A�d�q�����W�͏��O 2011.05.16
        
        If JGYOBU = AIRCON Or JGYOBU = OVEN Or JGYOBU = REIZOU Or JGYOBU = SHOKUSEN Then         '�G�A�R���A�d�q�����W�͏��O 2011.05.16 �①�ɒǉ� 2014.12.17 �H�� 2015.03.03
        Else
            If Trim(MUKE_CODE) = "00036003" Then
                Skip_Flg = True
            End If
        End If
    
        If Not Skip_Flg Then
                                '�o�ח\��d���`�F�b�N
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_SYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�`�[�h�c��" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Skip_Flg = True
                
                
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
                        Write #DUP_SYUKANo, , , "�o�׎捞�ُ݈탊�X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "�o�ד�", "�`�[��", "�x���溰��", "�q��/�r�r����", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"                  '2015.11.19
                        Write #DUP_SYUKANo, "�G���[���e", "�o�ד�", "�`�[��", "�o�א溰��", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"      '2015.11.19
                        Fast_Flg = False
                    End If
                
                
                    Write #DUP_SYUKANo, "���d����",
                    Write #DUP_SYUKANo, SYUKA_YMD,
                    Write #DUP_SYUKANo, DEN_NO,
                    Write #DUP_SYUKANo, MUKE_CODE,
                    Write #DUP_SYUKANo, MUKE_NAME,
                    Write #DUP_SYUKANo, CYU_KBN,
                    Write #DUP_SYUKANo, CYU_KBN_NAME,
                    Write #DUP_SYUKANo, HIN_NO,
                    Write #DUP_SYUKANo, SURYO,
                    Write #DUP_SYUKANo, ID_NO
                
                    Upd_com = BtOpUpdate
                
                Case BtErrKeyNotFound
                    Upd_com = BtOpInsert
                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "�o�ח\��", 0)                      '2016.06.23
                    Call File_Error(sts, BtOpGetEqual, "�o�ח\��", 1, Y_SYU_ID)             '2016.06.23
                    Exit Function
            End Select
    
    
    
    
            
            If Not Skip_Flg Then
                
                                '�i�ڃ}�X�^�̃`�F�b�N
'                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Soko_T(i, j).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode))
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Skip_Flg = True
'                        Call Log_Out(LOG_F, "�`�[ID=" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'                        Exit Function
'                End Select
                
                
                
                If Not Skip_Flg Then
                                                    '�g�����U�N�V�����J�n
                    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                        Exit Function
                    End If
                    
                    
                                                '�i�ڃ}�X�^�`�F�b�N
                    If Item_Check_Proc(Out_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_NO, , HIN_NAME) Then
                        GoTo Abort_Tran
                    End If
                    '2012.12.20
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                        Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                    End If
                    '2012.12.20
                                                                    
                    wkMUKE_CODE = MUKE_CODE
                                                                    
                                                                    
                    If Len(Trim(SS_CODE)) = 0 Or _
                        IsNumeric(Trim(SS_CODE)) Then
                    Else
                        SS_CODE = ""
                    End If
                                                                    
                                                                    
'-----------    2005.12.30
                    If JGYOBU = AIRCON Then
                        
                        'MTS���ނ̓ǂݑւ�
                        If GetIni(App.EXEName, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
                        Else
                            Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, Trim(c))
                            Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                        End If
                        
                        
                        
                        
                        '�G�A�R���������ꍇ������ɒ�������  2004.12.01
                        
                        If Trim(MUKE_CODE) = Trim(SS_CODE) Then
                            SS_CODE = ""
                        Else
                            If Len(Trim(SS_CODE)) <> 0 Then
                                MUKE_CODE = SS_CODE
                                SS_CODE = ""
                            End If
                        End If
                    Else
                        
                        '����@�̏ꍇ�A���l�P�𒼑���ɃZ�b�g 2006.03.25
                        If JGYOBU = SENTAKU And SYUKO_SYUSI = "S2" Then
                            If StrComp(ODER_NO, "FAX", vbTextCompare) Then
                            
                                wkSS = ""
                            
                                For k = 1 To Len(BIKOU1)
                                    If IsNumeric(Mid(BIKOU1, k, 1)) Then
                                        wkSS = wkSS & Mid(BIKOU1, k, 1)
                                    Else
                                        Exit For
                                    End If
                                Next k
                            
                                SS_CODE = wkSS
                            End If
                        
                        End If
                        
                        
                        
                        '���̎��ƕ��͌���̂܂�
                        If Len(Trim(SS_CODE)) = 0 Or _
                            IsNumeric(Trim(SS_CODE)) Then
                        Else
                            SS_CODE = ""
                        End If
                    End If
                        
                    Call UniCode_Conv(K0_MTS.MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                             
                             
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            
                            
                            If JGYOBU = AIRCON Then
                                '�G�A�R���������ꍇ������ɒ�����Ō�����Ͻ���V�K�쐬  2004.12.01
                            
                                Call UniCode_Conv(MTSREC.NAIGAI, Soko_T(i, j).NAIGAI)
                                Call UniCode_Conv(MTSREC.DATA_KBN, "")
                                Call UniCode_Conv(MTSREC.MUKE_CODE, MUKE_CODE)
                                Call UniCode_Conv(MTSREC.SS_CODE, "")
                                Call UniCode_Conv(MTSREC.MUKE_NAME, MUKE_NAME)
                                Call UniCode_Conv(MTSREC.SS_NAME, "")
                                Call UniCode_Conv(MTSREC.MUKE_DNAME, MUKE_NAME)
                                Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "")
                                Call UniCode_Conv(MTSREC.FILLER, "")
                                
                                Loop_Cnt = 0
                                
                                
                                Do
                                    sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                        
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                GoTo Abort_Tran
                                            End If
                                        
                                            DoEvents
                                            Sleep (500)
                                       
                                        
                                        
                                        Case Else
                                            'Call File_Error(sts, BtOpInsert, "������Ǘ�Ͻ�" & "key=" & StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode), 0)                      '2016.06.23
                                            Call File_Error(sts, BtOpInsert, "������Ǘ�Ͻ�" & "key=" & StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode), 1, MTS_ID)               '2016.06.23

'                                            Exit Function      '2015.11.19
                                            GoTo Abort_Tran     '2015.11.19
                                    End Select
                                Loop
                                                        
                                                        
                                                        
                                                        
                            
                            Else
                                '���̎��ƕ��͌���̂܂�
                                If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
                                    Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
                                    Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                                Else
                                    Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
                                    Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                                End If
                            End If
                            
                        
                        Case Else
                            'Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)                      '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 1, MTS_ID)               '2016.06.23
'                            Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                    End Select
                                                                    
                                                                    
                                    
                    If HAN_KBN = "2" Then
                        CYU_KBN = "E"
                    
                    
                    End If
                    '�����敪��6�͂Q��
                    If JGYOBU = SENTAKU And SYUKO_SYUSI = "S2" Then
                        
                        If CYU_KBN = "6" Then
                            CYU_KBN = "2"
                        
                        
                        End If
                    End If
                    
                    
                    '�����敪�ΏۊO��1�Ƃ���
                    If CYU_KBN = "1" Or _
                        CYU_KBN = "2" Or _
                        CYU_KBN = "3" Or _
                        CYU_KBN = "E" Then
                    Else
            
                        CYU_KBN = "1"
                    End If
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    
                    
                                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                    
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, JGYOBA)
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, DATA_KBN)
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, TORI_KBN)
                    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, KAIKEI_JGYOBA)
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, SHISAN_JGYOBA)
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                    Call UniCode_Conv(Y_SYUREC.SURYO, SURYO)
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, SYUKO_SYUSI)
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, HOJYO_SYUSI)
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
                    Call UniCode_Conv(Y_SYUREC.TANKA, TANKA)
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, ODER_NO)
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, ITEM_NO)
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, ODER_NO_R)
                    '20011.10.31
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, Left(KOSO_KEITAI, 10))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
                    
                    
                    If TANA_SPACE Then
                    
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                        
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, TANABAN1)
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, TANABAN2)
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, TANABAN3)
                    End If
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, MUKE_NAME)
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN)
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_NAME)
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, ORIGIN1)
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, ORIGIN2)
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, BIKOU2)
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, HAN_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CHOKU_KBN)
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, UNIT_ID_NO)
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, ZAIKO_HIKIATE)
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, GOKON_KANRI_NO)
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, JYUCHU_ZAN)
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, KYOKYU_KBN)
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, SHOHIN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, S_SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, S_HOJYO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, BIKOU1)
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, CHOHA_KBN)
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, JYU_HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, HIN_CHANGE_KBN)
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, MODULE_EXCHANGE)
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, ZAIKO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, ZAN_SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, ZAN_HOJYO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, NOUKI_YMD)
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, SERVICE_KANRI_NO)
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, KISHU_CODE)
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, ENVIRONMENT_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, KEPIN_KAIJYO)
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    
                    
                    If Upd_com = BtOpInsert Then    '2008.02.23
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
                        
                        
                        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")              '2006.09.07
                        Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")            '2006.09.07
                        
                        Call UniCode_Conv(Y_SYUREC.H_IO_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.H_SOKO_CODE, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
                    End If
            
                    Loop_Cnt = 0
    
                    Do
'                        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        sts = BTRV(Upd_com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    GoTo Abort_Tran
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            
                            Case BtErrDEAD_LOCK
                                GoTo Abort_Tran
                            Case Else
'                                Call File_Error(sts, BtOpInsert, "�o�ח\��")
                                'Call File_Error(sts, Upd_com, "�o�ח\��", 0)                           '2016.06.23
                                Call File_Error(sts, BtOpGetEqual, "�o�ח\��", 1, Y_SYU_ID)             '2016.06.23
                                GoTo Abort_Tran
                        End Select
                    Loop
    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
        
        
        
        
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                    DoEvents
        
        
        
        
        
        
        
                    If SYUKA_LOG_ON Then
                        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
                    End If
                
                
                
                    '���אU��   2009.04.14
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G11" Or Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                        
                        
                        
                        
                        
                        ' "00023410"��ǉ� 2009.06.25 "00021397"��ǉ��@2012.04.06
                        If (Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00023510" Or Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00023410" Or Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00021397") Then
                            If Trim(StrConv(Y_SYUREC.DATA_KBN, vbUnicode)) = "7" Then
                                If Trim(StrConv(Y_SYUREC.TORI_KBN, vbUnicode)) = "19" Then
'                                    If Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "00" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "01" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "07" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "08" Then
                                                                    '���׃f�[�^�쐬
                                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                                        Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, Right(ID_NO, 9))
                                
                                
                                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_NO)
                                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(SURYO), "0000000"))
                                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, SYUKA_YMD)
                                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, SYUKA_YMD)
                                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                                
                                
                                        Last_Proc_F = True              '���������ް��폜�����@���s�L��
                                
                                
                                        '���������ް��X�V
                                        Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
                                        Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                                        Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_NO)
                            
                                        WK_Y_QTY = CLng(SURYO)
                            
                            
                                        Loop_Cnt = 0
                                        
                                        Do
                                            sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                                        WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                                        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                                                
                                                        Loop_Cnt = 0
                                                        
                                                        Do
                                                        
                                                            sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                                            Select Case sts
                                                                Case BtNoErr
                                                                    Exit Do
                                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                                    Beep
'                                                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                                                    If ans = vbCancel Then
'                                                                        Exit Function
'                                                                    End If
                                                                
                                                                
                                                                    Loop_Cnt = Loop_Cnt + 1
                                                                    If Loop_Cnt > 5 Then
                                                                        GoTo Abort_Tran
                                                                    End If
                                                                
                                                                    DoEvents
                                                                    Sleep (500)
                                                                
                                                                
                                                                
                                                                Case BtErrDEAD_LOCK
                                                                    'Exit Function          '2015.11.19
                                                                    GoTo Abort_Tran         '2015.11.19
                                                                Case Else
                                                                    'Call File_Error(sts, BtOpUpdate, "���������ް�", 0)                        '2016.06.23
                                                                    Call File_Error(sts, BtOpUpdate, "���������ް�", 1, J_NYU_ID)               '2016.06.23
                                                                    'Exit Function          '2015.11.19
                                                                    GoTo Abort_Tran         '2015.11.19
                                                            End Select
                                                        
                                                        Loop
                                                        WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    Else
                                                        
                                                        Loop_Cnt = 0
                                                        Do
                                                            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                                            Select Case sts
                                                                Case BtNoErr
                                                                    Exit Do
                                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                                    Beep
'                                                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                                                    If ans = vbCancel Then
'                                                                        Exit Function
'                                                                    End If
                                                                
                                                                
                                                                    Loop_Cnt = Loop_Cnt + 1
                                                                    If Loop_Cnt > 5 Then
                                                                        GoTo Abort_Tran
                                                                    End If
                                                                
                                                                    DoEvents
                                                                    Sleep (500)
                                                                
                                                                
                                                                Case BtErrDEAD_LOCK
                                                                    
                                                                    'Exit Function      '2015.11.19
                                                                    GoTo Abort_Tran     '2015.11.19
                                                            Case Else
                                                                    'Call File_Error(sts, BtOpDelete, "���������ް�", 0)                        '2016.06.23
                                                                    Call File_Error(sts, BtOpDelete, "���������ް�", 1, J_NYU_ID)               '2016.06.23
                                                                    'Exit Function
                                                                    GoTo Abort_Tran     '2015.11.19
                                                            End Select
                                                        Loop
                                                        WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                                                    End If
                                            
                                                    Exit Do
                                                Case BtErrKeyNotFound
                                                    WK_E_QTY = 0
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                    Beep
'                                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                                    If ans = vbCancel Then
'                                                        Exit Function
'                                                   End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                               
                                                
                                                Case BtErrDEAD_LOCK
                                                    'Exit Function      '2015.11.19
                                                    GoTo Abort_Tran     '2015.11.19
                                                Case Else
                                                    'Call File_Error(sts, BtOpGetEqual, "���������ް�", 0)                  '2016.06.23
                                                    Call File_Error(sts, BtOpGetEqual, "���������ް�", 1, J_NYU_ID)         '2016.06.23
                                                    'Exit Function      '2015.11.19
                                                    GoTo Abort_Tran     '2015.11.19
                                            End Select
                                        Loop
                                                            '��s���א��i���׎��ѐ��j
                                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
                                
                                                            '�\�Z�P�ʌ�
                                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(Y_SYUREC.KEY_MUKE_CODE, vbUnicode))
                                                            '�\�Z�P�ʐ�
                                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, "")
                                                            '�W���I��
                                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, "")
                                                            'H�q�� 2006.10.17
                                        Call UniCode_Conv(Y_NYUREC.H_SOKO, StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode))
                        
                                                            '���׃��X�g�o�̓t���O   2007.06.12
                                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")
                        
                                        
                                        
                                
'                If Trim(ORIGIN1) = "" Then
                
'''2011.01.19
'''                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                    
                    
                    
                    
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_NO)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")


                    com = BtOpGetGreaterEqual
                    
                    GENSAN_CNT = 0
                    
                    
                    GENSANKOKU = ""
                    
                    Do
                        DoEvents
                    
                        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                If StrConv(GENSANREC.JGYOBU, vbUnicode) <> StrConv(Y_NYUREC.JGYOBU, vbUnicode) Or _
                                    StrConv(GENSANREC.NAIGAI, vbUnicode) <> StrConv(Y_NYUREC.NAIGAI, vbUnicode) Or _
                                    StrConv(GENSANREC.HIN_GAI, vbUnicode) <> StrConv(Y_NYUREC.HIN_NO, vbUnicode) Then
                                    Exit Do
                                End If
                            
                            
                                GENSAN_CNT = GENSAN_CNT + 1
                                If GENSAN_CNT > 1 Then
                                    GENSANKOKU = ""
                                    Exit Do
                                End If
                            
                            
                                GENSANKOKU = StrConv(GENSANREC.GENSANKOKU, vbUnicode)
                            
                            Case BtErrEOF
                                Exit Do
                            
                            Case Else
                                'Call File_Error(sts, com, "���Y���}�X�^", 0)               '2016.06.23
                                Call File_Error(sts, com, "���Y���}�X�^", 1, GENSAN_ID)     '2016.06.23
                                'Exit Function      '2015.11.19
                                GoTo Abort_Tran     '2015.11.19
                        End Select
                    
                        com = BtOpGetNext
                                    
                    Loop
                    
                    
                    
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, "")
                    If GENSAN_CNT = 1 Then
                    
                        Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)
                    End If
                    
'''2011.01.19
                    
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))



'               Else
'
'                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, ORIGIN1)
'                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, "")
'                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, "")
'
'                End If
                
                
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, "")
                                        
                Call UniCode_Conv(Y_NYUREC.ID_NO2, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, "")
                                        
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "0")
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "8")
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "0")
                
                
                
                WORK_SOKO = "90"
                
                Select Case JGYOBU
                    Case AIRCON, SENTAKU
                    
                    
                    Case Else
                        
                        WORK_SOKO = "81"
                        
                        '2009.06.25
                        If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                            WORK_SOKO = "80"
                        End If
                        '2009.06.25
                End Select
                
                
                
                
                
                
                
                
                
                
                
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, WORK_SOKO & "010101")
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, Format(WK_E_QTY, "00000000"))
                                        
                                        
                                        
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                                        
                                        
                                        
                                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                                        
                                        
                                        Loop_Cnt = 0
                                        
                                        
                                        Do
                                            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                    Beep
'                                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                                    If ans = vbCancel Then
'                                                        Exit Function
'                                                    End If
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                               
                                                
                                                Case Else
                                                    
                                                    
                                                    
                    '2010.05.24
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
                        Write #DUP_SYUKANo, , , "�o�׎捞�ُ݈탊�X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "�o�ד�", "�`�[��", "�x���溰��", "�q��/�r�r����", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"                  '2015.11.19
                        Write #DUP_SYUKANo, "�G���[���e", "�o�ד�", "�`�[��", "�o�א溰��", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"      '2015.11.19
                        Fast_Flg = False
                    End If
                    
                    Write #DUP_SYUKANo, "[���אU�փf�[�^]",
                    Write #DUP_SYUKANo, SYUKA_YMD,
                    Write #DUP_SYUKANo, DEN_NO,
                    Write #DUP_SYUKANo, MUKE_CODE,
                    Write #DUP_SYUKANo, MUKE_NAME,
                    Write #DUP_SYUKANo, CYU_KBN,
                    Write #DUP_SYUKANo, CYU_KBN_NAME,
                    Write #DUP_SYUKANo, HIN_NO,
                    Write #DUP_SYUKANo, SURYO,
                    Write #DUP_SYUKANo, ID_NO

'                    Write #DUP_SYUKANo, "[���אU�փf�[�^]"

'                    Call File_Error(sts, BtOpInsert, "���ח\��")
'                    Exit Function
                    Call File_Error(sts, BtOpInsert, "���ח\��", 0)
                    GoTo Loop_Proc
                    '2010.05.24
                                                    
                                                    
                                            End Select
                                        Loop
                                    
                        '------------ 2005.12.30
                                        WORK_SOKO = "90"
                                        
                                        Select Case JGYOBU
                                            Case AIRCON, SENTAKU
                                                Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
                                                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                                Select Case sts
                                                    Case BtNoErr
                                                    Case Else
                                                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                                        'Exit Function      '2015.11.19
                                                        GoTo Abort_Tran     '2015.11.19
                                                End Select
                                
                                                If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
                                
                                                    SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    MI_QTY = 0
                                                Else
                                                
                                                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                        MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                        SUMI_QTY = 0
                                                    Else
                                                        SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                        MI_QTY = 0
                                                    End If
                                                End If
                                                
                        '------------ 2005.12.30
                                            
                                            
                                            Case Else
                                                
                                                WORK_SOKO = "81"
                                                
                                                '2009.06.25
                                                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                                                    WORK_SOKO = "80"
                                                End If
                                                '2009.06.25
                                                
                                                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                    MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    SUMI_QTY = 0
                                                Else
                                                    SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    MI_QTY = 0
                                                End If
                                        End Select
                                        
                                
                        '                Wk_SOKO = KASO_NYUKA_SOKO
                        '                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
                        '                    Wk_SOKO = KASO_SMODOSHI_SOKO
                        '
                        '                End If
                                
                                        '���א��ō݌Ƀf�[�^�X�V�i�{�j
                                        If Nyuko_Update_Proc(JGYOBU, _
                                                            Soko_T(i, j).NAIGAI, _
                                                            HIN_NO, _
                                                            StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                                            (WORK_SOKO & "01" & "01" & "01"), _
                                                            YOIN_TU_NYUKA, _
                                                            SUMI_QTY, MI_QTY, _
                                                            WS_NO, WS_NO, 5, _
                                                            StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode) & " �`��:" & DEN_NO, , , , MENU_NO, , KAMOKU_FURIKAE, _
                                                            StrConv(Y_NYUREC.GENSANKOKU, vbUnicode), _
                                                            StrConv(Y_NYUREC.SHIIRE_WORK_CENTER, vbUnicode), _
                                                            StrConv(Y_NYUREC.ID_NO2, vbUnicode), _
                                                            StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode), _
                                                            StrConv(Y_NYUREC.YOSAN_TO, vbUnicode), Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))) Then
                                            'Exit Function      '2015.11.19
                                            GoTo Abort_Tran     '2015.11.19
                                    
                                        End If
                                    
                                        '�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
                                        If WK_E_QTY <> 0 Then
                                        '�݌Ƀf�[�^LOCK
                                            If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                                                JGYOBU, _
                                                                Soko_T(i, j).NAIGAI, _
                                                                HIN_NO, _
                                                                WS_NO, , , 5) Then
                                                'Exit Function      '2015.11.19
                                                GoTo Abort_Tran     '2015.11.19
                            
                                            End If
                                
                                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                MI_QTY = WK_E_QTY
                                            Else
                                                SUMI_QTY = WK_E_QTY
                                            End If
                                    
                                    
                                            If Syuko_Update_Proc(JGYOBU, _
                                                                Soko_T(i, j).NAIGAI, _
                                                                HIN_NO, _
                                                                StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                                                (WORK_SOKO & "01" & "01" & "01"), _
                                                                YOIN_MAE_SOUSAI, _
                                                                SUMI_QTY, MI_QTY, 0, _
                                                                WS_NO, WS_NO, 5) Then
                                                'Exit Function      '2015.11.19
                                                GoTo Abort_Tran     '2015.11.19
                                
                                            End If
                                    
                                    
                                    
                                    
                                    
                                    
                                        End If
 
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
'                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
Loop_Proc:
    
    Loop
    
        
    Close #DUP_SYUKANo
        
    Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    
    
    If Fast_Flg Then
        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
        Write #DUP_SYUKANo, , , "�o�׎捞�ُ݈탊�X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "�o�ד�", "�`�[��", "�x���溰��", "�q��/�r�r����", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"                  '2015.11.19
        Write #DUP_SYUKANo, "�G���[���e", "�o�ד�", "�`�[��", "�o�א溰��", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"      '2015.11.19
        Fast_Flg = False
    End If


    Write #DUP_SYUKANo, "��̧�ُo�ُ͈큄�@sts=" & sts,
    Write #DUP_SYUKANo, SYUKA_YMD,
    Write #DUP_SYUKANo, DEN_NO,
    Write #DUP_SYUKANo, MUKE_CODE,
    Write #DUP_SYUKANo, MUKE_NAME,
    Write #DUP_SYUKANo, CYU_KBN,
    Write #DUP_SYUKANo, CYU_KBN_NAME,
    Write #DUP_SYUKANo, HIN_NO,
    Write #DUP_SYUKANo, SURYO,
    Write #DUP_SYUKANo, ID_NO
    
    Close #DUP_SYUKANo          '2015.11.19
    
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
                                            '�w�b�_�[����i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "�������@�i�ԕύX���X�g�@������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '���׃w�b�_���
    Printer.Print "------- �i�ԁi�O���j-------";
    Printer.Print Tab(30);
    Printer.Print "------- �i�ԁi�����j-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "�`�[���t";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "���o�ɋ�";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "���o��";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "�q";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "�o�א�"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '���׈���i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = 0 To LBox_Hin.ListCount - 1
        
'        Ldata = LBox_Hin.List(i)
'
'                                        '�w�b�_�[�R���g���[��
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If
'
'                                        '���׈��
'        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '���ƕ��C÷�ć��C�����O�@���O'
'
'        Printer.Print Tab(MGN_L);
'        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�O���j
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '�O���i�ԕύX�H
'            Printer.Print Tab(MGN_L + 15);
'            Printer.Print Work;                                     '�}�X�^�i�ԁi�O���j
'        End If
'
'        Printer.Print Tab(MGN_L + 30);
'        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�����j
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "0" Then                               '�����i�ԕύX�H
'            Printer.Print Tab(MGN_L + 44);
'            Printer.Print Work;                                     '�}�X�^�i�ԁi�����j
'        End If
'
'        Printer.Print Tab(MGN_L + 58);                              '�`�[���t
'        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);
'
'        Printer.Print Tab(MGN_L + 69);                              '���o�ɋ敪
'        wk_IO = ChrCut(Ldata, 1)
'        Select Case wk_IO
'            Case IO_KBN_URI
'                Printer.Print wk_IO & " " & (IO_KBN_0);
'            Case IO_KBN_NYU
'                Printer.Print wk_IO & " " & (IO_KBN_1);
'            Case IO_KBN_SYU
'                Printer.Print wk_IO & " " & (IO_KBN_2);
'            Case IO_KBN_ZAT
'                Printer.Print wk_IO & " " & (IO_KBN_3);
'            Case Else
'                Printer.Print wk_IO;
'        End Select
'
'        Printer.Print Tab(MGN_L + 78);
'        Printer.Print ChrCut(Ldata, 6);                             '�`�[��
'
'        Printer.Print Tab(MGN_L + 85);                              '���o�ɐ�
'        Wqty = CLng(ChrCut(Ldata, 6))
'
'
'        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
'
'        Printer.Print Work;
'
'        Printer.Print Tab(MGN_L + 93);
'        Printer.Print ChrCut(Ldata, 2);                             '�q�ɋ敪�iνāj
'
'        Printer.Print Tab(MGN_L + 96);                              '�����敪
'        Select Case Left(Ldata, 1)
'            Case CYU_KBN_TUK
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
'            Case CYU_KBN_SPO
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
'            Case CYU_KBN_HJU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
'            Case CYU_KBN_BOU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
'            Case Else
'                Printer.Print ChrCut(Ldata, 1);
'        End Select
'
'        Printer.Print Tab(MGN_L + 103);
'        Printer.Print ChrCut(Ldata, 5);                             '�x����^�o�א�7
'
'        Printer.Print Tab(MGN_L + 110);                             '�ύX���b�Z�[�W
'        Select Case Left(Ldata, 1)
'            Case "0"
'                Printer.Print "�����ύX Ͻ��i�ԓ���";
'            Case "1"
'                Printer.Print "�O���ύX Ͻ��i�ԓ���";
'            Case "2"
'                Printer.Print "�݌ɗL�I�O���ύX�s��";
'        End Select
        
        Printer.Print LBox_Hin.List(i)
        
        Call LOG_OUT(LOG_F, LBox_Hin.List(i))
        
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    End If

End Sub

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String

    Call NG_File_Make_Proc

    '---------------------------------------------  ���ƕ������C�����[�v
    For i = 0 To UBound(JGYOBU_T)
        
        In_Cnt = 0
        Out_Cnt = 0

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents

        FileNo = FreeFile
        FileName = HS_IN_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

        On Error GoTo Error_Proc

        Open FileName For Input As #FileNo

        On Error GoTo 0


        If Nyuka_Update_Proc(JGYOBU_T(i).CODE) Then     '���ח\��f�[�^�X�V����

            Unload Me

        End If


        Close #FileNo

        '-----------------------------------------------
    
        FileNo = FreeFile
        FileName = HS_OUT_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
        
        On Error GoTo Error_Proc
        
'        Open fileName For Input As #FileNo
        Open FileName For Binary As #FileNo
    
        On Error GoTo 0
    
    
        If Syuka_Update_Proc(JGYOBU_T(i).CODE) Then  '�o�ח\��f�[�^�X�V����

            Unload Me
        End If
    
    
        Close #FileNo
    
    
    
    
    Next i

    If Not Err_FLg Then
        Call NG_File_Kill_Proc
    End If

    Unload Me

Error_Proc:

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer


Dim sBuffer     As String * 255
Dim com         As String
    
Dim Max_Soko    As Integer
    
Dim GENSAN_WK   As Variant              '2016.12.28
    
    
    If App.PrevInstance Then
 '       Beep
 '       MsgBox "����v���O�������s���ł��B"
        End
    End If


    F1020101.Caption = F1020101.Caption & Last_Update_Day


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
                                 '�o�׏d���f�[�^�o�̓t�@�C������荞��
    If GetIni("FILE", "DUP_SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "�o�׏d���f�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    DUP_SYUKA_DATA = Trim(c)
                               
    If JGYOB_TB_Set() Then      '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
    '---------------------------------------------- *
    '    SYS.INI -- > F102010.INI
    '   2015.03.04
    '---------------------------------------------- *
        
                                
                                
                                
                                '�q�ɍő吔����荞��
                                
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                    
                                
                                
                                
                                '�݌Ɏ�荞�ݗp�e�[�u���쐬
    ReDim Soko_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '�q�ɏ���荞��
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '�L���q�Ɋl��
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "�q�ɏ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
    
            If Trim(c) = "**" Then  '�q�Ɏw��I��
                Exit Do
            End If
    
    
'            ReDim Preserve JSOKO_T(i).JSOKO_T(0 To j)
            Soko_T(i, j).HS_SOKO = Trim(c)
                                '�����O���l��
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "�����O���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
            
            Soko_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                
                                
                                '���Ƀf�[�^�t�@�C�����̊l��
    If GetIni("FILE", "HS_SIJ_IN", "SYS", c) Then
        Beep
        MsgBox "���Ƀf�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    HS_IN_SIJ = Trim(c)
                                
                                '�o�Ƀf�[�^�t�@�C�����̊l��
    If GetIni("FILE", "HS_SIJ_OUT", "SYS", c) Then
        Beep
        MsgBox "�o�Ƀf�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    HS_OUT_SIJ = Trim(c)
                                
                                
                                
                                '�u�ʏ���ׁv�v���̊l��
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Beep
        MsgBox "�u�ʏ���ׁv�v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                '�u�O�ؑ��E�v�v���̊l��
    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
        Beep
        MsgBox "�u�O�ؑ��E�v�v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_MAE_SOUSAI = Trim(c)
                                
                                '���z���בq�ɂ̊l��
    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "���z���בq�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KASO_NYUKA_SOKO = Trim(c)
                                '���z�x���߂��q�ɂ̊l��
    If GetIni("SYSTEM", "KASO_SMODOSHI ", "SYS", c) Then
        Beep
        MsgBox "���z���בq�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KASO_SMODOSHI_SOKO = Trim(c)
                                
                                
                                '���̑�������i�����j�̊l��
    If GetIni(App.EXEName, "ETC_MTS_NAI", App.EXEName, c) Then
        Beep
        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ETC_MTS_NAI = Trim(c)
                                
                                '���̑�������i�C�O�j�̊l��
    If GetIni(App.EXEName, "ETC_MTS_GAI", App.EXEName, c) Then
        Beep
        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ETC_MTS_GAI = Trim(c)
                                
                                
                                
'---------------------------------------------- '�ƭ����̊l��    2007.11.06
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = RTrim(c)
    End If
                                
                                
                                '����@��p
    If GetIni(App.EXEName, "CENTER", App.EXEName, c) Then
        MyCenter = "O"
    Else
        MyCenter = Trim(c)
    End If
                                
                                
                                
'---------------------------------------------- '�ȖڐU�ւ̗v�� 2009.06.26
    KAMOKU_FURIKAE = YOIN_TU_NYUKA
    If GetIni(App.EXEName, "KAMOKU_FURIKAE", App.EXEName, c) Then
    Else
        KAMOKU_FURIKAE = RTrim(c)
    End If
                                
                                
                                '���̑�������̊l��
'    If GetIni(App.EXEName, "ETC_SS_NAI", "SYS", c) Then
'        Beep
'        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
'        End
'    End If
'    ETC_SS_NAI = Trim(c)
                                
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)


'---------------------------------------------- '�I�Ԑݒ���̊l��    2009.03.07
    If GetIni(App.EXEName, "TANA_SPACE", App.EXEName, c) Then
        TANA_SPACE = False
    Else
        If Trim(c) = "1" Then
            TANA_SPACE = True
        Else
            TANA_SPACE = False
        End If
    End If



                                '�i���ɂ�鏜�O 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                '�i���ɂ�鏜�O 2011.07.04

'---------------------------------------------- '���i����̫��    2012.12.20
    If GetIni(App.EXEName, "GOODS_F", App.EXEName, c) Then
        GOODS_F = "0"
    Else
        If Trim(c) = "1" Then
            GOODS_F = "1"
        Else
            GOODS_F = "0"
        End If
    End If
'---------------------------------------------- '���i����̫��    2012.12.20
    
'---------------------------------------------- '�X�V�Ώی��Y�� 2016.12.28
    If GetIni(App.EXEName, "GENSAN", App.EXEName, c) Then
        c = "*"
    End If
    GENSAN_WK = Split(Trim(c), ",", -1)

    For i = 0 To UBound(GENSAN_WK)
    
        ReDim Preserve GENSAN_T(0 To i)
        GENSAN_T(i) = GENSAN_WK(i)
    
    
    Next i
'---------------------------------------------- '�X�V�Ώی��Y�� 2016.12.28






                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�i�X�V�p���[�N�j�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���Y���}�X�^�n�o�d�m   2010.07.08
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                'PN�}�X�^�n�o�d�m   2010.09.01
    If PN_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    If Country_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�ƍ��p���ח\��n�o�d�m 2007.06.15
    If Y_GLICS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���������ް��n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
'���ԃ}�X�^�n�o�d�m ################################################################## 2005/05/16 Add ��
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
'#################################################################################### 2005/05/16 Add ��
                                
                                
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1020101.FontName
        .Size = F1020101.FontSize
    End With
    Set Printer.Font = NormalFont

    Last_Proc_F = False         '���������ް��폜�����@���s�L���t���O�N���A


    '�d������l��       2005.12.30
    i = -1
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    com = BtOpGetGreater
    SHIMUKE_Flg = False
    
    Do
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SHIMUKE_T(0 To i)
            
            
                SHIMUKE_Flg = True
            
                SHIMUKE_T(i).SHIMUKE_CODE = StrConv(P_CODEREC.C_Code, vbUnicode)
                SHIMUKE_T(i).JGYOBU = StrConv(P_CODEREC.OPTION1, vbUnicode)
                SHIMUKE_T(i).NAIGAI = StrConv(P_CODEREC.OPTION2, vbUnicode)
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
    
    
    '�d������l��       2005.12.30


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '���������ް��폜�����@���s�L��H
'        Call Last_Proc
'    End If

                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")              '2016.06.23
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^", 1, SOKO_ID)   '2016.06.23
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")              '2016.06.23
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^", 1, ITEM_ID)   '2016.06.23
        End If
    End If
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�\���}�X�^")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "�\���}�X�^", 1, P_COMPO_ID)    '2016.06.23
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^", ITEM_ID)          '2016.06.23
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")            '2016.06.23
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^", 1, MTS_ID)  '2016.06.23
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�v���}�X�^")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "�v���}�X�^", 1, YOIN_ID)         '2016.06.23
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�S���҃}�X�^")                '2016.06.23
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^", 1, TANTO_ID)      '2016.06.23
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")                '2016.06.23
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^", 1, P_CODE_ID)     '2016.06.23
        End If
    End If
                                            '���Y���}�X�^�b�k�n�r�d 2010.07.08
    sts = BTRV(BtOpClose, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "���Y���}�X�^")                '2016.06.23
            Call File_Error(sts, BtOpClose, "���Y���}�X�^", 1, GENSAN_ID)     '2016.06.23
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")                      '2016.06.23
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^", 1, ZAIKO_ID)          '2016.06.23
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�݌Ɉړ���")                      '2016.06.23
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���", 1, IDO_ID)            '2016.06.23
        End If
    End If
                                            '���ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "���ח\��")                        '2016.06.23
            Call File_Error(sts, BtOpClose, "���ח\��", 1, Y_NYU_ID)            '2016.06.23
        End If
    End If
                                            '�ƍ��p���ח\��b�k�n�r�d   2007.06.16
    sts = BTRV(BtOpClose, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�ƍ��p���ח\��")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "�ƍ��p���ח\��", 1, Y_GLICS_ID)    '2016.06.23
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "�o�ח\��")                        '2016.06.23
            Call File_Error(sts, BtOpClose, "�o�ח\��", 1, Y_SYU_ID)            '2016.06.23
        End If
    End If
                                            '���������ް��b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "���������ް�")                    '2016.06.23
            Call File_Error(sts, BtOpClose, "���������ް�", J_NYU_ID)           '2016.06.23
        End If
    End If
                                            '�a���������������Z�b�g
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020101 = Nothing

    End
End Sub

Private Function Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   �u�i�ڃ}�X�^�v�`�F�b�N���X�V����
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    
Dim Loop_Cnt    As Integer          '2011.01.19
    
    
    Item_Check_Proc = True

    HIN_CHANGE = 0
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                


                    If Mode = In_Mode Then          '�Γ��i�ԕύX�̃`�F�b�N
    '                Else
    
                        If Len(Trim(HIN_NAI)) <> 0 Then
                            If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                                HIN_CHANGE = NAI_CHANGE
                                BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '�����i�ԓ���ւ�
                                Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                            
                            
                                '�S���ҍX�V�ǉ� 2009.11.11
                                    
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            
                            
                            
                            
                            
                            
                            End If
                        End If
                    
                    
                    
                                        
                        '---------------    2010.07.08  ��
                        '���Y������ւ��`�F�b�N
                        If Len(Trim(GENSANKOKU)) <> 0 Or Len(Trim(GEN_GENSANKOKU)) <> 0 Or Len(Trim(SHIIRE_WORK_CENTER)) <> 0 Then
    '                        If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> Trim(GENSANKOKU) Then
                                '���Y������ւ�
                                
                            
                                If Trim(GENSANKOKU) <> "" Then
                                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                                Else
                                    Debug.Print
                                End If
                                
                                
    '                            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
    '
    '                                Call UniCode_Conv(ITEMREC.GENSANKOKU, GENSANKOKU)
    '
    '
    '                            End If
                                
                                
                                Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                                Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                            
                            
                                '�S���ҍX�V�ǉ� 2009.11.11
                                    
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                
                                
                            
    '                        End If
                        End If
                        '---------------    2010.07.08  ��
                    
                    
                        '---------------    2010.07.27  ��
                        '���敪�`�F�b�N
                        If Len(Trim(KANKYO_KBN)) <> 0 Or Len(Trim(KANKYO_KBN_ST)) <> 0 Or Len(Trim(KANKYO_KBN_SURYO)) <> 0 Then
                            
                            
                            
                            If Val(KANKYO_KBN_SURYO) = 0 Then
                            Else
                            
                                
                                '���敪����ւ�
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                                    
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                                
                                
                                '�S���ҍX�V�ǉ� 2009.11.11
                                        
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            End If
                                
                            
                        End If
                        '---------------    2010.07.08  ��
                    
                    
                    
                    End If
                
                
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '���ƕ�
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '�����O
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '�i�ԁi�O���j
                                                            '�i��
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '�V�K�i�Ԏ�*���Z�b�g2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '�i�ԁi�����j
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '�i�ԁi�����j
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '���l �z�X�g�q��
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '���l �z�X�g�I��
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '���ރR�[�h
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '��[�_
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '�����Ϗo�א�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          '�T���v����
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '�ŏI���ד��t
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '�r���t���O
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '�g�p�q�@�h�c
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '�g�p���v���O����
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '�ŏI�ƍ����t
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '�ŏI�ƍ����݌ɐ�
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '�������ƕ�
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '������l
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '������萔
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Jan�R�[�h
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '�i�ԓǂݑւ�
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '���i���L��
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '������
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          '��د���I��1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '�Ɩ��Ǘ��@ �d���敪
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           �̔��敪
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           ���x�P��
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           �g�����i
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           �W���e�������P���@9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           �W���e�������ݒ��
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           �W���e�������P��  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           �W���e�������ݒ��
                                            
                                            
                                                                        '           �d������
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             '����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '�d���P��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ۯĐ�
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ذ�����
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           �O���݌ɋ��z
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           ���ދ敪
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ���x���\�t
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��

'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '���i����   �i��
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           ���l
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           ��ЃR�[�h
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           �@��(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           �@��(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           ��
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           �v���X�`�b�N
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           ���i(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           ���i(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           ���i(3)
                
                
                                                                '           ���i(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           �K�p�@������
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           ��������
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           �K�p�@����l
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           ��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           ���l�R
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           ���ƕ��R�[�h
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           ���萔
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           �I��(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           �I��(2)
                
                
                
'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '���P�^�S���҃R�[�h
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           �O���݌ɐ���
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           �ŏI�o�א�
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS�݌�(S2) �܈�p
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS�݌�(P2) �܈�p
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '���`��
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               '�Ưĕ��i�敪
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '�����������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '�C�O�������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '�W���P��   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      '�Ưĕ��i�敪
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '�����������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '�C�O�������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '�W���P��   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '�ŏI�d����R�[�h   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '�ŏI�d���P��       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'Ұ������           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'Ұ������           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '�č���ϰ�          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '�ː�               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '�`��               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '�ގ�               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              '����ްف@����      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 '����ްٻ��ށiW�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 '����ްٻ��ށiD�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 '����ްٻ��ށiH�j   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '�������^���Ȃ�   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '���i���@�H��       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '���i���@�P���ݒ�� 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '���i���@���ތ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '���i���@���ޔ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '���i���@�P���ݒ�� 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '�A�����@�o���׸�   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '�g�p�e�[�v���     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '�g�p�e�[�v��       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '�I�ԃ}�[�N         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '�����P���@����     2008.04.15
    
                '2010.07.08 ��
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '���Y��             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '���Y��
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '���Y��
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 ��
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            'Call File_Error(sts, BtOpGetEqual, "COUNTRY")                  '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY", 1, Country_ID)    '2016.06.23
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '�O���P�� 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC���H�P��9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU���H�P��9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '���Y���b�g         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '�����[�g           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '�W������           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '�P���ݒ�S����     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '�d������           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '�����敪           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            '���x���\�薇��     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '���ތ���     �@    2008.08.20�ǉ�
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '��������           2008.08.20�ǉ�
         

'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                
                                
                
                
                
                '��2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '��2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '��ԋ敪           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '���iү���� 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '�����e 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '�ǉ��S����
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '�ǉ�����
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '�X�V�S����
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '�X�V����
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
           
            
            Case BtErrDEAD_LOCK
                Exit Function
                        
            Case Else
                'Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)            '2016.06.23
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 1, ITEM_ID)    '2016.06.23
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                'Call File_Error(sts, com, "�i�ڃ}�X�^", 0)
                Call File_Error(sts, com, "�i�ڃ}�X�^", 1, ITEM_ID)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '�\���}�X�^�̒ǉ�       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '�d�����溰��
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '���ƕ�
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '�����O
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '�i��
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            '�ް��敪
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '�ǔ�
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '��{�N���X
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '���l
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '�X�V�S����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '�X�V����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            Case BtErrDEAD_LOCK
                                Exit Function
                            Case Else
                                'Call File_Error(sts, BtOpInsert, "�\���}�X�^", 0)              '2016.06.23
                                Call File_Error(sts, BtOpInsert, "�\���}�X�^", 1, P_COMPO_ID)   '2016.06.23
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    Item_Check_Proc = False

End Function

Sub NG_File_Make_Proc()
'----------------------------------------------------------------------------
'                   �ُ�I���t�@�C���o�͏���
'----------------------------------------------------------------------------
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "�ُ�I���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    
    stream = FreeFile
    Open NG_FILE For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Print #stream, Buf
    Close stream
End Sub
Sub NG_File_Kill_Proc()
'----------------------------------------------------------------------------
'                   �ُ�I���t�@�C���폜����    2008.10.07
'----------------------------------------------------------------------------
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "�ُ�I���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    On Error GoTo Err_Proc
    Kill (NG_FILE)

Err_Proc:

End Sub


Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String, _
                                                                        TEXT_NO As String, _
                                                                        JGYOBU_Code As String, _
                                                                        CYOK_KBN As String, _
                                                                        DEN_DT As String, _
                                                                        IO_KBN As String, _
                                                                        PM_KBN As String, _
                                                                        DEN_SYU As String, _
                                                                        DEN_NO As String, _
                                                                        CYU_KBN As String, _
                                                                        HIN_GAI As String, _
                                                                        HIN_NAI As String, _
                                                                        HIN_NAME As String, _
                                                                        YOTEI_QTY As String, _
                                                                        YOSAN_FROM As String, _
                                                                        YOSAN_TO As String, _
                                                                        HOST_SOKO As String, _
                                                                        HOST_TANA As String, _
                                                                        SYUK_CODE As String, _
                                                                        SYUK_NAME As String, _
                                                                        GENSANKOKU As String, GEN_GENSANKOKU As String, SHIIRE_WORK_CENTER As String, KANKYO_KBN As String, KANKYO_KBN_ST As String, KANKYO_KBN_SURYO As String, ID_NO2 As String, AITESAKI_CODE As String, JYUCHU_YMD As String, SHITEI_NOUKI_YMD As String, MOTO_TEXT_NO As String) As Integer
'----------------------------------------------------------------------------
'           �ƍ��p���ח\��t�@�C���o�͏���
'           2007.06.15
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
        
        
Dim Loop_Cnt    As Integer  '2011.01.19
        
    Y_GLICS_PUT_PROC = True
        
    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, TEXT_NO)


    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
    
    Call UniCode_Conv(Y_GLICSREC.HIN_NO, HIN_GAI)
    Call UniCode_Conv(Y_GLICSREC.DEN_NO, DEN_NO)
    
    
    '2008.01.10 �}�C�i�X�̑Ή�
    If YOTEI_QTY >= 0 Then
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
    Else
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "000000"))
    End If
    
    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, CYU_KBN)
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, HIN_NAME)
    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
    
    
    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))


                        '��s���א��i���׎��ѐ��j
    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")

                        '�\�Z�P�ʌ�
    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, YOSAN_FROM)
                        '�\�Z�P�ʐ�
    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, YOSAN_TO)
                        '�W���I��
    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, HIN_NAI)
                        'H�q�� 2006.10.17
    Call UniCode_Conv(Y_GLICSREC.H_SOKO, HOST_SOKO)

                        '���׃��X�g�o�̓t���O   2007.06.12
    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
                        '�����敪
    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, CYOK_KBN)
                        '���o�ɋ敪
    Call UniCode_Conv(Y_GLICSREC.IO_KBN, IO_KBN)
                        '�ԍ��敪
    Call UniCode_Conv(Y_GLICSREC.PM_KBN, PM_KBN)
                        '�`�[���
    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, DEN_SYU)
                        '�x����^�o�א�
    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, SYUK_CODE)
                        '�x����^�o�א於
    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, SYUK_NAME)
                        '�}���N����
    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
    
    
    
    
    
    '----------------   2010.07.08 ��
    Call UniCode_Conv(Y_GLICSREC.GENSANKOKU, GENSANKOKU)                    '���Y����
    Call UniCode_Conv(Y_GLICSREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '�����\�����Y����
    Call UniCode_Conv(Y_GLICSREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '���ގd����ܰ�����
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN, KANKYO_KBN)                    '����ދ敪
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_ST, KANKYO_KBN_ST)              '����ދ敪�K�p�J�n
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)        '����ދ敪����
    Call UniCode_Conv(Y_GLICSREC.ID_NO2, ID_NO2)                            'ID_NO
    Call UniCode_Conv(Y_GLICSREC.AITESAKI_CODE, AITESAKI_CODE)              '����溰��
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_YMD, JYUCHU_YMD)                    '�󒍔N����
    Call UniCode_Conv(Y_GLICSREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)        '�w��[���N����
    Call UniCode_Conv(Y_GLICSREC.LIST_OUT_END_F, "")                        '����ؽďo��F
    Call UniCode_Conv(Y_GLICSREC.NYUKO_TANABAN, "")                         '���ɒI��
    Call UniCode_Conv(Y_GLICSREC.MAEGARI_SURYO, "")                         '�O�ؑ��E��
    '----------------   2010.07.08 ��
    
    
    
    '2011.03.23 �������v���O����
    Call UniCode_Conv(Y_GLICSREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
    '2011.03.23 ���e�L�X�g��
    If Trim(MOTO_TEXT_NO) = "" Then
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, "")
    Else
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
    End If
    
    
    
    
    
    
    
    Call UniCode_Conv(Y_GLICSREC.FILLER, "")


    Loop_Cnt = 0

    Do
        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                'Call File_Error(sts, BtOpInsert, "���ח\��", 0)                '2016.06.23
                Call File_Error(sts, BtOpInsert, "���ח\��", 1, Y_GLICS_ID)     '2016.06.23
                Exit Function
        End Select
    Loop

    Y_GLICS_PUT_PROC = False

End Function


Private Function MAEGARI_PROC(JGYOBU As String, HIN_GAI As String, YOTEI_QTY As String) As Integer
'----------------------------------------------------------------------------
'           �ƍ��p���ח\��t�@�C���o�͏���
'           2018.11.15
'----------------------------------------------------------------------------
Dim com             As Integer
Dim wkYOTEI_QTY     As Long
Dim sts             As Integer
        
    MAEGARI_PROC = True
    '���i�O�؏���
    Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_J_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
                
                                '�O�؂��ް��Ǎ���
    sts = BTRV(BtOpGetEqual, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    Select Case sts
        Case BtNoErr
            com = BtOpUpdate
        Case BtErrKeyNotFound
            com = BtOpInsert
            
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���׎��уf�[�^")
            Exit Function
    End Select
    
    If com = BtOpInsert Then
                                '�V�K�ǉ�
                                                '���ƕ�
        Call UniCode_Conv(J_NYUREC.JGYOBU, JGYOBU)
                                                '�����O
        Call UniCode_Conv(J_NYUREC.NAIGAI, NAIGAI_NAI)
                                                '�i�ځi�O���j
        Call UniCode_Conv(J_NYUREC.HIN_GAI, HIN_GAI)
                                                '���ѐ���
        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(Val(YOTEI_QTY), "00000000"))
                                                '�o�^��
        Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))
        
        
        
        Call UniCode_Conv(J_NYUREC.FILLER, "")
    Else
                                                '���ѐ���
        wkYOTEI_QTY = Val(YOTEI_QTY) + Val(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(wkYOTEI_QTY, "00000000"))
    End If
    '*------------------------------------------------------'�O�؂�f�[�^�o��
    sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, com, "���׎��уf�[�^")
            Exit Function
                
    End Select

    MAEGARI_PROC = False


End Function

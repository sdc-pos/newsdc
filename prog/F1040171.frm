VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1040171 
   Caption         =   "[�ً}����]�i�ڃ}�X�^�O�����z����ւ�����"
   ClientHeight    =   6675
   ClientLeft      =   2025
   ClientTop       =   -3510
   ClientWidth     =   8985
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
   ScaleHeight     =   6675
   ScaleWidth      =   8985
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   495
      Index           =   3
      Left            =   5145
      TabIndex        =   6
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�k�n�f"
      Height          =   495
      Index           =   2
      Left            =   1995
      TabIndex        =   5
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���s"
      Height          =   495
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   240
      Width           =   960
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   420
      TabIndex        =   3
      Top             =   1800
      Width           =   8310
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8085
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Q��"
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Top             =   1200
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   4845
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   1
      Left            =   5565
      TabIndex        =   8
      Top             =   5640
      Width           =   1065
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   5640
      Width           =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "��荞�݃f�[�^"
      Height          =   255
      Left            =   630
      TabIndex        =   0
      Top             =   1320
      Width           =   1800
   End
End
Attribute VB_Name = "F1040171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------   '�e�L�X�g��`

Private Const ptxTanto_Code% = 0            '�S���҃R�[�h
Private Const ptxTanto_Name% = 1            '�S���Җ���
Private Const ptxHin_Gai% = 2               '�i��
Private Const ptxHin_Name% = 3              '�i��

Private Const ptxST_SOKO% = 4               '�W���I�ԁ@ �q��
Private Const ptxST_RETU% = 5               '�W���I��   ��
Private Const ptxST_REN% = 6                '�W���I�ԁ@ �A
Private Const ptxST_DAN% = 7                '�W���I�ԁ@ �i

Private Const ptxBEF_SEI_LOT% = 8           '�ύX�O�@   ���b�g��
Private Const ptxBEF_SEI_RATE% = 9          '           �����[�g
Private Const ptxBEF_S_KOUSU% = 10          '           �����[�g
Private Const ptxBEF_S_KOUSU_GENKA% = 11    '           (����)���i���H��
Private Const ptxBEF_S_KOUSU_BAIKA% = 12    '           (����)���i���H��
Private Const ptxBEF_S_SHIZAI_GENKA% = 13   '           (����)����
Private Const ptxBEF_S_SHIZAI_BAIKA% = 14   '           (����)����

Private Const ptxBEF_S_GAISO_TANKA% = 165   '           �O���P��
Private Const ptxBEF_S_PPSC_KAKO_KOSU% = 161 '          PPSC���H�P��
Private Const ptxBEF_S_BU_KAKO_KOSU% = 162  '           BU���H�P��




Private Const ptxBEF_S_KOUSU_SET_DATE% = 15 '          �ݒ��
Private Const ptxBEF_SEI_TANKA_TANTO% = 16  '          �S����
Private Const ptxBEF_SE_TANKA_MEMO% = 17    '          ����

Private Const ptxAFT_SEI_LOT% = 18          '�ύX��@   ���b�g��
Private Const ptxAFT_SEI_RATE% = 19         '           �����[�g
Private Const ptxAFT_S_KOUSU% = 20          '           �H��
Private Const ptxAFT_S_KOUSU_GENKA% = 21    '           (����)���i���H��
Private Const ptxAFT_S_KOUSU_BAIKA% = 22    '           (����)���i���H��
Private Const ptxAFT_S_SHIZAI_GENKA% = 23   '           (����)����
Private Const ptxAFT_S_SHIZAI_BAIKA% = 24   '           (����)����




Private Const ptxAFT_S_GAISO_TANKA% = 166   '           �O���P��
Private Const ptxAFT_S_PPSC_KAKO_KOSU% = 163 '          PPSC���H�P��
Private Const ptxAFT_S_BU_KAKO_KOSU% = 164  '           BU���H�P��


Private Const ptxAFT_S_KOUSU_SET_DATE% = 25 '          �ݒ��
Private Const ptxAFT_SEI_TANKA_TANTO% = 26  '          �S����
Private Const ptxAFT_SE_TANKA_MEMO% = 27    '          ����


Private Const ptxZEN_AVE% = 28              '�����Ϗo�א�   �O�N�x�@����
Private Const ptxZEN_SYUKAQTY04% = 29       '�����Ϗo�א�   �O�N�x�@4��
Private Const ptxZEN_SYUKAQTY05% = 30       '�@                     5��
Private Const ptxZEN_SYUKAQTY06% = 31       '�@                     6��
Private Const ptxZEN_SYUKAQTY07% = 32       '�@                     7��
Private Const ptxZEN_SYUKAQTY08% = 33       '�@                     8��
Private Const ptxZEN_SYUKAQTY09% = 34       '�@                     9��
Private Const ptxZEN_SYUKAQTY10% = 35       '�@                     10��
Private Const ptxZEN_SYUKAQTY11% = 36       '�@                     11��
Private Const ptxZEN_SYUKAQTY12% = 37       '�@                     12��
Private Const ptxZEN_SYUKAQTY01% = 38       '�@                     1��
Private Const ptxZEN_SYUKAQTY02% = 39       '�@                     2��
Private Const ptxZEN_SYUKAQTY03% = 40       '�@                     3��

Private Const ptxTOU_AVE% = 41              '�����Ϗo�א�   ���N�x�@����
Private Const ptxTOU_SYUKAQTY04% = 42       '�����Ϗo�א�   ���N�x�@4��
Private Const ptxTOU_SYUKAQTY05% = 43       '�@                     5��
Private Const ptxTOU_SYUKAQTY06% = 44       '�@                     6��
Private Const ptxTOU_SYUKAQTY07% = 45       '�@                     7��
Private Const ptxTOU_SYUKAQTY08% = 46       '�@                     8��
Private Const ptxTOU_SYUKAQTY09% = 47       '�@                     9��
Private Const ptxTOU_SYUKAQTY10% = 48       '�@                     10��
Private Const ptxTOU_SYUKAQTY11% = 49       '�@                     11��
Private Const ptxTOU_SYUKAQTY12% = 50       '�@                     12��
Private Const ptxTOU_SYUKAQTY01% = 51       '�@                     1��
Private Const ptxTOU_SYUKAQTY02% = 52       '�@                     2��
Private Const ptxTOU_SYUKAQTY03% = 53       '�@                     3��





Private Const ptxBEF_KOUTEI_TANI01% = 54    '�O�H��01�@ �P��
Private Const ptxBEF_KOUTEI_QTY01% = 55     '           ����
Private Const ptxBEF_KOUTEI_KOUSU01% = 56   '           �H��
Private Const ptxBEF_KOUTEI_TANI02% = 57    '�O�H��02�@ �P��
Private Const ptxBEF_KOUTEI_QTY02% = 58     '           ����
Private Const ptxBEF_KOUTEI_KOUSU02% = 59   '           �H��
Private Const ptxBEF_KOUTEI_TANI03% = 60    '�O�H��03�@ �P��
Private Const ptxBEF_KOUTEI_QTY03% = 61     '           ����
Private Const ptxBEF_KOUTEI_KOUSU03% = 62   '           �H��
Private Const ptxBEF_KOUTEI_TANI04% = 63    '�O�H��04�@ �P��
Private Const ptxBEF_KOUTEI_QTY04% = 64     '           ����
Private Const ptxBEF_KOUTEI_KOUSU04% = 65   '           �H��
Private Const ptxBEF_KOUTEI_TANI05% = 66    '�O�H��05�@ �P��
Private Const ptxBEF_KOUTEI_QTY05% = 67     '           ����
Private Const ptxBEF_KOUTEI_KOUSU05% = 68   '           �H��
Private Const ptxBEF_KOUTEI_TANI06% = 69    '�O�H��06�@ �P��
Private Const ptxBEF_KOUTEI_QTY06% = 70     '           ����
Private Const ptxBEF_KOUTEI_KOUSU06% = 71   '           �H��
Private Const ptxBEF_KOUTEI_TANI07% = 72    '�O�H��07�@ �P��
Private Const ptxBEF_KOUTEI_QTY07% = 73     '           ����
Private Const ptxBEF_KOUTEI_KOUSU07% = 74   '           �H��
Private Const ptxBEF_KOUTEI_TANI08% = 75    '�O�H��08�@ �P��
Private Const ptxBEF_KOUTEI_QTY08% = 76     '           ����
Private Const ptxBEF_KOUTEI_KOUSU08% = 77   '           �H��
Private Const ptxBEF_KOUTEI_TANI09% = 78    '�O�H��09�@ �P��
Private Const ptxBEF_KOUTEI_QTY09% = 79     '           ����
Private Const ptxBEF_KOUTEI_KOUSU09% = 80   '           �H��

Private Const ptxBEF_KOUTEI_KEI1% = 81      '�O�H���v   �v

Private Const ptxBEF_KOUTEI_R_RATE% = 82    '�O�H���v   �]�T��

Private Const ptxBEF_KOUTEI_KEI2% = 83      '�O�H���v   (�b�^��)
Private Const ptxBEF_KOUTEI_KEI3% = 84      '�O�H���v   (���^��)
Private Const ptxBEF_KOUTEI_KEI4% = 85      '�O�H���v   (�~�^��)

Private Const ptxMAIN_KOUTEI_TANI01% = 86   '��ƍH��01 �P��
Private Const ptxMAIN_KOUTEI_QTY01% = 87    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU01% = 88  '           �H��
Private Const ptxMAIN_KOUTEI_TANI02% = 89   '��ƍH��02 �P��
Private Const ptxMAIN_KOUTEI_QTY02% = 90    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU02% = 91  '           �H��
Private Const ptxMAIN_KOUTEI_TANI03% = 92   '��ƍH��03 �P��
Private Const ptxMAIN_KOUTEI_QTY03% = 93    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU03% = 94  '           �H��
Private Const ptxMAIN_KOUTEI_TANI04% = 95   '��ƍH��04 �P��
Private Const ptxMAIN_KOUTEI_QTY04% = 96    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU04% = 97  '           �H��
Private Const ptxMAIN_KOUTEI_TANI05% = 98   '��ƍH��05 �P��
Private Const ptxMAIN_KOUTEI_QTY05% = 99    '           ����
Private Const ptxMAIN_KOUTEI_KOUSU05% = 100 '           �H��
Private Const ptxMAIN_KOUTEI_TANI06% = 101  '��ƍH��06 �P��
Private Const ptxMAIN_KOUTEI_QTY06% = 102   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU06% = 103 '           �H��
Private Const ptxMAIN_KOUTEI_TANI07% = 104  '��ƍH��07 �P��
Private Const ptxMAIN_KOUTEI_QTY07% = 105   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU07% = 106 '           �H��
Private Const ptxMAIN_KOUTEI_TANI08% = 107  '��ƍH��08 �P��
Private Const ptxMAIN_KOUTEI_QTY08% = 108   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU08% = 109 '           �H��
Private Const ptxMAIN_KOUTEI_TANI09% = 110  '��ƍH��09 �P��
Private Const ptxMAIN_KOUTEI_QTY09% = 111   '           ����
Private Const ptxMAIN_KOUTEI_KOUSU09% = 112 '           �H��

Private Const ptxMAIN_KOUTEI_KEI1% = 113    '��ƍH���v �v

Private Const ptxMAIN_KOUTEI_R_RATE% = 114  '��ƍH���v   �]�T��


Private Const ptxMAIN_KOUTEI_KEI2% = 115    '��ƍH���v  (�b�^��)
Private Const ptxMAIN_KOUTEI_KEI3% = 116    '��ƍH���v  (���^��)
Private Const ptxMAIN_KOUTEI_KEI4% = 117    '��ƍH���v  (�~�^��)

Private Const ptxAFT_KOUTEI_TANI01% = 118   '��H��01   �P��
Private Const ptxAFT_KOUTEI_QTY01% = 119    '           ����
Private Const ptxAFT_KOUTEI_KOUSU01% = 120  '           �H��
Private Const ptxAFT_KOUTEI_TANI02% = 121   '��H��02   �P��
Private Const ptxAFT_KOUTEI_QTY02% = 122    '           ����
Private Const ptxAFT_KOUTEI_KOUSU02% = 123  '           �H��
Private Const ptxAFT_KOUTEI_TANI03% = 124   '��H��03   �P��
Private Const ptxAFT_KOUTEI_QTY03% = 125    '           ����
Private Const ptxAFT_KOUTEI_KOUSU03% = 126  '           �H��
Private Const ptxAFT_KOUTEI_TANI04% = 127   '��H��04   �P��
Private Const ptxAFT_KOUTEI_QTY04% = 128    '           ����
Private Const ptxAFT_KOUTEI_KOUSU04% = 129  '           �H��
Private Const ptxAFT_KOUTEI_TANI05% = 130   '��H��05   �P��
Private Const ptxAFT_KOUTEI_QTY05% = 131    '           ����
Private Const ptxAFT_KOUTEI_KOUSU05% = 132  '           �H��
Private Const ptxAFT_KOUTEI_TANI06% = 133   '��H��06   �P��
Private Const ptxAFT_KOUTEI_QTY06% = 134    '           ����
Private Const ptxAFT_KOUTEI_KOUSU06% = 135  '           �H��
Private Const ptxAFT_KOUTEI_TANI07% = 136   '��H��07   �P��
Private Const ptxAFT_KOUTEI_QTY07% = 137    '           ����
Private Const ptxAFT_KOUTEI_KOUSU07% = 138  '           �H��
Private Const ptxAFT_KOUTEI_TANI08% = 139   '��H��08   �P��
Private Const ptxAFT_KOUTEI_QTY08% = 140    '           ����
Private Const ptxAFT_KOUTEI_KOUSU08% = 141  '           �H��
Private Const ptxAFT_KOUTEI_TANI09% = 142   '��H��09   �P��
Private Const ptxAFT_KOUTEI_QTY09% = 143    '           ����
Private Const ptxAFT_KOUTEI_KOUSU09% = 144  '           �H��

Private Const ptxAFT_KOUTEI_KEI1% = 145     '��H���v   �v

Private Const ptxAFT_KOUTEI_R_RATE% = 146   '��H���v   �]�T��



Private Const ptxAFT_KOUTEI_KEI2% = 147     '��H���v   (�b�^��)
Private Const ptxAFT_KOUTEI_KEI3% = 148     '��H���v   (���^��)
Private Const ptxAFT_KOUTEI_KEI4% = 149     '��H���v   (�~�^��)


Private Const ptxKOUTEI_KEI1% = 150         '�H���v   �v

Private Const ptxKOUTEI_R_RATE% = 151       '�H���v   �]�T��


Private Const ptxKOUTEI_KEI2% = 152         '�H���v   (�b�^��)
Private Const ptxKOUTEI_KEI3% = 153         '�H���v   (���^��)
Private Const ptxKOUTEI_KEI4% = 154         '�H���v   (�~�^��)


Private Const ptxS_CLASS_CODE% = 155        '���i���׽
Private Const ptxF_CLASS_CODE% = 156        '�t���׽
Private Const ptxN_CLASS_CODE% = 157        '���E�׽

Private Const ptxIO_TANKA_No% = 158         '�I�敪
Private Const ptxSE_Name% = 159             '�I�敪����





Private Const ptxSHIYOU_NO% = 167           '�d�l����       2009.06.02
Private Const ptxMITSUMORI_KBN% = 168       '���ς�敪     2009.06.02
'Private Const ptxTANKA_KIRIKAE_DT% = 169    '�P���ؑ֓��t   2009.06.02
Private Const ptxKIRIKAE_KBN% = 170         '�ؑ֋敪       2009.06.02
    







'------2009.07.24
Private Const ptxOLD_S_KOUSU_BAIKA% = 171       ' ��  (����)���i���H��
Private Const ptxOLD_S_SHIZAI_BAIKA% = 172      ' ��  (����)����

Private Const ptxOLD_S_GAISO_TANKA% = 173       ' ��  �O���P��
Private Const ptxOLD_S_PPSC_KAKO_KOSU% = 174    ' ��  PPSC���H�P��
Private Const ptxOLD_S_BU_KAKO_KOSU% = 175      ' ��  BU���H�P��
Private Const ptxTANKA_KIRIKAE_DT% = 176        ' ��  �P���ؑ֓��t
'------2009.07.24
Private Const ptxPLUS_KOUSU% = 177              ' �v���X���H��  2009.09.17




'------------------------------------   '�R���{��`
Private Const pcmbSHIMUKE% = 0          '�d������


'------------------------------------   '���b�`�e�L�X�g�{�b�N�X��`
Private Const prchBIKOU% = 0            '���l

Private Const prchM_BIKOU% = 1          '���Ϗ����l         2009.06.02



'------------------------------------   '�\���i
Private Const pGrdKOUSEI% = 0

Dim KOUSEI      As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row   As Integer                '�O���b�h�ő�\������

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 13             '�ő��

Private Const ColKO_JGYOBU% = 0         '���ƕ�
Private Const ColKO_NAIGAI% = 1         '�����O


Private Const ColKO_SYUBETSU% = 2       '���
Private Const ColKO_HIN_GAI% = 3        '�i��
Private Const ColKO_HIN_NAME% = 4       '�i��
Private Const ColKO_QTY% = 5            '����
Private Const ColG_ST_SHITAN% = 6       '�d����
Private Const ColG_ST_URITAN% = 7       '���し
Private Const ColG_ST_SHIKIN% = 8       '�d�����z
Private Const ColG_ST_URIKIN% = 9       '������z
Private Const ColS_KOUSU% = 10          '��Ǝ���
Private Const ColSEI_SYU_KON% = 11      '�W������
Private Const ColKO_BIKOU% = 12         '���l


                                        '���� ���z�o�͗p
Private Const ColG_ST_URIKIN_KUSATU% = 13



'-----------------------------------    �h���b�v�_�E��
Dim SYUBETSU        As New XArrayDB


'-----------------------------------

Dim KOSOU_KBN       As String * 2       '���敪
Dim GAISO_KBN       As String * 2       '�O���敪


Dim INV_IO_TANKA_No As String * 2       '�W���I���o�^���̏o�ɋ敪
Dim HIN_INV         As Boolean          '���o�^�i�Ԃ̓o�^��


Dim KUSATU_F        As Boolean          '�ΏۃZ���^�[�@���� OR ���ÈȊO


Dim SHIZAI_T        As Variant          '���ޑΏ�
Dim DOUKON_T        As Variant          '�����Ώ�
Dim KAKOU_T         As Variant          '���H�Ώ�

Dim BU_T            As Variant          'BU�t���Ώ�
Dim PPSC_T          As Variant          'PPSC�t���Ώ�

Private Const KUSATU_ETC$ = "���̑�"


Dim svHin_Gai       As String           '�i��
Dim svSHIMUKE_CODE  As String           '�d������


Dim FUTAI_KBN       As String * 2       '�t�э�� 2009.09.05

'-----------------------------------    �d�w�b�d�k �������Z��

Dim EX_NAME1        As String           '�����P
Dim EX_NAME2        As String           '�����Q

Dim EX_SYAMEI       As String           '���Ё@����
Dim EX_ADDR1        As String           '���Ё@�Z���P
Dim EX_ADDR2        As String           '���Ё@�Z���Q


Dim EX_CENTER_NAME  As String           '�Z���^�[   ����
Dim EX_CENTER_ADDR1 As String           '�Z���^�[   �Z���P
Dim EX_CENTER_ADDR2 As String           '�Z���^�[   �Z���Q

Dim EX_BIKOU1       As String           '���l�P
Dim EX_BIKOU2       As String           '���l�Q



'Dim EX_JIGYOBU      As String

'2009.06.02
Dim EX_SHIZAI_T     As Variant          '���ޑΏ�
Dim EX_SHIZAI_F     As Boolean          '���ޑΏ�

Dim EX_DOUKON_T     As Variant          '�����Ώ�
Dim EX_DOUKON_F     As Boolean          '�����Ώ�

Dim EX_FUKA_T       As Variant          '�t�����
Dim EX_FUKA_F       As Boolean          '�t�����


'2009.06.02

Dim EX_BCR_CODE     As String           '�ް�������ٺ���


Dim EXCEL_TEMPLATE  As String           'EXCEL����ڰ�

Private Const LAST_UPDATE_DAY$ = "2009.12.17 09:00"








Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String

    Select Case Index
        Case 0
            
            CommonDialog1.ShowOpen
            Text1(0).Text = Trim(CommonDialog1.fileName)
        
        
        
        
        Case 1
        
        
        
            ans = MsgBox("�i�ڃ}�X�^�O���J�z���z����ւ����������s���܂����H", vbYesNo, "�m�F����")
            
            If ans = vbYes Then
            
            
            
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            
            End If
        
        
        
        
        
        
        
        
        
        
        
        
        Case 2
        
            
            Label2(0).Caption = "���O�o�͒�"
            Label2(1).Caption = ""
            
            
            For i = 0 To List1.ListCount - 1
            
            
            
            
                Call Log_Out(LOG_F, List1.List(i))
            
            
            
            
            
            
            
            
            
            Next i
        
        
            Label2(0).Caption = "���O�o�͏I��"
            Label2(1).Caption = ""
        
        Case 3
            Unload Me
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






    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
                                
                                '�݌Ƀf�[�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

    F1040171.Caption = F1040171.Caption & " " & LAST_UPDATE_DAY


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    F1040171.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040171)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(F1040171)


    F1040171.MousePointer = vbDefault

End Sub




Private Function Update_Proc() As Integer
 
Dim In_File     As String
Dim tmp_File    As String
 
Dim i           As Long
Dim j           As Long
 
Dim rec         As String
 
Dim exl         As Object
 
 
Dim FileNo      As Long
 
Dim In_Rec      As String
Dim In_wk       As Variant
 
Dim sts         As Integer
 
Dim List_Wk     As String
 
 
Dim cnt         As Long
 
    Update_Proc = True
 
    Label2(0).Caption = "�f�[�^�ϊ���"
    cnt = 0
    
    In_File = Trim(Text1(0).Text)
    tmp_File = "c:\F104017.txt"
    
    Set exl = CreateObject("Excel.Application")
    '  exl.Application.Visible = True
    
    
    On Error GoTo Error_Proc
    exl.Application.Workbooks.Open fileName:=In_File
    
    FileNo = FreeFile
    Open tmp_File For Output As FileNo
    For j = 2 To 65536
        If exl.Cells(j, 1) = "" Then Exit For
            
            
            cnt = cnt + 1
            Label2(1).Caption = cnt
            DoEvents
            rec = ""
            For i = 1 To 256
                If exl.Cells(j, i) = "" Then Exit For
                rec = rec & exl.Cells(j, i) & vbTab
            Next
            Print #1, rec
    Next
    Close FileNo
    '  exl.Application.DisplayAlerts = False
    exl.Application.Quit
        
        
    cnt = 0
    FileNo = FreeFile
    Open tmp_File For Input As FileNo
    
    
    
    List1.Clear
    Label2(0).Caption = "�i�ڃ}�X�^�X�V"
    
    Do While Not EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, In_Rec
        
        In_wk = Split(In_Rec, vbTab, -1)
    
        cnt = cnt + 1
        Label2(1).Caption = cnt
        DoEvents
    
        If UBound(In_wk) < 7 Then
        Else
        
            If In_wk(7) <> "*" Then
            Else
        
                If IsNumeric(In_wk(5)) Then
        
        
        
                    
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(CStr(In_wk(0))))
                    
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(CStr(In_wk(1))))
            
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Trim(CStr(In_wk(2))))
            
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            
                            
                            If Val(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) = Val(CLng(Trim(CStr(In_wk(4))))) Then
                                            
                                            
                                List_Wk = CStr(In_wk(0)) & " " & CStr(In_wk(1)) & " " & CStr(In_wk(2)) & " " & StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode) & "->" & Format(CLng(CStr(In_wk(5))), "00000000000")
                                
                            
                                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Trim(CStr(In_wk(5)))), "00000000000"))
                            
                            
                            
                                List1.AddItem List_Wk
                                DoEvents
                            
                                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts Then
                                    Exit Function
                                End If
                        
                        
                            End If
                        
                        Case BtErrKeyNotFound
                        Case Else
                            Exit Function
                    End Select
            
            
                End If
            End If
        
        
        
        
        
        
        
        
        
        End If
    Loop
    
    Close FileNo
        
        
        
        
        
        
        
    Kill tmp_File
        

    Label2(0).Caption = "�����I��"
    Label2(1).Caption = ""



    Update_Proc = False
    
    Exit Function
Error_Proc:
 
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
Dim ans     As Integer
    
    
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("�h���C�u�܂��̓p�X��������܂���" & In_File, vbExclamation)
        Case ErrNotFound, 1004
            Beep
            ans = MsgBox("�t�@�C����������܂���" & In_File, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[  [" & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select


End Function

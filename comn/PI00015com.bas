Attribute VB_Name = "PI00015com"
Option Explicit

Private Type Item_Key_tag
    JGYOBU  As String * 1
    NAIGAI  As String * 1
End Type

Public K_Item_Tbl() As Item_Key_tag   '�����ޕi�ڏ��
Public G_Item_Tbl() As Item_Key_tag   '�O�����ޕi�ڏ��



Private Type D_Item_Tbl_Tag
    SYUBETSU    As String * 2               '���
    JGYOBU      As String * 1               '���ƕ�
    NAIGAI      As String * 1               '�����O
    HIN_GAI     As String * 20              '�i��
    QTY         As Double                   '����
    SHIJI_QTY   As Double                   '���ʁi�w�����j
    BIKOU       As String * 40              '���l�i���͒l�j
    ID_NO       As String * 12              'ID_No(�o�ח\��ID_No)
End Type



Public D_Item_Tbl()     As D_Item_Tbl_Tag   '�����^�\���i�ڏ��


Public Taget_Key        As String * 8       '�X�V�Ώۂ̎w�}�[��

Public Doukon_Tbl_No(0 To 19) _
                        As String * 1

Public Doukon_Start     As Integer          '��ʊJ�n�s��

Public POS_UMU          As Boolean

Public PRI_S_TANTO      As Boolean      '���x�^�S���҈�� OFF:����Ȃ� ON:�������
Public PRI_MAIN_BCR     As Boolean      'Ҳ��ް���� OFF:����Ȃ� ON:�������

Public PRI_BIKOU_BCR    As Integer      '���l���@0�F���͒l�@1:�o��BCR 2:�i��

Public PRI_DOUKON       As Boolean      '���i�������@���� OFF:����Ȃ� ON:�������

Public PRI_NYUKO_IN     As Boolean      '���Ɋ�����@���� OFF:����Ȃ� ON:�������
Public PRI_INPUT_IN     As Boolean      '���͊�����@���� OFF:����Ȃ� ON:�������

Public PRI_SAGYO_DAY    As Boolean      '��Ɠ��^���ʁ^�S�� OFF:����Ȃ� ON:������� 2007.05.22
Public PRI_HINBAN_BIKOU As Boolean      '�����@�i�ԁ^���^���� OFF:����Ȃ� ON:������� 2007.05.22


Public JISEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��
Public TASEKI_TITLE     As Variant      '���ӂ̖��̃^�C�g��

Public HIN_INV          As Boolean      '���o�^�i�ԉ�


'--------------------------------------------------- ���  ���ޑΉ��@2012.03.20
Public Jyogai_Soko_umu _
                        As Boolean              '���O�q�ɐݒ�L��

'--------------------------------------------------- ���  ���ޑΉ��@2012.03.20


'---------------------------------------------- *���i���w�}�ް��i�e�j�ʃ|�C���^
'�|�W�V���j���O
Public wP_SSHIJI_O_POS  As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_SSHIJI_O_REC  As P_SSHIJI_O_REC_Tag
'�L�[�E�f�[�^
Public K0_wP_SSHIJI_O   As KEY0_P_SSHIJI_O
Public K1_wP_SSHIJI_O   As KEY1_P_SSHIJI_O


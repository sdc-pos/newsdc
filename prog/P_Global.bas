Attribute VB_Name = "P_Global"
Option Explicit

'********************************************************************
'*                            �ϐ���`                              *
'*                                                                  *
'********************************************************************



'------------------------------------------ �Ǘ��}�X�^KEY��`
Public Const P_ST_KANRI_No$ = "01"          '����ް�KEY

Public Const P_ST_KANRI_DEF_No$ = "02"      '�����lKEY

'------------------------------------------ �R�}���h�{�^����`
Public Const P_CMD_Upd% = 0                 '�X�V
Public Const P_CMD_Ins% = 2                 '�V�K
Public Const P_CMD_DEL% = 3                 '�폜
Public Const P_CMD_DSP% = 4                 '����/�\��
Public Const P_CMD_OUT% = 7                 '�ް��o��
Public Const P_CMD_PRT% = 8                 '���

Public Const P_CMD_End% = 11                '�I��

'------------------------------------------ �R�[�h�}�X�^�敪��`
Public Const P_KBN01_CD$ = "01"             '�d���敪�@     �R�[�h
Public Const P_KBN01_NM$ = "�d���敪"       '          ����
Public Const P_KBN01_Len% = 2               '          �L������
Public Const P_KBN01_OP1% = True            '          ��߼��1
Public Const P_KBN01_OP2% = False           '          ��߼��2
Public Const P_KBN01_OP_NM1$ = "�W�v��"     '          ��߼�ݖ���1
Public Const P_KBN01_OP_NM2$ = ""           '          ��߼�ݖ���2



Public Const P_KBN02_CD$ = "02"             '�̔��敪�@     �R�[�h
Public Const P_KBN02_NM$ = "�̔��敪"       '          ����
Public Const P_KBN02_Len% = 2               '          �L������
Public Const P_KBN02_OP1% = True            '          ��߼��1
Public Const P_KBN02_OP2% = False           '          ��߼��2
Public Const P_KBN02_OP_NM1$ = "�W�v��"     '          ��߼�ݖ���1
Public Const P_KBN02_OP_NM2$ = ""           '          ��߼�ݖ���2


Public Const P_KBN03_CD$ = "03"             '���x�P�ʁ@     �R�[�h
Public Const P_KBN03_NM$ = "���x�P��"       '          ����
Public Const P_KBN03_Len% = 3               '          �L������
Public Const P_KBN03_OP1% = False           '          ��߼��1
Public Const P_KBN03_OP2% = False           '          ��߼��2
Public Const P_KBN03_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN03_OP_NM2$ = ""           '          ��߼�ݖ���2


Public Const P_KBN04_CD$ = "04"             '�d������@     �R�[�h
Public Const P_KBN04_NM$ = "�d������"       '          ����
Public Const P_KBN04_Len% = 2               '          �L������
Public Const P_KBN04_OP1% = True            '          ��߼��1
Public Const P_KBN04_OP2% = True            '          ��߼��2
Public Const P_KBN04_OP_NM1$ = "���ƕ�"     '          ��߼�ݖ���1
Public Const P_KBN04_OP_NM2$ = "�����O"     '          ��߼�ݖ���2

Public Const P_KBN05_CD$ = "05"             '���P/�S����    �R�[�h
Public Const P_KBN05_NM$ = "���P�^�S����"   '          ����
Public Const P_KBN05_Len% = 2               '          �L������
Public Const P_KBN05_OP1% = False           '          ��߼��1
Public Const P_KBN05_OP2% = False           '          ��߼��2
Public Const P_KBN05_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN05_OP_NM2$ = ""           '          ��߼�ݖ���2

Public Const P_KBN06_CD$ = "06"             '����           �R�[�h
Public Const P_KBN06_NM$ = "���"           '          ����
Public Const P_KBN06_Len% = 2               '          �L������
Public Const P_KBN06_OP1% = False           '          ��߼��1
Public Const P_KBN06_OP2% = False           '          ��߼��2
Public Const P_KBN06_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN06_OP_NM2$ = ""           '          ��߼�ݖ���2

Public Const P_KBN07_CD$ = "07"             '���/���ƕ�    �R�[�h
Public Const P_KBN07_NM$ = "���/���ƕ�"    '          ����
Public Const P_KBN07_Len% = 2               '          �L������
Public Const P_KBN07_OP1% = True            '          ��߼��1
Public Const P_KBN07_OP2% = True            '          ��߼��2
Public Const P_KBN07_OP_NM1$ = "���ƕ�"     '          ��߼�ݖ���1
Public Const P_KBN07_OP_NM2$ = "�����O"     '          ��߼�ݖ���2

Public Const P_KBN08_CD$ = "08"             '���ދ敪       �R�[�h
Public Const P_KBN08_NM$ = "���ދ敪"       '          ����
Public Const P_KBN08_Len% = 1               '          �L������
Public Const P_KBN08_OP1% = False           '          ��߼��1
Public Const P_KBN08_OP2% = False           '          ��߼��2
Public Const P_KBN08_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN08_OP_NM2$ = ""           '          ��߼�ݖ���2

Public Const P_KBN09_CD$ = "09"             '�o�c����       �R�[�h      2008.02.28
Public Const P_KBN09_NM$ = "�o�c����"       '          ����
Public Const P_KBN09_Len% = 2               '          �L������
Public Const P_KBN09_OP1% = False           '          ��߼��1
Public Const P_KBN09_OP2% = False           '          ��߼��2
Public Const P_KBN09_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN09_OP_NM2$ = ""           '          ��߼�ݖ���2

Public Const P_KBN10_CD$ = "10"             '����           �R�[�h      2008.02.28
Public Const P_KBN10_NM$ = "����"           '          ����
Public Const P_KBN10_Len% = 2               '          �L������
Public Const P_KBN10_OP1% = False           '          ��߼��1
Public Const P_KBN10_OP2% = False           '          ��߼��2
Public Const P_KBN10_OP_NM1$ = ""           '          ��߼�ݖ���1
Public Const P_KBN10_OP_NM2$ = ""           '          ��߼�ݖ���2





Public Const P_KBN_MAX% = 9                 '�敪���i����-�P�j

Public G_SCREEN_FLG As Integer              '��ʑJ�ڗp�̋��ʃt���O

Public Const G_SCREEN_INS% = 1              '�Ώۃ��R�[�h�Ȃ�
Public Const G_SCREEN_UPD% = 2              '�Ώۃ��R�[�h����

Public Const G_INPUT_OK& = &H80000005       '����OK̨����
Public Const G_INPUT_NG& = &H8000000F       '����NG̨����


Public P_YOIN_TU_NYUKA      As String * 2   '�u���ޒʏ���ׁv�̗v��
Public P_YOIN_MAE_SOUSAI    As String * 2   '�u���ޑO�؂葊�E�v�̗v��



'------------------------------------------ �g�����i
Public Const P_ASSEMBLY_OFF$ = "0"          '�g���ĂȂ�
Public Const P_ASSEMBLY_ON$ = "1"           '�g���Ă���
'------------------------------------------ ��
Public Const L_PAPER_OFF$ = "0"             'OFF
Public Const L_PAPER_ON$ = "1"              'ON
'------------------------------------------ �v���X�`�b�N
Public Const L_PLASTIC_OFF$ = "0"           'OFF
Public Const L_PLASTIC_ON$ = "1"            'ON
'------------------------------------------ �K�p�@�탉�x��
Public Const L_LABEL_OFF$ = "0"             'OFF
Public Const L_LABEL_ON$ = "1"              'ON
'------------------------------------------ �������x��
Public Const L_MAISU_OFF$ = "0"             'OFF
Public Const L_MAISU_ON$ = "1"              'ON
'------------------------------------------ ��/�O��/�����E�\��
Public Const P_HEAD$ = "0"                  'ͯ�ް

Public Const P_KOSOU$ = "1"                 '������
Public Const P_GAISOU$ = "2"                '�O������
Public Const P_DOUKON$ = "3"                '�����E�\��
'------------------------------------------ �����敪
Public Const P_TORI_GENERAL$ = "0"          '���
Public Const P_TORI_NAISYOKU$ = "1"         '���E
Public Const P_TORI_GENKIN$ = "2"           '����
Public Const P_TORI_SYANAI$ = "3"           '������
Public Const P_TORI_ANOTHER$ = "4"          '������
Public Const P_TORI_JIKYU$ = "5"            '���E(����)


Public Const P_TORI_GENERAL_N$ = "��@��"
Public Const P_TORI_NAISYOKU_N$ = "���@�E"
Public Const P_TORI_GENKIN_N$ = "���@��"
Public Const P_TORI_SYANAI_N$ = "������"
Public Const P_TORI_ANOTHER_N$ = "������"
Public Const P_TORI_JIKYU_N$ = "���@��"
'------------------------------------------ ���ٓ\��v��Ȃ�
Public Const P_G_LABEL_OFF$ = "0"           '�v�サ�Ȃ�
Public Const P_G_LABEL_ON$ = "1"            '�v�シ��
'------------------------------------------ �݌ɊǗ�
Public Const P_ZAIKO_F_ON$ = "1"            '�Ώ�
Public Const P_ZAIKO_F_OFF$ = "0"           '�ΏۊO

'------------------------------------------ ���{�쐬
Public Const P_SAMPLE_F_OFF$ = "0"          '�Ȃ�
Public Const P_SAMPLE_F_ON$ = "1"           '����
'------------------------------------------ �w���`��
Public Const P_SHIJI_F_NORMAL$ = "0"        '�Ȃ�
Public Const P_SHIJI_F_SPOT$ = "1"          '��߯�
Public Const P_SHIJI_F_KEPPIN$ = "2"        '���i����

Public Const P_SHIJI_F_SAIKON$ = "3"        '�č��� 2007.11.09


'------------------------------------------ �o�͑Ώہ@�w�}�[
Public Const P_PRI_SHIJI_OFF$ = "0"         '�Ȃ�
Public Const P_PRI_SHIJI_ON$ = "1"          '����
'------------------------------------------ �o�͑Ώہ@�߰�����
Public Const P_PRI_PARTS_OFF$ = "0"         '�Ȃ�
Public Const P_PRI_PARTS_ON$ = "1"          '����
'------------------------------------------ �o�͑Ώہ@�O������
Public Const P_PRI_GAISOU_OFF$ = "0"        '�Ȃ�
Public Const P_PRI_GAISOU_ON$ = "1"         '����
'------------------------------------------ �o�͑Ώہ@�@������
Public Const P_PRI_KISHU_OFF$ = "0"        '�Ȃ�
Public Const P_PRI_KISHU_ON$ = "1"         '����

'------------------------------------------ ����F
Public Const P_KAN_OFF$ = "0"               '����
Public Const P_KAN_ON$ = "1"                '����

'------------------------------------------ ��ݾ�F
Public Const P_CANCEL_OFF$ = "0"            '��
Public Const P_CANCEL_ON$ = "1"             '��ݾ�

'------------------------------------------ ���F
Public Const P_UKEIRE_CON$ = "0"            '�p���i�����j
Public Const P_UKEIRE_END$ = "1"            '�ŏI���

'------------------------------------------ ���F
Public Const P_PRINT_OFF$ = "0"             '�����
Public Const P_PRINT_ON$ = "1"              '�����

'------------------------------------------ ����F
Public Const P_SEIKYU_NON$ = "0"            '������
Public Const P_SEIKYU_PRI$ = "1"            '�����
Public Const P_SEIKYU_END$ = "9"            '���ߍ�


'------------------------------------------ �̔��敪
Public Const P_HN_HANBAI$ = "1"             '�̔�
Public Const P_HN_SEIZOU$ = "2"             '����
Public Const P_HN_YATIN$ = "3"              '�ƒ�
Public Const P_HN_ETC$ = "4"                '���̑�
Public Const P_HN_HAKEN$ = "5"              '�h��
                                            
                                            '*��L�ȊO�͑S�Ă��̑���
'------------------------------------------ �d���敪
Public Const P_SH_SHIIRE$ = "1"             '�d��
Public Const P_SH_SEIZOU$ = "2"             '����
Public Const P_SH_YATIN$ = "3"              '�ƒ�
Public Const P_SH_ETC$ = "4"                '���̑�
Public Const P_SH_HAKEN$ = "5"              '�h��
Public Const P_SH_KEIHI$ = "6"              '�o��
Public Const P_SH_ZEI$ = "7"                '�����


'------------------------------------------ �I�����ް����vKEY(�d������)
Public Const P_StokSum_Key$ = "zzz"
'------------------------------------------ ���Y���э��vKEY(�׽)
Public Const P_ClassSum_Key$ = "!!!!!!!!!!!!!!!!!!!!"



Attribute VB_Name = "Global"
Option Explicit

'   �E�C���h�E�Y�I���v��
    Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'   �E�C���h�E�Y�I���v��
    Declare Function ExitWindowsEx Lib "user32 " (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'   �������f
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'   �h�m�h�t�@�C����������
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'   �h�m�h�t�@�C���ǂݍ���
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'   �R���s���[�^���擾
    Declare Function GetComputerNameA Lib "kernel32" _
           (ByVal IpBuffer As String, nSize As Long) As Long
    
    Declare Function GetVersion Lib "kernel32.dll" () As Long
    Declare Function WriteProfileString Lib "kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

    Public Const HWND_BROADCAST  As Long = &HFFFF&
    Public Const WM_WININICHANGE As Long = &H1A&

    Public Const EM_GETLINECOUNT As Long = &HBA     '2016.01.05

    '2019.03.29
    Public Const CB_SHOWDROPDOWN = &H14F



'   �L�[�X�g���[�N�����֐�
    Declare Sub Keybd_Event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


    Declare Function GetDeviceCaps Lib "gdi32" _
        (ByVal hDC As Long, ByVal nIndex As Long) As Long

    Public Const HORZRES = 8           '���ۂ̃X�N���[���̕��i������̈�j
    Public Const VERTRES = 10          '���ۂ̃X�N���[���̍���
    Public Const PHYSICALWIDTH = 110   '�����I��(���p���T�C�Y�j
    Public Const PHYSICALHEIGHT = 111  '�����I����
    Public Const PHYSICALOFFSETX = 112 '����\�ȍ������̃}�[�W��
    Public Const PHYSICALOFFSETY = 113 '����\�ȏ�����̃}�[�W��
    
'********************************************************************
'*                            �ϐ���`                              *
'*                                                                  *
                   
'********************************************************************


'***** �V�X�e���ُ� ********** 97.01.08
Public Const SYS_ERR% = -100
Public Const SYS_CANCEL% = -200

'***** �V�X�e�����ʒ�` ******
                                    
'-----------------------------------'���ƕ��敪
Public Const SOJIKI$ = "7"          '�|���@
Public Const DENKA$ = "D"           '�d������
Public Const SUIHAN$ = "4"          '���ъ�
Public Const SENTAKU$ = "1"         '����@�i�A�C�����j
Public Const AIRCON$ = "A"          '�G�A�R��           2004.12.01
Public Const REIZOU$ = "R"          '�①��             2007.05.24

Public Const SETSUBI$ = "B"         '�ݔ�   2007.03.28

Public Const SHIZAI$ = "S"          '����   2005.11.16
Public Const BUZAI$ = "C"           '����   2012.03.23
Public Const BLBU$ = "5"            '�ޭ�è���ݸ�   2012.04.06
Public Const OVEN$ = "6"            '�d�q�����W     2012.05.16
Public Const YUKADAN$ = "Y"         '���g�[         2013.06.06
Public Const JCS$ = "J"             'JCS            2015.01.22
Public Const SHOKUSEN$ = "2"        '�H��           2015.03.03

Public Const JGYOBU_NON$ = "0"      '���ƕ��敪�Ȃ�
                                   
'-----------------------------------'�q�ɋ敪
Public Const BUN_JITU$ = "0"        '���q��
Public Const BUN_KASO$ = "1"        '�V�X�e���ŗL
Public Const BUN_AUTO$ = "2"        '�����q��
                                   
Public Const SOKO_BUN0$ = "���q��  "
Public Const SOKO_BUN1$ = "�V�X�e��"
Public Const SOKO_BUN2$ = "�����q��"
'-----------------------------------'�����O
Public Const NAIGAI_NON$ = "0"      '�Ȃ�
Public Const NAIGAI_NAI$ = "1"      '����
Public Const NAIGAI_GAI$ = "2"      '�C�O
                                   
Public Const NAIGAI0$ = "�Ȃ�"
Public Const NAIGAI1$ = "����"
Public Const NAIGAI2$ = "�C�O"
'-----------------------------------'�q�Ɂ^�I�@�g�p��
Public Const KAHI_KBN_OK$ = "0"     '�g�p��
Public Const KAHI_KBN_NG$ = "1"     '�g�p�s��
                                   
Public Const KAHI_KBN0$ = "�g�p�@��"
Public Const KAHI_KBN1$ = "�g�p�s��"
'-----------------------------------'�q�Ɂ^�I�@�݌ɏƍ� 2004.02
Public Const ZAIKO_SHOGO_FLG_OK$ = "0"      '�ƍ��L
Public Const ZAIKO_SHOGO_FLG_NG$ = "1"      '�ƍ���
                                   
Public Const ZAIKO_SHOGO0$ = "�΁@��"
Public Const ZAIKO_SHOGO1$ = "�ΏۊO"
'-----------------------------------'�q�� ���ڋ敪
Public Const KONS_KBN_OK$ = "0"     '���ډ�
Public Const KONS_KBN_NG$ = "1"     '���ڕs��

Public Const KONS_KBN0$ = "���ډ�  "
Public Const KONS_KBN1$ = "���ڕs��"
'-----------------------------------'������@�l�s�r�敪
Public Const MUKE_MTS$ = "1"        '�l�s�r
Public Const MUKE_SS$ = "2"         '�r�r

'-----------------------------------'�o�ח\��^�݌Ɂ@�g�p��
Public Const LOCK_OFF$ = "0"        '�g�p��
Public Const LOCK_ON$ = "1"         '�g�p��
'-----------------------------------'���i�ς݁^�����i�̎���2004.04
Public Const GOODS_ON$ = "0"        '���i��
Public Const GOODS_OFF$ = "1"       '�����i
'-----------------------------------'�o�ׁ^���ׂ̊����t���O
Public Const KAN_KBN_UN$ = "0"      '������
Public Const KAN_KBN_FIN$ = "9"     '�����ς�
'-----------------------------------'�o�ח\�� �����敪
'Public Const KAN_SOFF_POFF_KOFF$ = "0"      '�����敪�����o�Ɂ^������^�����i
'Public Const KAN_SING_POFF_KOFF$ = "1"      '�����敪���o�ɒ��^������^�����i
'Public Const KAN_SOFF_PON_KOFF$ = "2"       '�����敪�����o�Ɂ^����ρ^�����i
'Public Const KAN_SING_PON_KOFF$ = "3"       '�����敪���o�ɒ��^����ρ^�����i
'Public Const KAN_SON_POFF_KOFF$ = "4"       '�����敪���o�ɍρ^������^�����i
'Public Const KAN_SON_PON_KOFF$ = "5"        '�����敪���o�ɍρ^����ρ^�����i
'Public Const KAN_SON_PNON_KON$ = "6"        '�����敪���o�ɍρ^�\�^���i��
'Public Const KAN_SNO_PNO_KNO$ = "9"         '�����敪���o�ɕs�^����s�^���i�s��

'Public Const KAN_L_SOFF_POFF_KOFF$ = "A"    '�����敪�����o�Ɂ^������^�����i
'Public Const KAN_L_SING_POFF_KOFF$ = "B"    '�����敪���o�ɒ��^������^�����i
'Public Const KAN_L_SOFF_PON_KOFF$ = "C"     '�����敪�����o�Ɂ^����ρ^�����i
'Public Const KAN_L_SING_PON_KOFF$ = "D"     '�����敪���o�ɒ��^����ρ^�����i

'-----------------------------------'��Ɓ^�v���̎���
Public Const ACT_ZAITEI_IN$ = "1"       '�ݒ��i�{�j
Public Const ACT_ZAITEI_OUT$ = "2"      '�ݒ��i�|�j
Public Const ACT_NYUKA$ = "3"           '����
Public Const ACT_SYUKA_KEI$ = "4"       '�o��(�o�ח\��L��)
Public Const ACT_SYUKA_HYO$ = "5"       '�o��(�o�ɕ\)
Public Const ACT_SYUKA_GAI$ = "6"       '�o��(�o�ח\�薳��)
Public Const ACT_IDO_IN$ = "7"          '�ړ�����
Public Const ACT_IDO_OUT$ = "8"         '�ړ��o��
Public Const ACT_DENPYO_ID$ = "9"       '�`�[�h�c   2004.02
Public Const ACT_KENPIN$ = "A"          '���i
Public Const ACT_WEL_ETC$ = "B"         'WEL��p

Public Const ACT_KENPIN_MTS$ = "C"      '������ǂݍ��ݗp
Public Const ACT_GOODS_ONFF$ = "D"      '���i���������i�؂�ւ��p

Public Const ACT_SPECIAL_PROCESS$ = "E" '���ꏈ��

Public Const ACT_KENPIN_DEN$ = "F"      '���i�i���PC�j 2006.12.07

Public Const ACT_SYUKA_HYO_OSAKA$ = "G" '�o�ɕ\�o�Ɂi���PC�j 2007.03.16

Public Const ACT_IN_KENPIN_OSAKA$ = "H" '���Ɍ��i�i���PC�j 2007.06.07
Public Const ACT_IN_TANA_OSAKA$ = "I"   '���I���Ɂi���PC�j 2007.06.07

Public Const ACT_FURIKAE$ = "J"         '���ސU�ցi���PC�j 2007.06.28


Public Const ACT_BINNO$ = "K"           '�և��i�ڊǗp�j 2009.03.11


Public Const ACT_KENPIN_GAI$ = "L"      '���i�C�O   2009.08.05


'Public Const ACT_SAI_SU$ = "M"          '�ː��^����   2010.01.21


Public Const ACT_SHOUHINKA$ = "M"       '���i��   2010.09.03

Public Const ACT_LotNo$ = "N"         '���g�[�@����   2013.06.06

Public Const ACT_MODULE$ = "O"         '���W���[��   2014.06.24


Public Const ACT_DENPYO_ID2$ = "P"      '�`�[�h�c   2015.02.21

Public Const ACT_KENPIN_Drct$ = "Q"     '�������i   2016.10.03

Public Const ACT_BCR_PRINT$ = "R"       '�o�[�R�[�h�󎚁@2017.04.10

Public Const ACT_NEW_KENPIN$ = "S"      '�V���i 2018.11.05
Public Const ACT_NEW_KENPIN_MTS$ = "T"  '�V������ǂݍ��ݗp 2018.11.05





Public Const ACT_SYSTEM$ = "Z"      '�V�X�e����p

Public YOIN_TU_NYUKA        As String * 2       '�u�ʏ���ׁv�̗v��
Public YOIN_MAEGARI         As String * 2       '�u�O�؂���ׁv�̗v��
Public YOIN_MAE_SOUSAI      As String * 2       '�u�O�؂葊�E�v�̗v��
Public YOIN_FURIKAE         As String * 2       '�u�����O�U�ւ��v�̗v��
Public YOIN_FURIKAE_OUT     As String * 2       '�u�����O�U�ւ����̏o�Ɂv�̗v��
Public YOIN_FURIKAE_IN      As String * 2       '�u�����O�U�ւ����̓��Ɂv�̗v��

Public YOIN_TANASHOGO       As String * 2       '�u�I�ƍ��v�̗v��
Public YOIN_TANAHINSHOGO    As String * 2       '�u�I�i�ƍ��v�̗v��


Public YOIN_HIN_SHOGO       As String * 2       '�u�i�ԏƍ��v�̗v�� 2011.02.03



'-----------------------------------'�z�X�g�f�[�^���o�ɋ敪
Public Const IO_KBN_URI$ = "0"      '���グ
Public Const IO_KBN_NYU$ = "1"      '����
Public Const IO_KBN_SYU$ = "2"      '�o��
Public Const IO_KBN_ZAT$ = "3"      '�݌ɒ���
Public Const IO_KBN_SYU_JITU$ = "4" '�o�׎���
Public Const IO_KBN_HENPIN$ = "5"   '�Ǖi�ԕi

Public Const IO_KBN_0$ = "���グ"
Public Const IO_KBN_1$ = "���@��"
Public Const IO_KBN_2$ = "�o�@��"
Public Const IO_KBN_3$ = "�݁@��"
Public Const IO_KBN_4$ = "�o�׎�"
Public Const IO_KBN_5$ = "�Ǖi��"
'-----------------------------------'�����敪
Public Const CYU_KBN_HSP$ = "0"      '��[�E�X�|�b�g
Public Const CYU_KBN_TUK$ = "1"      '����
Public Const CYU_KBN_SPO$ = "2"      '�X�|�b�g(�Ǒւ����恁�O)
Public Const CYU_KBN_HJU$ = "3"      '��[(�Ǒւ����恁�O)
Public Const CYU_KBN_TOK$ = "4"      '����(�Ǒւ����恁�O)
Public Const CYU_KBN_BOU$ = "E"      '�f��
Public Const CYU_KBN_KIN$ = "T"      '�������ً}�i�v�d�k��p�j

Public Const CYU_KBN_0$ = "��X"
Public Const CYU_KBN_1$ = "����"
'''Public Const CYU_KBN_2$ = "�X�|"      2003.06.03
Public Const CYU_KBN_2$ = "�ً}"        '2003.06.03
Public Const CYU_KBN_3$ = "��["
Public Const CYU_KBN_4$ = "����"
'Public Const CYU_KBN_4$ = "���"       '2005.11.16 ����c�b�́u��āv��L���ɂ���

Public Const CYU_KBN_E$ = "�f��"
Public Const CYU_KBN_T$ = "��O"        '2004.05.18
'-----------------------------------'�v���֌W
Public Const SUM_KBN_IN$ = "1"      '����
Public Const SUM_KBN_OT$ = "2"      '�o��
Public Const SUM_KBN_ZT$ = "3"      '�ݒ��}
Public Const SUM_KBN_MV$ = "4"      '�ړ�
Public Const SUM_KBN_NON$ = "0"     '�Ȃ�

Public Const SUM_KBN_I$ = "���Ɂ@"
Public Const SUM_KBN_O$ = "�o�Ɂ@"
Public Const SUM_KBN_Z$ = "�ݒ��}"
Public Const SUM_KBN_M$ = "�ړ��@"
Public Const SUM_KBN_N$ = "�Ȃ��@"

Public Const NORMAL_YOIN$ = "0"     '�ʏ�v��
Public Const SYSTEM_YOIN$ = "1"     '�V�X�e���v��

Public Const NORMAL_YOIN_N$ = "�ʏ�@�@"
Public Const SYSTEM_YOIN_N$ = "�V�X�e��"

'-----------------------------------'���̑��@���ʒ�`
Public Const ETS_MTS$ = "ZZZZZ"     '���̑�������
'-----------------------------------'�v���ݒ�
'Public Const ALL_YOIN$ = "0"        '�X�L���i�^��ʎg�p��
'-----------------------------------'�S�S���ҋ��ʃR�[�h
Public Const ALL_TANTO_CODE$ = "ZZZZZ"

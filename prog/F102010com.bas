Attribute VB_Name = "F102010com"
Option Explicit

Type wkSyukaRec_tag
    JGYOBA(0 To 7)              As Byte             '���Ə�
    DATA_KBN(0 To 0)            As Byte             '�f�[�^�敪
    TORI_KBN(0 To 1)            As Byte             '����敪
    ID_NO(0 To 11)              As Byte             'ID-NO
    KAIKEI_JGYOBA(0 To 7)       As Byte             '��v�p���Ə꺰��
    SHISAN_JGYOBA(0 To 7)       As Byte             '���Y�Ǘ����Ə꺰��
    HIN_NO(0 To 19)             As Byte             '�i�ڔԍ�
    DEN_NO(0 To 9)              As Byte             '�`�[�ԍ�
    SURYO(0 To 6)               As Byte             '�o�ɐ���
    MUKE_CODE(0 To 7)           As Byte             '�o�ɐ�
    SYUKO_SYUSI(0 To 7)         As Byte             '�o�Ɏ��x
    SHISAN_SYUSI(0 To 7)        As Byte             '���Y�Ǘ��p�݌Ɏ��x����
    HOJYO_SYUSI(0 To 7)         As Byte             '�⏕�݌Ɏ��x����
    SYUKO_YMD(0 To 7)           As Byte             '�o�ɓ��t
    TANKA(0 To 9)               As Byte             '�P��
    ODER_NO(0 To 11)            As Byte             '�I�[�_�[�ԍ�
    ITEM_NO(0 To 4)             As Byte             '�A�C�e���ԍ�
    ODER_NO_R(0 To 4)           As Byte             '�I�[�_�[����
    KOSO_KEITAI(0 To 13)        As Byte             '���`��       10-->14 2011.10.31
    SYUKA_YMD(0 To 7)           As Byte             '�o�ד�
    TANABAN1(0 To 9)            As Byte             '�I�ԂP
    TANABAN2(0 To 9)            As Byte             '�I�ԂQ
    TANABAN3(0 To 9)            As Byte             '�I�ԂR
    MUKE_NAME(0 To 23)          As Byte             '�o�ɐ於��
    CYU_KBN(0 To 0)             As Byte             '�����敪
    CYU_KBN_NAME(0 To 39)       As Byte             '�����敪����
    ORIGIN1(0 To 9)             As Byte             '���Y���P
    ORIGIN2(0 To 9)             As Byte             '���Y���Q
    BIKOU2(0 To 39)             As Byte             '���l�Q
    HAN_KBN(0 To 0)             As Byte             '�̔��敪
    CHOKU_KBN(0 To 0)           As Byte             '�����敪
    UNIT_ID_NO(0 To 11)         As Byte             '�ƯďC��ID-NO
    ZAIKO_HIKIATE(0 To 2)       As Byte             '�݌Ɉ�������
    GOKON_KANRI_NO(0 To 7)      As Byte             '�����Ǘ��ԍ�
    JYUCHU_ZAN(0 To 6)          As Byte             '�󒍎c����
    KYOKYU_KBN(0 To 0)          As Byte             '�����敪
    SHOHIN_SYUSI(0 To 7)        As Byte             '���i���[������x
    S_SHISAN_SYUSI(0 To 7)      As Byte             '���i���[�i���Y�Ǘ����x����
    S_HOJYO_SYUSI(0 To 7)       As Byte             '���i���[�i�⏕���x����
    BIKOU1(0 To 39)             As Byte             '���l�P
    CHOHA_KBN(0 To 0)           As Byte             '���[�敪
    JYU_HIN_NO(0 To 39)         As Byte             '�󒍕i�ڔԍ�
    HIN_NAME(0 To 39)           As Byte             '�i��
    HIN_CHANGE_KBN(0 To 0)      As Byte             '�i�ԕύX�敪
    MODULE_EXCHANGE(0 To 0)     As Byte             '���W���[�������敪
    ZAIKO_SYUSI(0 To 7)         As Byte             '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
    ZAN_SHISAN_SYUSI(0 To 7)    As Byte             '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
    ZAN_HOJYO_SYUSI(0 To 7)     As Byte             '�c�݌ɂ܂Ƃߕ⏕���x����
    NOUKI_YMD(0 To 7)           As Byte             '�w��[��
    SERVICE_KANRI_NO(0 To 8)    As Byte             '�T�[�r�X��ЊǗ��ԍ�
    KISHU_CODE(0 To 2)          As Byte             '�@��i�ڃR�[�h
    ENVIRONMENT_KBN(0 To 0)     As Byte             '���K�i���i�敪
    SS_CODE(0 To 7)             As Byte             '������R�[�h
    KEPIN_KAIJYO(0 To 0)        As Byte             '���i�����敪
'    FILLER(0 To 3)              As Byte
    CRLF(0 To 1)                As Byte             'CRLF
End Type

Public RYOHEN      As String * 2       '�Ǖi�ԕi�̗v�� 2009.07.10


Public Const WEL_MAEGARI_TANA_S_OSAKA$ = "H2"       '�uWEL ���ޑO�ؓ��Ɂv�̗v�� 2016.05.30


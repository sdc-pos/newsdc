Attribute VB_Name = "F1100801bas"
Option Explicit

Type INREC_Tag
    JGYOBA(0 To 7)                    As Byte     '���Ə�R�[�h
    SISAN_JGYOBA(0 To 7)    As Byte     '���Y�Ǘ����Ə�R�[�h
    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
    KISYU_HIN(0 To 2)       As Byte     '��\�@��i�ڃR�[�h
    HINMOKU_CD(0 To 2)      As Byte     '�i�ڃR�[�h
    SOKO_CD(0 To 1)         As Byte     '�q�ɃR�[�h
    KOSO_CD(0 To 9)         As Byte     '���`�ԃR�[�h
    BUHIN_SIZ(0 To 0)       As Byte     '���i�T�C�Y�敪
    KONPO_SAISU(0 To 13)    As Byte     '���i����ː�
    LABEL_HAKKO(0 To 0)     As Byte     '�K�p�@�탉�x�����s�敪
    KOBAI_TANTO(0 To 4)     As Byte     '�w���S���҃R�[�h
    UNIT_BUHIN(0 To 0)      As Byte     '���j�b�g���i�敪
    NAI_BUHIN(0 To 0)       As Byte     '�����������i�敪
    GAI_BUHIN(0 To 0)       As Byte     '�C�O�������i�敪
    HIN_BETU_NM(0 To 19)    As Byte     '�i�ڕʖ�
    HIN_NAME(0 To 19)       As Byte     '�i�ږ�
    U_TANKA2(0 To 9)        As Byte     '����P���Q
    U_TANKA3(0 To 9)        As Byte     '����P���R
    U_TANKA4(0 To 9)        As Byte     '����P���S
    LOC_NO1(0 To 9)         As Byte     '۹���ݔԍ��P
    LOC_NO2(0 To 9)         As Byte     '۹���ݔԍ��Q
    LOC_NO3(0 To 9)         As Byte     '۹���ݔԍ��R
    HIN_NAI(0 To 19)        As Byte     '�H��i�ڔԍ��i�����i�ԁj
    GENSANKOKU(0 To 9)      As Byte     '�����\�����Y����
    HYO_TANKA(0 To 9)       As Byte     '�W���P��
    FILLER(0 To 3)          As Byte     'FILLER
'    JGYOBA(0 To 7)          As Byte     '���Ə�R�[�h
'    SISAN_JGYOBA(0 To 7)    As Byte     '���Y�Ǘ����Ə�R�[�h
'    HIN_GAI(0 To 19)        As Byte     '�i�ԁi�O���j
'    KISYU_HIN(0 To 2)       As Byte     '��\�@��i�ڃR�[�h
'    HINMOKU_CD(0 To 2)      As Byte     '�i�ڃR�[�h
'    SOKO_CD(0 To 1)         As Byte     '�q�ɃR�[�h
'    KOSO_CD(0 To 9)         As Byte     '���`�ԃR�[�h
'    BUHIN_SIZ(0 To 0)       As Byte     '���i�T�C�Y�敪
'    KONPO_SAISU(0 To 13)    As Byte     '���i����ː�
'    LABEL_HAKKO(0 To 0)     As Byte     '�K�p�@�탉�x�����s�敪
'    KOBAI_TANTO(0 To 4)     As Byte     '�w���S���҃R�[�h
'    UNIT_BUHIN(0 To 0)      As Byte     '���j�b�g���i�敪
'    NAI_BUHIN(0 To 0)       As Byte     '�����������i�敪
'    GAI_BUHIN(0 To 0)       As Byte     '�C�O�������i�敪
'    HIN_BETU_NM(0 To 19)    As Byte     '�i�ڕʖ�
'    HIN_NAME(0 To 19)       As Byte     '�i�ږ�
'    U_TANKA2(0 To 9)        As Byte     '����P���Q
'    U_TANKA3(0 To 9)        As Byte     '����P���R
'    U_TANKA4(0 To 9)        As Byte     '����P���S
'    LOC_NO1(0 To 9)         As Byte     '۹���ݔԍ��P
'    LOC_NO2(0 To 9)         As Byte     '۹���ݔԍ��Q
'    LOC_NO3(0 To 9)         As Byte     '۹���ݔԍ��R
'    HIN_NAI(0 To 19)        As Byte     '�H��i�ڔԍ��i�����i�ԁj
'    GENSANKOKU(0 To 9)      As Byte     '�����\�����Y����
'    HYO_TANKA(0 To 9)       As Byte     '�W���P��
'    FILLER(0 To 3)          As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public INREC    As INREC_Tag


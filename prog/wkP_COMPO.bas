Attribute VB_Name = "wkP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const wkP_COMPO_ID$ = "wkP_COMPO"

'�y�[�W�T�C�Y
Private Const wkP_COMPO_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public wkP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Public Type wkP_COMPOREC_Tag
    
    
    SHIMUKE(0 To 2)         As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
    KO_SYUBETSU(0 To 1)     As Byte         '�q�@���
    KO_JGYOBU(0 To 0)       As Byte         '�q�@���ƕ�
    KO_NAIGAI(0 To 0)       As Byte         '�q�@�����O
    KO_HIN_GAI(0 To 19)     As Byte         '�q�@�i��
    KO_QTY(0 To 5)          As Byte         '�q�@����(999V99)
    KO_BIKOU(0 To 39)       As Byte         '�q�@���l
    FILLER(0 To 137)        As Byte         'Filler
    UPD_TANTO(0 To 4)       As Byte         '�X�V�@�S����
    UPD_DATETIME(0 To 13)   As Byte         '�X�V�@����

End Type
'�f�[�^�E�o�b�t�@
Public wkP_COMPOREC         As wkP_COMPOREC_Tag

'�L�[��`

Type KEY0_wkP_COMPO                         '�j�d�x�O
    SHIMUKE(0 To 2)         As Byte         '�d������
    JGYOBU(0 To 0)          As Byte         '���ƕ�
    NAIGAI(0 To 0)          As Byte         '�����O
    HIN_GAI(0 To 19)        As Byte         '�e�i��
    DATA_KBN(0 To 0)        As Byte         '�ް��敪
    SEQNO(0 To 2)           As Byte         '�ǔ�
End Type
    
'�L�[�E�f�[�^
Public K0_wkP_COMPO         As KEY0_wkP_COMPO

Type wkP_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private wkP_COMPO_Speck     As wkP_COMPO_FSpeck

Public Function wkP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �\���}�X�^�i�ʃ|�W�V���j���O�j  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wkP_COMPO_Open = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", wkP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [wkP_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wkP_COMPO_POS, wkP_COMPOREC, Len(wkP_COMPOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                Exit Function
        End Select
    Loop
    
    wkP_COMPO_Open = False

End Function

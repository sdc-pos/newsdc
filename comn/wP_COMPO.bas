Attribute VB_Name = "wP_COMPO"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const wP_COMPO_ID$ = "P_COMPO"

'�y�[�W�T�C�Y
Private Const wP_COMPO_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public wP_COMPO_POS      As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
'�f�[�^�E�o�b�t�@
Public wP_COMPO_O_REC        As P_COMPO_O_REC_Tag


'�f�[�^�E�o�b�t�@
Public wP_COMPO_K_REC        As P_COMPOREC_K_Tag

'�L�[��`

    
    
    
'�L�[�E�f�[�^
Public wK0_P_COMPO           As KEY0_P_COMPO

Public wK1_P_COMPO           As KEY1_P_COMPO             '2014.06.23

Public wK2_P_COMPO           As KEY2_P_COMPO             '2018.0.220



Type P_COMPO_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
    ks2                     As BtKeySpeck   ' �� ��߯��\����
    ks3                     As BtKeySpeck   ' �� ��߯��\����
    ks4                     As BtKeySpeck   ' �� ��߯��\����
    ks5                     As BtKeySpeck   ' �� ��߯��\����

    ks6                     As BtKeySpeck   ' �� ��߯��\����    '2014.06.23

    ks7                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks8                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks9                     As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks10                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks11                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20
    ks12                    As BtKeySpeck   ' �� ��߯��\����    '2018.02.20

End Type

Private wP_COMPO_Speck       As P_COMPO_FSpeck

Public Function wP_COMPO_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �\���}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    wP_COMPO_Open = True
                                            '�\���}�X�^�t���p�X�捞��
    sts = GetIni("FILE", wP_COMPO_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_COMPO]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, wP_COMPO_POS, wP_COMPO_O_REC, Len(wP_COMPO_O_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                Exit Function
        End Select
    Loop
    
    wP_COMPO_Open = False

End Function

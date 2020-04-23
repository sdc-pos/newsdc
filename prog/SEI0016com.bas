Attribute VB_Name = "SEI0016com"

Option Explicit


Public KOUSEI      As New XArrayDB

'Public KO_KOUSEI    As New XArrayDB



'********************************************************************
'*                                                                  *
'*              �\���}�X�^  �t�@�C����`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�t�@�C���h�c
Public Const wP_COMPO_ID$ = "P_COMPO"

'�|�W�V�����E�u���b�N
Public wP_COMPO_POS         As POSBLK
'�f�[�^�E�o�b�t�@
Public wP_COMPO_O_REC        As P_COMPO_O_REC_Tag
'�f�[�^�E�o�b�t�@
Public wP_COMPO_K_REC        As P_COMPOREC_K_Tag
    
'�L�[�E�f�[�^
Public K0_wP_COMPO           As KEY0_P_COMPO




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
            Case Else
                Call File_Error(sts, BtOpOpen, "�\���}�X�^")
                Exit Function
        End Select
    Loop
    
    wP_COMPO_Open = False

End Function

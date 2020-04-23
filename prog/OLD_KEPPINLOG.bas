Attribute VB_Name = "OLD_KEPPINLOG"
Option Explicit
'********************************************************************
'*
'*              ���i�h�~�x�����O�@�t�@�C����`
'*
'*          CREATE 2004.05.08
'********************************************************************
'�t�@�C���h�c
Public Const OLD_KEPPINLOG_ID$ = "OLD_KEPPINLOG"

'�y�[�W�T�C�Y
Public Const OLD_KEPPINLOG_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public OLD_KEPPINLOG_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                              *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type OLD_KEPPINLOGREC_Tag
    JGYOBU(0 To 0)          As Byte     '���ƕ��敪
    NAIGAI(0 To 0)          As Byte     '�����O
    HIN_GAI(0 To 12)        As Byte     '�i�ԁi�O���j
    CREATE_DT(0 To 7)       As Byte     '�쐬���t
    FILLER(0 To 16)         As Byte     'FILLER
End Type

'�f�[�^�E�o�b�t�@
Public OLD_KEPPINLOGREC     As OLD_KEPPINLOGREC_Tag

'�L�[��`
Private Type KEY0_OLD_KEPPINLOG     '�j�d�x�O
    JGYOBU(0 To 0)      As Byte     '���ƕ��敪
    NAIGAI(0 To 0)      As Byte     '�����O
    HIN_GAI(0 To 12)    As Byte     '�i�ԁi�O���j
End Type


'�L�[�E�f�[�^
Public K0_OLD_KEPPINLOG As KEY0_OLD_KEPPINLOG



Function OLD_KEPPINLOG_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              ���i�h�~�x�����O�@�n�o�d�m                          *
'*                                                                  *
'*      ��  ��:Open Mode(Btrieve�Q��)                               *
'*      �߂�l:false ����                                           *
'*             true  �ُ�                                           *
'*                                                                  *
'*          CREATE 2004.05.08                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    OLD_KEPPINLOG_Open = True
                                            '���i�h�~�x�����O�t���p�X�捞��
    sts = GetIni("FILE", OLD_KEPPINLOG_ID, "CONV2006", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "CONV2006.INI[OLD_KEPPINLOG] �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)
    
    Do
        sts = BTRV(BtOpOpen, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
            
                OLD_KEPPINLOG_Open = sts
                Exit Function
            Case Else
                Call File_Error(sts, BtOpOpen, "�i���j���i�h�~�x�����O")
                Exit Function
        End Select
    Loop

    OLD_KEPPINLOG_Open = False

End Function



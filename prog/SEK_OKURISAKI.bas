Attribute VB_Name = "SEK_OKURISAKI"
Option Explicit
'********************************************************************
'*
'*              �ϐ������@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const SEK_OKURISAKI_ID$ = "SEK_OKURISAKI"

'�y�[�W�T�C�Y
Public Const SEK_OKURISAKI_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public SEK_OKURISAKI_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SEK_OKURISAKIREC_Tag
    OKURISAKI_CD(0 To 7)        As Byte     '���Ӑ�R�[�h
    
    MUKE_NAME(0 To 39)          As Byte     '���Ӑ於��


    JYUSHO(0 To 159)            As Byte     '�Z��       2009.11.19
    
    TEL_NO(0 To 19)             As Byte     '�d�b�ԍ�   2010.01.21
    YUBIN_NO(0 To 7)            As Byte     '�X�֔ԍ�   2010.01.21



    FILLER(0 To 147)            As Byte     'FILLER





End Type

'�f�[�^�E�o�b�t�@
Public SEK_OKURISAKIREC         As SEK_OKURISAKIREC_Tag

'�L�[��`
Type KEY0_SEK_OKURISAKI                     '�j�d�x�O
    OKURISAKI_CD(0 To 7)        As Byte     '���Ӑ�R�[�h
End Type


'�L�[�E�f�[�^
Public K0_SEK_OKURISAKI         As KEY0_SEK_OKURISAKI

Type SEK_OKURISAKI_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
End Type

Private SEK_OKURISAKI_Speck     As SEK_OKURISAKI_FSpeck

Private Function SEK_OKURISAKI_Create() As Integer
'********************************************************************
'*
'*              �ϐ������@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SEK_OKURISAKI_Create = True
                                            '�ϐ������t���p�X�捞��
    sts = GetIni(App.EXEName, SEK_OKURISAKI_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [SEK_OKURISAKI]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    SEK_OKURISAKI_Speck.fs.recoleng = Len(SEK_OKURISAKIREC)     ' ���R�[�h��
    SEK_OKURISAKI_Speck.fs.PageSize = SEK_OKURISAKI_PG_SIZ      ' �y�[�W�T�C�Y
    SEK_OKURISAKI_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    SEK_OKURISAKI_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    SEK_OKURISAKI_Speck.fs.reserve = &H0                        ' �\��ς�
                                                    
'---------------------------------------------------
                                                        ' �L�[�O
    SEK_OKURISAKI_Speck.ks0.keypos = 1                          ' �L�[�|�W�V����
    SEK_OKURISAKI_Speck.ks0.keyleng = 8                         ' �L�[��
    SEK_OKURISAKI_Speck.ks0.keyflag = BtKfExt + BtKfChg         ' �L�[�t���O
    SEK_OKURISAKI_Speck.ks0.keytype = Chr(BtKtString)           ' �L�[�^�C�v
    SEK_OKURISAKI_Speck.ks0.reserve = &H0                       ' �\��ς�

    sts = BTRV(BtOpCreate, SEK_OKURISAKI_POS, SEK_OKURISAKI_Speck, Len(SEK_OKURISAKI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�ϐ������")
        Exit Function
    End If

    SEK_OKURISAKI_Create = False

End Function

Public Function SEK_OKURISAKI_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              �ϐ������@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SEK_OKURISAKI_Open = True
                                            '�ϐ������t���p�X�捞��
    sts = GetIni(App.EXEName, SEK_OKURISAKI_ID, App.EXEName, c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [�ϐ������]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SEK_OKURISAKI_Create()    '�ϐ������쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SEK_OKURISAKI_POS, SEK_OKURISAKIREC, Len(SEK_OKURISAKIREC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�ϐ������")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�ϐ������")
                Exit Function
        End Select
    Loop

    SEK_OKURISAKI_Open = False

End Function

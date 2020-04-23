Attribute VB_Name = "SE_MIN_URIAGE"
Option Explicit
'********************************************************************
'*
'*              �~�j�}���������  �t�@�C����`
'*
'*          CREATE 2008.02.28
'********************************************************************
'�t�@�C���h�c
Public Const SE_MIN_URIAGE_ID$ = "SE_MIN_URIAGE"

'�y�[�W�T�C�Y
Public Const SE_MIN_URIAGE_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public SE_MIN_URIAGE_POS         As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type SE_MIN_URIAGEREC_Tag
    JITU_DATE(0 To 7)           As Byte     '���ѓ��t
    DEN_NO(0 To 7)              As Byte     '�`�[��
    GYO_NO(0 To 2)              As Byte     '�s��
    KEIJYO_YM(0 To 5)           As Byte     '�v��N��(YYYYMM)
    UKEHARAI_CODE(0 To 4)       As Byte     '�󕥐溰��
    
    SE_KBN(0 To 1)              As Byte     '�����敪
    MANA_KBN(0 To 1)            As Byte     '�o�c����
    POST_CODE(0 To 1)           As Byte     '����
    SUB_ITEM(0 To 39)           As Byte     '�������ځi��o�p�j
    SDC_ITEM(0 To 39)           As Byte     '�������ځi�r�c�b�p�j
        
    SURYO(0 To 11)              As Byte     '����   S9(8)V99
    TANKA(0 To 10)              As Byte     '����   9(8)V99
        
    URI_KIN(0 To 8)             As Byte     '���z   S9(9)
    ZEI_KIN(0 To 8)             As Byte     '����� S9(9)
        
    TEKIYO(0 To 39)             As Byte     '�E�v
        
        
    UPD_TANTO(0 To 4)           As Byte     '�X�V�@�S����
    UPD_DATETIME(0 To 13)       As Byte     '�X�V�@����
    
    
    FILLER(0 To 103)            As Byte     'FILLER




End Type
'�f�[�^�E�o�b�t�@
Public SE_MIN_URIAGEREC         As SE_MIN_URIAGEREC_Tag

'�L�[��`

Type KEY0_SE_MIN_URIAGE         '�j�d�x�O
    JITU_DATE(0 To 7)           As Byte     '���ѓ��t
    DEN_NO(0 To 7)              As Byte     '�`�[��
    GYO_NO(0 To 2)              As Byte     '�s��
End Type





'�L�[�E�f�[�^
Public K0_SE_MIN_URIAGE         As KEY0_SE_MIN_URIAGE

Type SE_MIN_URIAGE_FSpeck
    fs      As BtFileSpeck                 ' ̧�� ��߯��\����
    ks0     As BtKeySpeck                 ' �� ��߯��\����
    ks1     As BtKeySpeck
    ks2     As BtKeySpeck
End Type

Private SE_MIN_URIAGE_Speck     As SE_MIN_URIAGE_FSpeck
Private Function SE_MIN_URIAGE_Create() As Integer
'********************************************************************
'*
'*              �~�j�}���������  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    SE_MIN_URIAGE_Create = True
                                            '�~�j�}��������уt���p�X�捞��
    sts = GetIni("FILE", SE_MIN_URIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_MIN_URIAGE]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    SE_MIN_URIAGE_Speck.fs.recoleng = Len(SE_MIN_URIAGEREC)     ' ���R�[�h��
    SE_MIN_URIAGE_Speck.fs.PageSize = SE_MIN_URIAGE_PG_SIZ      ' �y�[�W�T�C�Y
    SE_MIN_URIAGE_Speck.fs.idexnumb = 1                         ' �C���f�b�N�X��
    SE_MIN_URIAGE_Speck.fs.fileflag = 0                         ' �t�@�C���t���O
    SE_MIN_URIAGE_Speck.fs.reserve = &H0                        ' �\��ς�

'-----------------------------------------------
                                                ' �L�[�P
    SE_MIN_URIAGE_Speck.ks0.keypos = 1                  ' �L�[�|�W�V����
    SE_MIN_URIAGE_Speck.ks0.keyleng = 8                 ' �L�[��
    SE_MIN_URIAGE_Speck.ks0.keyflag = BtKfExt + _
                                        BtKfSeg
    SE_MIN_URIAGE_Speck.ks0.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SE_MIN_URIAGE_Speck.ks0.reserve = &H0               ' �\��ς�
                                                
    SE_MIN_URIAGE_Speck.ks1.keypos = 9                  ' �L�[�|�W�V����
    SE_MIN_URIAGE_Speck.ks1.keyleng = 8                 ' �L�[��
    SE_MIN_URIAGE_Speck.ks1.keyflag = BtKfExt + _
                                        BtKfSeg
    SE_MIN_URIAGE_Speck.ks1.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SE_MIN_URIAGE_Speck.ks1.reserve = &H0               ' �\��ς�

                                                
    SE_MIN_URIAGE_Speck.ks2.keypos = 17                  ' �L�[�|�W�V����
    SE_MIN_URIAGE_Speck.ks2.keyleng = 3                 ' �L�[��
    SE_MIN_URIAGE_Speck.ks2.keyflag = BtKfExt
    SE_MIN_URIAGE_Speck.ks2.keytype = Chr(BtKtString)   ' �L�[�^�C�v
    SE_MIN_URIAGE_Speck.ks2.reserve = &H0               ' �\��ς�


'-----------------------------------------------



    sts = BTRV(BtOpCreate, SE_MIN_URIAGE_POS, SE_MIN_URIAGE_Speck, Len(SE_MIN_URIAGE_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "�~�j�}���������")
        Exit Function
    End If

    SE_MIN_URIAGE_Create = False

End Function

Public Function SE_MIN_URIAGE_Open(mode As Integer) As Integer
'********************************************************************
'*
'*              �~�j�}���������  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    SE_MIN_URIAGE_Open = True
                                            '�~�j�}��������уt���p�X�捞��
    sts = GetIni("FILE", SE_MIN_URIAGE_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [SE_MIN_URIAGE]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), ByVal FullPath, Len(FullPath), mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = SE_MIN_URIAGE_Create()    '�~�j�}��������э쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, SE_MIN_URIAGE_POS, SE_MIN_URIAGEREC, Len(SE_MIN_URIAGEREC), ByVal FullPath, Len(FullPath), mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "�~�j�}���������")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "�~�j�}���������")
                Exit Function
        End Select
    Loop

    SE_MIN_URIAGE_Open = False
    

End Function



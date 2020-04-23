Attribute VB_Name = "P_SEISAN_SUM"
Option Explicit

'********************************************************************
'*
'*              ���Y���яW�v�ް�  �t�@�C����`
'*
'*          CREATE 2005.11.11
'********************************************************************
'�t�@�C���h�c
Public Const P_SEISAN_SUM_ID$ = "P_SEISAN_SUM"

'�y�[�W�T�C�Y
Private Const P_SEISAN_SUM_PG_SIZ% = 1024

'�|�W�V�����E�u���b�N
Public P_SEISAN_SUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           �\���̒�`                             *
'*                                                                  *
'********************************************************************
'*************************** ���ږ���` *****************************
Private Type UCHIWAKE_TBL_Tag
    NAI_TANKA(0 To 10)      As Byte         '�����@�P��
    GAI_TANKA(0 To 10)      As Byte         '�O���@�P��
End Type

'���R�[�h��`
Public Type P_SEISAN_SUM_REC_Tag
    
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    CLASS_CODE(0 To 19)     As Byte         '�N���X�i�i�ԁj
        
    GK_NAI_CNT(0 To 4)      As Byte         '�������Y�@����
    GK_NAI_SURYO(0 To 10)   As Byte         '�������Y  ����
    GK_GAI_CNT(0 To 4)      As Byte         '�O�����Y�@����
    GK_GAI_SURYO(0 To 10)   As Byte         '�O�����Y  ����
                                               
    GK_TANKA(0 To 10)       As Byte         '���v�P��
    
                                            '���Y�@����
    UCHIWAKE_TBL(0 To 2)    As UCHIWAKE_TBL_Tag

    KO_GENKA(0 To 10)       As Byte         '���@����
    GA_GENKA(0 To 10)       As Byte         '�O���@����
    GK_GENKA(0 To 10)       As Byte         '�O���H��

End Type
'�f�[�^�E�o�b�t�@
Public P_SEISAN_SUM_REC     As P_SEISAN_SUM_REC_Tag

'�L�[��`
Public Type KEY0_P_SEISAN_SUM               '�j�d�x�O
    SHIMUKE_CODE(0 To 1)    As Byte         '�d������
    CLASS_CODE(0 To 19)     As Byte         '�N���X�i�i�ԁj
End Type
    
'�L�[�E�f�[�^
Public K0_P_SEISAN_SUM      As KEY0_P_SEISAN_SUM

Type P_SEISAN_SUM_FSpeck
    fs                      As BtFileSpeck  ' ̧�� ��߯��\����
    ks0                     As BtKeySpeck   ' �� ��߯��\����
    ks1                     As BtKeySpeck   ' �� ��߯��\����
End Type

Private P_SEISAN_SUM_Speck  As P_SEISAN_SUM_FSpeck
Private Function P_SEISAN_SUM_Create() As Integer
'********************************************************************
'*
'*              ���Y���яW�v�ް�  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_SEISAN_SUM_Create = True
                                            '���Y���яW�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_SEISAN_SUM_Speck.fs.recoleng = Len(P_SEISAN_SUM_REC)  ' ���R�[�h��
    P_SEISAN_SUM_Speck.fs.PageSize = P_SEISAN_SUM_PG_SIZ    ' �y�[�W�T�C�Y
    P_SEISAN_SUM_Speck.fs.idexnumb = 1                      ' �C���f�b�N�X��
    P_SEISAN_SUM_Speck.fs.fileflag = 0                      ' �t�@�C���t���O
    P_SEISAN_SUM_Speck.fs.reserve = &H0                     ' �\��ς�
    '--------------------------------------------------- �L�[�O ��
    P_SEISAN_SUM_Speck.ks0.keypos = 1                       ' �L�[�|�W�V����
    P_SEISAN_SUM_Speck.ks0.keyleng = 2                      ' �L�[��
    P_SEISAN_SUM_Speck.ks0.keyflag = BtKfExt + BtKfSeg      ' �L�[�t���O
    P_SEISAN_SUM_Speck.ks0.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_SEISAN_SUM_Speck.ks0.reserve = &H0                    ' �\��ς�
    
    
    P_SEISAN_SUM_Speck.ks1.keypos = 3                       ' �L�[�|�W�V����
    P_SEISAN_SUM_Speck.ks1.keyleng = 20                     ' �L�[��
    P_SEISAN_SUM_Speck.ks1.keyflag = BtKfExt                ' �L�[�t���O
    P_SEISAN_SUM_Speck.ks1.keytype = Chr(BtKtString)        ' �L�[�^�C�v
    P_SEISAN_SUM_Speck.ks1.reserve = &H0                    ' �\��ς�
    
    '--------------------------------------------------- �L�[�O ��
    
    
    
    sts = BTRV(BtOpCreate, P_SEISAN_SUM_POS, P_SEISAN_SUM_Speck, Len(P_SEISAN_SUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���Y���яW�v�ް�")
        Exit Function
    End If
    
    P_SEISAN_SUM_Create = False

End Function

Public Function P_SEISAN_SUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���Y���яW�v�ް�  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_SEISAN_SUM_Open = True
                                            '���Y���яW�v�ް��t���p�X�捞��
    sts = GetIni("FILE", P_SEISAN_SUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_SEISAN_SUM]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_SEISAN_SUM_Create() '���Y���яW�v�ް��쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���Y���яW�v�ް�")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���Y���яW�v�ް�")
                Exit Function
        End Select
    Loop
    
    P_SEISAN_SUM_Open = False

End Function


Attribute VB_Name = "PN_M"
Option Explicit
'********************************************************************
'*
'*              PN�}�X�^ �t�@�C����`
'*
'*          CREATE 2009.05.29
'********************************************************************
'�t�@�C���h�c
Public Const PN_M_ID = "PN_M"

'�y�[�W�T�C�Y
Public Const PN_M_PG_SIZ% = 4096

'�|�W�V�����E�u���b�N
Public PN_M_POS As POSBLK
'=
'=
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type PN_MREC_Tag
    JCode(0 To 7)           As Byte     '���Ə�R�[�h
    ShisanJCode(0 To 7)     As Byte     '���Y�Ǘ����Ə�R�[�h
    PN(0 To 19)             As Byte     '�i�ڔԍ�
    DModel(0 To 2)          As Byte     '��\�@��i�ڃR�[�h
    HINMOKU(0 To 2)         As Byte     '�i�ڃR�[�h
    SOKO(0 To 1)            As Byte     '�q�ɃR�[�h
    KKeitai_10(0 To 9)      As Byte     '���`�ԃR�[�h         2012.03.06 10��
    Size_Kbn(0 To 0)        As Byte     '���i�T�C�Y�敪
    Saisu(0 To 13)          As Byte     '���i����ː�
    TekiLabel(0 To 0)       As Byte     '�K�p�@�탉�x�����s�敪
    KobaiTanto_5(0 To 4)      As Byte     '�w���S���҃R�[�h     2012.03.06�@5��
    UnitKbn(0 To 0)         As Byte     '���j�b�g���i�敪
    NaiKbn(0 To 0)          As Byte     '�����������i�敪
    GaiKbn(0 To 0)          As Byte     '�C�O�������i�敪
    PnBetsu(0 To 39)        As Byte     '�i�ڕʖ�                                       2009.07.14 Byte���g�� 20 �� 40
    PName(0 To 39)          As Byte     '�i�ږ�                                         2009.07.14 Byte���g�� 20 �� 40
    Tanka2(0 To 9)          As Byte     '����P���Q
    Tanka3(0 To 9)          As Byte     '����P���R
    Tanka4(0 To 9)          As Byte     '����P���S
    Loc1(0 To 9)            As Byte     '���P�[�V�����ԍ��P
    Loc2(0 To 9)            As Byte     '���P�[�V�����ԍ��Q
    Loc3(0 To 9)            As Byte     '���P�[�V�����ԍ��R
    SPn(0 To 19)            As Byte     '�H��i�ڔԍ�
    MadeIn(0 To 19)          As Byte     '�����\�����Y��                        2009.07.14 Byte���g�� 10 �� 20
    HyoTan(0 To 9)          As Byte     '�W���P��
    Syutan(0 To 1)                      As Byte     '�I�[����
        INS_ID(0 To 9)                  As Byte     '�o�^ID                                             2009.07.14 �ǉ�
        INS_TM(0 To 11)                 As Byte     '�o�^���� yyyymmddhhmm              2009.07.14 �ǉ�
        UPD_ID(0 To 9)                  As Byte     '�X�VID                                             2009.07.14 �ǉ�
        UPD_TM(0 To 11)                 As Byte     '�X�V���� yyyymmddhhmm              2009.07.14 �ǉ�
    
    
    MadeInCode(0 To 2)      As Byte                 '2010.08.20
    GENSANKOKU(0 To 19)     As Byte                 '���Y���@2012.02.06
        
    
    KKeitai(0 To 13)         As Byte     '���`��               2012.03.07
    KobaiTanto(0 To 7)      As Byte     '�w���S���Һ���         2012.03.07
    NaiModel(0 To 19)       As Byte     '�����@��i�ڔԍ�       2012.03.07
    NaiModelNew(0 To 19)    As Byte     '�����ŐV�@��i�ڔԍ�   2012.03.07
    GaiModel(0 To 19)       As Byte     '�A�o�@��i�ڔԍ�       2012.03.07
    GaiModelNew(0 To 19)    As Byte     '�A�o�ŐV�@��i�ڔԍ�   2012.03.07
    PNameEngA(0 To 39)      As Byte     '�p�� �i�ڕʖ�          2012.03.07
    PNameEng(0 To 39)       As Byte     '�p�� �i�ږ�            2012.03.07
    NaiDisconYm(0 To 5)     As Byte     '���������ŐؔN��       2012.03.07
    GaiDisconYm(0 To 5)     As Byte     '�C�O�����ŐؔN��       2012.03.07
    
    
    
    
    
    
    
    'FILLER(0 To 10)         As Byte
    

End Type
'�f�[�^�E�o�b�t�@
Public PN_MREC           As PN_MREC_Tag


'�L�[��`
Type KEY0_PN_M                       '�j�d�x�O
    JCode(0 To 7)           As Byte     '���Ə�R�[�h
    ShisanJCode(0 To 7)     As Byte     '���Y�Ǘ����Ə�R�[�h
    PN(0 To 19)             As Byte     '�i�ڔԍ�
End Type

Type KEY1_PN_M                       '�j�d�x�P
    JCode(0 To 7)           As Byte     '���Ə�R�[�h
    ShisanJCode(0 To 7)     As Byte     '���Y�Ǘ����Ə�R�[�h
    SPn(0 To 19)            As Byte     '�H��i�ڔԍ�
End Type


'�L�[�E�f�[�^
Public K0_PN_M           As KEY0_PN_M
Public K1_PN_M           As KEY1_PN_M

Private Type PN_M_FSpeck
    fs  As BtFileSpeck              ' ̧�� ��߯��\����
    ks0 As BtKeySpeck               ' �� ��߯��\����
    ks1 As BtKeySpeck               ' �� ��߯��\����
    ks2 As BtKeySpeck               ' �� ��߯��\����
End Type

Private PN_M_Speck    As PN_M_FSpeck
Private Function PN_M_Create() As Integer
'********************************************************************
'*
'*              PN_M�Ǘ��W�v�t�@�C��  �b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.04.22
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PN_M_Create = True
                                            'PN_M�Ǘ��W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", PN_M_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI[PN_M] �ǂݍ��݃G���[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PN_M_Speck.fs.recoleng = Len(PN_MREC)             ' ���R�[�h��
    PN_M_Speck.fs.PageSize = PN_M_PG_SIZ              ' �y�[�W�T�C�Y
    PN_M_Speck.fs.idexnumb = 2                       ' �C���f�b�N�X��
    PN_M_Speck.fs.fileflag = 0                       ' �t�@�C���t���O
    PN_M_Speck.fs.reserve = &H0                      ' �\��ς�

'---------------------------------------------------' �L�[�O
    PN_M_Speck.ks0.keypos = 1                        ' �L�[�|�W�V����
    PN_M_Speck.ks0.keyleng = 8 + 8 + 20              ' �L�[��
                                                     ' �L�[�t���O
    PN_M_Speck.ks0.keyflag = BtKfExt + BtKfChg
    PN_M_Speck.ks0.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PN_M_Speck.ks0.reserve = &H0                     ' �\��ς�

'---------------------------------------------------' �L�[�P
    PN_M_Speck.ks1.keypos = 1                        ' �L�[�|�W�V����
    PN_M_Speck.ks1.keyleng = 8 + 8                     ' �L�[��
    PN_M_Speck.ks1.keyflag = BtKfExt + BtKfSeg + BtKfDup + BtKfChg     ' �L�[�t���O
    PN_M_Speck.ks1.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PN_M_Speck.ks1.reserve = &H0                     ' �\��ς�

    PN_M_Speck.ks2.keypos = 219                      ' �L�[�|�W�V����
    PN_M_Speck.ks2.keyleng = 20                      ' �L�[��
    PN_M_Speck.ks2.keyflag = BtKfExt + BtKfDup + BtKfChg                ' �L�[�t���O
    PN_M_Speck.ks2.keytype = Chr(BtKtString)         ' �L�[�^�C�v
    PN_M_Speck.ks2.reserve = &H0                     ' �\��ς�
    
    
    sts = BTRV(BtOpCreate, PN_M_POS, PN_M_Speck, Len(PN_M_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "PN�}�X�^")
        Exit Function
    End If

    PN_M_Create = False

End Function

Function PN_M_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              PN�}�X�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'*          CREATE 2004.04.22
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PN_M_Open = True
                                            'PN_M�Ǘ��W�v�t�@�C���t���p�X�捞��
    sts = GetIni("FILE", PN_M_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PN_M_POS, PN_MREC, Len(PN_MREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PN_M_Create()        'PN_M�Ǘ��W�v�t�@�C���쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PN_M_POS, PN_MREC, Len(PN_MREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "PN�}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "PN�}�X�^")
                Exit Function
        End Select
    Loop
    PN_M_Open = False
End Function

Public Function PN_M_GET(JG As String, PN As String, Locked As Integer) As Integer
'           ����
'   JG      ���ƕ�
'   PN      �i�ڔԍ�

'   Locked  �f�����k������

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim JC          As String       '���Ə�R�[�h
Dim SC          As String       '���Y�Ǘ����Ə�R�[�h


    PN_M_GET = True
    
    '           ���Ə�R�[�h�@�ݒ�
    JC = String(UBound(PN_MREC.JCode) + 1, "0")
    If GetIni("JCODE", JG, "PN_JCode", JC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] ���ƕ�[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(JC) = "" Then
        JC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    '           ���Y�Ǘ����Ə�R�[�h�@�ݒ�
    SC = String(UBound(PN_MREC.ShisanJCode) + 1, "0")
    If GetIni("ShisanJCode", JG, "PN_JCode", SC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] ���ƕ�[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(SC) = "" Then
        SC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    
    Call UniCode_Conv(K0_PN_M.JCode, JC)        '���Ə�R�[�h
    Call UniCode_Conv(K0_PN_M.ShisanJCode, SC)  '���Y�Ǘ����Ə�R�[�h
    Call UniCode_Conv(K0_PN_M.PN, PN)           '�i�ڔԍ�
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, PN_M_POS, PN_MREC, Len(PN_MREC), K0_PN_M, Len(K0_PN_M), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'MsgBox "�w�肳�ꂽ�f�[�^������܂���B"
                
                
                Call UniCode_Conv(PN_MREC.PN, PN)
                Call UniCode_Conv(PN_MREC.SPn, PN)
                
                                  
                
                Call UniCode_Conv(PN_MREC.PName, "")                '�i�ږ�
                Call UniCode_Conv(PN_MREC.Tanka2, "0000000.00")     '����P���Q
                Call UniCode_Conv(PN_MREC.Tanka3, "0000000.00")     '����P���R
                Call UniCode_Conv(PN_MREC.Tanka4, "0000000.00")     '����P���S

                Call UniCode_Conv(PN_MREC.SPn, "")                  '�H��i�ڔԍ�
                
                
                
                PN_M_GET = BtErrKeyNotFound
                
                
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<PN_M>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "PN_M")
                Exit Function
        End Select
    Loop
    
    PN_M_GET = False

End Function

Public Function PN_M_GET2(JG As String, PN As String, Locked As Integer) As Integer
'           ����
'   JG      ���ƕ�
'   PN      �i�ڔԍ�

'   Locked  �f�����k������

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim JC          As String       '���Ə�R�[�h
Dim SC          As String       '���Y�Ǘ����Ə�R�[�h


    PN_M_GET2 = True
    
    If Trim(PN) = "" Then
        'MsgBox "�i�ڔԍ����󔒁@���@�w�肳�ꂽ�f�[�^������܂���B"
        Exit Function
    End If
    
    
    '           ���Ə�R�[�h�@�ݒ�
    JC = String(UBound(PN_MREC.JCode) + 1, "0")
    If GetIni("JCODE", JG, "PN_JCode", JC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] ���ƕ�[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(JC) = "" Then
        JC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    '           ���Y�Ǘ����Ə�R�[�h�@�ݒ�
    SC = String(UBound(PN_MREC.ShisanJCode) + 1, "0")
    If GetIni("ShisanJCode", JG, "PN_JCode", SC) Then
        Call LOG_OUT(LOG_F, "[PN_JCode.INI] [JCODE] ���ƕ�[" & JG & "] READ ERROR")
        Exit Function
    End If
    If Trim(SC) = "" Then
        SC = String(UBound(PN_MREC.JCode) + 1, "0")
    End If
    
    
    Call UniCode_Conv(K1_PN_M.JCode, JC)        '���Ə�R�[�h
    Call UniCode_Conv(K1_PN_M.ShisanJCode, SC)  '���Y�Ǘ����Ə�R�[�h
    Call UniCode_Conv(K1_PN_M.SPn, PN)          '�H��i�ڔԍ�
    
    com = BtOpGetEqual + Locked
    Do
        sts = BTRV(com, PN_M_POS, PN_MREC, Len(PN_MREC), K1_PN_M, Len(K1_PN_M), 1)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound       '���R�[�h����
                'MsgBox "�w�肳�ꂽ�f�[�^������܂���B"
                
                
                Call UniCode_Conv(PN_MREC.PN, PN)
                Call UniCode_Conv(PN_MREC.SPn, PN)
                
                Call UniCode_Conv(PN_MREC.PName, "")                '�i�ږ�
                Call UniCode_Conv(PN_MREC.Tanka2, "0000000.00")     '����P���Q
                Call UniCode_Conv(PN_MREC.Tanka3, "0000000.00")     '����P���R
                Call UniCode_Conv(PN_MREC.Tanka4, "0000000.00")     '����P���S
                
                
                
                PN_M_GET2 = BtErrKeyNotFound

                
                
                
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<PN_M>" & Chr(13) & Chr(10) & _
                            "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then Exit Function
            Case Else
                Call File_Error(sts, com, "PN_M")
                Exit Function
        End Select
    Loop
    
    PN_M_GET2 = False

End Function


Public Sub Rclr_PN_MREC()
'********************************************************************
'*
'*              PN�}�X�^  ���R�[�h������
'*
'********************************************************************

    Call UniCode_Conv(PN_MREC.JCode, "")                '���Ə�R�[�h
    Call UniCode_Conv(PN_MREC.ShisanJCode, "")          '���Y�Ǘ����Ə�R�[�h
    Call UniCode_Conv(PN_MREC.PN, "")                   '�i�ڔԍ�
    Call UniCode_Conv(PN_MREC.DModel, "")               '��\�@��i�ڃR�[�h
    Call UniCode_Conv(PN_MREC.HINMOKU, "")              '�i�ڃR�[�h
    Call UniCode_Conv(PN_MREC.SOKO, "")                 '�q�ɃR�[�h
    Call UniCode_Conv(PN_MREC.KKeitai, "")              '���`�ԃR�[�h
    Call UniCode_Conv(PN_MREC.Size_Kbn, "")             '���i�T�C�Y�敪
    Call UniCode_Conv(PN_MREC.Saisu, "")                '���i����ː�
    Call UniCode_Conv(PN_MREC.TekiLabel, "")            '�K�p�@�탉�x�����s�敪
    
    Call UniCode_Conv(PN_MREC.KobaiTanto, "")           '�w���S���҃R�[�h
    Call UniCode_Conv(PN_MREC.UnitKbn, "")              '���j�b�g���i�敪
    Call UniCode_Conv(PN_MREC.NaiKbn, "")               '�����������i�敪
    Call UniCode_Conv(PN_MREC.GaiKbn, "")               '�C�O�������i�敪
    Call UniCode_Conv(PN_MREC.PnBetsu, "")              '�i�ڕʖ�
    Call UniCode_Conv(PN_MREC.PName, "")                '�i�ږ�
    Call UniCode_Conv(PN_MREC.Tanka2, "")                '����P���Q
    Call UniCode_Conv(PN_MREC.Tanka3, "")                '����P���R
    Call UniCode_Conv(PN_MREC.Tanka4, "")                '����P���S
    Call UniCode_Conv(PN_MREC.Loc1, "")                '���P�[�V�����ԍ��P
    
    Call UniCode_Conv(PN_MREC.Loc2, "")                '���P�[�V�����ԍ��Q
    Call UniCode_Conv(PN_MREC.Loc3, "")                '���P�[�V�����ԍ��R
    Call UniCode_Conv(PN_MREC.SPn, "")                '�H��i�ڔԍ�
    Call UniCode_Conv(PN_MREC.MadeIn, "")                '�����\�����Y��
    Call UniCode_Conv(PN_MREC.HyoTan, "")                '�W���P��
    Call UniCode_Conv(PN_MREC.Syutan, "")                '�I�[����
    'Call UniCode_Conv(PN_MREC.FILLER, "")                '


End Sub


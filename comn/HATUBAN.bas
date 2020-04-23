Attribute VB_Name = "HATUBN"
Option Explicit
'********************************************************************
'*
'*              ���ԃ}�X�^�@�t�@�C����`
'*
'********************************************************************
'�t�@�C���h�c
Public Const HATUBAN_ID$ = "HATUBAN"

'�y�[�W�T�C�Y
Public Const HATUBAN_PG_SIZ% = 512

'�|�W�V�����E�u���b�N
Public HATUBAN_POS As POSBLK
'********************************************************************
'*
'*                           �\���̒�`
'*
'********************************************************************
'*************************** ���ږ���` *****************************
'���R�[�h��`
Type HATUBANREC_Tag
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
    NYK_KBN(0 To 0)         As Byte         '���ד`�[���敪
    NYK_DEN_NO(0 To 4)      As Byte         '�����ד`�[��
    SYK_KBN(0 To 0)         As Byte         '�o�ד`�[���敪
    SYK_DEN_NO(0 To 4)      As Byte         '���o�ד`�[��
    NYK_ID_KBN(0 To 0)      As Byte         '����ID���敪
    NYK_ID_NO(0 To 7)       As Byte         '������ID��
    SYK_ID_KBN(0 To 0)      As Byte         '�o��ID���敪
    SYK_ID_NO(0 To 10)      As Byte         '���o��ID��         2006.05.23 7-->11

    OPC_ID_KBN(0 To 0)      As Byte         '���PCID���敪     2006.12.11
    OPC_ID_NO(0 To 5)       As Byte         '���PC���o��ID��   2006.12.11

    OPC_DEN_KBN(0 To 0)     As Byte         '���PC�`�[���敪   2006.12.11
    OPC_DEN_NO(0 To 5)      As Byte         '���PC�`�[��       2006.12.11

    OPC_SYU_NO(0 To 11)     As Byte         '���PC�o�ɕ\��     2007.03.15


    FILLER(0 To 19)         As Byte         'FILLER             2006.12.11
End Type

'�f�[�^�E�o�b�t�@
Public HATUBANREC           As HATUBANREC_Tag

'�L�[��`
Type KEY0_HATUBAN            '�j�d�x�O
    JGYOBU(0 To 0)          As Byte         '���ƕ��敪
End Type

'�L�[�E�f�[�^
Public K0_HATUBAN           As KEY0_HATUBAN

Type HATUBAN_FSpeck
    fs      As BtFileSpeck                  '̧�� ��߯��\����
    ks0     As BtKeySpeck                   '�� ��߯��\����
End Type

Private HATUBAN_Speck As HATUBAN_FSpeck

Private Function HATUBAN_Create() As Integer
'********************************************************************
'*
'*              ���ԃ}�X�^�@�b�q�d�`�s�d
'*
'*      ��  ��:�Ȃ�
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HATUBAN_Create = True
                                            '���ԃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HATUBAN]�ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    HATUBAN_Speck.fs.recoleng = Len(HATUBANREC)     ' ���R�[�h��
    HATUBAN_Speck.fs.PageSize = HATUBAN_PG_SIZ      ' �y�[�W�T�C�Y
    HATUBAN_Speck.fs.idexnumb = 1                   ' �C���f�b�N�X��
    HATUBAN_Speck.fs.fileflag = 0                   ' �t�@�C���t���O
    HATUBAN_Speck.fs.reserve = &H0                  ' �\��ς�
                                                    ' �L�[�O
    HATUBAN_Speck.ks0.keypos = 1                    ' �L�[�|�W�V����
    HATUBAN_Speck.ks0.keyleng = 1                   ' �L�[��
    HATUBAN_Speck.ks0.keyflag = BtKfExt             ' �L�[�t���O
    HATUBAN_Speck.ks0.keytype = Chr(BtKtString)     ' �L�[�^�C�v
    HATUBAN_Speck.ks0.reserve = &H0                 ' �\��ς�

    sts = BTRV(BtOpCreate, HATUBAN_POS, HATUBAN_Speck, Len(HATUBAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "���ԃ}�X�^")
        Exit Function
    End If

    HATUBAN_Create = False

End Function

Public Function HATUBAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ���ԃ}�X�^�@�n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HATUBAN_Open = True
                                            '���ԃ}�X�^�t���p�X�捞��
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [HATUBAN]�ǂݍ��݃G���[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HATUBAN_Create()        '���ԃ}�X�^�쐬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "���ԃ}�X�^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "���ԃ}�X�^")
                Exit Function
        End Select
    Loop

    HATUBAN_Open = False

End Function

Public Function Den_No_Set_Proc(Mode As Integer, JGYOBU As String, DEN_NO As String, Optional MSG As Integer = 1, Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      �u�o�ׁ^�o�ɏ��� ���ׁ^���ɏ��� ���ʁv
'*          �v��O�`�[ �`�[�����ԏ���
'*          ���o�b�`�[���ǉ�      2006.12.11
'*          ���o�b�o�ɕ\���ǉ�    2007.03.15
'*
'*  �v��O�̓`�[���̎捞��
'*  (���ԃ}�X�^��OPEN/CLOSE�͌Ăь���)
'*  �����F  ���[�h�i�ȗ��s�� 10:���ד`�[�� 11:���׃e�L�X�g���@20:�o�ד`�[�� 21:�o�ׂh�c�� 30:���PC�o�ד`�[�� 31:���PC�o��ID�� 32:���PC�o�ɕ\���j
'*          ���ƕ�(�ȗ��s��)
'*          �`�[��(�ȗ��s��)
'*          ���b�Z�[�W�\��(�ȗ��@0:�\�������@1:�\���L��)
'*          ���g���C(���g���C��(0�`99 0:����))
'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer
Dim wk_No       As Long
Dim W_Cnt       As Integer
    
Dim NYU_KBN     As String * 1
Dim SYU_KBN     As String * 1

Dim NYU_ID_KBN  As String * 1
Dim SYU_ID_KBN  As String * 1

Dim OPC_ID_KBN  As String * 1
Dim OPC_DEN_KBN As String * 1


Dim c           As String * 128

    
    Den_No_Set_Proc = True
    
    DEN_NO = ""
    W_Cnt = 0
    '*------------------------------------------------------'���ԃ}�X�^�ǂݍ���
    Call UniCode_Conv(K0_HATUBAN.JGYOBU, JGYOBU)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ԃ}�X�^", 0)
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    If com = BtOpInsert Then
                                                            '��P���ڂ̋敪
        If GetIni("DEN_KBN", "NYU_DEN_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_DEN_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "NYU_ID_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_ID_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_ID_KBN", "SYS", c) Then
            Call LOG_OUT(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_ID_KBN = Trim(c)
        
        
        '���o�b �ǉ�  2006.12.11
        If GetIni("DEN_KBN", "OSAKA_ID_KBN", "SYS", c) Then
            OPC_ID_KBN = ""
        Else
            OPC_ID_KBN = Trim(c)
        End If

        If GetIni("DEN_KBN", "OSAKA_DEN_KBN", "SYS", c) Then
            OPC_DEN_KBN = ""
        Else
            OPC_DEN_KBN = Trim(c)
        End If


        
        
        
        Call UniCode_Conv(HATUBANREC.JGYOBU, JGYOBU)            '���ƕ�
        Call UniCode_Conv(HATUBANREC.NYK_KBN, NYU_KBN)          '���ד`�[�敪
        Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, "00000")       '���ד`�[��
        Call UniCode_Conv(HATUBANREC.SYK_KBN, SYU_KBN)          '�o�ד`�[�敪
        Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, "00000")       '�o�ד`�[��
        
        Call UniCode_Conv(HATUBANREC.NYK_ID_KBN, NYU_ID_KBN)    '���ׂh�c�敪
        Call UniCode_Conv(HATUBANREC.NYK_ID_NO, "00000000")     '���׃e�L�X�g��
        Call UniCode_Conv(HATUBANREC.SYK_ID_KBN, SYU_ID_KBN)    '�o�ׂh�c�敪
        Call UniCode_Conv(HATUBANREC.SYK_ID_NO, "00000000000")  '�o�ׂh�c��
        
        
        
        '���PC 2006.12.17
        If Trim(OPC_ID_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, OPC_ID_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "000000")
        
        End If
        
        If Trim(OPC_DEN_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, OPC_DEN_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "000000")
        
        End If
        
        
        Call UniCode_Conv(HATUBANREC.OPC_SYU_NO, "000000000000")
        
        Call UniCode_Conv(HATUBANREC.FILLER, "")
    End If
    
    Select Case Mode
        Case 10
                                    '���ד`�[��
            If StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, Format(wk_No, "00000"))
    
        Case 11
                                    '���ׂh�c��
            If StrConv(HATUBANREC.NYK_ID_NO, vbUnicode) = "99999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000")
            Call UniCode_Conv(HATUBANREC.NYK_ID_NO, Format(wk_No, "00000000"))
                                
        Case 20
                                '�o�ד`�[��
            If StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, Format(wk_No, "00000"))
        Case 21
                                    '�o�ׂh�c��
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "99999999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000000")
            Call UniCode_Conv(HATUBANREC.SYK_ID_NO, Format(wk_No, "00000000000"))
    
        Case 31
                                    '���h�c��
            If GetIni("DEN_KBN", "SYU_ID_KBN", "SYS", c) Then
            
                SYU_ID_KBN = ""
            Else
                SYU_ID_KBN = Trim(c)
            
            End If
    
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "999999" Then
                wk_No = 1
            Else
                
                If Not IsNumeric(StrConv(HATUBANREC.OPC_ID_NO, vbUnicode)) Then
                    wk_No = 1
                Else
                    wk_No = CLng(StrConv(HATUBANREC.OPC_ID_NO, vbUnicode)) + 1
                End If
            End If
        
            
            If Trim(StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode)) = "" Then
                Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, SYU_ID_KBN)
            End If
            DEN_NO = StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode) & Format(wk_No, "000000")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, Format(wk_No, "000000"))
        Case 32
            If GetIni("DEN_KBN", "SYU_DEN_KBN", "SYS", c) Then
                SYU_KBN = ""
            Else
                SYU_KBN = Trim(c)
            End If
            
            If Trim(StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode)) = "" Then
                Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, SYU_KBN)
            End If
            
            If StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode) = "999999" Then
                wk_No = 1
            Else
                If Not IsNumeric(StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode)) Then
                    wk_No = 1
                Else
                    wk_No = CLng(StrConv(HATUBANREC.OPC_DEN_NO, vbUnicode)) + 1
                End If
            End If
        
            DEN_NO = StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode) & Format(wk_No, "000000")
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, Format(wk_No, "000000"))
            
        Case 33
                                    '���o�ɕ\��
            If StrConv(HATUBANREC.OPC_SYU_NO, vbUnicode) = "999999999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.OPC_SYU_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = Format(wk_No, "000000000000")
            Call UniCode_Conv(HATUBANREC.OPC_SYU_NO, Format(wk_No, "000000000000"))
    
    
    End Select
    '*------------------------------------------------------'���ԃ}�X�^�o��
    W_Cnt = 0
    Do
        sts = BTRV(com, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, com, "���ԃ}�X�^")
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    Den_No_Set_Proc = False          '����I��

End Function



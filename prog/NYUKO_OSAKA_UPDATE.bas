Attribute VB_Name = "NYUKO_OSAKA_UPDATE"
Option Explicit

Public DAITO_SOKO_NO       As String * 2


Public Function Nyuko_OSAKA_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    GYO_INS As String, _
                                    DEN_NO As String, _
                                    SEQ_NO As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional MENU_NO As String = "  ") As Integer
'****************************************************
'*      �u���ׁ^���ɏ����v�݌Ƀf�[�^�X�V
'*
'*  �݌Ƀf�[�^�̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  ���g�����U�N�V�����������K�v�ȏꍇ�͌Ăь��ōs����
'*  �g�p̧��    :   �݌Ƀf�[�^
'*                  �i�ڃ}�X�^
'*                  �v���}�X�^
'*                  �݌Ɉړ���
'*                  ���׎���
'*                  �q�Ƀ}�X�^
'*
'*  �����F  ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          TO��iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          ���ד�(YYYYMMDD �ȗ��s��)
'*          �v��(�ȗ��s��)
'*          ���i���ςݎ��ѐ��i���ꂩ����K�{�j
'*          �����i���ѐ��@�@�i�@�@�V�@�@�@�@�j
'*          ID(�ȗ��s��)
'*          �S���ҁi�ȗ��s�j
'*          ���ɍ쐬(�ȗ��s�� 0:�X�V 1:�ǉ�)
'*          �`�[��(�ȗ��s��)
'*          SEQNO(�ȗ��s�@ں��އ�)
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          ����(�ȗ��� �����ɏo�͂��郁�����e)
'*          �ƭ���ٰ�߁i�����Ǘ����ځj  2006.01.30
'*  �߂�l: false       :����
'*          true        :�p���\�Ȉُ�
'*          SYS_ERR     :�p���ł��Ȃ��ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts             As Integer
Dim com             As Integer

Dim RETRY_CNT       As Integer
Dim MESG_FLG        As Integer
Dim RETRY_SU        As Integer
    
Dim ans             As Integer
    
Dim Ins_DateTime    As String * 14                  '2004.12.09
    
    
    Nyuko_OSAKA_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")    '2004.12.09
    '*------------------------------------------------------'�i��Ͻ��̊m��
    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)               '���ƕ�
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)               '���O
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)             '�i�ԁi�O���j
        
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                If MESG_FLG = 1 Then
                    Beep
                    MsgBox "�i�ڃR�[�h�����݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                End If
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents                                                       '2016.01.26
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Nyuko_OSAKA_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
'============================================================
'************************************************************ ���i���ςݍX�V
    If SUMI_JITU_QTY <> 0 Then
    '*------------------------------------------------------'�݌Ƀf�[�^�Ǎ���
        Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '�q�ɇ�
        Call UniCode_Conv(K0_ZAIKO.Retu, Mid(TO_LOCATION, 3, 2))    '��
        Call UniCode_Conv(K0_ZAIKO.Ren, Mid(TO_LOCATION, 5, 2))     '�A
        Call UniCode_Conv(K0_ZAIKO.Dan, Mid(TO_LOCATION, 7, 2))     '�i
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                  '���ƕ�
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                  '���O
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                '�i�ԁi�O���j
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                   '���i�^�����i
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)              '���ד�
    
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌��ް�", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents                                                       '2016.01.26
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
            
        Loop
                                                                                
        If com = BtOpInsert Then
                                        '�V�K�ǉ�
            Call UniCode_Conv(ZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '�q�ɇ�
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))    '��
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))     '�A
            Call UniCode_Conv(ZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))     '�i
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                  '���ƕ�
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                  '���O
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                '�i�ԁi�O���j
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                   '���i�^�����i
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, NYUKA_DT)              '���ד�
                                                                        '���ɓ�
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, Format(Date, "yyyymmdd"))
                                                                        '�i�ԁi�����j
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                        '�L���݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(SUMI_JITU_QTY, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
                                                                        '���i�����t
            Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '�v��N��
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '�݌ɐ��X�V
                                                                        '�݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + SUMI_JITU_QTY, "00000000"))
                                                                    
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '�v��N��
        
        
        End If
    
        RETRY_CNT = 0
    '*------------------------------------------------------'�݌Ƀf�[�^�o��
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                            Call File_Error(sts, com, "�݌��ް�", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o��
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    SUMI_JITU_QTY, _
                                    0, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    , , , MENU_NO, , , , , DEN_NO, Format(Now, "YYYYMMDD"))
        If sts Then
            Nyuko_OSAKA_Update_Proc = sts
            Exit Function
        End If
    End If
'************************************************************ �����i�X�V
    If MI_JITU_QTY <> 0 Then
    '*------------------------------------------------------'�݌Ƀf�[�^�Ǎ���
        Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '�q�ɇ�
        Call UniCode_Conv(K0_ZAIKO.Retu, Mid(TO_LOCATION, 3, 2))    '��
        Call UniCode_Conv(K0_ZAIKO.Ren, Mid(TO_LOCATION, 5, 2))     '�A
        Call UniCode_Conv(K0_ZAIKO.Dan, Mid(TO_LOCATION, 7, 2))     '�i
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                  '���ƕ�
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                  '���O
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                '�i�ԁi�O���j
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                   '���i�^�����i
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, NYUKA_DT)              '���ד�
    
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    com = BtOpInsert
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌��ް�", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
            
        Loop
                                                                                
        If com = BtOpInsert Then
                                        '�V�K�ǉ�
            Call UniCode_Conv(ZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2)) '�q�ɇ�
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))    '��
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))     '�A
            Call UniCode_Conv(ZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))     '�i
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                  '���ƕ�
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                  '���O
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                '�i�ԁi�O���j
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                   '���i�^�����i
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, NYUKA_DT)              '���ד�
                                                                        '���ɓ�
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, Format(Date, "yyyymmdd"))
                                                                        '�i�ԁi�����j
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                        '�L���݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(MI_JITU_QTY, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '�v��N��
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '�݌ɐ��X�V
                                                                        '�݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + MI_JITU_QTY, "00000000"))
        
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                 '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                   '�v��N��
        
        
        End If
    
        RETRY_CNT = 0
    '*------------------------------------------------------'�݌Ƀf�[�^�o��
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                            Call File_Error(sts, com, "�݌��ް�", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                    
                        End If
                
                    End If
                
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o��
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    0, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    , , , MENU_NO, , , , , DEN_NO, Format(Now, "YYYYMMDD"))
        If sts Then
            Nyuko_OSAKA_Update_Proc = sts
            Exit Function
        End If
    End If
'============================================================
'============================================================
                                        '�q��Ͻ��Ǎ���
    Call UniCode_Conv(K0_SOKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound           '�L��Ƃ܂������G���[�ɂ��Ȃ�
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�")
            Nyuko_OSAKA_Update_Proc = SYS_ERR
            Exit Function
    End Select
    
    If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                                        '�W���I��
'        If Last_JGYOBU = SOJIKI Or _
'            Last_JGYOBU = SENTAKU Then
                                        '�|���@�͐ݒ�ς݂̏㏑�������Ȃ�
            If StrConv(ITEMREC.ST_SET_DT, vbUnicode) = Space(8) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "yyyymmdd"))
                Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(TO_LOCATION, 1, 2))
                Call UniCode_Conv(ITEMREC.ST_RETU, Mid(TO_LOCATION, 3, 2))
                Call UniCode_Conv(ITEMREC.ST_REN, Mid(TO_LOCATION, 5, 2))
                Call UniCode_Conv(ITEMREC.ST_DAN, Mid(TO_LOCATION, 7, 2))
            End If
'        Else
'            Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "yyyymmdd"))
'            Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(TO_LOCATION, 1, 2))
'            Call UniCode_Conv(ITEMREC.ST_RETU, Mid(TO_LOCATION, 3, 2))
'            Call UniCode_Conv(ITEMREC.ST_REN, Mid(TO_LOCATION, 5, 2))
'            Call UniCode_Conv(ITEMREC.ST_DAN, Mid(TO_LOCATION, 7, 2))
'        End If
                                        '�O����ɒI
        Call UniCode_Conv(ITEMREC.BEF_SOKO, Mid(TO_LOCATION, 1, 2))
        Call UniCode_Conv(ITEMREC.BEF_RETU, Mid(TO_LOCATION, 3, 2))
        Call UniCode_Conv(ITEMREC.BEF_REN, Mid(TO_LOCATION, 5, 2))
        Call UniCode_Conv(ITEMREC.BEF_DAN, Mid(TO_LOCATION, 7, 2))
    End If
                                        '�ŏI���ɓ�
    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, Format(Date, "yyyymmdd"))
                                        '�ŏI���ד��t
    If StrConv(ITEMREC.LAST_INP_DT, vbUnicode) < NYUKA_DT Then
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, NYUKA_DT)
    End If
    '*------------------------------------------------------'�i�ڃ}�X�^�o��
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^", 0)
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                    
                End If
                
                If MESG_FLG = 0 Then
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Nyuko_OSAKA_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================

'============================================================   ���׃f�[�^�X�V�^�쐬
    If GYO_INS = "9" Then '2007.09.12
    Else
    
        If GYO_INS = "0" Then
            com = BtOpUpdate
        
            Call UniCode_Conv(K0_Y_NYU_O.SEQ_NO, SEQ_NO)
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "���ח\�肪���݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ח\��", 0)
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ח\��", 0)
                                Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ח\��")
                        Nyuko_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
        
        
        
        
        Else
            com = BtOpInsert
            sts = BTRV(BtOpGetLast, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            Select Case sts
                Case BtNoErr
                    SEQ_NO = StrConv(Y_NYU_O_REC.SEQ_NO, vbUnicode)
                Case BtErrEOF
                    SEQ_NO = "000"
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ח\���ް�")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
            SEQ_NO = Format(CInt(SEQ_NO) + 1, "000")
                
            Call UniCode_Conv(Y_NYU_O_REC.JGYOBU, JGYOBU)
            Call UniCode_Conv(Y_NYU_O_REC.SOKO_NO, DAITO_SOKO_NO)
            Call UniCode_Conv(Y_NYU_O_REC.SEQ_NO, SEQ_NO)
            Call UniCode_Conv(Y_NYU_O_REC.NYUKO_YMD, Format(Now, "YYYYMMDD"))
            Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, "000000")
        
            Call UniCode_Conv(Y_NYU_O_REC.MAKER_CODE, StrConv(ITEMREC.MAKER_CODE, vbUnicode))
            Call UniCode_Conv(Y_NYU_O_REC.NAIGAI, NAIGAI)
        
            Call UniCode_Conv(Y_NYU_O_REC.HIN_NO, HIN_GAI)
        
            Call UniCode_Conv(Y_NYU_O_REC.Y_SURYO, "00000000")
        
            Call UniCode_Conv(Y_NYU_O_REC.ORDER_NO, "")
        
            Call UniCode_Conv(Y_NYU_O_REC.FILLER, "")
        End If
    
        If Trim(DEN_NO) <> "" And DEN_NO <> "000000" Then
            Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, DEN_NO)
        End If
        Call UniCode_Conv(Y_NYU_O_REC.J_SURYO, Format(MI_JITU_QTY + SUMI_JITU_QTY, "00000000"))
    
        Call UniCode_Conv(Y_NYU_O_REC.TANTO_CODE, TANTO_CODE)
        
        Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, "1")
        
        Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
        Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
        
            
        RETRY_CNT = 0
        Do
            sts = BTRV(com, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                        
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                            '�񐔃I�[�o�[
                            Call File_Error(sts, com, "���ח\��", 0)
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    
                    End If
                    
                    If MESG_FLG = 0 Then
'                        DoEvents
                        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                            DoEvents                                                    '2016.01.26
                        End If                                                          '2016.01.26
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Nyuko_OSAKA_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "���ח\��")
                    Nyuko_OSAKA_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    
    End If
    
    
    Nyuko_OSAKA_Update_Proc = False
    
    
End Function

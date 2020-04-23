Attribute VB_Name = "NYUKO_UPDATE"
Option Explicit

Public Function Nyuko_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional SHIIRE_CODE As String = "     ", _
                                    Optional SHIIRE_TANKA As String = "           ", _
                                    Optional KEIJYO_YM As String = "      ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0, _
                                    Optional KAMOKU_FURIKAE As String = "  ", _
                                    Optional GENSANKOKU As String = "                    ", _
                                    Optional SHIIRE_WORK_CENTER As String = "        ", _
                                    Optional ID_NO2 As String = "            ", _
                                    Optional YOSAN_FROM As String = "     ", _
                                    Optional YOSAN_TO As String = "     ", _
                                    Optional wkMTS As String = "        ") As Integer
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
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          ����(�ȗ��� �����ɏo�͂��郁�����e)
'*          �d���溰�ށi���ޗp�ȗ��@2006.01.05�j
'*          �d���P���i���ޗp�ȗ��@2006.01.05�j
'*          �v��N���i���ޗp�ȗ��@2006.01.05�j
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
    
    
    Nyuko_Update_Proc = True
                                                                      
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
                        Nyuko_Update_Proc = SYS_CANCEL
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
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            
            
            Case BtErrDEAD_LOCK
                Nyuko_Update_Proc = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Nyuko_Update_Proc = SYS_ERR
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
                            Nyuko_Update_Proc = SYS_CANCEL
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
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Nyuko_Update_Proc = SYS_ERR
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
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '�v��N��
            
            
            
            
            '------------   2010.07.08 ��
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '���Y��
                                                                        '���ގd����ܰ�����
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '�\�Z�@��
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '�\�Z�@��
            '------------   2010.07.08 ��
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '�݌ɐ��X�V
                                                                        '�݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + SUMI_JITU_QTY, "00000000"))
                                                                    
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '�v��N��
        
        
        
        
            '------------   2010.07.08 ��
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '���Y��
                                                                        '���ގd����ܰ�����
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '�\�Z�@��
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '�\�Z�@��
            '------------   2010.07.08 ��
        
        
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
                            Nyuko_Update_Proc = SYS_CANCEL
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
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o��
'        sts = IDOREKI_OUTPUT_PROC(Space(8), _
'                                    TO_LOCATION, _
'                                    JGYOBU, _
'                                    NAIGAI, _
'                                    HIN_GAI, _
'                                    NYUKA_DT, _
'                                    YOIN, _
'                                    SUMI_JITU_QTY, _
'                                    0, _
'                                    ID, _
'                                    TANTO_CODE, _
'                                    RETRY, , _
'                                    MEMO, _
'                                    Ins_DateTime, _
'                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO)
'        If sts Then
'            Nyuko_Update_Proc = sts
'            Exit Function
'        End If
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
                            Nyuko_Update_Proc = SYS_CANCEL
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
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Nyuko_Update_Proc = SYS_ERR
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
            
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '�v��N��
            
            '------------   2010.07.08 ��
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '���Y��
                                                                        '���ގd����ܰ�����
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '�\�Z�@��
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '�\�Z�@��
            '------------   2010.07.08 ��
            
            Call UniCode_Conv(ZAIKOREC.FILLER, "")
        Else
                                        '�݌ɐ��X�V
                                                                    '�݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) + MI_JITU_QTY, "00000000"))
        
            Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, SHIIRE_CODE)        '�d���溰��
            Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, SHIIRE_TANKA)      '�d���P��
            Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, KEIJYO_YM)            '�v��N��
            '------------   2010.07.08 ��
            Call UniCode_Conv(ZAIKOREC.GENSANKOKU, GENSANKOKU)          '���Y��
                                                                        '���ގd����ܰ�����
            Call UniCode_Conv(ZAIKOREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
            Call UniCode_Conv(ZAIKOREC.ID_NO2, ID_NO2)                  'ID_NO
            Call UniCode_Conv(ZAIKOREC.YOSAN_FROM, YOSAN_FROM)          '�\�Z�@��
            Call UniCode_Conv(ZAIKOREC.YOSAN_TO, YOSAN_TO)              '�\�Z�@��
            '------------   2010.07.08 ��
        
        
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
                            Nyuko_Update_Proc = SYS_CANCEL
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
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                
                Case BtErrDEAD_LOCK
                    Nyuko_Update_Proc = SYS_CANCEL
                    Exit Function
                
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Nyuko_Update_Proc = SYS_ERR
                    Exit Function
                
            End Select
        Loop
'============================================================
'    '*------------------------------------------------------'�݌Ɉړ����o��
'        sts = IDOREKI_OUTPUT_PROC(Space(8), _
'                                    TO_LOCATION, _
'                                    JGYOBU, _
'                                    NAIGAI, _
'                                    HIN_GAI, _
'                                    NYUKA_DT, _
'                                    YOIN, _
'                                    0, _
'                                    MI_JITU_QTY, _
'                                    ID, _
'                                    TANTO_CODE, _
'                                    RETRY, , _
'                                    MEMO, _
'                                    Ins_DateTime, _
'                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO)
'        If sts Then
'            Nyuko_Update_Proc = sts
'            Exit Function
'        End If
    End If
'============================================================
    '*------------------------------------------------------'�݌Ɉړ����o�� 2008.08.08 �o�͉ӏ��ړ�
        
    If Trim(KAMOKU_FURIKAE) = "" Then       '2009.06.26
    
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    YOIN, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO, , , , , , , , GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM, YOSAN_TO, wkMTS)
    
    Else
    
        sts = IDOREKI_OUTPUT_PROC(Space(8), _
                                    TO_LOCATION, _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    KAMOKU_FURIKAE, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, , _
                                    MEMO, _
                                    Ins_DateTime, _
                                    SHIIRE_CODE, SHIIRE_TANKA, KEIJYO_YM, MENU_NO, , , , , , , , GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM, YOSAN_TO, wkMTS)
    
    
    End If
    
    
    If sts Then
        Nyuko_Update_Proc = sts
        Exit Function
    End If
    
    
    
'    If YOIN = YOIN_MAEGARI Then        2016.05.30
    If YOIN = YOIN_MAEGARI Or YOIN = WEL_MAEGARI_TANA_S_OSAKA Then  '2016.05.30
    '*------------------------------------------------------'�O�؂�f�[�^�Ǎ���
        Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
            Case SHIZAI
                '���ޑO�؏���
                Call UniCode_Conv(K0_P_NYU.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Now, "YYYYMMDD"))
                
                
                RETRY_CNT = 0
                Do
                                                '�O�؂��ް��Ǎ���
                    sts = BTRV(BtOpGetEqual + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
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
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޑO��", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                        
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ޑO�؃f�[�^")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                    
                Loop
        
                If com = BtOpInsert Then
                                            '�V�K�ǉ�
                                                            '���ƕ�
                    Call UniCode_Conv(P_NYUREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '�����O
                    Call UniCode_Conv(P_NYUREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '�i�ځi�O���j
                    Call UniCode_Conv(P_NYUREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                            '���ד�
                    Call UniCode_Conv(P_NYUREC.NYUKA_DT, Format(Now, "YYYYMMDD"))
                                                            '���ѐ���
                    Call UniCode_Conv(P_NYUREC.NYUKA_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
                                                            '���E���t
                    Call UniCode_Conv(P_NYUREC.SOUSAI_DT, "")
                                                            '���E��
                    Call UniCode_Conv(P_NYUREC.SOUSAI_DT, "00000000")
                                                            '�o�^�[��
                    Call UniCode_Conv(P_NYUREC.WS_ID, ID)
                                        
                                                            '�d����
                    Call UniCode_Conv(P_NYUREC.SHIIRE_CODE, SHIIRE_CODE)
                                            
                    Call UniCode_Conv(P_NYUREC.SHIIRE_TANKA, SHIIRE_TANKA)
                    
                    
                    Call UniCode_Conv(P_NYUREC.FILLER, "")
                
                                                            '�o�^����
                    Call UniCode_Conv(P_NYUREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Else
                                                            '���ѐ���
                    SUMI_JITU_QTY = SUMI_JITU_QTY + MI_JITU_QTY + CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode))
                    Call UniCode_Conv(P_NYUREC.NYUKA_QTY, Format(SUMI_JITU_QTY, "00000000"))
                End If
            '*------------------------------------------------------'�O�؂�f�[�^�o��
                RETRY_CNT = 0
                Do
                    sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then
                                
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '�񐔃I�[�o�[
                                    Call File_Error(sts, com, "���ޑO��", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                            
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        
                        Case Else
                            Call File_Error(sts, com, "���ޑO��")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                                
                    End Select
                Loop
            
            
            
            
            
            
            
            Case Else
                '���i�O�؏���
                Call UniCode_Conv(K0_J_NYU.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_J_NYU.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
                RETRY_CNT = 0
                Do
                                                '�O�؂��ް��Ǎ���
                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
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
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���׎���", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                        
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���׎��уf�[�^")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                    End Select
                    
                Loop
        
                If com = BtOpInsert Then
                                            '�V�K�ǉ�
                                                            '���ƕ�
                    Call UniCode_Conv(J_NYUREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                            '�����O
                    Call UniCode_Conv(J_NYUREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                            '�i�ځi�O���j
                    Call UniCode_Conv(J_NYUREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                            '���ѐ���
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
                    Call UniCode_Conv(J_NYUREC.FILLER, "")
                
                
                    Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))            '2019.01.12
                Else
                                                            '���ѐ���
                    SUMI_JITU_QTY = SUMI_JITU_QTY + MI_JITU_QTY + CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                    Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(SUMI_JITU_QTY, "00000000"))
                End If
            
            
            
                          
            '*------------------------------------------------------'�O�؂�f�[�^�o��
                RETRY_CNT = 0
                Do
                    sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then
                                
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                                '�񐔃I�[�o�[
                                    Call File_Error(sts, com, "���׎���", 0)
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                            
                                End If
                            
                            End If
                        
                            If MESG_FLG = 0 Then
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                            Else
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Nyuko_Update_Proc = SYS_CANCEL
                                    Exit Function
                                End If
                            End If
                        
                        
                        Case BtErrDEAD_LOCK
                            Nyuko_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        Case Else
                            Call File_Error(sts, com, "���׎��уf�[�^")
                            Nyuko_Update_Proc = SYS_ERR
                            Exit Function
                                
                    End Select
                Loop
        End Select
    End If
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
            Nyuko_Update_Proc = SYS_ERR
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
                        Nyuko_Update_Proc = SYS_CANCEL
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
                        Nyuko_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case BtErrDEAD_LOCK
                Nyuko_Update_Proc = SYS_CANCEL
                Exit Function
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Nyuko_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================

    
    Nyuko_Update_Proc = False
    
    
End Function

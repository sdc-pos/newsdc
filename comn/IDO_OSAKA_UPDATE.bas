Attribute VB_Name = "IDO_OSAKA_UPDATE"
Option Explicit

Public Function IDO_OSAKA_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    FROM_LOCATION As String, _
                                    TO_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    DEN_NO As String, _
                                    SEQ_NO As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional MENU_NO As String = "  ") As Integer
'****************************************************
'*      �u�ړ������v�݌Ƀf�[�^�X�V
'*
'*  �݌Ƀf�[�^�̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  �g�p̧��    :   �݌Ƀf�[�^
'*                  �i�ڃ}�X�^
'*                  �v���}�X�^
'*                  �݌Ɉړ���
'*                  �q�Ƀ}�X�^
'*  �����F  ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          FROM��iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          TO��iXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          ���ד�(YYYYMMDD �ȗ��� �ȗ���FIFO)
'*          �v��(�ȗ��s��)
'*          ���i���ςݎ��ѐ��i���ꂩ����K�{�j
'*          �����i���ѐ��@�@�i�@�@�V�@�@�@�@�j
'*          ID(�ȗ��s��)
'*          �S���ҁi�ȗ��s�j
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
Dim sts         As Integer
Dim com         As Integer
Dim Upd_com     As Integer


Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim ans         As Integer
    
Dim Zan_Qty     As Long
Dim WK_Qty      As Long
    
Dim TO_NAIGAI   As String * 1
    
Dim IDO_GOODS_ON_F  As String * 1
Dim IDO_GOODS_YMD   As String * 8
    
Dim Ins_DateTime    As String * 14              '2004.12.09


    IDO_OSAKA_Update_Proc = True
                                                                      
                                                                      
                                                                      
                                                                      
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")    '2004.12.09
                                        
    '*------------------------------------------------------'�i��Ͻ��iFROM���j�̊m��
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                IDO_OSAKA_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop


    '*------------------------------------------------------'���ח\��̊m��
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���ח\��")
                IDO_OSAKA_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
                                        
'*------------------------------------------------------'�q��Ͻ��Ǎ���
    Call UniCode_Conv(K0_SOKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound           '�L��Ƃ܂������G���[�ɂ��Ȃ�
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�")
            IDO_OSAKA_Update_Proc = SYS_ERR
            Exit Function
    End Select

    IDO_GOODS_ON_F = "1"
    IDO_GOODS_YMD = ""
'    If Left(YOIN, 1) = ACT_IDO_OUT Then
    If JGYOBU <> SHIZAI Then
    '���ޕi�͐U�ւ��Ȃ�2006.01.10
        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = "0" Then
            IDO_GOODS_ON_F = "0"
            IDO_GOODS_YMD = Format(Now, "YYYYMMDD")
        End If

    End If
'    End If
'============================================================
    If YOIN = YOIN_FURIKAE Then     '�����O�U�ւ͓��O�𔽓]
        If NAIGAI = NAIGAI_NAI Then
            TO_NAIGAI = NAIGAI_GAI
        Else
            TO_NAIGAI = NAIGAI_NAI
        End If
    Else
        TO_NAIGAI = NAIGAI
    End If
    
    
    '*------------------------------------------------------'���ד��w�薳�� �݌Ƀf�[�^�Ǎ��݁iFROM���̏����j
    '*
    
    
    '���۸ޏo��    '2008.08.06
    If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                ID, _
                                JGYOBU, _
                                NAIGAI, _
                                MENU_NO, _
                                YOIN, _
                                HIN_GAI, _
                                SUMI_JITU_QTY, _
                                MI_JITU_QTY, _
                                FROM_LOCATION, _
                                TO_LOCATION) Then
        Exit Function
    End If
    
    
    
    
    
    
    
    '*--------------------  ���i���ς݂̏���
    If SUMI_JITU_QTY <> 0 Then
    
        Zan_Qty = SUMI_JITU_QTY

        Do

            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM�q�ɇ�
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM��
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM�A
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM�i
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "0")                       '���i�^�����i
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                                            '�I�{�i�u���[�N
                        If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                            StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                            JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                            NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                            StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "0" Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function
                        End If
                    
                        If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            Upd_com = BtOpUpdate
                            WK_Qty = Zan_Qty
                        Else
                            Upd_com = BtOpDelete
                            WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        End If

                        Exit Do
                    Case BtErrEOF

                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Exit Function

                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If Upd_com = BtOpUpdate Then
                                                                            '�L���݌ɐ�
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
            
            End If


            RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop
'============================================================
    '*------------------------------------------------------'���ד��w�薳�� �݌Ƀf�[�^�Ǎ��݁iTO���̏����j
            Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO�q�ɇ�
            Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO��
            Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO�A
            Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO�i
            Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '���ƕ�
            Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '���O
            Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '�i�ԁi�O���j
            Call UniCode_Conv(K0_wZAIKO.GOODS_ON, "0")                      '���i�^�����i
                                                                            '���ד�
            Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr

                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If Upd_com = BtOpInsert Then
                                                '�V�K�ǉ�
                Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '�q�ɇ�
                Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '��
                Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '�A
                Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '�i
                Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '���ƕ�
                Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '���O
                Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '�i�ԁi�O���j
                Call UniCode_Conv(wZAIKOREC.GOODS_ON, "0")                      '���i�^�����i
                                                                                '���ד�
                Call UniCode_Conv(wZAIKOREC.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD")) '���ɓ�
                                                                                '�i�ԁi�����j
                Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                '�L���݌ɐ�
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
                Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '�r���t���O
                Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '�g�p���q�@ID
                Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '�g�p����۸���
                                                                                '�d���溰��2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                '�d����P��2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                '�v��N��2006.01.08
                Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
                
                
                Call UniCode_Conv(wZAIKOREC.FILLER, "")
            Else
                                            '�݌ɐ��X�V
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + WK_Qty, "00000000"))
            End If

            RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop
        '*------------------------------------------------------'�݌Ɉړ����o��
            sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            TO_LOCATION, _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            WK_Qty, 0, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, , MEMO, _
                                            Ins_DateTime, , , , , , , , , DEN_NO, Format(Now, "YYYYMMDD"))
            If sts Then
                IDO_OSAKA_Update_Proc = sts
                Exit Function
            End If

            Zan_Qty = Zan_Qty - WK_Qty

            If Zan_Qty <= 0 Then
                Exit Do                     '�������Ƃ��I��
            End If

        Loop
                
    End If
'================================================================================
    '*
    '*--------------------  �����i���̏���
    If MI_JITU_QTY <> 0 Then
    
        Zan_Qty = MI_JITU_QTY

        Do

            Call UniCode_Conv(K0_ZAIKO.SOKO_NO, Mid(FROM_LOCATION, 1, 2))   'FROM�q�ɇ�
            Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      'FROM��
            Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       'FROM�A
            Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       'FROM�i
            Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
            Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
            Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
            Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "1")                       '���i�^�����i
            Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                                            '�I�{�i�u���[�N
                        If FROM_LOCATION <> (StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                            StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                            JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                            NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                            StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> "1" Then
                            If MESG_FLG = 1 Then
                                Beep
                                MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                            Exit Function
                        End If
                    
                        If Zan_Qty < CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            Upd_com = BtOpUpdate
                            WK_Qty = Zan_Qty
                        Else
                            Upd_com = BtOpDelete
                            WK_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        End If

                        Exit Do
                    Case BtErrEOF

                        If MESG_FLG = 1 Then
                            Beep
                            MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                        Exit Function

                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If Upd_com = BtOpUpdate Then
                                                                            '�L���݌ɐ�
                Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
            
            End If


            RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop
'============================================================
    '*------------------------------------------------------'���ד��w�薳�� �݌Ƀf�[�^�Ǎ��݁iTO���̏����j
            Call UniCode_Conv(K0_wZAIKO.SOKO_NO, Mid(TO_LOCATION, 1, 2))    'TO�q�ɇ�
            Call UniCode_Conv(K0_wZAIKO.Retu, Mid(TO_LOCATION, 3, 2))       'TO��
            Call UniCode_Conv(K0_wZAIKO.Ren, Mid(TO_LOCATION, 5, 2))        'TO�A
            Call UniCode_Conv(K0_wZAIKO.Dan, Mid(TO_LOCATION, 7, 2))        'TO�i
            Call UniCode_Conv(K0_wZAIKO.JGYOBU, JGYOBU)                     '���ƕ�
            Call UniCode_Conv(K0_wZAIKO.NAIGAI, TO_NAIGAI)                  '���O
            Call UniCode_Conv(K0_wZAIKO.HIN_GAI, HIN_GAI)                   '�i�ԁi�O���j
            Call UniCode_Conv(K0_wZAIKO.GOODS_ON, IDO_GOODS_ON_F)           '���i�^�����i
                                                                            '���ד�
            Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))

            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr

                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If Upd_com = BtOpInsert Then
                                                '�V�K�ǉ�
                Call UniCode_Conv(wZAIKOREC.SOKO_NO, Mid(TO_LOCATION, 1, 2))    '�q�ɇ�
                Call UniCode_Conv(wZAIKOREC.Retu, Mid(TO_LOCATION, 3, 2))       '��
                Call UniCode_Conv(wZAIKOREC.Ren, Mid(TO_LOCATION, 5, 2))        '�A
                Call UniCode_Conv(wZAIKOREC.Dan, Mid(TO_LOCATION, 7, 2))        '�i
                Call UniCode_Conv(wZAIKOREC.JGYOBU, JGYOBU)                     '���ƕ�
                Call UniCode_Conv(wZAIKOREC.NAIGAI, TO_NAIGAI)                  '���O
                Call UniCode_Conv(wZAIKOREC.HIN_GAI, HIN_GAI)                   '�i�ԁi�O���j
                Call UniCode_Conv(wZAIKOREC.GOODS_ON, IDO_GOODS_ON_F)           '���i�^�����i
                                                                                '���ד�
                Call UniCode_Conv(wZAIKOREC.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
                Call UniCode_Conv(wZAIKOREC.NYUKO_DT, Format(Date, "YYYYMMDD")) '���ɓ�
                                                                                '�i�ԁi�����j
                Call UniCode_Conv(wZAIKOREC.HIN_NAI, StrConv(ITEMREC.HIN_NAI, vbUnicode))
                                                                                '�L���݌ɐ�
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
                Call UniCode_Conv(wZAIKOREC.LOCK_F, LOCK_OFF)                   '�r���t���O
                Call UniCode_Conv(wZAIKOREC.WEL_ID, "")                         '�g�p���q�@ID
                Call UniCode_Conv(wZAIKOREC.PRG_ID, "")                         '�g�p����۸���

                Call UniCode_Conv(wZAIKOREC.GOODS_YMD, IDO_GOODS_YMD)           '���i����
                
                                                                                '�d���溰��2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_CODE, StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                                                                                '�d����P��2006.01.08
                Call UniCode_Conv(wZAIKOREC.SHIIRE_TANKA, StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode))
                                                                                '�v��N��2006.01.08
                Call UniCode_Conv(wZAIKOREC.KEIJYO_YM, StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode))
                
                
                Call UniCode_Conv(wZAIKOREC.FILLER, "")
            Else
                                            '�݌ɐ��X�V
                Call UniCode_Conv(wZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(wZAIKOREC.YUKO_Z_QTY, vbUnicode)) + WK_Qty, "00000000"))
            End If

            RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(Upd_com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                IDO_OSAKA_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                        IDO_OSAKA_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop
        '*------------------------------------------------------'�݌Ɉړ����o��
            If IDO_GOODS_ON_F = "0" Then
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            TO_LOCATION, _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            WK_Qty, 0, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, , MEMO & "���i�U��", _
                                            Ins_DateTime, _
                                            , MENU_NO, , , , , , , DEN_NO, Format(Now, "YYYYMMDD"))
            Else
                sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                            TO_LOCATION, _
                                            JGYOBU, _
                                            NAIGAI, _
                                            HIN_GAI, _
                                            StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                            YOIN, _
                                            0, WK_Qty, _
                                            ID, _
                                            TANTO_CODE, _
                                            RETRY, , MEMO, _
                                            Ins_DateTime, _
                                            , MENU_NO, , , , , , , DEN_NO, Format(Now, "YYYYMMDD"))
            End If
                
            If sts Then
                IDO_OSAKA_Update_Proc = sts
                Exit Function
            End If

            Zan_Qty = Zan_Qty - WK_Qty

            If Zan_Qty <= 0 Then
                Exit Do                     '�������Ƃ��I��
            End If

        Loop
                
    End If
    
    
'============================================================
    
    If Trim(DEN_NO) <> "" And DEN_NO <> "000000" Then
        Call UniCode_Conv(Y_NYU_O_REC.DEN_NO, DEN_NO)
    End If
    
    Call UniCode_Conv(Y_NYU_O_REC.KENPIN_F, KAN_KBN_FIN)
    
    Call UniCode_Conv(Y_NYU_O_REC.WEL_ID, "")
    Call UniCode_Conv(Y_NYU_O_REC.PRG_ID, "")
    
        
    RETRY_CNT = 0
    Do
        sts = BTRV(BtOpUpdate, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYU_O.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, com, "���ח\��")
                IDO_OSAKA_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
    
    
    If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                                        '�W���I��
'        If Last_JGYOBU = SOJIKI Or _
'            Last_JGYOBU = SENTAKU Then
                                        '�|���@�͐ݒ�ς݂̏㏑�������Ȃ�
'''�S�Z���^�[�ݒ�ϕW���I�Ԃ͕ύX���Ȃ��B2004.04.10
            If StrConv(ITEMREC.ST_SET_DT, vbUnicode) = Space(8) Then
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Date, "YYYYMMDD"))
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
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
                        IDO_OSAKA_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                IDO_OSAKA_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    IDO_OSAKA_Update_Proc = False
    
End Function

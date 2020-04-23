Attribute VB_Name = "SYUKO_SEK_UPDATE"
Option Explicit


Public Function Syuko_SEK_Update_Proc(JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    NYUKA_DT As String, _
                                    FROM_LOCATION As String, _
                                    YOIN As String, _
                                    SUMI_JITU_QTY As Long, _
                                    MI_JITU_QTY As Long, _
                                    SYUKA_QTY As Long, _
                                    ID As String, _
                                    TANTO_CODE As String, _
                                    Optional RETRY As Integer = 10, _
                                    Optional MEMO As String = "          ", _
                                    Optional CYU_KBN As String = " ", _
                                    Optional MUKE_CODE As String = "                ", _
                                    Optional SYUKA_YMD As String = "        ", _
                                    Optional DEN_NO As String = "          ", _
                                    Optional ID_NO As String = "        ", _
                                    Optional MENU_NO As String = "  ", _
                                    Optional BIN_NO As String = "  ", _
                                    Optional LOG_NON As Integer = 0, _
                                    Optional Ins_DateTime As String, _
                                    Optional mode As Integer = 0) As Integer
'****************************************************
'*      �u�o�ׁ^�o�ɏ����v�݌Ƀf�[�^�X�V
'*      ���o�b�p�@2007.03.17
'*  �݌Ƀf�[�^�̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  ���g�����U�N�V�����������K�v�ȏꍇ�͌Ăь��ōs����
'*  �g�p̧��    :   �݌Ƀf�[�^
'*                  �i�ڃ}�X�^
'*                  �v���}�X�^
'*                  ������}�X�^
'*                  �o�ח\��f�[�^
'*                  �݌Ɉړ���
'*  �����F  ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          FROM�I�ԁiXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          ���ד�(YYYYMMDD �ȗ��� �ȗ���FIFO)
'*          �v��(�ȗ��s��)
'*          ���i���ςݎ��ѐ��i�����ꂩ�K�{�j
'*          �����i���ѐ�    �i�@�@�@�V�@�@�j
'*          �o�א���        �i            �j
'*          ID(�ȗ��s��)
'*          �S���ҁi�ȗ��s�j
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*          ����(�ȗ��� �����ɏo�͂��郁�����e)
'*          �����敪�i�o�׎��K�{�j
'*          �`�[�h�c�i�o�׎��K�{�j
'*          �ƭ���ٰ�߁i�����Ǘ����ځj  2006.01.30
'*          �և�                        2007.05.16
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
    
Dim GET_DEN_NO  As String * 6
Dim GET_ID_NO   As String * 12
    
Dim JITU_QTY    As Long
Dim GOODS_F     As String * 1
    
Dim Wk_SUMI_JITU_QTY    As Long
Dim Wk_MI_JITU_QTY      As Long
Dim Wk_SYUKA_QTY        As Long

'Dim Ins_DateTime    As String * 14
    
Dim wkYOIN      As String * 2
    
    
    
'2010.01.06
Dim svSYUKA_YMD As String
Dim svDEN_NO    As String
Dim svTOK_KBN   As String
Dim svID_NO     As String
'2010.01.06
    
    
'''''   2011.04.11
Dim Total_SUMI_JITU_QTY     As Long
Dim Total_MI_JITU_QTY       As Long
'''''   2011.04.11
    
    
    
    
    Syuko_SEK_Update_Proc = True
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
                                        
                                        
    Wk_SUMI_JITU_QTY = SUMI_JITU_QTY
    Wk_MI_JITU_QTY = MI_JITU_QTY
    Wk_SYUKA_QTY = SYUKA_QTY
                                        
'    Ins_DateTime = Format(Now, "YYYYMMDDHHMMSS")
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
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                If RETRY_SU <> 0 Then
                    
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                
                End If
                
                If MESG_FLG = 0 Then
                    DoEvents
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Syuko_SEK_Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
'============================================================
'*------------------------------------------------------'���ד��w�薳�� �݌Ƀf�[�^�Ǎ��݁i�X�L���i�����j
'
'---------------------------------------'���i���ς݁`�����i�������������Ă�

    If LOG_NON = 1 Then
    Else
        If SUMI_JITU_QTY = 0 And MI_JITU_QTY = 0 Then
            
            
            If Left(YOIN, 1) = ACT_DENPYO_ID Or _
                Left(YOIN, 1) = ACT_SYUKA_HYO Or _
                Left(YOIN, 1) = ACT_SYUKA_HYO_OSAKA Then
                wkYOIN = ACT_SYUKA_KEI & CYU_KBN
            End If
            Total_SUMI_JITU_QTY = 0
            Total_MI_JITU_QTY = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.04
        Else
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
                                        "") Then
                Exit Function
            End If
        End If
    End If




    Zan_Qty = SYUKA_QTY
    Do
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Mid(FROM_LOCATION, 1, 2))   '�q�ɇ�
        Call UniCode_Conv(K0_ZAIKO.Retu, Mid(FROM_LOCATION, 3, 2))      '��
        Call UniCode_Conv(K0_ZAIKO.Ren, Mid(FROM_LOCATION, 5, 2))       '�A
        Call UniCode_Conv(K0_ZAIKO.Dan, Mid(FROM_LOCATION, 7, 2))       '�i
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)                      '���ƕ�
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)                      '���O
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")                        '���i�^�����i
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")                        '���ד�
        
        
        SUMI_JITU_QTY = 0
        MI_JITU_QTY = 0
        
        RETRY_CNT = 0

        Do
            sts = BTRV(BtOpGetGreater + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                                        '�I�{�i�{���i�^�����i�u���[�N
                    If FROM_LOCATION <> (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(ZAIKOREC.Dan, vbUnicode)) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        
'                        If MESG_FLG = 1 Then
'                            Beep
'                            MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
'                        End If
'                        Exit Function
                    
                        mode = 1
                        GoTo SYUKA_UPDATE
                    
                    
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

'                    If MESG_FLG = 1 Then
'                        Beep
'                        MsgBox "�݌ɐ����s�����Ă��܂��B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
'                    End If
                    
'                    Exit Function
                    
                    mode = 1
                    GoTo SYUKA_UPDATE
                    

                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 1)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If

                    End If

                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop

        If Upd_com = BtOpUpdate Then
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)) - Zan_Qty, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                '�r���t���O
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                      '�g�p���q�@ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                      '�g�p���v���O����
        End If

        If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
            SUMI_JITU_QTY = SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
            Total_SUMI_JITU_QTY = Total_SUMI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
        Else
            MI_JITU_QTY = MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
            Total_MI_JITU_QTY = Total_MI_JITU_QTY + WK_Qty
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
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
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                
                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        Loop
    '============================================================
        '*------------------------------------------------------'�݌Ɉړ����o��
        
        
        
        
        '�o�ח\�聕�����悩������N���A�擾 2007.06.02
            
            
        '2010.01.06
        svSYUKA_YMD = StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode)
        svDEN_NO = StrConv(Y_SYUREC.DEN_NO, vbUnicode)
        svTOK_KBN = StrConv(Y_SYUREC.TOK_KBN, vbUnicode)
        svID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
        '2010.01.06
            
            
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, "")
        Call UniCode_Conv(Y_SYUREC.DEN_NO, "")
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
        Call UniCode_Conv(IDOREC.TOKU_MARK, "")
        Call UniCode_Conv(Y_SYUREC.ID_NO, "")
            
    
        Call UniCode_Conv(MTSREC.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
        Call UniCode_Conv(MTSREC.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
                    
        Call UniCode_Conv(MTSREC.MUKE_NAME, StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
        Call UniCode_Conv(MTSREC.SS_NAME, "")
            
        Call UniCode_Conv(MTSREC.MUKE_DNAME, StrConv(Y_SYUREC.MUKE_NAME, vbUnicode))
            
            
        
        
        
        
        
        
        
        
        sts = IDOREKI_OUTPUT_PROC(FROM_LOCATION, _
                                    Space(8), _
                                    JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), _
                                    YOIN, _
                                    SUMI_JITU_QTY, _
                                    MI_JITU_QTY, _
                                    ID, _
                                    TANTO_CODE, _
                                    RETRY, _
                                    CYU_KBN, _
                                    MEMO, _
                                    Ins_DateTime, _
                                    StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode), _
                                    StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode), _
                                    StrConv(ZAIKOREC.KEIJYO_YM, vbUnicode), _
                                    MENU_NO, MUKE_CODE, , ID_NO, _
                                    BIN_NO, _
                                    DEN_NO, SYUKA_YMD, 1)

        If sts Then
            Syuko_SEK_Update_Proc = sts
            Exit Function
        End If
        
        Zan_Qty = Zan_Qty - WK_Qty

        If Zan_Qty <= 0 Then
            Exit Do                     '�������Ƃ��I��
        End If
    Loop

SYUKA_UPDATE:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11
    If LOG_NON = 1 Then
    Else
        If SYUKA_QTY <> 0 Then
            If P_SAGYO_LOG_OUTPUT_PROC(TANTO_CODE, _
                                        ID, _
                                        JGYOBU, _
                                        NAIGAI, _
                                        MENU_NO, _
                                        wkYOIN, _
                                        HIN_GAI, _
                                        Total_SUMI_JITU_QTY, _
                                        Total_MI_JITU_QTY, _
                                        FROM_LOCATION, _
                                        "", _
                                        ID_NO, _
                                        MUKE_CODE) Then
                Exit Function
            End If
        End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2011.04.11


'============================================================
'   �o�ח\��(νĲҰ��)--���o�ח\��̍X�V
'============================================================
    
    
    '����o�א���KEEP
    Wk_SYUKA_QTY = Wk_SYUKA_QTY + Wk_SUMI_JITU_QTY + Wk_MI_JITU_QTY
    
    Call UniCode_Conv(K4_Y_SYU_H.ID_NO, ID_NO)
    com = BtOpGetGreaterEqual
    
    
    Do
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
            Select Case sts
                Case BtNoErr
                    
                                        
                    If StrConv(Y_SYU_HREC.ID_NO, vbUnicode) <> ID_NO Then
                        sts = BtErrEOF
                    End If
                    
                    Exit Do
                Case BtErrEOF
                    If MESG_FLG = 1 Then
                        
                        
                        If Wk_SYUKA_QTY <> 0 Then
                            Beep
                            MsgBox "�o�ח\�肪���݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                        End If
                    End If
                    
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
                        
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                            '�񐔃I�[�o�[
                            Call File_Error(sts, com + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        
                        End If
                    
                    End If
                    
                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�o�ח\��(νĲҰ��)")
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
        Loop
        
        '�����I��
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        
        If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) = "1" Then
            '��ݾٕ��͖�����
        
            '2010.01.06
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, svSYUKA_YMD)
            Call UniCode_Conv(Y_SYUREC.DEN_NO, svDEN_NO)
            Call UniCode_Conv(Y_SYUREC.TOK_KBN, svTOK_KBN)
            Call UniCode_Conv(Y_SYUREC.ID_NO, svID_NO)
            '2010.01.06
        
        
        Else
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrEOF
                        If MESG_FLG = 1 Then
                            
                            
                            If Wk_SYUKA_QTY <> 0 Then
                                Beep
                                MsgBox "�o�ח\�肪���݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                        End If
                        
                        Exit Function
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '�񐔃I�[�o�[
                                Call File_Error(sts, com + BtSNoWait, "�o�ח\��", 0)
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��")
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
            
            
            WK_Qty = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
            If Wk_SYUKA_QTY >= WK_Qty Then
                                  
                                  
                                  
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
                                                                '�������t
                Call UniCode_Conv(Y_SYUREC.KAN_YMD, Format(Date, "yyyymmdd"))
                                                                '�m�萔��
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(Y_SYUREC.SURYO, vbUnicode))
                        
            
            
            Else
            
                
                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) + Wk_SYUKA_QTY, "0000000"))
            
            End If
        
        
                    
            Call UniCode_Conv(Y_SYU_HREC.J_SURYO, StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
        
            Wk_SYUKA_QTY = Wk_SYUKA_QTY - WK_Qty
        
        
        
        
        
        
        
        
        
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")              '�g�p�[��ID�i���󔒁j
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")              '�g�p���v���O�����i���󔒁j
    
            '*------------------------------------------------------'�o�ח\��o��
            RETRY_CNT = 0
            Do
                
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
    
                            End If
    
                        End If
    
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", MESG_FLG)
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
    
                End Select
            Loop
        
        
        
        
        
            '*------------------------------------------------------'�����f�[�^�o��
            Call UniCode_Conv(K2_Y_SYU_TEI.KEN_NO, StrConv(Y_SYU_HREC.SEK_KEN_NO, vbUnicode))
            Call UniCode_Conv(K2_Y_SYU_TEI.HIN_NO, StrConv(Y_SYU_HREC.SEK_HIN_NO, vbUnicode))
        
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                                            
                    Case BtErrKeyNotFound
                        If MESG_FLG = 1 Then
                            
                            
                            If Wk_SYUKA_QTY <> 0 Then
                                Beep
                                MsgBox "�����f�[�^�����݂��܂���B�X�V�����𒆎~���܂��B", vbOKOnly, "�m�F����"
                            End If
                        End If
                        
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                                '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
                            
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                                '�񐔃I�[�o�[
                                Call File_Error(sts, com + BtSNoWait, "�����ް�")
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            
                            End If
                        
                        End If
                        
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "�����ް�")
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
                End Select
            
            Loop
        
        
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_TEI_REC.SHOGO_DATETIME, Ins_DateTime)
        
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Ins_DateTime)
        
        
        
        
        
            RETRY_CNT = 0
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K2_Y_SYU_TEI, Len(K2_Y_SYU_TEI), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then
    
                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, BtOpUpdate, "�����ް�")
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
    
                            End If
    
                        End If
    
                        If MESG_FLG = 0 Then
                            DoEvents
                        Else
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Syuko_SEK_Update_Proc = SYS_CANCEL
                                Exit Function
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, com, "�o�ח\��", MESG_FLG)
                        Syuko_SEK_Update_Proc = SYS_ERR
                        Exit Function
    
                End Select
            Loop
        
        
        
        
        
        
        
        End If
        
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, Ins_DateTime)
    
        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Ins_DateTime)

        '*------------------------------------------------------'�o�ח\��(νĲҰ��)�o��
        RETRY_CNT = 0
        Do
            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    If MESG_FLG = 0 Then
                        DoEvents
                    Else
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYU_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Syuko_SEK_Update_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Case Else
                    Call File_Error(sts, com, "�o�ח\��", MESG_FLG)
                    Syuko_SEK_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop



        If Wk_SYUKA_QTY <= 0 Then
            Exit Do
        End If


        com = BtOpGetNext
    Loop

'============================================================
    If mode = 1 Then
        sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts Then
            Call File_Error(sts, BtOpUnlock, "�i��Ͻ�", MESG_FLG)
            Syuko_SEK_Update_Proc = SYS_ERR
            Exit Function
        End If
        Syuko_SEK_Update_Proc = False
        Exit Function
    End If
                                        
                                        '�ŏI�o�ɓ�
    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Format(Date, "yyyymmdd"))
    
    Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, Format(SUMI_JITU_QTY + MI_JITU_QTY, "00000000"))
    
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
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    
                    End If
                    
                End If
                
                If MESG_FLG = 0 Then
                    DoEvents
                Else
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Syuko_SEK_Update_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                Syuko_SEK_Update_Proc = SYS_ERR
                Exit Function
                        
        End Select
    Loop
'============================================================
    
    Syuko_SEK_Update_Proc = False
    
End Function



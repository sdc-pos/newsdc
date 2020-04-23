Attribute VB_Name = "LotNoProc"
Public Function LOTNO_IN_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@���׌��i�����x
'
'       2013.06.06
'-------------------------------------------------------


Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20

Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant



    LOTNO_IN_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)         '2014.07.24
                '    Case LCD_LotNo_BCR      'BC                                '2014.07.24
    
                Select Case i                                                   '2014.07.24
                    Case 1      'BC                                             '2014.07.24
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_IN_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_IN_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        
                        
                        '------------------ ���g�Ǘ��f�[�^�Ǎ���
                        sts = LotNo_Check_Proc(CStr(IN_BCR(0)), _
                                                CStr(IN_BCR(1)), _
                                                0)
                        
                        Select Case sts
                            Case False
                            
                                If Trim(StrConv(LOTNOREC.IDt, vbUnicode)) <> "" Then
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���i�ςł��B", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    
                                    
                                    LOTNO_IN_CHECK_PROC = False
                                    Exit Function
                                
                                End If
                            
                            Case BtErrKeyNotFound
                            '-------------------------- �f�[�^���o�^
                            
                            
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�Y�������ް��Ȃ�", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                
                                
                                LOTNO_IN_CHECK_PROC = False
                                Exit Function
                            
                            
                            
                            
                            
                            
                            
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                                
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�����ް��g�p��", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                
                                
                                LOTNO_IN_CHECK_PROC = False
                                Exit Function
                            
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                
                        '-------------------------------------------    ���g�Ǘ��f�[�^�X�V
                        '�݌ɐ�
                        Call UniCode_Conv(LOTNOREC.SQty, StrConv(LOTNOREC.IQty, vbUnicode))
                        '���ד�
                        Call UniCode_Conv(LOTNOREC.IDt, Format(Now, "YYYYMMDD"))
                        '���גS����
                        Call UniCode_Conv(LOTNOREC.ITantoCode, ID_KANRI_TBL(ING_No).TANTO_CODE)
                        '�X�V�h�c
                        Call UniCode_Conv(LOTNOREC.UpdID, App.EXEName)
                        '�X�V����
                        Call UniCode_Conv(LOTNOREC.UpdDtm, Format(Now, "YYYYMMDDHHMMSS"))
                        '��������   -------------
                        sts = BTRV(BtOpUpdate, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End Select
                        '-------------------------------------------    ��ƃ��O�o��
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                        If Trim(MENU_NO) = "" Then
                        Else
                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                                (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                                ID_KANRI_TBL(ING_No).NAIGAI, _
                                                                MENU_NO, _
                                                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                                 CStr(IN_BCR(0)), Val(IN_BCR(2)), , , , _
                                                                 CStr(IN_BCR(1))) Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                        End If
                        '-------------------------------------------    �X�V����
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                        
                        
                        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DEF, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�����׌��i�n�j", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                        
                        
                End Select
            Next i
    End Select

    LOTNO_IN_CHECK_PROC = False
    



End Function

Public Function LOTNO_OUT_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@�o�׌��i�����x
'
'       2013.06.06
'-------------------------------------------------------

Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20

Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant


Dim wkNum           As Long


    LOTNO_OUT_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)             '2014.07.24
                '    Case LCD_LotNo_BCR      'BC                                    '2014.07.24
    
                Select Case i                                                       '2014.07.24
                    Case 1                                                          '2014.07.24
    
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_OUT_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_OUT_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        
                        
                        '------------------ ���g�Ǘ��f�[�^�Ǎ���
                        sts = LotNo_Check_Proc(CStr(IN_BCR(0)), _
                                                CStr(IN_BCR(1)), _
                                                0)
                        
                        Select Case sts
                            Case False
                            
                                If Val(StrConv(LOTNOREC.SQty, vbUnicode)) = 0 Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���i�ρ@�݌ɐ����O", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    
                                    
                                    LOTNO_OUT_CHECK_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
'2014.07.19                                If Val(StrConv(LOTNOREC.OQty, vbUnicode)) > Val(StrConv(LOTNOREC.SQty, vbUnicode)) Then
                                If Val(IN_BCR(2)) > Val(StrConv(LOTNOREC.SQty, vbUnicode)) Then     '2014.07.19
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���i�ρ@�݌ɐ���" & StrConv(Val(StrConv(LOTNOREC.SQty, vbUnicode)), vbWide), _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    LOTNO_OUT_CHECK_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                            Case BtErrKeyNotFound
                            '-------------------------- �f�[�^���o�^
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                LOTNO_OUT_CHECK_PROC = False
                                Exit Function
                            
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�ް��g�p��", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                LOTNO_OUT_CHECK_PROC = False
                                Exit Function
                                
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                
                        '-------------------------------------------    ���g�Ǘ��f�[�^�X�V
                        
                        '�o�א�(�o�א�+����)
                        wkNum = Val(StrConv(LOTNOREC.OQty, vbUnicode))
                        wkNum = wkNum + Val(IN_BCR(2))
                        Call UniCode_Conv(LOTNOREC.OQty, Format(wkNum, "000000"))
                        '�݌ɐ�(�݌ɐ�-����)
                        wkNum = Val(StrConv(LOTNOREC.SQty, vbUnicode))
                        wkNum = wkNum - Val(IN_BCR(2))
                        Call UniCode_Conv(LOTNOREC.SQty, Format(wkNum, "000000"))
                        '�o�ד�
                        Call UniCode_Conv(LOTNOREC.ODt, Format(Now, "YYYYMMDD"))
                        '�o�גS����
                        Call UniCode_Conv(LOTNOREC.OTantoCode, ID_KANRI_TBL(ING_No).TANTO_CODE)
                        '�X�V�h�c
                        Call UniCode_Conv(LOTNOREC.UpdID, App.EXEName)
                        '�X�V����
                        Call UniCode_Conv(LOTNOREC.UpdDtm, Format(Now, "YYYYMMDDHHMMSS"))
                        '��������   -------------
                        sts = BTRV(BtOpUpdate, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End Select
                        
                        
                        '-------------------------------------------    ��ƃ��O�o��
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                        If Trim(MENU_NO) = "" Then
                        Else
                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                                (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                                ID_KANRI_TBL(ING_No).NAIGAI, _
                                                                MENU_NO, _
                                                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                                 CStr(IN_BCR(0)), Val(IN_BCR(2)), , , , _
                                                                 CStr(IN_BCR(1))) Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                        End If
                        '-------------------------------------------    �X�V����
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DEF, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "���o�׌��i�n�j", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                End Select
            Next i
        
            
            
            
            
    
    
    End Select

    LOTNO_OUT_CHECK_PROC = False


End Function

Public Function LOTNO_OUT_CANCEL_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@�o�׃L�����Z�������x
'
'       2013.06.06
'-------------------------------------------------------

Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20

Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant


Dim wkNum           As Long


    LOTNO_OUT_CANCEL_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)         '2014.07.24
                '    Case LCD_LotNo_BCR      'BC                                '2014.07.24
    
                Select Case i                                                   '2014.07.24
                    Case 1                  'BC                                 '2014.07.24
    
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_OUT_CANCEL_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_OUT_CANCEL_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        
                        
                        '------------------ ���g�Ǘ��f�[�^�Ǎ���
                        sts = LotNo_Check_Proc(CStr(IN_BCR(0)), _
                                                CStr(IN_BCR(1)), _
                                                0)
                        
                        Select Case sts
                            Case False
                            
                                If Val(StrConv(LOTNOREC.OQty, vbUnicode)) = 0 Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���o��", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    LOTNO_OUT_CANCEL_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                                If Val(StrConv(LOTNOREC.OQty, vbUnicode)) < Val(IN_BCR(2)) Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~�o�א��s���i" & StrConv(Val(StrConv(LOTNOREC.OQty, vbUnicode)), vbWide), _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    LOTNO_OUT_CANCEL_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                            Case BtErrKeyNotFound
                            '-------------------------- �f�[�^���o�^
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                LOTNO_OUT_CANCEL_PROC = False
                                Exit Function
                            
                            
                            
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�ް��g�p��", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                LOTNO_OUT_CANCEL_PROC = False
                                Exit Function
                            
                            
                            
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        ID_KANRI_TBL(ING_No).Model = CStr(IN_BCR(0))
                        ID_KANRI_TBL(ING_No).PLotNo = CStr(IN_BCR(1))
                        ID_KANRI_TBL(ING_No).SURYO = Val(IN_BCR(2))
                        
                        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DEF, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    , "��낵���ł����H", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "�݌ɐ�:" & Val(StrConv(LOTNOREC.SQty, vbUnicode)) & "��" & Val(IN_BCR(2)))
                        
                        
                End Select
            Next i
        
            
            
            
            
        Case Step_Sagyo2_RES        '�R��ڂ̎�M�iAny Key�j
            
            
            '-------------------------------------------    ���g�Ǘ��f�[�^�X�V
            sts = LotNo_Check_Proc(ID_KANRI_TBL(ING_No).Model, _
                                    ID_KANRI_TBL(ING_No).PLotNo, _
                                    0)
            Select Case sts
                Case False
                
                    If Val(StrConv(LOTNOREC.OQty, vbUnicode)) = 0 Then
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                            Buzzer_DOUBLE, _
                                            , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                            TYPE_BCANK, "�~���o��", _
                                            , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                            , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                            , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                        
                        LOTNO_OUT_CANCEL_PROC = False
                        Exit Function
                    End If
                
                
                    If Val(StrConv(LOTNOREC.OQty, vbUnicode)) < ID_KANRI_TBL(ING_No).SURYO Then
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                            Buzzer_DOUBLE, _
                                            , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                            TYPE_BCANK, "�~�o�א��s���i" & StrConv(Val(StrConv(LOTNOREC.OQty, vbUnicode)), vbWide), _
                                            , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                            , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                            , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                        
                        LOTNO_OUT_CANCEL_PROC = False
                        Exit Function
                        
                    End If
                
                
                Case BtErrKeyNotFound
                '-------------------------- �f�[�^���o�^
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
        
                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DOUBLE, _
                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                        TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                    
                    LOTNO_OUT_CANCEL_PROC = False
                    Exit Function
                
                
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
        
                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DOUBLE, _
                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                        TYPE_BCANK, "�~�ް��g�p��", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                    
                    LOTNO_OUT_CANCEL_PROC = False
                    Exit Function
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
            
            
            '�o�א�(�o�א�+����)
            wkNum = Val(StrConv(LOTNOREC.OQty, vbUnicode))
            wkNum = wkNum - ID_KANRI_TBL(ING_No).SURYO
            Call UniCode_Conv(LOTNOREC.OQty, Format(wkNum, "000000"))
            '�݌ɐ�(�݌ɐ�-����)
            wkNum = Val(StrConv(LOTNOREC.SQty, vbUnicode))
            wkNum = wkNum + ID_KANRI_TBL(ING_No).SURYO
            Call UniCode_Conv(LOTNOREC.SQty, Format(wkNum, "000000"))
            '�o�ד�/�o�גS���҂͂��̂܂�
                            
            '�X�V�h�c
            Call UniCode_Conv(LOTNOREC.UpdID, App.EXEName)
            '�X�V����
            Call UniCode_Conv(LOTNOREC.UpdDtm, Format(Now, "YYYYMMDDHHMMSS"))
            '��������   -------------
            sts = BTRV(BtOpUpdate, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End Select
            '-------------------------------------------    ��ƃ��O�o��
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
            If Trim(MENU_NO) = "" Then
            Else
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     RTrim(ID_KANRI_TBL(ING_No).Model), ID_KANRI_TBL(ING_No).SURYO, , , , _
                                                     RTrim(ID_KANRI_TBL(ING_No).PLotNo)) Then
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            End If
            '-------------------------------------------    �X�V����
            
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            
            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DEF, _
                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                        TYPE_BCANK, "���o�׷�ݾ�OK", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "�݌ɐ�:" & wkNum)
    
    End Select

    LOTNO_OUT_CANCEL_PROC = False



End Function

Public Function LOTNO_LABEL_PRINT_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@���x�����s�����x
'
'       2013.06.06
'-------------------------------------------------------
Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20

Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant


Dim wkNum           As Long

Dim Mai_Su          As Long

Dim FileName        As String

Dim wkTEXT          As Variant

    LOTNO_LABEL_PRINT_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)         '2014.07.24
                '    Case LCD_LotNo_BCR      'BC                                '2014.07.24
    
    
                Select Case i                                                   '2014.07.24
                    Case 1              'BC                                     '2014.07.24
    
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function
                        
                        End If
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        ID_KANRI_TBL(ING_No).Model = CStr(IN_BCR(0))
                        ID_KANRI_TBL(ING_No).PLotNo = CStr(IN_BCR(1))
                        ID_KANRI_TBL(ING_No).SURYO = Val(IN_BCR(2))
                        
                        '-----------------------------------------------�w�b�_�[
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                    
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '���l�����\��
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(0).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(0).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                                                                                'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model))
                                                                                '���l�����\��
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(1).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo))
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                                'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "��     ��:" & Space(10 - Len(Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))) & Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "��     ��:" & Space(10 - Len(Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))) & Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))

                                                                                '���l�����\��
                        Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))) & Format(ID_KANRI_TBL(ING_No).SURYO, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(ID_KANRI_TBL(ING_No).SURYO, "#0"))) & Format(ID_KANRI_TBL(ING_No).SURYO, "#0")
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = "15"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "15"
                                                                                '���͌���
                        Send_Text.Box_Type(3).Max_Size = "06"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "06"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                                'BOX����
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "��     ��:" & Space(10 - Len(Format(1, "#0"))) & Format(1, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "��     ��:" & Space(10 - Len(Format(1, "#0"))) & Format(1, "#0"))

                                                                                '���l�����\��
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(1, "#0"))) & Format(1, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(1, "#0"))) & Format(1, "#0")
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = "15"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "15"
                                                                                '���͌���
                        Send_Text.Box_Type(4).Max_Size = "06"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "06"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                        Sendbuf = Text_Create_Proc()
                        
                End Select
            Next i
        
            
            
            
            
        Case Step_Sagyo2_RES        '�R��ڂ̎�M�iAny Key�j
            For i = 0 To M_Gyo - 1
                
                Select Case i
                    Case 1
                    
                        wkTEXT = Split(ID_KANRI_TBL(ING_No).Recv_text(i), " ", -1)
                    
                        wkTEXT = Split(wkTEXT(0), ":", -1)

                        If UBound(wkTEXT) > 0 Then
                    
                            ID_KANRI_TBL(ING_No).Model = wkTEXT(1)
                        Else
                            ID_KANRI_TBL(ING_No).Model = wkTEXT(0)
                        End If
                    Case 2
                        wkTEXT = Split(ID_KANRI_TBL(ING_No).Recv_text(i), " ", -1)
                    
                        wkTEXT = Split(wkTEXT(0), ":", -1)
                    
                        If UBound(wkTEXT) > 0 Then
                    
                            ID_KANRI_TBL(ING_No).PLotNo = wkTEXT(1)
                        Else
                            ID_KANRI_TBL(ING_No).PLotNo = wkTEXT(0)
                        End If
                    
                    Case 3
                                            
                                            
                                            
                                            
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                        
                    
            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~���ʃG���[", _
                                                    , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function

                        End If
                    
                        ID_KANRI_TBL(ING_No).SURYO = Val(RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                        If ID_KANRI_TBL(ING_No).SURYO <= 0 Then
                    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~���ʃG���[", _
                                                    , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function
                        End If
                    
                    
                    Case 4
                
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                        
                    
            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�����G���[", _
                                                    , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function

                        End If
                
                        Mai_Su = Val(RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                
                        If Mai_Su <= 0 Then
                    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�����G���[", _
                                                    , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                            LOTNO_LABEL_PRINT_PROC = False
                            Exit Function
                        End If
                
                
                        '-------------------------------------------    ��ƃ��O�o��
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                        If Trim(MENU_NO) = "" Then
                        Else
                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                                (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                                ID_KANRI_TBL(ING_No).NAIGAI, _
                                                                MENU_NO, _
                                                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                                 RTrim(ID_KANRI_TBL(ING_No).Model), ID_KANRI_TBL(ING_No).SURYO, , , , _
                                                                 RTrim(ID_KANRI_TBL(ING_No).PLotNo)) Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                        End If
                        '-------------------------------------------    �X�V����
                
                
                
                        FileName = N4_SendFile
                        
                        Call LotNo_Print_File_Make_Proc(FileName, Mai_Su)
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                                                    
                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                    
                                                    
                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                    
                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                    
                        '-----------------------------------------------�w�b�_�[
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_LABEL                   '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_LABEL
                    
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                    
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                    
                        Send_Text.FileName = FileName                           '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = FileName
                    
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                    
                        Sendbuf = Text_Create_Proc()

                
                
                End Select
           Next i
            
            
        Case Step_PRINT_RES         '�Q��ڂ̎�M�i����I���j
            
            
            
            
            '���̍�Ɨv��
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- �G���[���b�Z�[�W�쐬
                Case Else
                '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
                    Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
    
    
    End Select

    LOTNO_LABEL_PRINT_PROC = False


End Function


Private Sub Wel_LOTNO_Make_PROC(Sendbuf As String, _
                                Wel_Buzzer As String, _
                                Optional Wel_TYPE1 As String = TYPE_REF, _
                                Optional Wel_LCD1 As String = " ", _
                                Optional Wel_TYPE2 As String = TYPE_REF, _
                                Optional Wel_LCD2 As String = " ", _
                                Optional Wel_TYPE3 As String = TYPE_REF, _
                                Optional Wel_LCD3 As String = " ", _
                                Optional Wel_TYPE4 As String = TYPE_REF, _
                                Optional Wel_LCD4 As String = " ", _
                                Optional Wel_TYPE5 As String = TYPE_REF, _
                                Optional Wel_LCD5 As String = " ")
'-------------------------------------------------------
'
'   �w���g�[�������@���M÷�č쐬�����x
'
'       2013.06.06
'-------------------------------------------------------
                        
                        
        '-----------------------------------------------�w�b�_�[
        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK

        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF

        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only

        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"

        Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""

        Send_Text.Buzzer = Wel_Buzzer                           '�u�U�[���@�W��
        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Wel_Buzzer
                        
        '-----------------------------------------------�P�s��
                                                                'BOX����
        Send_Text.Box_Type(0).Box_Type = Wel_TYPE1
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = Wel_TYPE1
                                                                '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Wel_LCD1)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, Wel_LCD1)
                                                                '���l�����\��
        Send_Text.Box_Type(0).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                '�����J�[�\���ʒu
        Send_Text.Box_Type(0).Start_Pos = "01"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = "01"
                                                                '���͌���
        Send_Text.Box_Type(0).Max_Size = "20"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "20"
                                                                
        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
        '-----------------------------------------------�Q�s��
                                                                'BOX����
        Send_Text.Box_Type(1).Box_Type = Wel_TYPE2
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = Wel_TYPE2
                                                                '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Wel_LCD2)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Wel_LCD2)
                                                                '���l�����\��
        Send_Text.Box_Type(1).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                '�����J�[�\���ʒu
        Send_Text.Box_Type(1).Start_Pos = "01"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                '���͌���
        Send_Text.Box_Type(1).Max_Size = "20"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "20"
                                                                
        Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------�R�s��
                                                                'BOX����
        Send_Text.Box_Type(2).Box_Type = Wel_TYPE3
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = Wel_TYPE3
                                                                '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Wel_LCD3)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Wel_LCD3)
                                                                '���l�����\��
        Send_Text.Box_Type(2).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                '�����J�[�\���ʒu
        Send_Text.Box_Type(2).Start_Pos = "01"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                
        Send_Text.Box_Type(2).Max_Size = "20"                       '2007.07.21
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"  '2007.07.21
                                                                
                                                                
                                                                
        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
        '-----------------------------------------------�S�s��
                                                                'BOX����
        Send_Text.Box_Type(3).Box_Type = Wel_TYPE4
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = Wel_TYPE4
                                                                '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Wel_LCD4)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Wel_LCD4)
                                                                '���l�����\��
        Send_Text.Box_Type(3).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                '�����J�[�\���ʒu
        Send_Text.Box_Type(3).Start_Pos = "01"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                '���͌���
        Send_Text.Box_Type(3).Max_Size = "20"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                
        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
        '-----------------------------------------------�S�s��
                                                                'BOX����
        Send_Text.Box_Type(4).Box_Type = Wel_TYPE5
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = Wel_TYPE4
                                                                '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Wel_LCD5)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Wel_LCD5)
                                                                '���l�����\��
        Send_Text.Box_Type(4).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                '�����J�[�\���ʒu
        Send_Text.Box_Type(4).Start_Pos = "01"                    '���l�͂T���Œ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                '���͌���
         Send_Text.Box_Type(4).Max_Size = "20"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                
        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

        
        
        
        Sendbuf = Text_Create_Proc()

End Sub

Public Function LotNo_Check_Proc(Model As String, _
                                    PLotNo As String, _
                                    Locked As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@���g�Ǘ��f�[�^�Ǎ��ݏ����x
'
'       2013.06.06
'-------------------------------------------------------
Dim sts         As Integer
Dim RETRY_CNT   As Integer



    LotNo_Check_Proc = True


    Call UniCode_Conv(K0_LOTNO.Model, Model)
    Call UniCode_Conv(K0_LOTNO.PLotNo, PLotNo)
                
                
    RETRY_CNT = 0
                
                
    Do
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(BtOpGetEqual + Locked, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                RETRY_CNT = RETRY_CNT + 1
                If RETRY_CNT > FILE_RETRY Then
                    Call File_Error(sts, BtOpGetEqual + Locked, "���g�Ǘ��f�[�^", 0)
                    Exit Do
                End If
            Case Else
                Exit Do
        End Select
            
    Loop

    LotNo_Check_Proc = sts

End Function


Private Sub LotNo_Print_File_Make_Proc(FileName As String, Mai_Su As Long)
'-------------------------------------------------------
'
'   �w���g�[������ ���ٗp�f�[�^�t�@�C���쐬�x
'       2013.06.06
'
'-------------------------------------------------------
Dim sts         As Integer


Dim FileNo      As Long

Dim FullPath    As String

Dim wkPrint    As String * 20
Dim wklen       As Long
Dim wkmod       As Long

    
    
'-------------------------  2012.03.19
    If Right(F1100101.CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = F1100101.CtrsWsk1.SendFolder & "\" & FileName & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    Else
        FullPath = F1100101.CtrsWsk1.SendFolder & FileName & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    End If
'-------------------------  2012.03.19




    On Error Resume Next
    Kill (FullPath)             '���M�p�t�@�C���폜
    On Error GoTo 0
        
    FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
    Open FullPath For Output As #FileNo







    Print #FileNo, "#"
    Print #FileNo, "JOB"
    
'------------------------------------------------------------------ ���x���v�����^�[�󎚔Z�x    2016.02.15
'    Print #FileNo, "DEF MK=1,DK=8,MD=3,PW=384,PH=344,XO=8,UM=24,BM=0,AF=1"
    Print #FileNo, "DEF MK=1,DK=" & Format(DK_DEF, "0") & ",MD=3,PW=384,PH=344,XO=8,UM=24,BM=0,AF=1"
'------------------------------------------------------------------ ���x���v�����^�[�󎚔Z�x    2016.02.15
    Print #FileNo, "START"
    
'    Print #FileNo, "FONT TP=7,CS=0,LG=80,WD=45,LS=0"
'    Print #FileNo, "TEXT X=15,Y=30,L=1"
    
    
    
    
    wkPrint = RTrim(ID_KANRI_TBL(ING_No).Model) & " " & RTrim(ID_KANRI_TBL(ING_No).PLotNo) & " " & Format(ID_KANRI_TBL(ING_No).SURYO, "00")
'    If Len(Trim(wkPrint)) < 20 Then
'        wklen = 20 - Len(Trim(wkPrint))
'
'
'        wklen = ToRoundDown(CCur(wklen / 2), 0)
'        If wklen < 1 Then
'        Else
'            wkPrint = Space(wklen) & Trim(wkPrint)
'        End If
'    End If
    
'    Print #FileNo, Left(Trim(Hinban), 14)
'    Print #FileNo, " "

    
'    Print #FileNo, "BCD TP=7,X=40,Y=120,DR=0,NW=1,RA=,HT=100,HR=0,MG=0,NS=0,NE=0,NZ=0"
    Print #FileNo, "BCD TP=7,X=20,Y=120,DR=0,NW=1,RA=1,HT=80,HR=1,MG=0,NS=0,NE=0,NZ=0"
    Print #FileNo, wkPrint
    
    
    Print #FileNo, "QTY P=" & Format(Mai_Su, "#0")
    Print #FileNo, "END"


    
    Print #FileNo, "JOBE"
    

    Close #FileNo

    FileName = N4_SendFile & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"




End Sub




Public Function INVNO_OUT_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@�o�׌��i����(���󇂕t��)�x
'
'       2014.07.01
'-------------------------------------------------------

Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20

Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant


Dim wkNum           As Long


    INVNO_OUT_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i���󇂁j
                
        
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)     '2014.07.24
                '    Case LCD_InvNo_BCR      '����                        '2014.07.24
                Select Case i                                               '2014.07.24
                    Case 1                                                  '2014.07.24
        
        
                        If Not IsNumeric(RTrim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���󇂃G���[", "���l�̂݉�", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            INVNO_OUT_CHECK_PROC = False
                            Exit Function
                        
                        End If
        
        
                        ID_KANRI_TBL(ING_No).INVNO = (RTrim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
        
        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                            Buzzer_DOUBLE, _
                                            , "����:" & ID_KANRI_TBL(ING_No).INVNO, _
                                            TYPE_BCANK, LCD_LotNo_BCR, _
                                            , "�i��:", _
                                            , "����:", _
                                            , "����:")

        
        
                End Select
        
        
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
'                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))

'                    Case LCD_LotNo_BCR      'BC
                Select Case i
                
                    Case 1
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            INVNO_OUT_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc() - 1
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            INVNO_OUT_CHECK_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        
                        
                        '------------------ ���g�Ǘ��f�[�^�Ǎ���
                        sts = LotNo_Check_Proc(CStr(IN_BCR(0)), _
                                                CStr(IN_BCR(1)), _
                                                0)
                        
                        Select Case sts
                            Case False
                            
                                If Val(StrConv(LOTNOREC.SQty, vbUnicode)) = 0 Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���i�ρ@�݌ɐ����O", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    
                                    
                                    INVNO_OUT_CHECK_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
 '2014.07.19                               If Val(StrConv(LOTNOREC.OQty, vbUnicode)) > Val(StrConv(LOTNOREC.SQty, vbUnicode)) Then
                                If Val(IN_BCR(2)) > Val(StrConv(LOTNOREC.SQty, vbUnicode)) Then '2014.07.19
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                        TYPE_BCANK, "�~���i�ρ@�݌ɐ���" & StrConv(Val(StrConv(LOTNOREC.SQty, vbUnicode)), vbWide), _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    INVNO_OUT_CHECK_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                            Case BtErrKeyNotFound
                            '-------------------------- �f�[�^���o�^
                            
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                INVNO_OUT_CHECK_PROC = False
                                Exit Function
                            
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            
                                
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_DNAME, _
                                                    TYPE_BCANK, "�~�ް��g�p��", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                INVNO_OUT_CHECK_PROC = False
                                Exit Function
                                
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                
                        '-------------------------------------------    ���g�Ǘ��f�[�^�X�V
                        
                        '�o�א�(�o�א�+����)
                        wkNum = Val(StrConv(LOTNOREC.OQty, vbUnicode))
                        wkNum = wkNum + Val(IN_BCR(2))
                        Call UniCode_Conv(LOTNOREC.OQty, Format(wkNum, "000000"))
                        
                        Call UniCode_Conv(INVNOREC.OQty, Format(wkNum, "000000"))           '�o�א�(InvNo)
                        
                        '�݌ɐ�(�݌ɐ�-����)
                        wkNum = Val(StrConv(LOTNOREC.SQty, vbUnicode))
                        wkNum = wkNum - Val(IN_BCR(2))
                        Call UniCode_Conv(LOTNOREC.SQty, Format(wkNum, "000000"))
                        '�o�ד�
                        Call UniCode_Conv(LOTNOREC.ODt, Format(Now, "YYYYMMDD"))
                        '�o�גS����
                        Call UniCode_Conv(LOTNOREC.OTantoCode, ID_KANRI_TBL(ING_No).TANTO_CODE)
                        '�X�V�h�c
                        Call UniCode_Conv(LOTNOREC.UpdID, App.EXEName)
                        '�X�V����
                        Call UniCode_Conv(LOTNOREC.UpdDtm, Format(Now, "YYYYMMDDHHMMSS"))
                        '��������   -------------
                        sts = BTRV(BtOpUpdate, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "���g�Ǘ��f�[�^", 0)
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                        
                        '-------------------------------------------    ���g���󇂃f�[�^�X�V
                        Call UniCode_Conv(INVNOREC.INVNO, ID_KANRI_TBL(ING_No).INVNO)       '����
                        Call UniCode_Conv(INVNOREC.Model, CStr(IN_BCR(0)))                  '�i��
                        Call UniCode_Conv(INVNOREC.LotNo, CStr(IN_BCR(1)))                  '�����ԍ�
                        Call UniCode_Conv(INVNOREC.ODt, Format(Now, "YYYYMMDD"))            '�o�ד�
                        
                        
                        Call UniCode_Conv(INVNOREC.FILLER, "")                              '�o�^ID
                        
                        Call UniCode_Conv(INVNOREC.EntID, App.EXEName)                      '�o�^ID
                        Call UniCode_Conv(INVNOREC.EntDtm, Format(Now, "YYYYMMDDHHMMSS"))   '�o�^����
                        Call UniCode_Conv(INVNOREC.UpdID, "")                               '�X�VID
                        Call UniCode_Conv(INVNOREC.UpdDtm, "")                              '�X�V����
                        
                        
                        '��������   -------------
                        sts = BTRV(BtOpInsert, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                Call File_Error(sts, BtOpInsert, "���g���󇂃f�[�^", 0)
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                        
                        '-------------------------------------------    ��ƃ��O�o��
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                        If Trim(MENU_NO) = "" Then
                        Else
                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                                (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                                ID_KANRI_TBL(ING_No).NAIGAI, _
                                                                MENU_NO, _
                                                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                                 CStr(IN_BCR(0)), Val(IN_BCR(2)), , , , _
                                                                 CStr(IN_BCR(1))) Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                        End If
                        '-------------------------------------------    �X�V����
                        
                        
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DEF, _
                                                    , "����:" & ID_KANRI_TBL(ING_No).INVNO, _
                                                    TYPE_BCANK, "���o�׌��i�n�j", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                End Select
            Next i
        
            
            
            
            
    
    
    End Select

    INVNO_OUT_CHECK_PROC = False


End Function

Public Function INVNO_OUT_CANCEL_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���g�[�������@�o�׃L�����Z������(���󇂕t��)�x
'
'       2014.07.01
'-------------------------------------------------------

Dim sts             As Integer

Dim Model           As String * 20
Dim PLotNo          As String * 20
Dim INVNO           As String * 20


Dim i               As Integer
Dim MENU_NO         As String * 2


Dim IN_BCR          As Variant


Dim wkNum           As Long


Dim com             As Integer

Dim Rec_cnt         As Long             '2014.07.30


    INVNO_OUT_CANCEL_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁ������������ʁj
            For i = 0 To M_Gyo - 1
                
                'Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)     '2014.07.24
                '    Case LCD_LotNo_BCR      'BC                            '2014.07.24
    
                Select Case i                                               '2014.07.24
                    Case 1      'BC                                         '2014.07.24
    
    
                        IN_BCR = Split(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), " ", -1)
                        
                        If UBound(IN_BCR) < 2 Then
                    
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            INVNO_OUT_CANCEL_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        If Not IsNumeric(IN_BCR(2)) Then
                        
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�[�R�[�h�ُ�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            INVNO_OUT_CANCEL_PROC = False
                            Exit Function
                        
                        End If
                        
                        
                        
                        
                        '------------------ ���g�Ǘ��f�[�^�Ǎ���
                        sts = LotNo_Check_Proc(CStr(IN_BCR(0)), _
                                                CStr(IN_BCR(1)), _
                                                0)
                        
                        Select Case sts
                            Case False
                            
                                If Val(StrConv(LOTNOREC.OQty, vbUnicode)) = 0 Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                        TYPE_BCANK, "�~���o��", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    INVNO_OUT_CANCEL_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                                If Val(StrConv(LOTNOREC.OQty, vbUnicode)) < Val(IN_BCR(2)) Then
                                    
                                    
                                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        
                                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                        TYPE_BCANK, "�~�o�א��s���i" & StrConv(Val(StrConv(LOTNOREC.OQty, vbUnicode)), vbWide) & ")", _
                                                        , "�i��:" & CStr(IN_BCR(0)), _
                                                        , "����:" & CStr(IN_BCR(1)), _
                                                        , "����:" & Val(IN_BCR(2)))
                                    
                                    INVNO_OUT_CANCEL_PROC = False
                                    Exit Function
                                    
                                End If
                            
                            
                            Case BtErrKeyNotFound
                            '-------------------------- �f�[�^���o�^
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                    TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                INVNO_OUT_CANCEL_PROC = False
                                Exit Function
                            
                            
                            
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            
                            
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    
                                Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DOUBLE, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                    TYPE_BCANK, "�~�ް��g�p��", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "����:" & Val(IN_BCR(2)))
                                
                                INVNO_OUT_CANCEL_PROC = False
                                Exit Function
                            
                            
                            
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                        
                        
                                                
                        
                        
                        
                        
                        
                        
                        
'                        If Val(StrConv(LOTNOREC.IQty, vbUnicode)) > 1 Then         2014.07.30
'
'                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
'
'                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
'                                                Buzzer_DOUBLE, _
'                                                , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
'                                                TYPE_BCANK, "�~�o�א����P", _
'                                                , "�i��:" & CStr(IN_BCR(0)), _
'                                                , "����:" & CStr(IN_BCR(1)), _
'                                                , "����:" & Val(IN_BCR(2)))
'
'                            INVNO_OUT_CANCEL_PROC = False
'                            Exit Function
'
'                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        ID_KANRI_TBL(ING_No).Model = CStr(IN_BCR(0))
                        ID_KANRI_TBL(ING_No).PLotNo = CStr(IN_BCR(1))
                        ID_KANRI_TBL(ING_No).SURYO = Val(IN_BCR(2))
                        
                        ID_KANRI_TBL(ING_No).SQty = Val(StrConv(LOTNOREC.SQty, vbUnicode))
                        
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                    Buzzer_DEF, _
                                                    , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                    , "��낵���ł����H", _
                                                    , "�i��:" & CStr(IN_BCR(0)), _
                                                    , "����:" & CStr(IN_BCR(1)), _
                                                    , "�݌ɐ�:" & Val(StrConv(LOTNOREC.SQty, vbUnicode)) & "��" & Val(StrConv(LOTNOREC.SQty, vbUnicode)) + Val(IN_BCR(2)))
                        
                        
                End Select
            Next i
        
            
            
            
            
        Case Step_Sagyo2_RES        '�R��ڂ̎�M�iAny Key�j
            
            
            '-------------------------------------------    ���g�Ǘ��f�[�^�X�V
            sts = LotNo_Check_Proc(ID_KANRI_TBL(ING_No).Model, _
                                    ID_KANRI_TBL(ING_No).PLotNo, _
                                    0)
            Select Case sts
                Case False
                
                    If Val(StrConv(LOTNOREC.OQty, vbUnicode)) = 0 Then
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                            Buzzer_DOUBLE, _
                                            , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                            TYPE_BCANK, "�~���o��", _
                                            , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                            , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                            , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                        
                        INVNO_OUT_CANCEL_PROC = False
                        Exit Function
                    End If
                
                
                    If Val(StrConv(LOTNOREC.OQty, vbUnicode)) < ID_KANRI_TBL(ING_No).SURYO Then
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
                        Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                            Buzzer_DOUBLE, _
                                            , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                            TYPE_BCANK, "�~�o�א��s���i" & StrConv(Val(StrConv(LOTNOREC.OQty, vbUnicode)), vbWide), _
                                            , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                            , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                            , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                        
                        INVNO_OUT_CANCEL_PROC = False
                        Exit Function
                        
                    End If
                
                
                Case BtErrKeyNotFound
                '-------------------------- �f�[�^���o�^
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
        
                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DOUBLE, _
                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                        TYPE_BCANK, "�~�Y���ް��Ȃ�", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                    
                    INVNO_OUT_CANCEL_PROC = False
                    Exit Function
                
                
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
        
                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DOUBLE, _
                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                        TYPE_BCANK, "�~�ް��g�p��", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                    
                    INVNO_OUT_CANCEL_PROC = False
                    Exit Function
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
            
            
            
            
            
            
            
            
            
            
'            If Val(StrConv(LOTNOREC.IQty, vbUnicode)) > 1 Then                 '2014.07.30
'
'                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
'
'                Call Wel_LOTNO_Make_PROC(Sendbuf, _
'                                    Buzzer_DOUBLE, _
'                                    , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
'                                    TYPE_BCANK, "�~�o�א����P", _
'                                    , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
'                                    , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
'                                    , "����:" & ID_KANRI_TBL(ING_No).SURYO)
'
'                INVNO_OUT_CANCEL_PROC = False
'                Exit Function
'
'            End If
            
            
            '�o�א�(�o�א�+����)
            wkNum = Val(StrConv(LOTNOREC.OQty, vbUnicode))
            wkNum = wkNum - ID_KANRI_TBL(ING_No).SURYO
            Call UniCode_Conv(LOTNOREC.OQty, Format(wkNum, "000000"))
            '�݌ɐ�(�݌ɐ�-����)
            wkNum = Val(StrConv(LOTNOREC.SQty, vbUnicode))
            wkNum = wkNum + ID_KANRI_TBL(ING_No).SURYO
            Call UniCode_Conv(LOTNOREC.SQty, Format(wkNum, "000000"))
            '�o�ד�/�o�גS���҂͂��̂܂�
                            
            '�X�V�h�c
            Call UniCode_Conv(LOTNOREC.UpdID, App.EXEName)
            '�X�V����
            Call UniCode_Conv(LOTNOREC.UpdDtm, Format(Now, "YYYYMMDDHHMMSS"))
            '��������   -------------
            sts = BTRV(BtOpUpdate, LOTNO_POS, LOTNOREC, Len(LOTNOREC), K0_LOTNO, Len(K0_LOTNO), 0)
            Select Case sts
                Case BtNoErr
                Case Else
                    Call File_Error(sts, BtOpUpdate, "���g�Ǘ��f�[�^", 0)
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
            '-------------------------------------------    ���g���󇂃f�[�^�폜        2014.07.30 �X�V�ʒu���ړ�
        
'            Call UniCode_Conv(K0_INVNO.Model, RTrim(ID_KANRI_TBL(ING_No).Model))
'            Call UniCode_Conv(K0_INVNO.LotNo, RTrim(ID_KANRI_TBL(ING_No).PLotNo))
'
'            sts = BTRV(BtOpGetEqual, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
'            Select Case sts
'                Case BtNoErr
'                Case BtErrKeyNotFound
'                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                Sendbuf = Text_Create_Proc()
'                Exit Function
'            End Select
'
'
'
'            sts = BTRV(BtOpDelete, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
'            Select Case sts
'                Case BtNoErr
'                Case BtErrKeyNotFound
'                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                Sendbuf = Text_Create_Proc()
'                Exit Function
'            End Select
            
            
            
            
            
            
            '-------------------------------------------    ��ƃ��O�o��
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
            If Trim(MENU_NO) = "" Then
            Else
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     RTrim(ID_KANRI_TBL(ING_No).Model), ID_KANRI_TBL(ING_No).SURYO, , , , _
                                                     RTrim(ID_KANRI_TBL(ING_No).PLotNo)) Then
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            End If
            
            
            
            '-------------------------------------------    ���g���󇂃f�[�^�폜        2014.07.30 �X�V�ʒu���ړ�
            
            '--------- > �����`�F�b�N   2014.07.30
            
            If InvNo_Rec_Cnt(Rec_cnt, ID_KANRI_TBL(ING_No).Model, ID_KANRI_TBL(ING_No).PLotNo) Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Select Case Rec_cnt
                Case 0
                
                    Call LOG_OUT(LOG_F, "Model=" & RTrim(ID_KANRI_TBL(ING_No).Model) & " PLotNo=" & ID_KANRI_TBL(ING_No).PLotNo & " (���g�o�׷�ݾ�)���g�����ް����o�^")
                
                Case 1
                
                    Call UniCode_Conv(K0_INVNO.Model, RTrim(ID_KANRI_TBL(ING_No).Model))
                    Call UniCode_Conv(K0_INVNO.LotNo, RTrim(ID_KANRI_TBL(ING_No).PLotNo))
        
                    sts = BTRV(BtOpGetEqual, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "���g�Ǘ��f�[�^", 0)
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                    End Select
        
        
        
                    sts = BTRV(BtOpDelete, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                           Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                    End Select
                
                
                Case Else
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                
                    Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DOUBLE, _
                                        , "�����ް� " & Rec_cnt & " ��", _
                                        TYPE_BCANK, LCD_InvNo_BCR, _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                    
                    INVNO_OUT_CANCEL_PROC = False
                    Exit Function
            
            
            
            End Select
            
            
            '-------------------------------------------    �X�V����
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            
            
            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                        Buzzer_DEF, _
                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                        TYPE_BCANK, "���o�׷�ݾ�OK", _
                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                        , "�݌ɐ�:" & ID_KANRI_TBL(ING_No).SQty + ID_KANRI_TBL(ING_No).SURYO)
    
    
    
    
    
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i����o�[�R�[�h�j         2014.07.30
            
            For i = 0 To M_Gyo - 1
                
    
                Select Case i
                    Case 1
    
    
                        INVNO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
    
    
                        Call UniCode_Conv(K0_INVNO.Model, ID_KANRI_TBL(ING_No).Model)
                        Call UniCode_Conv(K0_INVNO.LotNo, ID_KANRI_TBL(ING_No).PLotNo)
    
                        com = BtOpGetGreaterEqual
    
    
                        Rec_cnt = 0
                        
                        Do
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                            sts = BTRV(com, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                    If RTrim(StrConv(INVNOREC.Model, vbUnicode)) <> RTrim(ID_KANRI_TBL(ING_No).Model) Or _
                                        RTrim(StrConv(INVNOREC.LotNo, vbUnicode)) <> RTrim(ID_KANRI_TBL(ING_No).PLotNo) Then
                                        Exit Do
                                    End If
                                
                                    If RTrim(INVNO) = RTrim(StrConv(INVNOREC.INVNO, vbUnicode)) Then
                                        sts = BTRV(BtOpDelete, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case Else
                                                Call File_Error(sts, BtOpDelete, "���g���󇂃f�[�^", 0)
                                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                Exit Function
                                        End Select
                                    
                                        Rec_cnt = Rec_cnt + 1
                                
                                        Exit Do
                                
                                    End If
                                
                                
                                
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                
                                    Call File_Error(sts, com, "���g���󇂃f�[�^", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            
                            End Select
                        
                            com = BtOpGetNext
                        
                        
                        Loop
    
    
                        If Rec_cnt = 0 Then
    
                            Call LOG_OUT(LOG_F, "����=" & INVNO & " Model=" & RTrim(ID_KANRI_TBL(ING_No).Model) & " PLotNo=" & ID_KANRI_TBL(ING_No).PLotNo & " (���g�o�׷�ݾ�)���g�����ް����o�^")
    
    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DOUBLE, _
                                                        , "�Y������f�[�^�Ȃ�", _
                                                        TYPE_BCANK, "����:" & INVNO, _
                                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                        , "����:" & ID_KANRI_TBL(ING_No).SURYO)
                        
                        Else
                            
    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            wkNum = ID_KANRI_TBL(ING_No).SURYO + ID_KANRI_TBL(ING_No).SQty
                            Call Wel_LOTNO_Make_PROC(Sendbuf, _
                                                        Buzzer_DEF, _
                                                        , ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                        TYPE_BCANK, "���o�׷�ݾ�OK", _
                                                        , "�i��:" & RTrim(ID_KANRI_TBL(ING_No).Model), _
                                                        , "����:" & RTrim(ID_KANRI_TBL(ING_No).PLotNo), _
                                                        , "�݌ɐ�:" & wkNum)
    
    
    
                        End If
                        
                End Select
            Next i
            
    
    
    End Select

    INVNO_OUT_CANCEL_PROC = False



End Function

Private Function InvNo_Rec_Cnt(Rec_cnt As Long, Model As String, LotNo As String) As Integer
'-------------------------------------------------------
'
'   �w���󇂃f�[�^�@�Ǎ��݁x
'       2014.07.30
'-------------------------------------------------------
Dim sts     As Integer
Dim com     As Integer
                        
    InvNo_Rec_Cnt = True
                        
    Call UniCode_Conv(K0_INVNO.Model, Model)
    Call UniCode_Conv(K0_INVNO.LotNo, LotNo)
                                                
    com = BtOpGetGreaterEqual

    Rec_cnt = 0

    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(com, INVNO_POS, INVNOREC, Len(INVNOREC), K0_INVNO, Len(K0_INVNO), 0)
        Select Case sts
            Case BtNoErr
                
                If RTrim(StrConv(INVNOREC.Model, vbUnicode)) <> RTrim(Model) Or _
                    RTrim(StrConv(INVNOREC.LotNo, vbUnicode)) <> RTrim(LotNo) Then
                    Exit Do
                End If
                
                Rec_cnt = Rec_cnt + 1
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select

        com = BtOpGetNext
    
    Loop

    InvNo_Rec_Cnt = False

End Function


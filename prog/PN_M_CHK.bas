Attribute VB_Name = "PN_M_CHK"


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> MT_2009.06.01
Function PN_CHK(PN_CODE As String, NAIGAI As String, InsTanto As String, Optional Inv_Mode As Integer = 0, Optional Message = 0)

'   �����FPN_CODE           �`�F�b�N�Ώەi��
'         NaiGai            G�i�O���i�ԁj�AN�i�����i�ԁj
'         InsTanto          �h�s�d�l�ǉ��S���ҁi�������̓v���O�������j

Dim yn          As Integer
Dim W_Msg       As String
Dim W_STR       As String
    
Dim sts         As Integer
    
Dim sts1        As Integer
    
    
    PN_CHK = True
    
    
    
    
    
    Select Case NAIGAI
        Case "G"
            sts = PN_M_GET(Last_JGYOBU, PN_CODE, 0)
            If sts Then
                If sts = BtErrKeyNotFound Then
                    If Inv_Mode = 1 Then
                        MsgBox "���͂������ڂ̓G���[�ł��B�i�O���i�ԁj"
                        Exit Function
                    Else
                    End If
                Else
                    MsgBox "���͂������ڂ̓G���[�ł��B�i�O���i�ԁj"
                    Exit Function
                End If
            End If
        Case Else
            sts = PN_M_GET2(Last_JGYOBU, PN_CODE, 0)
            If sts Then
                
                If sts = BtErrKeyNotFound Then
                    If Inv_Mode = 1 Then
                        MsgBox "���͂������ڂ̓G���[�ł��B�i�����i�ԁj"
                        Exit Function
                    Else
                    End If
                Else
                    MsgBox "���͂������ڂ̓G���[�ł��B�i�����i�ԁj"
                    Exit Function
                End If
            
            Else
            
            
                If Message = 1 Then
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))
                    sts1 = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts1
                        
                        Case BtNoErr
                        
                            W_Msg = "�Γ��i�Ԃ�ύX���܂����H" & Chr(13) & Chr(10)
                            W_Msg = W_Msg & "�ΊO�i�ԁF" & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & Chr(13) & Chr(10)
                            W_Msg = W_Msg & "�Γ��i�ԁF" & Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) & "��" & Trim(PN_CODE) & Chr(13) & Chr(10)

                        
                        
                            yn = MsgBox(W_Msg, vbYesNo + vbDefaultButton2, "�Γ��i�ԕύX�m�F")
                            If yn = vbYes Then
                    
                    
                    
                                Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))     '�i�ԁi�����j
            
    
    
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, InsTanto)        '�ǉ��@�S����
        
                                W_Date = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))  '�ǉ��@����
                            
                            
                                Do
                                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                            If ans = vbCancel Then
                                                PN_CHK = False
                                                Exit Function
                                            End If
                                            
                                        Case Else
                                            Call File_Error(sts, BtOpInsert, "�i�ڃ}�X�^")
                                            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                                            Exit Function
                                    End Select
                                Loop
                    
                            End If
                                
                            PN_CHK = False
                            Exit Function
                        
                        Case BtErrKeyNotFound
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                            Exit Function
                    End Select
                End If
            End If
    
    End Select
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Y���Ή�  2012.01.23  �폜 2012.02.06
'    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
'    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
'    Select Case sts
'        Case BtNoErr
'
'        Case BtErrKeyNotFound
'            Call UniCode_Conv(CountryREC.CountryName, "")
'            Call UniCode_Conv(CountryREC.CountryName2, "")
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "Country�}�X�^")
'            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
'            Exit Function
'    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Y���Ή�  2012.01.23
    
    W_Msg = ""
                                
    W_Msg = W_Msg & "�Y������i�Ԃ�����܂���B" & Chr(13) & Chr(10)
    W_Msg = W_Msg & "���L�̕i�Ԃ�V�K�o�^���܂����H" & Chr(13) & Chr(10)
    W_Msg = W_Msg & " " & Chr(13) & Chr(10)
    W_Msg = W_Msg & "���ƕ��@�F" & Last_JGYOBU & Chr(13) & Chr(10)
    W_Msg = W_Msg & "�����O�@�F1 ����" & Chr(13) & Chr(10)
    W_Msg = W_Msg & "�i�@�ԁ@�F" & RTrim(StrConv(PN_MREC.PN, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "�Γ��i�ԁF" & RTrim(StrConv(PN_MREC.SPn, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "�i�@���@�F" & RTrim(StrConv(PN_MREC.PName, vbUnicode)) & Chr(13) & Chr(10)
    
    
    
    
    
    
    
    If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka2, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "�P���P�@�F" & W_STR & Chr(13) & Chr(10)
    
    If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka3, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "�P���Q�@�F" & W_STR & Chr(13) & Chr(10)
    
    If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
        W_STR = Format(CDbl(Trim(StrConv(PN_MREC.Tanka4, vbUnicode))), "###,##0.00")
    Else
        W_STR = "0"
    End If
    W_Msg = W_Msg & "�P���R�@�F" & W_STR & Chr(13) & Chr(10)
    W_Msg = W_Msg & "�����\�����Y�� �F" & RTrim(StrConv(PN_MREC.MadeIn, vbUnicode)) & Chr(13) & Chr(10)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Y���Ή�  2012.01.23  --> ���Y��  2012.02.06
'    W_Msg = W_Msg & "MadeInCode�F" & RTrim(StrConv(PN_MREC.MadeInCode, vbUnicode)) & Chr(13) & Chr(10)
    W_Msg = W_Msg & "���Y���F" & RTrim(StrConv(PN_MREC.GENSANKOKU, vbUnicode)) & Chr(13) & Chr(10)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   ���Y���Ή�  2012.01.23
                                
    yn = MsgBox(W_Msg, vbYesNo + vbExclamation, "�m�F����")
    If yn = vbNo Then
        
        Exit Function
    End If

    If Item_PUT_Proc(InsTanto) Then
        W_STR = "�ǉ��Ɏ��s���܂����B" & Chr(13) & Chr(10) & "�Ď��s�肢�܂��B"
        MsgBox W_STR
        
        Exit Function
    End If
    
    
    PN_CHK = False
    
    
End Function


'                                           MT_2009.06.01
Function Item_PUT_Proc(InsTanto As String) As Integer

Dim sts         As Integer
Dim ans         As Integer
Dim W_Date      As String

    Item_PUT_Proc = True
    
    
    
    
    
    
    
    Call Rclr_ITEMREC

    Call UniCode_Conv(ITEMREC.JGYOBU, Last_JGYOBU)                          '���ƕ��敪
    Call UniCode_Conv(ITEMREC.NAIGAI, "1")                                  '�����O

    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))      '�i�ԁi�O���j
    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(PN_MREC.PName, vbUnicode))  '�i��
    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))     '�i�ԁi�����j

    Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(PN_MREC.SOKO, vbUnicode))    '�W�����ɑq�� �q��
    
    Call UniCode_Conv(ITEMREC.ST_RETU, "")          '             ��
    Call UniCode_Conv(ITEMREC.ST_REN, "")           '             �A
    Call UniCode_Conv(ITEMREC.ST_DAN, "")           '             �i
    
    Call UniCode_Conv(ITEMREC.JAN_CODE, "")         'Jan�R�[�h
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")      '�O���b�N�X�I�ԂP
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")      '�O���b�N�X�I�ԂQ
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")      '�O���b�N�X�I�ԂR


''*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
'    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, StrConv(PN_MREC.PName, vbUnicode))      '���i����   �i��

                                                    '           ���i(1)
    
        
    
    
    If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
    End If
                                                    
                                                    
                                                    '           ���i(2)
    If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
    End If
                                                    
                                                    
                                                    '           ���i(3)
    If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
        Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
    Else
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
    End If
    
    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")    '           �K�p�@����l(���@��i�R�j)
    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")    '           ��Ǝw��
    Call UniCode_Conv(ITEMREC.L_BIKOU3, "")         '           ���l�R
                                                    
                                                    
                                                    
'Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, Last_JGYOBU)
                                                    '           ���萔
    Call UniCode_Conv(ITEMREC.L_IRI_QTY, String(UBound(ITEMREC.L_IRI_QTY) + 1, "0"))
''*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
    Call UniCode_Conv(ITEMREC.S_TANTO, "")          '���P�^�S���҃R�[�h


    
    
    Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))      '���Y��
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.23 ���Y����-->PN���ɕύX  2012.02.06
'    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(PN_MREC.GENSANKOKU, vbUnicode))
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2012.01.23 ���Y����
    


    Call UniCode_Conv(ITEMREC.INS_TANTO, InsTanto)          '�ǉ��@�S����
    
    W_Date = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    Call UniCode_Conv(ITEMREC.Ins_DateTime, W_Date)         '�ǉ��@����





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.12.14
    Call UniCode_Conv(ITEMREC.D_MODEL, StrConv(PN_MREC.DModel, vbUnicode))          '��\�@��i�ں���
    Call UniCode_Conv(ITEMREC.HINMOKU, StrConv(PN_MREC.HINMOKU, vbUnicode))         '�i�ں���
    Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '�W���P��
    Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))        '���`��
    Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      '�Ưċ敪
    Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '���������敪
    Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '���O�����敪
    Call UniCode_Conv(ITEMREC.GLICS1_TANA, StrConv(PN_MREC.Loc1, vbUnicode))        '�I��1
    Call UniCode_Conv(ITEMREC.GLICS2_TANA, StrConv(PN_MREC.Loc2, vbUnicode))        '�I��2
    Call UniCode_Conv(ITEMREC.GLICS3_TANA, StrConv(PN_MREC.Loc3, vbUnicode))        '�I��3
    Call UniCode_Conv(ITEMREC.L_KISHU1, StrConv(PN_MREC.NaiModel, vbUnicode))       '��\�@��1
    Call UniCode_Conv(ITEMREC.L_KISHU2, StrConv(PN_MREC.GaiModel, vbUnicode))       '��\�@��2
    Call UniCode_Conv(ITEMREC.CS_TANTO_CD, StrConv(PN_MREC.KobaiTanto, vbUnicode))  '�w���S����
    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, StrConv(PN_MREC.PNameEngA, vbUnicode))  '�p��i��
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.12.14
    




'---------------------------------------------------------------------------------------------
    Do
        sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Item_PUT_Proc = False
                    Exit Function
                End If
            Case BtErrDuplicates
                MsgBox "���łɒǉ�����Ă��܂��B"
                Exit Function
                
            Case Else
                Call File_Error(sts, BtOpInsert, "�i�ڃ}�X�^")
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                Exit Function
        End Select
    Loop
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)                      '���ƕ��敪
    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")                              '�����O
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(PN_MREC.PN, vbUnicode))  '�i�ԁi�O���j
    Do
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
                
            Case BtErrKeyNotFound
                MsgBox "�i�ڃ}�X�^�@�ǉ����s�I"
                Exit Function
            
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Item_PUT_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                Exit Function
        End Select
    Loop
    
    Item_PUT_Proc = False
    
    
End Function


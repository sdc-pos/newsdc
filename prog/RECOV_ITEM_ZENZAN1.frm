VERSION 5.00
Begin VB.Form RECOV_ITEM_ZENZAN1 
   Caption         =   "�i�ڃ}�X�^�@�O���c���@�������� (2011.03.01 14:00)"
   ClientHeight    =   10296
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14988
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10296
   ScaleWidth      =   14988
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "RECOV_ITEM_ZENZAN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    RECOV_ITEM_ZENZAN1.MousePointer = vbHourglass

    Call Ctrl_Lock(RECOV_ITEM_ZENZAN1)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(RECOV_ITEM_ZENZAN1)


    RECOV_ITEM_ZENZAN1.MousePointer = vbDefault

End Sub


Private Sub Form_Activate()

Dim yn  As Integer


    yn = MsgBox("�i�ڃ}�X�^�O���c�����������@���s���܂����H", vbYesNo + vbDefaultButton2, "�m�F����")

    If yn = vbNo Then
        Unload Me
    End If

    If Next_Proc() Then
    End If

    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub


Private Sub Form_Load()

Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer

'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If


    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)







                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '���ޒI�����n�o�d�m
    If P_STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If





End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer


                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If

                                            '�݌��ް��b�k�n�r�d

                                            '���ޒI�����b�k�n�r�d
    sts = BTRV(BtOpClose, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޒI��")
        End If
    End If

    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set RECOV_ITEM_ZENZAN1 = Nothing


    End
End Sub



Private Function Next_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���ތJ�z����
'----------------------------------------------------------------------------
Dim sts                     As Integer
Dim com                     As Integer


Dim i                       As Integer

Dim wk_VAL                  As Long

Dim Skip_Flg                As Boolean

Dim SYUSHI_ON               As Boolean          '2007.11.13

Dim Sum_Zen_Zaiko           As Long
Dim Sum_Zaiko               As Long
Dim Sum_Nyuko               As Long
Dim Sum_Syuko               As Long

Dim Sum_Zaiko_Kin           As Long




Dim svJGYOBU                As String * 1
Dim svNAIGAI                As String * 1
Dim svHIN_GAI               As String * 20




    Next_Proc = True
    RECOV_ITEM_ZENZAN1.MousePointer = vbHourglass






    '-------------------------------------  �������c��Ă���

    svJGYOBU = ""
    svNAIGAI = ""
    svHIN_GAI = ""


    com = BtOpGetFirst

    Do

        DoEvents

        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K0_P_STOCK, Len(K0_P_STOCK), 0)

        Select Case sts
            Case BtNoErr

            Case BtErrEOF
                Exit Do

            Case Else
                Call File_Error(sts, com, "���ޒI���ް�")
                Exit Function

        End Select


        If Trim(svJGYOBU) = "" Then
            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)

            Sum_Zen_Zaiko = 0
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0

            Sum_Zaiko_Kin = 0
        End If


        If svJGYOBU <> StrConv(P_STOCK_REC.JGYOBU, vbUnicode) Or _
            svNAIGAI <> StrConv(P_STOCK_REC.NAIGAI, vbUnicode) Or _
            svHIN_GAI <> StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) Then

            If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Nyuko = 0 And Sum_Syuko = 0 Then


            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)

                Skip_Flg = False
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                Select Case sts
                    Case BtNoErr

                    Case BtErrKeyNotFound
                        Skip_Flg = True

                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function

                End Select


                If Not Skip_Flg Then

                    If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                    End If

                    wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))
                    wk_VAL = 0
                
                    wk_VAL = wk_VAL + Sum_Zaiko_Kin

                    If wk_VAL < 0 Then
                        wk_VAL = 0
                    End If
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(wk_VAL, "00000000000"))

                    If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "000000000")
                    End If

                    wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
                    wk_VAL = 0
                                
                    wk_VAL = wk_VAL + Sum_Zaiko

                    If wk_VAL < 0 Then
                        wk_VAL = 0
                    End If
                    
                    
                    
If wk_VAL <> Val(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
    Call LOG_OUT(LOG_F, "[ITEM_No]" & StrConv(ITEMREC.HIN_GAI, vbUnicode) & "[�C���O]" & Val(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) & "[�C����]" & wk_VAL)
End If
                    
                    
                    
                    
                    
                    
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(wk_VAL, "00000000"))




                    Do
                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                        Select Case sts
                            Case BtNoErr
                                Exit Do

                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                DoEvents

                            Case Else
                                Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                                Exit Function

                        End Select
                    Loop

                End If

            End If

            svJGYOBU = StrConv(P_STOCK_REC.JGYOBU, vbUnicode)
            svNAIGAI = StrConv(P_STOCK_REC.NAIGAI, vbUnicode)
            svHIN_GAI = StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)


            Sum_Zen_Zaiko = 0
            Sum_Zaiko = 0
            Sum_Nyuko = 0
            Sum_Syuko = 0


            Sum_Zaiko_Kin = 0

        End If


        If Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" And Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" Then

            If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                Sum_Zen_Zaiko = Sum_Zen_Zaiko + CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))

            End If

        Else
            If IsNumeric(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode)) Then
                Sum_Nyuko = Sum_Nyuko + CLng(StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
            End If

            If IsNumeric(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode)) Then
                Sum_Syuko = Sum_Syuko + CLng(StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))
            End If


            If IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
                Sum_Zaiko = Sum_Zaiko + CLng(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
            End If



            If IsNumeric(StrConv(P_STOCK_REC.TANKA, vbUnicode)) And IsNumeric(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then

                Sum_Zaiko_Kin = Sum_Zaiko_Kin + ToRoundUp(CCur(StrConv(P_STOCK_REC.TANKA, vbUnicode)) * _
                                CCur(StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode)), 0)
            End If

        End If






        com = BtOpGetNext

    Loop



    If Trim(svJGYOBU) <> "" Then

        If Sum_Zen_Zaiko = Sum_Zaiko And Sum_Nyuko = 0 And Sum_Syuko = 0 Then

        Else
            Call UniCode_Conv(K0_ITEM.JGYOBU, svJGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, svNAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, svHIN_GAI)


            Skip_Flg = False
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)


            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound


                    Skip_Flg = True
                Case Else

                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select


            If Not Skip_Flg Then


                If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "00000000000")
                End If

                wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode))

                wk_VAL = wk_VAL + Sum_Zaiko_Kin

                If wk_VAL < 0 Then
                    wk_VAL = 0
                End If
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(wk_VAL, "00000000000"))


                If Not IsNumeric(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "000000000")
                End If


                wk_VAL = CLng(StrConv(ITEMREC.G_ZEN_ZAIKO_QTY, vbUnicode))
                wk_VAL = wk_VAL + Sum_Zaiko

                If wk_VAL < 0 Then
                    wk_VAL = 0
                End If
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, Format(wk_VAL, "00000000"))


                Do
                    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            DoEvents
                        Case Else

                            Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                            Exit Function
                    End Select


                Loop

            End If

        End If
    End If




    MsgBox "�����������I�����܂����B"

    Next_Proc = False

End Function


VERSION 5.00
Begin VB.Form F1020901 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�������Y�����i���ɊǗ����X�g�@(F102090 2011.07.14 12:00)"
   ClientHeight    =   3312
   ClientLeft      =   2028
   ClientTop       =   2268
   ClientWidth     =   10932
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3312
   ScaleWidth      =   10932
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1020901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NormalFont      As New StdFont              '����t�H���g
Dim NormalBoldFont  As New StdFont              '����t�H���g
Dim SmallFont       As New StdFont              '����t�H���g
Dim SmallBoldFont   As New StdFont              '����t�H���g
Dim LargeFont       As New StdFont              '����t�H���g
Dim LargeUnderFont  As New StdFont              '����t�H���g

Private Const MGN_L% = 1                        '���]���i�����F�P����j
Private Const MGN_U% = 2                        '��]���i�s���F�P����j
Private Const LMax% = 6





Private Sub Form_Activate()

    If Print_Proc() Then
        Unload Me
    End If


    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)


                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

                                '���׃f�[�^�t�@�C��OPEN
    If Y_NYU_Open(0) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C��OPEN
    If ZAIKO_Open(0) Then
        Unload Me
    End If
                                '�����σf�[�^�t�@�C��OPEN
    If AVE_SYUKA_Open(0) Then
        Unload Me
    End If
    Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '���׃f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���׃f�[�^�t�@�C��")
        End If
    End If
                                            '�݌Ƀf�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^�t�@�C��")
        End If
    End If
                                            '�����Ϗo�ׂb�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo��")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020901 = Nothing

    End
End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1020901.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020901)


    F1020901.MousePointer = vbDefault

End Sub
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   �������
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim i           As Integer

Dim Lcnt        As Integer
    
Dim Print_Now   As String
    
    
Dim SUMI_QTY()      As Long
Dim MI_QTY()        As Long
Dim GENSANKOKU()    As String * 20
    
Dim svSoko      As String * 2
Dim svRetu      As String * 2
Dim svRen       As String * 2
Dim svDan       As String * 2
    
    
Dim Maru_Suuji  As String
    
Dim Read_next   As Integer
    
    
    Print_Proc = True

    Print_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    Printer.Orientation = vbPRORLandscape
    Lcnt = 99
    
    With NormalFont
        .NAME = F1020901.Font.NAME
        .Size = 11
        .Bold = False
    
    End With
    With NormalBoldFont
        .NAME = F1020901.Font.NAME
        .Size = 11
        .Bold = True
    End With
    
    With SmallFont
        .NAME = F1020901.Font.NAME
        .Size = 9
        .Bold = False
    End With
    
    With SmallBoldFont
        .NAME = F1020901.Font.NAME
        .Size = 9
        .Bold = True
    End With
    
    
    With LargeFont
        .NAME = F1020901.Font.NAME
        .Size = 14
        .Bold = True
    End With
    
    With LargeUnderFont
        .NAME = F1020901.Font.NAME
        .Size = 14
        .Bold = True
        .Underline = True
    End With
    
    
    
    com = BtOpGetGreater


    Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
    Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
    Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")



    Do
        DoEvents
        
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        
        
        Select Case sts
            Case BtNoErr
                If StrConv(Y_NYUREC.LIST_OUT_END_F, vbUnicode) <> "0" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
        
        
            Case Else
                Call File_Error(sts, com, "���ח\��f�[�^", 0)
                Exit Function

        End Select
        
        
        svSoko = ""
        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) <> "0" Then
        Else
        
        
    
    
            Call UniCode_Conv(K4_ZAIKO.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
            Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
            Call UniCode_Conv(K4_ZAIKO.Retu, "")
            Call UniCode_Conv(K4_ZAIKO.Ren, "")
            Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    
    
            com = BtOpGetGreater
    
            
            
            
            
            Do
            
                DoEvents
                sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
                
                
                Select Case sts
                    Case BtNoErr
                    
                    
                        If Trim(StrConv(Y_NYUREC.JGYOBU, vbUnicode)) <> Trim(StrConv(ZAIKOREC.JGYOBU, vbUnicode)) Or _
                            Trim(StrConv(Y_NYUREC.NAIGAI, vbUnicode)) <> Trim(StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                            Trim(StrConv(Y_NYUREC.HIN_NO, vbUnicode)) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                            
                            Exit Do
                        End If
                                        
                    
                    Case BtErrEOF
                        Exit Do
                
                
                    Case Else
                        Call File_Error(sts, com, "�݌Ƀf�[�^", 0)
                        Exit Function
        
                End Select
            
                If StrConv(Y_NYUREC.ID_NO2, vbUnicode) = StrConv(ZAIKOREC.ID_NO2, vbUnicode) Then
                
                Else
            
                    If Trim(svSoko) = "" Then
                        svSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
                        svRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
                        svRen = StrConv(ZAIKOREC.Ren, vbUnicode)
                        svDan = StrConv(ZAIKOREC.Dan, vbUnicode)
                        
                        Erase SUMI_QTY
                        Erase MI_QTY
                        Erase GENSANKOKU
                        
                        ReDim Preserve SUMI_QTY(0 To 0)
                        ReDim Preserve MI_QTY(0 To 0)
                        ReDim Preserve GENSANKOKU(0 To 0)
                        GENSANKOKU(0) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(0) = 0
                        MI_QTY(0) = 0
                    End If
                
                    If svSoko <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        svRetu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        svRen <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        svDan <> StrConv(ZAIKOREC.Dan, vbUnicode) Then
                
                
                
                        For i = 0 To UBound(GENSANKOKU)
                
                
                            
                            If Head_Print_Proc(Print_Now, Lcnt) Then
                                Exit Function
                            End If
        
                            Lcnt = Lcnt + 1
                            
                            Select Case Lcnt
                            
                            
                            
                                Case 1
                                    Maru_Suuji = "�@"
                                Case 2
                                    Maru_Suuji = "�A"
                                Case 3
                                    Maru_Suuji = "�B"
                                Case 4
                                    Maru_Suuji = "�C"
                                Case 5
                                    Maru_Suuji = "�D"
                                Case 6
                                    Maru_Suuji = "�E"
                                Case 7
                                    Maru_Suuji = "�F"
                                Case 8
                                    Maru_Suuji = "�G"
                                Case 9
                                    Maru_Suuji = "�H"
                                Case 10
                                    Maru_Suuji = "�I"
                                Case 11
                                    Maru_Suuji = "�J"
                                Case 12
                                    Maru_Suuji = "�K"
                                Case 13
                                    Maru_Suuji = "�L"
                                Case 14
                                    Maru_Suuji = "�M"
                                Case 15
                                    Maru_Suuji = "�N"
                                Case 16
                                    Maru_Suuji = "�O"
                                Case 17
                                    Maru_Suuji = "�P"
                                Case 18
                                    Maru_Suuji = "�Q"
                                Case 19
                                    Maru_Suuji = "�R"
                                Case 20
                                    Maru_Suuji = "�S"
                                Case Else
                                    Maru_Suuji = "�`"
                            End Select
                            
                            
                            
                            '------------------------------------   17�s��
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 13);
                            Printer.Print "���Y�����F";
                            Printer.Print Tab(MGN_L + 56);      '60-->56
                            Printer.Print "����" & Maru_Suuji & "(���i���� - �����i)";
                            Printer.Print Tab(MGN_L + 85);
                            Printer.Print "�I��" & Maru_Suuji;
                            Printer.Print Tab(MGN_L + 105);
                            Printer.Print "���i�[";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 141);
                            Printer.Print "���ٔ��s";
                            Printer.Print Tab(MGN_L + 151);
                            Printer.Print "��p����";
                            Printer.Print Tab(MGN_L + 161);
                            Set Printer.Font = NormalFont
                            Printer.Print "�@�@���l"
                            '------------------------------------   18�s��
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 5);
                            
                            If Lcnt = 1 Then
                                Printer.Print "�݌ɕi";
                            Else
                                Printer.Print "�@�@�@";
                            End If
                            Printer.Print Tab(MGN_L + 13);
                            Set Printer.Font = NormalBoldFont
                            Printer.Print StrConv(Trim(GENSANKOKU(i)), vbWide);  'trim
                            Printer.Print Tab(MGN_L + 49);      '53-->49
                            Printer.Print StrConv(Space(8 - Len(Format(SUMI_QTY(i) + MI_QTY(i), "#,0"))) & _
                                            Format(SUMI_QTY(i) + MI_QTY(i), "#,0") & "(" & _
                                            Format(SUMI_QTY(i), "#,0") & "-" & _
                                            Format(MI_QTY(i), "#,0") & ")", vbWide);
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 85);
                            Printer.Print svSoko & "-" & svRetu & "-" & svRen & "-" & svDan;
                            Printer.Print Tab(MGN_L + 107);
                            Printer.Print "�|";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 143);     '114-->143
                            Set Printer.Font = NormalFont
                            Printer.Print "��";
                            Set Printer.Font = SmallFont
                            Printer.Print Tab(MGN_L + 153);     '121-->153
                            Set Printer.Font = NormalFont
                            Printer.Print "��"
                            '------------------------------------   19�s��
                            Set Printer.Font = NormalFont
                            Printer.Print Tab(MGN_L + 5);
                            Printer.Print String(70, "��")
        
        
        
                        Next i
                    
                    
                    
                    
                    
                    
                        svSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
                        svRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
                        svRen = StrConv(ZAIKOREC.Ren, vbUnicode)
                        svDan = StrConv(ZAIKOREC.Dan, vbUnicode)
                        
                        Erase SUMI_QTY
                        Erase MI_QTY
                        Erase GENSANKOKU
                        
                        ReDim Preserve SUMI_QTY(0 To 0)
                        ReDim Preserve MI_QTY(0 To 0)
                        ReDim Preserve GENSANKOKU(0 To 0)
                        GENSANKOKU(0) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(0) = 0
                        MI_QTY(0) = 0
                    
                    
                    
                    
                    
                    End If
                    
                    
                    For i = 0 To UBound(GENSANKOKU)
                    
                    
                        If Trim(StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)) = Trim(GENSANKOKU(i)) Then
                            Exit For
                        End If
                    
                    Next i
                    
                    
                    
                    If i > UBound(GENSANKOKU) Then
                    
                    
                        ReDim Preserve SUMI_QTY(0 To i)
                        ReDim Preserve MI_QTY(0 To i)
                        ReDim Preserve GENSANKOKU(0 To i)
                        GENSANKOKU(i) = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
                        SUMI_QTY(i) = 0
                        MI_QTY(i) = 0
                    
                    End If
                    
                    
                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                        SUMI_QTY(i) = SUMI_QTY(i) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                        MI_QTY(i) = MI_QTY(i) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                
                End If
                
                
                com = BtOpGetNext
    
            Loop
    
    
            If Trim(svSoko) <> "" Then
            
                For i = 0 To UBound(GENSANKOKU)
        
        
                    If Head_Print_Proc(Print_Now, Lcnt) Then
                        Exit Function
                    End If
    
                    Lcnt = Lcnt + 1
                    
                    
                    Select Case Lcnt
                    
                    
                    
                        Case 1
                            Maru_Suuji = "�@"
                        Case 2
                            Maru_Suuji = "�A"
                        Case 3
                            Maru_Suuji = "�B"
                        Case 4
                            Maru_Suuji = "�C"
                        Case 5
                            Maru_Suuji = "�D"
                        Case 6
                            Maru_Suuji = "�E"
                        Case 7
                            Maru_Suuji = "�F"
                        Case 8
                            Maru_Suuji = "�G"
                        Case 9
                            Maru_Suuji = "�H"
                        Case 10
                            Maru_Suuji = "�I"
                        Case 11
                            Maru_Suuji = "�J"
                        Case 12
                            Maru_Suuji = "�K"
                        Case 13
                            Maru_Suuji = "�L"
                        Case 14
                            Maru_Suuji = "�M"
                        Case 15
                            Maru_Suuji = "�N"
                        Case 16
                            Maru_Suuji = "�O"
                        Case 17
                            Maru_Suuji = "�P"
                        Case 18
                            Maru_Suuji = "�Q"
                        Case 19
                            Maru_Suuji = "�R"
                        Case 20
                            Maru_Suuji = "�S"
                        Case Else
                            Maru_Suuji = "�`"
                    End Select
                    
                    '------------------------------------   17�s��
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 13);
                    Printer.Print "���Y�����F";
                    Printer.Print Tab(MGN_L + 56);      '60-->56
                    Printer.Print "����" & Maru_Suuji & "(���i���� - �����i)";
                    Printer.Print Tab(MGN_L + 85);
                    Printer.Print "�I��" & Maru_Suuji;
                    Printer.Print Tab(MGN_L + 105);
                    Printer.Print "���i�[";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 141);
                    Printer.Print "���ٔ��s";
                    Printer.Print Tab(MGN_L + 151);
                    Printer.Print "��p����";
                    Printer.Print Tab(MGN_L + 161);
                    Set Printer.Font = NormalFont
                    Printer.Print "�@�@���l"
                    '------------------------------------   18�s��
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 5);
                    
                    If Lcnt = 1 Then
                        Printer.Print "�݌ɕi";
                    Else
                        Printer.Print "�@�@�@";
                    End If
                    Printer.Print Tab(MGN_L + 13);
                    Set Printer.Font = NormalBoldFont
                    Printer.Print StrConv(Trim(GENSANKOKU(i)), vbWide);
                    Printer.Print Tab(MGN_L + 49);      '53-->49
                    Printer.Print StrConv(Space(8 - Len(Format(SUMI_QTY(i) + MI_QTY(i), "#,0"))) & _
                                    Format(SUMI_QTY(i) + MI_QTY(i), "#,0") & "(" & _
                                    Format(SUMI_QTY(i), "#,0") & "-" & _
                                    Format(MI_QTY(i), "#,0") & ")", vbWide);
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 85);
                    Printer.Print svSoko & "-" & svRetu & "-" & svRen & "-" & svDan;
                    Printer.Print Tab(MGN_L + 107);
                    Printer.Print "�|";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 143);     '114-->143
                    Set Printer.Font = NormalFont
                    Printer.Print "��";
                    Set Printer.Font = SmallFont
                    Printer.Print Tab(MGN_L + 153);     '121-->153
                    Set Printer.Font = NormalFont
                    Printer.Print "��"
                    '------------------------------------   19�s��
                    Set Printer.Font = NormalFont
                    Printer.Print Tab(MGN_L + 5);
                    Printer.Print String(70, "��")
    
    
                Next i
    
    
    
    
    
            End If
        
        
        

        
        End If
        







        If Trim(svSoko) <> "" Then
            Lcnt = 7
        
        Else
        
        End If


        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "0" Then
            Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "9")
        
            If Trim(svSoko) = "" Then
                If Head_Print_Proc(Print_Now, Lcnt) Then
                    Exit Function
                End If
            
                Lcnt = 7
        
        
            End If
        
        End If
        
        If StrConv(Y_NYUREC.LIST_NYU_KANRI_F, vbUnicode) = "9" And StrConv(Y_NYUREC.LIST_NYU_CHECK_F, vbUnicode) = "9" Then
            Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "9")
            
            
                        
            
            Read_next = BtOpGetGreater
        
        
        Else
            Read_next = BtOpGetNext
        End If
        
        
        
        sts = BTRV(BtOpUpdate, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        If sts <> BtNoErr Then
        
            Call File_Error(sts, BtOpUpdate, "���ח\��f�[�^", 0)
            Exit Function

        End If


        If Read_next = BtOpGetGreater Then

            Call UniCode_Conv(K4_Y_NYU.LIST_OUT_END_F, "0")
            Call UniCode_Conv(K4_Y_NYU.JGYOBU, "")
            Call UniCode_Conv(K4_Y_NYU.NAIGAI, "")
            Call UniCode_Conv(K4_Y_NYU.HIN_NO, "")
    
        End If
        com = Read_next
    Loop

    Printer.EndDoc

    Print_Proc = False

End Function


Private Function Head_Print_Proc(Print_Now As String, Lcnt As Integer) As Integer
'----------------------------------------------------------------------------
'                   �w�b�_�[����
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
    Head_Print_Proc = True


    If Lcnt < 6 Then
        Head_Print_Proc = False
        Exit Function
    End If

    If Lcnt = 99 Then
    Else
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    
    '------------------------------------   1�s��
    Printer.Print Tab(MGN_L + 55);
    Set Printer.Font = LargeFont

    Printer.Print "�������Y�����i���ɊǗ����X�g";

    Printer.Print Tab(MGN_L + 100);
    Set Printer.Font = SmallFont
    Printer.Print Print_Now

    '------------------------------------   2�s��
    Printer.Print
    '------------------------------------   3�s��
    Printer.Print
    
    
    '------------------------------------   6�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "����������������������������������";
    Printer.Print
    
    
    
    '------------------------------------   6�s��
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 7);
    Printer.Print "�i��";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "�i��";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "�����Ϗo�א�";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "���@���@�F�@��@���@��Ɗ�����@��";
    Printer.Print
    '------------------------------------   6�s��
    
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "����������������������������������";
    Printer.Print
    '------------------------------------   7�s��
    Set Printer.Font = LargeUnderFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print Trim(StrConv(StrConv(Y_NYUREC.HIN_NO, vbUnicode), vbWide));
    
    Set Printer.Font = LargeFont
    Printer.Print Tab(MGN_L + 30);
    Printer.Print StrConv(Y_NYUREC.HIN_NAME, vbUnicode);
    Set Printer.Font = LargeFont
    Printer.Print Tab(MGN_L + 70);
    
    
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    
    
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "0")
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�", 0)
            Exit Function

    End Select
    Printer.Print Space(12 - Len(Format(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode), "#,##0"))) & _
                        Format(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode), "#,##0");
    
    
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "���@�@�@�@�@�@�@���@�@�@�@�@�@�@��";
    Printer.Print
    '------------------------------------   8�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "���@�@�@�@�@�@�@���@�@�@�@�@�@�@��";
    Printer.Print
    '------------------------------------   9�s��
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "���@�@�@�@�@�@�@���@�@�@�@�@�@�@��";
    Printer.Print
    '------------------------------------   10�s��
    Printer.Print Tab(MGN_L + 149);
    Printer.Print "����������������������������������";
    Printer.Print
    '------------------------------------   11�s��
    Printer.Print
    '------------------------------------   12�s��
    Printer.Print
    '------------------------------------   13�s��
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "���Y�����F";
    Printer.Print Tab(MGN_L + 56);      '60-->56
    Printer.Print "���ʁF";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "�I�ԁF";
    Printer.Print Tab(MGN_L + 105);
    Printer.Print "���i�[";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 141);
    Printer.Print "���ٔ��s";
    Printer.Print Tab(MGN_L + 151);
    Printer.Print "��p����";
    Printer.Print Tab(MGN_L + 161);
    Set Printer.Font = NormalFont
    Printer.Print "���ɒI��"
    '------------------------------------   14�s��
'    Printer.Print
    '------------------------------------   15�s��
    Set Printer.Font = NormalBoldFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print "���וi";
    Printer.Print Tab(MGN_L + 13);
    Set Printer.Font = NormalBoldFont
    Printer.Print StrConv(Trim(StrConv(Y_NYUREC.GENSANKOKU, vbUnicode)), vbWide);           'trim
    Printer.Print Tab(MGN_L + 49);      '53-->49
    Printer.Print StrConv(Space(8 - Len(Format(StrConv(Y_NYUREC.SURYO, vbUnicode), "#,##0"))) & _
                        Format(StrConv(Y_NYUREC.SURYO, vbUnicode), "#,##0"), vbWide);
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 85);
    Printer.Print Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 1, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 3, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 5, 2) & "-" & _
                    Mid(StrConv(Y_NYUREC.NYUKO_TANABAN, vbUnicode), 7, 2);
    Printer.Print Tab(MGN_L + 107);
    Printer.Print "��";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 143);     '114-->143
    Set Printer.Font = NormalFont
    Printer.Print "��";
    Set Printer.Font = SmallFont
    Printer.Print Tab(MGN_L + 153);     '121-->153
    Set Printer.Font = NormalFont
    Printer.Print "��"
    '------------------------------------   16�s��
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 5);
    Printer.Print String(70, "��")
    '------------------------------------   18�s��


    Lcnt = 0

    Head_Print_Proc = False

End Function

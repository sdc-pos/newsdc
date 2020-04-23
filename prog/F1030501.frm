VERSION 5.00
Begin VB.Form F1030501 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�o�ɕ\���(�������)"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2430
   ClientWidth     =   7320
   ControlBox      =   0   'False
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�o�ɕ\�������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
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
Attribute VB_Name = "F1030501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const LMAX% = 46                    '�œ��ő�s��
Private Const MGN_L% = 5                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Dim Pdate As String                         '����J�n���t�iͯ�ް�p�j
Dim Ptime As String                         '����J�n�����iͯ�ް�p�j
'Dim PRT_CAN As Boolean                      '����r���L�����Z���v��


Dim NormalFont As New StdFont               '����t�H���g
Dim Code39Font As New StdFont               '����t�H���g

Private KASO_NYUKA_SOKO As String * 2       '���z�@���בq�ɔԍ�
Private KASO_SYOHN_SOKO As String * 2       '���z�@���i���q�ɔԍ�
Private KASO_NAI_SOKO As String * 2         '���z�@���E�q�ɔԍ�


Private Type Select_Tbl_tag                 '��������p�e�[�u��
    JGYOBU          As String * 1
    MUKE_CODE()     As String * 10
    CYU_KBN         As String * 1
    TITLE           As String
End Type

Dim Select_Tbl()    As Select_Tbl_tag

Dim Yuko_Day        As Integer

Dim Start_YMD       As String * 8
Dim End_YMD         As String * 8

Private Function Y_Syu_Get(com As Integer, Cnt As Integer) As Integer

Dim sts As Integer
Dim OP  As Integer
Dim ans As Integer
Dim i   As Integer

    
    If com = BtOpGetGreaterEqual Then
                                        '�ŏ��̂j�d�x�Z�b�g
        Call UniCode_Conv(K6_Y_SYU.JGYOBU, Select_Tbl(Cnt).JGYOBU)
        
        If Select_Tbl(Cnt).CYU_KBN = "*" Then
            Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, "")
        Else
            Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, Select_Tbl(Cnt).CYU_KBN)
        End If
        Call UniCode_Conv(K6_Y_SYU.HTANABAN, "")
        Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
        Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    End If
    
    OP = com + BtSNoWait
    
    Do
        
        Do
            sts = BTRV(OP, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Select_Tbl(Cnt).JGYOBU Or _
                        (Select_Tbl(Cnt).CYU_KBN <> "*" And _
                        StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Select_Tbl(Cnt).CYU_KBN) Then
                                                        '���ƕ��C�����敪�u���[�N
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                            Y_Syu_Get = sts
                            Exit Function
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    
                    End If
                                                        '�f�[�^��������������H
                    If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_UN And _
                        Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then

                                                        '������S���w��Ȃ�n�j
                            For i = 0 To UBound(Select_Tbl(Cnt).MUKE_CODE)
                                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(Select_Tbl(Cnt).MUKE_CODE(i)) Then
                                    Exit For
                                End If
                            Next i
                            If i > UBound(Select_Tbl(Cnt).MUKE_CODE) Then
                                OP = BtOpGetNext + BtSNoWait
                                Exit Do
                            End If
                                        '�f�[�^�n�j
                            Y_Syu_Get = BtNoErr
                            Exit Function
                        
                        

                    End If

                    OP = BtOpGetNext + BtSNoWait
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                Case BtErrEOF
                    Y_Syu_Get = sts
                    Exit Function
                Case Else
                    Call File_Error(sts, OP + BtSNoWait, "�o�ח\��t�@�C��")
                    Y_Syu_Get = sts
                    Exit Function
            End Select
        Loop
    Loop
End Function

Private Function Print_Proc(Cnt As Integer) As Integer

Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 2
Dim PRI_HIN_GAI     As String * 13
Dim Betu_LOCATION   As String * 8

Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long

Dim RetBuf          As String
    
Dim SUMI_QTY        As Long
Dim MI_QTY        As Long
    
    Print_Proc = True

    
'    PRT_CAN = False
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            '�o�ח\��f�[�^�ǂݍ���
        sts = Y_Syu_Get(com, Cnt)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select

        If Lcnt = 99 Then
            SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
        Else
                                            '�q�ɂ̃u���[�N
            If SAVE_SOKO_No <> Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) Then
                Lcnt = LMAX + 1
                SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
            End If
        End If

        If Lcnt > LMAX Then                 '�w�b�_�[�R���g���[��
            If Head_Proc(Lcnt, Cnt) Then
                Exit Function
            End If
            PRI_HIN_GAI = ""
        End If
                                            
        If StrConv(Y_SYUREC.HIN_NO, vbUnicode) <> PRI_HIN_GAI Then
            PRI_HIN_GAI = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                            '���׈��
            Printer.Print Tab(MGN_L);
                                            '�W���I��
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);

            Printer.Print Tab(MGN_L + 10);
                                            '�i��(�O)
            Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);

            Printer.Print Tab(MGN_L + 24);
                                            '�W���I�@�݌ɐ�
            If Len(Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode))) = 0 Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    StrConv(Y_SYUREC.HTANABAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
            
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;

                                            '�ʒu�I����
            If Tana_Kensaku(Betu_LOCATION) Then
                Print_Proc = True
                Exit Function
            End If
            
            SUMI_QTY = 0
            MI_QTY = 0
            
            If Len(Trim(Betu_LOCATION)) = 0 Then
            Else
                                            '�ʒu�I�@�݌ɐ�
                Printer.Print Tab(MGN_L + 35);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
                
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        Betu_LOCATION) Then
                    Exit Function
                End If
            End If
            
            Printer.Print Tab(MGN_L + 46);
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '���i�������E�݌ɐ�
            Printer.Print Tab(MGN_L + 55);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NAI_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '���בq�ɍ݌�
            Printer.Print Tab(MGN_L + 64);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
                        
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
        End If
                                            '�`�[��
        Printer.Print Tab(MGN_L + 77);
        Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);
                                            '�����溰��
        Printer.Print Tab(MGN_L + 86);
        Printer.Print StrConv(Y_SYUREC.MUKE_CODE, vbUnicode);
                                            '�����於��
        Printer.Print Tab(MGN_L + 95);
        Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
        Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Printer.Print StrConv(MTSREC.MUKE_DNAME, vbUnicode);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                Exit Function
        End Select


        Printer.Print Tab(MGN_L + 105);
        TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
        RetBuf = Format(TEMP_QTY, "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;

        Printer.Print Tab(MGN_L + 115);
                                                '����t�H���g�ݒ�i�b�������R�X�j
        Set Printer.Font = Code39Font
                            '�o�[�R�[�h(*�`�[ID*)
        Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                '����t�H���g�ݒ�i�ʏ�j
        Set Printer.Font = NormalFont
        
        Printer.Print
        Printer.Print
        
        Lcnt = Lcnt + 3



                                                '������t�ݒ�X�V
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            
        Do
        
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Print_Proc = SYS_CANCEL
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                    Print_Proc = SYS_ERR
                    Exit Function
                    
            End Select
        
        
        Loop

        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If




    Print_Proc = False

End Function
                                    
Private Function Head_Proc(Lcnt As Integer, Cnt As Integer) As Integer
Dim i As Integer
Dim sts As Integer

    Head_Proc = True

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    
'    Printer.Print Tab(MGN_L + 20); "�O��� ";
'                                        '����t�H���g�ݒ�
'    Set Printer.Font = Code39Font
'    Printer.Print "*LAST*";
'    Set Printer.Font = NormalFont
    
    Printer.Print Tab(MGN_L + 41);
    
    Printer.Print Select_Tbl(Cnt).TITLE & "�o�ɕ\";
    
    
    Printer.Print Tab(MGN_L + 91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L + 5);
    Printer.Print "�q�ɁF";
    Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2);
    Printer.Print Tab(MGN_L + 15);
    Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
            Exit Function
    End Select
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 23);
    Printer.Print "�W���I�݌�";
    Printer.Print Tab(MGN_L + 35);
    Printer.Print "�ʒu�I��";
    Printer.Print Tab(MGN_L + 47);
    Printer.Print "�ʒu�݌�";
    Printer.Print Tab(MGN_L + 56);
    Printer.Print "���i����";
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "���בq��";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "�o �� ��";
    Printer.Print Tab(MGN_L + 105);
    Printer.Print "�o�א�";
    Printer.Print

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(Y_SYUREC.HIN_NO, vbUnicode) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2) Then
                                                '�V�X�e���q�ɂ̔���
                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '�l�����Ȃ��̂œǂݔ�΂�
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "�q�Ƀ}�X�^")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "�݌Ƀf�[�^")
                Exit Function
        End Select
            
            
    Loop
    
    Tana_Kensaku = False

End Function

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c           As String * 128
Dim i           As Integer
Dim j           As Integer
Dim Get_Data    As String * 10
Dim Work_Date   As String * 8
     
     
     
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

                                
                                '���׉��z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '���i�����z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '���E���z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                '������捞��
    i = -1
    Do
                                '���s�p�����[�^�捞��
        If GetIni(App.EXEName, "JGYO" & Format(i + 2, "00"), "SYS", c) Then
            Beep
            MsgBox "�o�ɕ\�p����p�����[�^�̎捞�݂Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        
        
        
        If Trim(c) = "END" Then
            Exit Do
        End If
        
        
        
        i = i + 1
        ReDim Preserve Select_Tbl(i)
        Select_Tbl(i).JGYOBU = Trim(c)
                                
                                '���s�p�����[�^�捞��
        If GetIni(App.EXEName, "MUKE" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "�o�ɕ\�p����p�����[�^�̎捞�݂Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        
        
        For j = 0 To 99
            Call Data_Select(Trim(c), j + 1, 99, Get_Data)
            If Len(Trim(Get_Data)) = 0 Then
                Exit For
            End If
        
            ReDim Preserve Select_Tbl(i).MUKE_CODE(j)
    
            Select_Tbl(i).MUKE_CODE(j) = Get_Data
        
        Next j
                                
                                
        If GetIni(App.EXEName, "CYU" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "�o�ɕ\�p����p�����[�^�̎捞�݂Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
                                
        Select_Tbl(i).CYU_KBN = Trim(c)
                                
        If GetIni(App.EXEName, "TITLE" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "�o�ɕ\�p����p�����[�^�̎捞�݂Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
                                
        Select_Tbl(i).TITLE = Trim(c)
                                
                                
    Loop
                                
    If i = (-1) Then            '����w���Ȃ�
        End
    End If
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '��ƊǗ��}�X�^�n�o�d�m
    If SAGYO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��t�@�C���n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1030501.FontName
        .Size = 10
    End With
                                '����t�H���g�ݒ�i�o�[�R�[�h�j
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
                                
        
    If GetIni(App.EXEName, "YUKO_DAY", "SYS", c) Then
        Yuko_Day = 0
    Else
        Yuko_Day = CInt(Trim(c))
    End If


    
    If Yuko_Day = 0 Then
        Start_YMD = Format(Now, "YYYYMMDD")
        End_YMD = Format(Now, "YYYYMMDD")

    Else

        Start_YMD = Format(DateAdd("d", Yuko_Day, Date), "YYYYMMDD")
        End_YMD = "99991231"

    End If
    

    For i = 0 To UBound(Select_Tbl)
        If Print_Proc(i) Then
            Unload Me
        End If
    Next i

    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
'    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
'        End If
'    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '��ƊǗ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SAGYO_POS, SAGYOREC, Len(SAGYOREC), K0_SAGYO, Len(K0_SAGYO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "��ƊǗ��}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030501 = Nothing

    End
End Sub


VERSION 5.00
Begin VB.Form F1090101 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�݌ɏW�v����[�e109010]�@2019.01.08 13:30"
   ClientHeight    =   4740
   ClientLeft      =   2310
   ClientTop       =   2610
   ClientWidth     =   7350
   ControlBox      =   0   'False
   Enabled         =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7350
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���s���I"
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
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�݌ɍ��ك`�F�b�N�p�݌ɏW�v����"
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
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7200
   End
End
Attribute VB_Name = "F1090101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SOKO_TBL()      As String * 2
Dim NON_FLG         As Boolean

Dim ZENKAI_YMD      As String

Private Function SumZ_Update() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�쐬����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Upd_Com     As Integer
    
Dim ans         As Integer
    
''Dim Save_Hin    As String * 15
Dim Save_Hin    As String * 22



Dim Zaiko_Qty   As Long
Dim Syuka_Qty   As Long
    
Dim i           As Integer
Dim SKIP_FLG    As Boolean
    
    SumZ_Update = True
    
    Zaiko_Qty = 0
    
    com = BtOpGetFirst
    Do
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Exit Function
        End Select
        
If Left(StrConv(ZAIKOREC.HIN_GAI, vbUnicode), 14) = "1250201-125040" Then
Debug.Print
End If
        
        
        SKIP_FLG = False
    
        If Not NON_FLG Then
        Else
            '���O�q�Ɍ���
            For i = 0 To UBound(SOKO_TBL)
                If SOKO_TBL(i) = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) Then
                    SKIP_FLG = True
                    Exit For
                End If
            Next i
        End If
        If Not SKIP_FLG Then
            '���O�I�Ԍ���
            Call UniCode_Conv(K0_TANA.SOKO_NO, StrConv(ZAIKOREC.SOKO_NO, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, StrConv(ZAIKOREC.Retu, vbUnicode))
            Call UniCode_Conv(K0_TANA.Ren, StrConv(ZAIKOREC.Ren, vbUnicode))
            Call UniCode_Conv(K0_TANA.Dan, StrConv(ZAIKOREC.Dan, vbUnicode))
    
            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(TANAREC.ZAIKO_SHOGO_FLG, vbUnicode) = ZAIKO_SHOGO_FLG_NG Then
                        SKIP_FLG = True
                    End If
                Case BtErrKeyNotFound
                    '�ُ킾����ΏۊO
                    SKIP_FLG = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                    Exit Function
            End Select
        End If
    
    
        If SKIP_FLG Then
        Else
            If com = BtOpGetFirst Then
                Save_Hin = StrConv(ZAIKOREC.JGYOBU, vbUnicode) & _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) & _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            End If
        
            If Save_Hin <> (StrConv(ZAIKOREC.JGYOBU, vbUnicode) & _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) & _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                                        '�݌ɏW�v�f�[�^�ǂݍ���
                Call UniCode_Conv(K0_SUMZ.JGYOBU, Left(Save_Hin, 1))
                Call UniCode_Conv(K0_SUMZ.NAIGAI, Mid(Save_Hin, 2, 1))
''                Call UniCode_Conv(K0_SUMZ.HIN_GAI, Right(Save_Hin, 13))
                Call UniCode_Conv(K0_SUMZ.HIN_GAI, Right(Save_Hin, 20))
                    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_Com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_Com = BtOpInsert
                            Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                       Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                            Exit Function
                    End Select
                Loop
                    
                If Upd_Com = BtOpInsert Then
                    Call UniCode_Conv(SUMZREC.JGYOBU, Left(Save_Hin, 1))    '���ƕ�
                    Call UniCode_Conv(SUMZREC.NAIGAI, Mid(Save_Hin, 2, 1))  '�����O
''                    Call UniCode_Conv(SUMZREC.HIN_GAI, Right(Save_Hin, 13)) '�i�ԁi�O���j
                    Call UniCode_Conv(SUMZREC.HIN_GAI, Right(Save_Hin, 20)) '�i�ԁi�O���j
                                                                                        
                
                
                    Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")            '�݌ɑ���
                    Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, "00000000")          '�O�����݌ɑ���
                    Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")            '�o�ɍςݐ�
                    Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")           '���ח\�萔
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")            'νč݌ɐ�
                    Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, "00000000")        '�O��νč݌ɐ�
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")              '���ِ�
                    Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v��
                    
                    
                    '2009.02.09
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SAI_YMD, "")
                    
                    
                    Call UniCode_Conv(SUMZREC.FILLER, "")
                End If
                
                
                
                Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v��
                
                
'2019.01.08                Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + Zaiko_Qty, "00000000"))
                Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(Val(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + Zaiko_Qty, "00000000"))
                                        '�݌ɏW�v�f�[�^��������
                Do
                    sts = BTRV(Upd_Com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                       Case Else
                            Call File_Error(sts, Upd_Com, "�݌ɏW�v�f�[�^")
                            Exit Function
                    End Select
                Loop
                
                Save_Hin = StrConv(ZAIKOREC.JGYOBU, vbUnicode) & _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) & _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                Zaiko_Qty = 0
                Syuka_Qty = 0
            End If
        
        
'2019.01.08            Zaiko_Qty = Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Zaiko_Qty = Zaiko_Qty + Val(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        
        End If
        
        com = BtOpGetNext
    Loop


    If Zaiko_Qty <> 0 Then
                                        
                                        '�݌ɏW�v�f�[�^�ǂݍ���
        Call UniCode_Conv(K0_SUMZ.JGYOBU, Left(Save_Hin, 1))
        Call UniCode_Conv(K0_SUMZ.NAIGAI, Mid(Save_Hin, 2, 1))
''        Call UniCode_Conv(K0_SUMZ.HIN_GAI, Right(Save_Hin, 13))
        Call UniCode_Conv(K0_SUMZ.HIN_GAI, Right(Save_Hin, 20))
                    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Upd_Com = BtOpUpdate
                    Exit Do
                Case BtErrKeyNotFound
                    Upd_Com = BtOpInsert
                    Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^")
                    Exit Function
            End Select
        Loop
                    
        If Upd_Com = BtOpInsert Then
            Call UniCode_Conv(SUMZREC.JGYOBU, Left(Save_Hin, 1))    '���ƕ�
            Call UniCode_Conv(SUMZREC.NAIGAI, Mid(Save_Hin, 2, 1))  '�����O
''            Call UniCode_Conv(SUMZREC.HIN_GAI, Right(Save_Hin, 13)) '�i�ԁi�O���j
            Call UniCode_Conv(SUMZREC.HIN_GAI, Right(Save_Hin, 20)) '�i�ԁi�O���j
                                                                                        
                
            Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")            '�݌ɑ���
            Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, "00000000")          '�O�����݌ɑ���
            Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")            '�o�ɍςݐ�
            Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")           '���ח\�萔
            Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")            'νč݌ɐ�
            Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, "00000000")        '�O��νč݌ɐ�
            Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")              '���ِ�
'''2007.05.17            Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v��
            Call UniCode_Conv(SUMZREC.FILLER, "")
        End If
            
        Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v�� 2007.05.17
        
'2019.01.08        Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + Zaiko_Qty, "00000000"))
        Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(Val(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + Zaiko_Qty, "00000000"))
                                        '�݌ɏW�v�f�[�^��������
        Do
            sts = BTRV(Upd_Com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, Upd_Com, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
    End If
                                        '�W���I�ԃZ�b�g
                
                
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
                        '���̏����ł͖{�����肦�Ȃ��I�I
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌ɏW�v�f�[�^")
                Exit Function
        End Select
    
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SUMZREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SUMZREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SUMZREC.HIN_GAI, vbUnicode))
                                                            '�W�����ɒI
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
            Case BtErrKeyNotFound
                Call UniCode_Conv(SUMZREC.ST_SOKO, "")
                Call UniCode_Conv(SUMZREC.ST_RETU, "")
                Call UniCode_Conv(SUMZREC.ST_REN, "")
                Call UniCode_Conv(SUMZREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
    
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
    
    
    
        com = BtOpGetNext
    Loop
                
                
                





    SumZ_Update = False

End Function
Private Function SumZ_Init() As Integer
'----------------------------------------------------------------------------
'                   �u�݌ɏW�v�f�[�^�v�N���A�[����
'----------------------------------------------------------------------------

Dim sts As Integer
Dim com As Integer
    
Dim ans As Integer
    
    SumZ_Init = True
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com + BtSNoWait, "�݌ɏW�v�f�[�^")
                Exit Function
        End Select
        
        
        
If Trim(StrConv(SUMZREC.HIN_GAI, vbUnicode)) = "1250201-12504" Then
Debug.Print
End If
        
        If ZENKAI_YMD <> Format(Now, "YYYY/MM/DD") Then
            Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, StrConv(SUMZREC.T_Zai_Qty, vbUnicode))
        
        
            '���񍷈ف��O�񍷈� 2009.02.09
            Call UniCode_Conv(SUMZREC.ZEN_SAI_QTY, StrConv(SUMZREC.SAI_QTY, vbUnicode))
        
        
        
        End If
        
        Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")
        Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")
        Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")
        
        '2009.02.09
        Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
        If Not IsNumeric(StrConv(SUMZREC.SAI_YMD, vbUnicode)) Then
            Call UniCode_Conv(SUMZREC.SAI_YMD, "")
        End If
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    SumZ_Init = False

End Function
Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                '�O�񏈗�����荞��
    If GetIni(App.EXEName, "ZENKAI_YMD", "SYS", c) Then
        ZENKAI_YMD = ""
    Else
        ZENKAI_YMD = RTrim(c)
    End If
                                
                                '���O�q�Ɏ�荞��
    
    i = 0
    Do
        
        If GetIni(App.EXEName, "SOKO" & Format(i + 1, "0"), "SYS", c) Then
            Exit Do
        End If
    
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve SOKO_TBL(i)
        SOKO_TBL(i) = Trim(c)
        i = i + 1
    
    Loop
                                
    If i = 0 Then
        NON_FLG = False         '���O�q�ɂȂ�
    Else
        NON_FLG = True          '���O�q�ɂ���
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
    
    Call Main_Proc

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
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
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
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�݌ɏW�v�f�[�^")
    End If

    Set F1090101 = Nothing

    End
End Sub

Private Sub Main_Proc()
'----------------------------------------------------------------------------
'                   ���C������
'----------------------------------------------------------------------------
    
    
    DoEvents
    
    If SumZ_Init() Then             '�݌ɏW�v�f�[�^�N���A�[
        Unload Me
    End If
    
    
    If SumZ_Update() Then           '�݌ɏW�v
        Unload Me
    End If
    
    If Syuka_Modishi() Then         '��s�o�ו��̖߂�
        Unload Me
    End If
                                    '�h�m�h�������t�o��
    If WriteIni(App.EXEName, "ZENKAI_YMD", "SYS", Format(Now, "YYYY/MM/DD")) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & " ZENKAI_YMD")
        Unload Me
    End If
    
    
    Unload Me

End Sub
Private Function Syuka_Modishi() As Integer
'----------------------------------------------------------------------------
'                   ��s�o�ו��߂�����
'----------------------------------------------------------------------------

Dim com     As Integer
Dim sts     As Integer
Dim ans     As Integer

Dim Upd_Com As Integer
    
    Syuka_Modishi = True
    
    com = BtOpGetFirst

    Do
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��")
                Exit Function
        End Select
    
    
    
'2019.01.08        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) >= Format(Now, "YYYYMMDD") And _
'2019.01.08            CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
            
        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) >= Format(Now, "YYYYMMDD") And _
            Val(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
            
            
                                        '�݌ɏW�v�f�[�^�ǂݍ���
            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                    
            Do
                DoEvents
                sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_Com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_Com = BtOpInsert
                        Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                   Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌ɏW�v�f�[�^")
                        Exit Function
                End Select
            Loop
                    
            If Upd_Com = BtOpInsert Then
                                                                        '���ƕ�
                Call UniCode_Conv(SUMZREC.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                                                                        '�����O
                Call UniCode_Conv(SUMZREC.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
                                                                        '�i�ԁi�O���j
                Call UniCode_Conv(SUMZREC.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                                                                                        
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
                                                                        '�W�����ɒI
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                        Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SUMZREC.ST_SOKO, "")
                        Call UniCode_Conv(SUMZREC.ST_RETU, "")
                        Call UniCode_Conv(SUMZREC.ST_REN, "")
                        Call UniCode_Conv(SUMZREC.ST_DAN, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
                
                
                Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")            '�݌ɑ���
                Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")            '�o�ɍςݐ�
                Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")           '���ח\�萔
                Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")            'νč݌ɐ�
                Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")              '���ِ�
'''2007.05.17                Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v��
                
                
                
                
                
                
                
                
                '2009.02.09
                Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                Call UniCode_Conv(SUMZREC.SAI_YMD, "")
                
                
                
                
                
                Call UniCode_Conv(SUMZREC.FILLER, "")
            End If
                                                                            
            Call UniCode_Conv(SUMZREC.SUM_DT, Format(Now, "yyyymmdd"))  '�W�v�� 2007.05.17
                                                                            
                                                                            '��s�o�ו����Z
'2019.01.08            Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
'2019.01.08            Call UniCode_Conv(SUMZREC.SYK_E_QTY, Format(CLng(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)) + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
                                        
                                        
            Call UniCode_Conv(SUMZREC.T_Zai_Qty, Format(Val(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) + Val(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
            Call UniCode_Conv(SUMZREC.SYK_E_QTY, Format(Val(StrConv(SUMZREC.SYK_E_QTY, vbUnicode)) + Val(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
                                        
                                        
                                        '�݌ɏW�v�f�[�^��������
            Do
                sts = BTRV(Upd_Com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                            '���̏����ł͖{�����肦�Ȃ��I�I
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                   Case Else
                        Call File_Error(sts, Upd_Com, "�݌ɏW�v�f�[�^")
                        Exit Function
                End Select
            Loop
        End If

        com = BtOpGetNext
    
    Loop


    Syuka_Modishi = False

End Function


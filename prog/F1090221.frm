VERSION 5.00
Begin VB.Form F1090221 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�݌ɍ��ك`�F�b�N�����i�܈�p�j"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
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
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�z�X�g�f�[�^�W�v��"
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
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�f�[�^��������"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�݌ɍ��ك`�F�b�N����"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   4800
   End
End
Attribute VB_Name = "F1090221"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type YUKO_SOKO_TBL             '�L��νđq�Ɏ�荞�݃e�[�u��
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_T() As YUKO_SOKO_TBL

Dim HS_NaiG As String                   '�����O�i������e�j�����@ν��ް����e�ɂ��ݒ�
Private ZENKAI_YMD      As String       '�O�񏈗��N����
Private ZENZENKAI_YMD   As String       '�O�X�񏈗��N����
Private Function SumZ_Init() As Integer

Dim sts As Integer
Dim com As Integer
Dim ans As Integer

    SumZ_Init = True
    
    com = BtOpGetFirst

    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(SUMZREC.T_Zai_Qty, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
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
        
        DoEvents
    Loop

    SumZ_Init = False

End Function
Private Function SumZ_Update() As Integer

Dim i           As Integer
Dim sts         As Integer
Dim Upd_com     As Integer
Dim T_Zai_Qty   As Long
Dim SAI_QTY     As Long
        
Dim ans         As Integer
        
    SumZ_Update = True
    
    Do
        DoEvents
        sts = HS_ZAI_Get1           'ν��ް� �Ǎ���
        If sts Then
                        
            Exit Do
        End If
        If StrConv(HS_ZAIREC.JGYOBU, vbUnicode) < " " Then
            Exit Do
        End If
                                    '�����O�敪�̐ݒ�
        For i = 0 To UBound(SOKO_T)
            If SOKO_T(i).HS_SOKO = "  " Then
                Exit For
            End If
            If RTrim(StrConv(HS_ZAIREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T(i).HS_SOKO) Then
                HS_NaiG = SOKO_T(i).NAIGAI
                Exit For
            End If
        Next i
                                    
        If Len(Trim(StrConv(HS_ZAIREC.HIN_NAI, vbUnicode))) = 0 Then
        Else
            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(HS_ZAIREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(HS_NaiG, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(HS_ZAIREC.HIN_GAI, vbUnicode))
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^")
                        Exit Function
                End Select
            Loop

            If CLng(StrConv(HS_ZAIREC.ZEN_Z_QTY, vbUnicode)) = 0 Or _
                StrConv(HS_ZAIREC.QTY_SIGN, vbUnicode) <> " " Then
            Else
                If Upd_com = BtOpInsert Then
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(HS_ZAIREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(HS_NaiG, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_ZAIREC.HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Call UniCode_Conv(ITEMREC.ST_RETU, "")
                            Call UniCode_Conv(ITEMREC.ST_REN, "")
                            Call UniCode_Conv(ITEMREC.ST_DAN, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                    
                    
                    
                    Call UniCode_Conv(SUMZREC.JGYOBU, StrConv(HS_ZAIREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(SUMZREC.NAIGAI, StrConv(HS_NaiG, vbUnicode))
                    Call UniCode_Conv(SUMZREC.HIN_GAI, StrConv(HS_ZAIREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")
                    Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))
                    Call UniCode_Conv(SUMZREC.FILLER, "")
                End If
                
                Call UniCode_Conv(SUMZREC.HS_ZAIQTY, Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)) + CLng(StrConv(HS_ZAIREC.ZEN_Z_QTY, vbUnicode)), "00000000"))
            
                SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                If SAI_QTY < 0 Then
                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
                Else
                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
                End If
                
                Do
                    sts = BTRV(Upd_com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '����͂Ȃ�
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌ɏW�v�f�[�^")
                            Exit Function
                    End Select
                Loop
            End If
        End If                                              '1999.09.09
        DoEvents
    Loop
End Function
Private Function HS_ZAI_INIT() As Integer
Dim Work As String
    
    HS_ZAI_INIT = True
    
    
    If HS_ZAI_Open1(0, Work) = False Then   '̧�ٖ����Ȃ珈�����Ȃ�
            
        Close #HS_ZAI_No

        If HS_ZAI_Open1(1, Work) Then       '����@���ƕ�ν��ް� OPEN
            Exit Function
        End If

        If SumZ_Update() Then               '����@���ƕ�  �݌ɐݒ�f�[�^�捞��
            Exit Function
        End If

        Close #HS_ZAI_No                    'ν��ް� CLOSE
    End If

End Function

Private Sub Form_Activate()

Dim ans As Integer

    
    If ZENKAI_YMD = Format(Now, "YYYY/MM/DD") Then
        ans = MsgBox("�{���̍݌Ɏ�荞�ݏ����͏I�����Ă��܂��B���s���܂����H", vbYesNo, "�m�F����")
    Else
        ans = MsgBox("�u�݌ɍ��ك`�F�b�N�����v���s���܂����H", vbYesNo, "�m�F����")
    End If

    If ans = vbNo Then
        Unload Me
    End If
    If ans = vbNo Then
        Unload Me
    End If

        
    F1090221.MousePointer = vbHourglass
    Label1(1).Visible = True
    
    If SumZ_Init() Then             '���كf�[�^������
        Unload Me
    End If
    
    Label1(1).Visible = False
    Label1(2).Visible = True
    
    If HS_ZAI_INIT() Then           '���ك`�F�b�N����
        Unload Me
    End If

    If WriteIni(App.EXEName, "ZENZENKAI_YMD", "SYS", ZENKAI_YMD) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & " ZENZENKAI_YMD")
        Unload Me
    End If

    If WriteIni(App.EXEName, "ZENKAI_YMD", "SYS", Format(Now, "YYYY/MM/DD")) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & " ZENKAI_YMD")
        Unload Me
    End If


    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim c   As String * 128
Dim sts As Integer
Dim i   As Integer
Dim j   As Integer
    
Dim Max_Soko    As Integer
    
    
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
    If JGYOB_TB_Set() Then      '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                '�q�ɍő吔����荞��
                                
    If GetIni(App.EXEName, "MAX_SOKO", "SYS", c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                '�݌Ɏ�荞�ݗp�e�[�u���쐬
    ReDim SOKO_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '�q�ɏ���荞��
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '�L���q�Ɋl��
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "�q�ɏ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
    
            If Trim(c) = "**" Then  '�q�Ɏw��I��
                Exit Do
            End If
    
    
'            ReDim Preserve SOKO_T(0 To i, 0 To j)
            SOKO_T(i, j).HS_SOKO = Trim(c)
                                '�����O���l��
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), "SYS", c) Then
                Beep
                MsgBox "�����O���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
            
            SOKO_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                '�O�񏈗����̊l��
    If GetIni(App.EXEName, "ZENZENKAI_YMD", "SYS", c) Then
        ZENZENKAI_YMD = ""
    Else
        ZENZENKAI_YMD = RTrim(c)
    End If
                                
                                '�O�X�񏈗����̊l��
    If GetIni(App.EXEName, "ZENKAI_YMD", "SYS", c) Then
        ZENKAI_YMD = ""
    Else
        ZENKAI_YMD = RTrim(c)
    End If

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
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
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1090221 = Nothing

    End
End Sub



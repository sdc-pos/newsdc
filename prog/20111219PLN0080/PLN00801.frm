VERSION 5.00
Begin VB.Form PLN00801 
   Caption         =   "[PLN0080]���i���\��o�ɕ\���s"
   ClientHeight    =   5685
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   9975
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
   OLEDropMode     =   1  '�蓮
   ScaleHeight     =   5685
   ScaleWidth      =   9975
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   3
      Left            =   6000
      TabIndex        =   11
      Top             =   3120
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   2
      Left            =   5160
      TabIndex        =   9
      Top             =   3120
      Width           =   372
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   2400
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1332
   End
   Begin VB.ComboBox Combo1 
      Height          =   336
      Index           =   0
      ItemData        =   "PLN00801.frx":0000
      Left            =   4200
      List            =   "PLN00801.frx":000A
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "�������I�����܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "������������s���܂�"
      Top             =   0
      Width           =   1380
   End
   Begin VB.Label lblBC 
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8160
      TabIndex        =   12
      Top             =   5040
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "�`"
      Height          =   252
      Index           =   4
      Left            =   5640
      TabIndex        =   10
      Top             =   3240
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "�W���I��"
      Height          =   252
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "�`"
      Height          =   252
      Index           =   0
      Left            =   5640
      TabIndex        =   6
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "���i���\���"
      Height          =   252
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "�a�t"
      Height          =   252
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   492
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "�����I��"
      Begin VB.Menu SHORI 
         Caption         =   "���"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "�I��"
         Index           =   1
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "PLN00801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbBU% = 0


Private Const ptxYOTEI_DT_S% = 0
Private Const ptxYOTEI_DT_E% = 1
Private Const ptxST_SOKO_S% = 2
Private Const ptxST_SOKO_E% = 3



Dim NormalFont As New StdFont               '����t�H���g
Dim Code39Font As New StdFont               '����t�H���g

Private KASO_NYUKA_SOKO As String * 2       '���z�@���בq�ɔԍ�
Private KASO_SYOHN_SOKO As String * 2       '���z�@���i���q�ɔԍ�
Private KASO_NAI_SOKO As String * 2         '���z�@���E�q�ɔԍ�


Private Const LMAX% = 56                    '�œ��ő�s��
Private Const MGN_L% = 2                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Dim Pdate As String                         '����J�n���t�iͯ�ް�p�j
Dim Ptime As String                         '����J�n�����iͯ�ް�p�j


Private Const LAST_UPDATE_DAY$ = " 2011.12.19 14:00"


Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '�������


            If Trim(Text1(ptxYOTEI_DT_S).Text) = "" Then
                Text1(ptxYOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxYOTEI_DT_S).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�ē��͂��ĉ������i���i���\����@�J�n�j"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If


            If Trim(Text1(ptxYOTEI_DT_E).Text) = "" Then
                Text1(ptxYOTEI_DT_E).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxYOTEI_DT_E).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�ē��͂��ĉ������i���i���\����@�I���j"
                Text1(ptxYOTEI_DT_E).SetFocus
                Exit Sub
            End If


            If Text1(ptxYOTEI_DT_S).Text > Text1(ptxYOTEI_DT_E).Text Then
                MsgBox "���͂������ڂ̓G���[�ł��B�ē��͂��ĉ������i���i���\����j"
                Text1(ptxYOTEI_DT_S).SetFocus
                Exit Sub
            End If


            If Trim(Text1(ptxST_SOKO_E).Text) = "" Then
                Text1(ptxST_SOKO_E).Text = "zz"
            End If

            If Text1(ptxST_SOKO_S).Text > Text1(ptxST_SOKO_E).Text Then
                MsgBox "���͂������ڂ̓G���[�ł��B�ē��͂��ĉ������i�W���I�ԁj"
                Text1(ptxST_SOKO_S).SetFocus
                Exit Sub
            End If


            If Print_Proc() Then
                Unload Me
            End If

        Case 1          '�I��

            Unload Me
    End Select



    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[���i���v��V�X�e��]���i�����X�g���", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)

                                
                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                




    PLN00801.Caption = PLN00801.Caption & " " & LAST_UPDATE_DAY


    Call Bu_Set_Proc
                                
                                
                                
                                '���׉��z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", App.EXEName, c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '���i�����z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", App.EXEName, c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '���E���z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NAI_SOKO", App.EXEName, c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                
                                
                                '���i���p���ח\��t�@�C���n�o�d�m
    If PLN_S_YOTEI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '���i���w�}�f�[�^�i�e�j�n�o�d�m
    If P_SSHIJI_O_Open(BtOpenRead) Then
        Unload Me
    End If


                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = PLN00801.FontName
        .Size = 10
    End With
                                '����t�H���g�ݒ�i�o�[�R�[�h�j
    With Code39Font
        .NAME = lblBC.FontName
        .Size = lblBC.FontSize
    End With

    Text1(ptxYOTEI_DT_S).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxYOTEI_DT_E).Text = Format(Now, "YYYY/MM/DD")


End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���\��t�@�C��")
        End If
    End If
    
    
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K0_PLN_S_YOTEI, Len(K0_PLN_S_YOTEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set PLN00801 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
    
    End Select



End Sub


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u���i�����X�g�v���׈������
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    

    
Dim Skip_FLg        As Boolean
    
    
Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 2
Dim Betu_LOCATION   As String * 8


Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long
Dim RetBuf          As String
    
Dim i               As Integer

Dim c               As String * 128
Dim SHIMUKE_CODE    As String * 2

    Print_Proc = True
    
    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���\��o�ɕ\���s�@�����J�n�I�I", Me.hwnd, 0)
    
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    
    
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_SOKO, Text1(ptxST_SOKO_S).Text)
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_RETU, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_REN, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.ST_DAN, "")
    
    Call UniCode_Conv(K5_PLN_S_YOTEI.JGYOBU, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.NAIGAI, "")
    Call UniCode_Conv(K5_PLN_S_YOTEI.HIN_GAI, "")
    
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, PLN_S_YOTEI_POS, PLN_S_YOTEI_R, Len(PLN_S_YOTEI_R), K5_PLN_S_YOTEI, Len(K5_PLN_S_YOTEI), 5)
    
        Select Case sts
            Case BtNoErr
                
                
                
                If StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) > Text1(ptxST_SOKO_E).Text Then
                    Exit Do
                End If
            
            
                Skip_FLg = False
                If Trim(Right(Combo1(pcmbBU).Text, 1)) <> "" Then
                    If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) <> Right(Combo1(pcmbBU).Text, 1) Then
                        Skip_FLg = True
                    End If
                End If
            
            
                If StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) < Format(Text1(ptxYOTEI_DT_S).Text, "YYYYMMDD") Or _
                    StrConv(PLN_S_YOTEI_R.YOTEI_DT, vbUnicode) > Format(Text1(ptxYOTEI_DT_E).Text, "YYYYMMDD") Then
                    Skip_FLg = True
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "���i���\��f�[�^")
                Exit Function
        End Select
            
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.15 �w�}�[���s���t�������i���[�j���t�̊l��
        If GetIni(App.EXEName, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), App.EXEName, c) Then
            SHIMUKE_CODE = ""
        Else
            SHIMUKE_CODE = Trim(c)
        End If
        
        
        
        
        Call UniCode_Conv(K4_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
        Call UniCode_Conv(K4_P_SSHIJI_O.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K4_P_SSHIJI_O.Print_datetime, StrConv(PLN_S_YOTEI_R.Ins_DateTime, vbUnicode))
        
        com = BtOpGetGreaterEqual
        
        Call UniCode_Conv(PLN_S_YOTEI_R.SASIZU_DateTime, "")
        Call UniCode_Conv(PLN_S_YOTEI_R.S_KAN_DateTime, "")
        
        
        Do
            DoEvents
            sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K4_P_SSHIJI_O, Len(K4_P_SSHIJI_O), 4)
            Select Case sts
                Case BtNoErr
                        
                    If StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Then
                        Exit Do
                    End If
                        
                    If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                        StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                        Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                        
                        
                    If StrConv(P_SSHIJI_O_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
                    Else
                        Skip_FLg = True
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "���i���w�}�f�[�^�i�e�j")
                    Exit Function
            End Select
        
            com = BtOpGetNext
        
        Loop
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>2011.11.15 �w�}�[���s���t�������i���[�j���t�̊l��
        If Skip_FLg Then
        Else
            
            
 '           If Lcnt = 99 Then
 '               SAVE_SOKO_No = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)
 '           Else
 '                                               '�q�ɂ̃u���[�N
 '               If SAVE_SOKO_No <> StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) Then
 '                   Lcnt = LMAX + 1
 '                   SAVE_SOKO_No = StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)
 '               End If
 '           End If
        
        
        
            If Lcnt > LMAX Then                 '�w�b�_�[�R���g���[��
                If Head_Proc(Lcnt, SAVE_SOKO_No) Then
                    Exit Function
                End If
            End If
        
        
            Printer.Print Tab(MGN_L);
                                            '�W���I��
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & "-";
            Printer.Print StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode);
                                            'BU
            Printer.Print Tab(MGN_L + 13);
            For i = 0 To UBound(JGYOBU_T)
                If StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
                    Printer.Print Left(JGYOBU_T(i).NAME, 10);
                End If
            Next i
                                            '�i��
            Printer.Print Tab(MGN_L + 28);
            Printer.Print Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode));
                                            '�W���I�@�݌ɐ�
            If Trim(StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode)) = "" Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                        StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) & _
                                        StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
            Printer.Print Tab(MGN_L + 43);
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
                                            '�ʒu�I�@�݌ɐ�
            If Trim(Betu_LOCATION) <> "" Then
                Printer.Print Tab(MGN_L + 54);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
            
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    Betu_LOCATION) Then
                    Exit Function
                End If
                Printer.Print Tab(MGN_L + 65);
                ZAIKO_QTY = SUMI_QTY + MI_QTY
                RetBuf = Format(ZAIKO_QTY, "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
            End If
                                            '���i���������E�@�݌ɐ�
            Printer.Print Tab(MGN_L + 80);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
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
            Printer.Print Tab(MGN_L + 90);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode), _
                                    StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '�\�萔
            Printer.Print Tab(MGN_L + 100);
            RetBuf = Format(Val(StrConv(PLN_S_YOTEI_R.YOTEI_QTY, vbUnicode)), "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
            Printer.Print Tab(MGN_L + 113);
                                
                                '����t�H���g�ݒ�iCODE39�j
            Set Printer.Font = Code39Font
                                '�o�[�R�[�h(*�i��*)
            Printer.Print "*" & Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) & "*";
                                '����t�H���g�ݒ�i�ʏ�j
            Set Printer.Font = NormalFont
        
        
        
            Printer.Print
            Printer.Print
        
        
            Lcnt = Lcnt + 2
        
        End If
            
            
            
    
        com = BtOpGetNext
    
    
    Loop


    Printer.EndDoc


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i���\��o�ɕ\���s�@�����I���I�I", Me.hwnd, 0)





    Print_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function


Private Function Head_Proc(Lcnt As Integer, SAVE_SOKO_No As String) As Integer
'----------------------------------------------------------------------------
'                   �u���i�����X�g�v�w�b�_�[�������
'----------------------------------------------------------------------------
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
    'Printer.Print Tab(3);                  '97.10.14
    
    Printer.Print Tab(MGN_L + 41);
    
    Printer.Print "�w���i���\��o�ɕ\�x     �@�@���i���\����F";
    Printer.Print Text1(ptxYOTEI_DT_S).Text & "�`" & Text1(ptxYOTEI_DT_E).Text;
    
    Printer.Print Tab(MGN_L + 110);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

'    Printer.Print Tab(MGN_L + 5);
'    Printer.Print "�q�ɁF";
'    Printer.Print SAVE_SOKO_No;
'    Printer.Print Tab(MGN_L + 15);
'    Call UniCode_Conv(K0_SOKO.Soko_No, SAVE_SOKO_No)
'    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'    Select Case sts
'        Case BtNoErr
'            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
'        Case BtErrKeyNotFound
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
'            Exit Function
'    End Select
'
'    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "�a�t";
    Printer.Print Tab(MGN_L + 28);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 42);
    Printer.Print "�W���I�݌�";
    Printer.Print Tab(MGN_L + 54);
    Printer.Print "�ʒu�I��";
    Printer.Print Tab(MGN_L + 81);
    Printer.Print "���i����";
    Printer.Print Tab(MGN_L + 91);
    Printer.Print "���בq��";
    Printer.Print Tab(MGN_L + 101);
    Printer.Print "�@�\�萔";
    Printer.Print

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function

Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer
'----------------------------------------------------------------------------
'                   �ʒu���I�Ԍ���
'----------------------------------------------------------------------------

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(PLN_S_YOTEI_R.JGYOBU, vbUnicode) Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(PLN_S_YOTEI_R.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(PLN_S_YOTEI_R.HIN_GAI, vbUnicode)) Then
                    
                    Exit Do
                
                End If
                
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_SOKO, vbUnicode) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_RETU, vbUnicode) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_REN, vbUnicode) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(PLN_S_YOTEI_R.ST_DAN, vbUnicode) Then
                                                
                                                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
'                                                '�V�X�e���q�ɂ̔���
'                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
'                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
'                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
'                                                StrConv(ZAIKOREC.Dan, vbUnicode)
'                                Exit Do
'
'                            End If
'                        Case BtErrKeyNotFound
'                                                '�l�����Ȃ��̂œǂݔ�΂�
'                        Case Else
'                            Call File_Error(sts, BtOpGetGreater, "�q�Ƀ}�X�^")
'                            Exit Function
'                    End Select
                                        
'                                                '�V�X�e���q�ɂ�L���Ƃ���
                    If KASO_NYUKA_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
                        KASO_SYOHN_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
                        KASO_NAI_SOKO <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Then
                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do

                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "�݌Ƀf�[�^")
                Exit Function
        End Select
'
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2011.12.17
            
    Loop
    
    Tana_Kensaku = False

End Function



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00801.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00801)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00801)


    PLN00801.MousePointer = vbDefault

End Sub



Private Sub Bu_Set_Proc()
'----------------------------------------------------------------------------
'                   ��ʍ��ځi�a�t�j�̃Z�b�g
'----------------------------------------------------------------------------
Dim i   As Integer




    Combo1(pcmbBU).Clear


        Combo1(pcmbBU).AddItem "�S��" & "          " & " "
        



    For i = 0 To UBound(JGYOBU_T)
            
        Combo1(pcmbBU).AddItem JGYOBU_T(i).NAME & "          " & JGYOBU_T(i).CODE
            
            
    Next i

    Combo1(pcmbBU).ListIndex = 0
End Sub



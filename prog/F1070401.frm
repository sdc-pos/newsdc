VERSION 5.00
Begin VB.Form F1070401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���ރZ���^�[�z�I���\���([F107040] 2013.03.25 08:00)"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2250
   ClientWidth     =   11295
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
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   8040
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   7
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   6
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   5
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   4
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������f"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ް�"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   31
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   30
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   29
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "�`"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   28
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   27
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   26
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   25
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "�I�Ԕ͈�"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   24
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�Ώ۔N��"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   23
      Top             =   2160
      Width           =   975
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
      TabIndex        =   22
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1070401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxYM% = 0                            '�Ώ۔N��


Private Const ptxS_SOKO_No% = 1                     '�J�n   �q�ɇ�
Private Const ptxS_Retu% = 2                        '�@     ��
Private Const ptxS_Ren% = 3                         '�@     �A
Private Const ptxS_Dan% = 4                         '�@     �i
Private Const ptxE_SOKO_No% = 5                     '�J�n   �q�ɇ�
Private Const ptxE_Retu% = 6                        '�@     ��
Private Const ptxE_Ren% = 7                         '�@     �A
Private Const ptxE_Dan% = 8                         '�@     �i

Private Const Text_Max% = 8                         '��ʍ��ڕʍő���ޯ��


Private Print_Jgyobu            As Variant          '����Ώێ��ƕ�
Private Print_Jgyobu_T()        As String * 1


Private Print_Yoin_PLUS         As Variant          '����Ώۗv��(�����v��)
Private Print_Yoin_PLUS_T()     As String * 2

Private Print_Yoin_MINUS        As Variant          '����Ώۗv��(�����v��)
Private Print_Yoin_MINUS_T()    As String * 2

Private Print_JYOGAI_SOKO       As Variant          '������O�q��(���O�q��)
Private Print_JYOGAI_SOKO_T()   As String * 2
        
        
Private Print_SHIME_BI          As String * 2       '���ߓ�
Private Print_DATE_S            As String * 8       '�g�p���͈́@�J�n
Private Print_DATE_E            As String * 8       '�g�p���͈́@�I��
        


Private Const LMAX% = 44                            '�œ��ő�s��
Private Const MGN_L% = 3                            '���]���i�����F�P����j
Private Const MGN_U% = 1                            '��]���i�s���F�P����j

Private Pdate                   As String           '����J�n���t�iͯ�ް�p�j
Private Ptime                   As String           '����J�n�����iͯ�ް�p�j

Private NormalFont              As New StdFont      '����t�H���g

Private PRT_CAN                 As Boolean          '����r���L�����Z���v��


Private wkDateTime              As String

Private F107040CSV              As String           'CSV�o�̓t�@�C��


Private Function Print_Proc() As Integer
'-------------------------------------------------------------------
'
'   �I���f�[�^�@�W�v�@���@���
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean



    Print_Proc = True

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�N���A�[��", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���o�b�@�z�I���e")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���o�b�@�z�I���e")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�W�v��", Me.hwnd, 0)


    wkDateTime = Format(Now, "YYYYMMDDHHMMSS")


'������́u������f�v�ȊO�̃C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time



    PRT_CAN = False

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ���o�ɏ��W�v
    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Print_DATE_S)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "�z�I���\������f", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Print_DATE_E Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ɉړ���")
                    Exit Function
            End Select
    
    
            For j = 0 To UBound(Print_Yoin_PLUS_T)
                If Print_Yoin_PLUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(1) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
    
            For j = 0 To UBound(Print_Yoin_MINUS_T)
                If Print_Yoin_MINUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(2) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
            com = BtOpGetNext
    
        Loop

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �݌ɏ��W�v
        
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Text1(ptxS_SOKO_No).Text)
        Call UniCode_Conv(K0_ZAIKO.Retu, Text1(ptxS_Retu).Text)
        Call UniCode_Conv(K0_ZAIKO.Ren, Text1(ptxS_Ren).Text)
        Call UniCode_Conv(K0_ZAIKO.Dan, Text1(ptxS_Dan).Text)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Text1(ptxS_Dan).Text)
    
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "�z�I���\������f", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Exit Function
            End Select
    
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = Print_Jgyobu_T(i) Then
                If Update_Proc(3) Then
                    Exit Function
                End If
            End If
    
            com = BtOpGetNext
                
        Loop
                
    Next i
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ��������J�n
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�����", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
                
                
                
    Do
    
        DoEvents
                                            '������f�v��
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '��ʍ��ڃ��b�N����
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "�z�I���\������f", Me.hwnd, 0)
            Command1.Visible = False
            Print_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���o�b�@�z�I���e")
                Exit Function
        End Select
    
    
        '�w�b�_�[�R���g���[��
        If LCNT > LMAX Then
            Call Print_Head(LCNT)
        End If
    
    
        '���ƕ�
        For j = 0 To UBound(JGYOBU_T)
            If JGYOBU_T(j).CODE = StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode) Then
                Exit For
            End If
        Next j
        Printer.Print Tab(MGN_L);
        If j <= UBound(JGYOBU_T) Then
            Call Moji_Cut_Proc(JGYOBU_T(j).NAME, RetBuf, 10)
            Printer.Print RetBuf;
        End If
        '�I��
        Printer.Print Tab(MGN_L + 15);
        Printer.Print StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode);
        '�i��
        Printer.Print Tab(MGN_L + 30);
        Printer.Print Left(StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode), 12);
        '�i��
        Printer.Print Tab(MGN_L + 45);
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OSAKA_PSTOCKREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
        Call Moji_Cut_Proc(StrConv(ITEMREC.HIN_NAME, vbUnicode), RetBuf, 20)
        Printer.Print RetBuf;
        '���ɐ�
        Printer.Print Tab(MGN_L + 77);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;
        '�o�ɐ�
        Printer.Print Tab(MGN_L + 87);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;
        '�݌ɐ�
        Printer.Print Tab(MGN_L + 97);
        RetBuf = Format(CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode)), "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf
        Printer.Print
        
        
        LCNT = LCNT + 2
        
        com = BtOpGetNext
    
    Loop
                

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    
    If WriteIni(App.EXEName, "LAST_PRINT_DateTime", App.EXEName, Now) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_PRINT_DateTime=")
        Unload Me
    End If
    
    
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\����I��", Me.hwnd, 0)

    Print_Proc = False
End Function

Private Sub Print_Head(LCNT As Integer)
                                        
Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    Printer.Print Text1(ptxYM).Text; " ����"
    Printer.Print Tab(36);
    Printer.Print "������  �u���ރZ���^�[�v�z�I���\  ������";
    Printer.Print Tab(100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    Printer.Print
                                        '���׈��
    Printer.Print Tab(MGN_L);
    Printer.Print "���ƕ�";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "�I    ��";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "�i  ��";
    Printer.Print Tab(MGN_L + 41);
    Printer.Print "�@�@�i             ��";
    Printer.Print Tab(MGN_L + 80);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 100);
    Printer.Print "���݌�"
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1070401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070401)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070401)


    F1070401.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �b�r�u�o��  2012.04.19
        Case 7
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�Ώ۔N���j"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_SOKO_No).Text) = "" Then
                Text1(ptxE_SOKO_No).Text = "zz"
            End If
            If Trim(Text1(ptxE_Retu).Text) = "" Then
                Text1(ptxE_Retu).Text = "zz"
            End If
            If Trim(Text1(ptxE_Ren).Text) = "" Then
                Text1(ptxE_Ren).Text = "zz"
            End If
            If Trim(Text1(ptxE_Dan).Text) = "" Then
                Text1(ptxE_Dan).Text = "zz"
            End If
            If (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) > _
                (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�I�Ԕ͈́j"
                Text1(ptxS_SOKO_No).SetFocus
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("�u�z�I���\�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Output_Proc() Then
                    Unload Me
                End If
            End If
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �b�r�u�o��
        
        
        
        
        
        
        
        
        
        
        
        Case 8                              '���
            If Not IsDate(Text1(ptxYM).Text & "/01") Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�Ώ۔N���j"
                Text1(ptxYM).SetFocus
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_SOKO_No).Text) = "" Then
                Text1(ptxE_SOKO_No).Text = "zz"
            End If
            If Trim(Text1(ptxE_Retu).Text) = "" Then
                Text1(ptxE_Retu).Text = "zz"
            End If
            If Trim(Text1(ptxE_Ren).Text) = "" Then
                Text1(ptxE_Ren).Text = "zz"
            End If
            If Trim(Text1(ptxE_Dan).Text) = "" Then
                Text1(ptxE_Dan).Text = "zz"
            End If
            If (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) > _
                (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i�I�Ԕ͈́j"
                Text1(ptxS_SOKO_No).SetFocus
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("�u�z�I���\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
            End If
                    
        Case 11                             '�I��
            Unload Me
    End Select
End Sub
Private Sub Command1_Click()
    PRT_CAN = True
End Sub
Private Sub Form_DblClick()
    PrintForm
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

Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer

Dim wkYY        As Integer
Dim wkMM        As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\��� ", Me.hwnd, 0)
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
    LOG_F = Trim(c)
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                
                                '����Ώێ��ƕ�
    If GetIni(App.EXEName, "JGYOBU_CODE", App.EXEName, c) Then
        MsgBox "����Ώێ��ƕ��̊l���Ɏ��s���܂���(JGYOBU_CODE=)�B�����𒆎~���܂��B"
        End
    Else
        Print_Jgyobu = Split(Trim(c), ",", -1)
        Erase Print_Jgyobu_T
        
        For i = 0 To UBound(Print_Jgyobu)
        
            ReDim Preserve Print_Jgyobu_T(0 To i)
            Print_Jgyobu_T(i) = Print_Jgyobu(i)
        Next i
        
        
    End If
                                '���ߓ�
    If GetIni(App.EXEName, "SHIME_BI", App.EXEName, c) Then
        MsgBox "���ߓ��̊l���Ɏ��s���܂���(SHIME_BI=)�B�����𒆎~���܂��B"
        End
    Else
        Print_SHIME_BI = Trim(c)
        If Not IsNumeric(Print_SHIME_BI) Then
            MsgBox "���ߓ��̊l���Ɏ��s���܂���(SHIME_BI=)�B�����𒆎~���܂��B"
            End
        End If
        
        
        '�J�n��
        If Mid(Format(Date, "YYYYMMDD"), 7, 2) > Print_SHIME_BI Then
            Print_DATE_S = Mid(Format(Date, "YYYYMMDD"), 1, 6) & Format(Val(Print_SHIME_BI) + 1, "00")
        Else
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
            wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) - 1
            If wkMM < 1 Then
                wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) - 1
                wkMM = 12
            End If
            Print_DATE_S = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(Print_SHIME_BI) + 1, "00")
        End If
        '�I����
        If Mid(Format(Date, "YYYYMMDD"), 7, 2) <= Print_SHIME_BI Then
            Print_DATE_E = Mid(Format(Date, "YYYYMMDD"), 1, 6) & Print_SHIME_BI
        Else
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
            wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) + 1
            If wkMM > 12 Then
                wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) + 1
                wkMM = 1
            End If
            Print_DATE_E = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(Print_SHIME_BI) + 1, "00")
        End If
    End If
                                
                                '����Ώۗv���i�����j
    If GetIni(App.EXEName, "YOIN_CODE_PLUS", App.EXEName, c) Then
        c = "**"
    End If
    Print_Yoin_PLUS = Split(Trim(c), ",", -1)
    Erase Print_Yoin_PLUS_T
    For i = 0 To UBound(Print_Yoin_PLUS)
    
        ReDim Preserve Print_Yoin_PLUS_T(0 To i)
        Print_Yoin_PLUS_T(i) = Print_Yoin_PLUS(i)
    Next i
                                
                                
                                
                                '����Ώۗv���i�����j
    If GetIni(App.EXEName, "YOIN_CODE_MINUS", App.EXEName, c) Then
        c = "**"
    End If
    Print_Yoin_MINUS = Split(Trim(c), ",", -1)
    
    Erase Print_Yoin_MINUS_T
        
    For i = 0 To UBound(Print_Yoin_MINUS)
    
        ReDim Preserve Print_Yoin_MINUS_T(0 To i)
        Print_Yoin_MINUS_T(i) = Print_Yoin_MINUS(i)
    Next i
                                '������O�q��
    If GetIni(App.EXEName, "JYOGAI_SOKO", App.EXEName, c) Then
        c = "**"
    End If
    
    
    Print_JYOGAI_SOKO = Split(Trim(c), ",", -1)
    Erase Print_JYOGAI_SOKO_T
        
    For i = 0 To UBound(Print_JYOGAI_SOKO)
    
        ReDim Preserve Print_JYOGAI_SOKO_T(0 To i)
        Print_JYOGAI_SOKO_T(i) = Print_JYOGAI_SOKO(i)
    Next i
                                
                                
                                '�b�r�u̧��
    If GetIni(App.EXEName, "F107040CSV", App.EXEName, c) Then
    Else
        F107040CSV = Trim(c)
        Command(7).Enabled = True
    End If
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenRead) Then
        Unload Me
    End If
                                '���o�b�@�z�I���e�n�o�d�m
    If OSAKA_PSTOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1070401.FontName
        .Size = F1070401.FontSize
    End With
    Set Printer.Font = NormalFont
    
    Text1(ptxYM).Text = Left(Format(Now, "YYYY/MM/DD"), 7)
    
    Text1(ptxYM).SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
    
    
    
    yn = MsgBox("[�z�I���\���]�������I�����܂����H", vbYesNo, "�m�F����")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '���o�b�@�z�I���e�b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���o�b�@�z�I���e")
        End If
    End If
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070401 = Nothing

    End
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text1(i).Enabled And Text1(i).Visible And Text1(i).TabStop Then
            Text1(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Update_Proc(Mode As Integer) As Integer
'-------------------------------------------------------------------
'
'   �I���f�[�^�@�쐬
'
'   mode :  1:���ɍX�V
'           2:�o�ɍX�V
'           3:�݌ɍX�V
'
'-------------------------------------------------------------------
Dim sts         As Integer

Dim wkSoko      As String * 2
Dim wkRetu      As String * 2
Dim wkRen       As String * 2
Dim wkDan       As String * 2


Dim wkJGYOBU    As String * 1
Dim wkNaigai    As String * 1
Dim wkHin_GAI   As String * 20

Dim wkQTY       As Long

Dim com         As Integer
Dim i           As Integer

    Update_Proc = True


    Select Case Mode
        Case 1
            If (StrConv(IDOREC.TO_SOKO, vbUnicode) & StrConv(IDOREC.TO_RETU, vbUnicode) & StrConv(IDOREC.TO_REN, vbUnicode) & StrConv(IDOREC.TO_DAN, vbUnicode)) < (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) Or _
                (StrConv(IDOREC.TO_SOKO, vbUnicode) & StrConv(IDOREC.TO_RETU, vbUnicode) & StrConv(IDOREC.TO_REN, vbUnicode) & StrConv(IDOREC.TO_DAN, vbUnicode)) > (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
    
                Update_Proc = False
                Exit Function
    
            Else
            
                wkSoko = StrConv(IDOREC.TO_SOKO, vbUnicode)
                wkRetu = StrConv(IDOREC.TO_RETU, vbUnicode)
                wkRen = StrConv(IDOREC.TO_REN, vbUnicode)
                wkDan = StrConv(IDOREC.TO_DAN, vbUnicode)
            
                wkJGYOBU = StrConv(IDOREC.JGYOBU, vbUnicode)
                wkNaigai = StrConv(IDOREC.NAIGAI, vbUnicode)
                wkHin_GAI = StrConv(IDOREC.HIN_GAI, vbUnicode)
            
            End If
    
        Case 2

            If (StrConv(IDOREC.FROM_SOKO, vbUnicode) & StrConv(IDOREC.FROM_RETU, vbUnicode) & StrConv(IDOREC.FROM_REN, vbUnicode) & StrConv(IDOREC.FROM_DAN, vbUnicode)) < (Text1(ptxS_SOKO_No).Text & Text1(ptxS_Retu).Text & Text1(ptxS_Ren).Text & Text1(ptxS_Dan).Text) Or _
                (StrConv(IDOREC.FROM_SOKO, vbUnicode) & StrConv(IDOREC.FROM_RETU, vbUnicode) & StrConv(IDOREC.FROM_REN, vbUnicode) & StrConv(IDOREC.FROM_DAN, vbUnicode)) > (Text1(ptxE_SOKO_No).Text & Text1(ptxE_Retu).Text & Text1(ptxE_Ren).Text & Text1(ptxE_Dan).Text) Then
    
                Update_Proc = False
                Exit Function
    
            Else
            
                wkSoko = StrConv(IDOREC.FROM_SOKO, vbUnicode)
                wkRetu = StrConv(IDOREC.FROM_RETU, vbUnicode)
                wkRen = StrConv(IDOREC.FROM_REN, vbUnicode)
                wkDan = StrConv(IDOREC.FROM_DAN, vbUnicode)
    
    
                wkJGYOBU = StrConv(IDOREC.JGYOBU, vbUnicode)
                wkNaigai = StrConv(IDOREC.NAIGAI, vbUnicode)
                wkHin_GAI = StrConv(IDOREC.HIN_GAI, vbUnicode)
    
    
            End If
    
    
        Case 3
    
            wkSoko = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            wkRetu = StrConv(ZAIKOREC.Retu, vbUnicode)
            wkRen = StrConv(ZAIKOREC.Ren, vbUnicode)
            wkDan = StrConv(ZAIKOREC.Dan, vbUnicode)
    
            wkJGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode)
            wkNaigai = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            wkHin_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
    
    End Select

    
    For i = 0 To UBound(Print_JYOGAI_SOKO_T)
    
        If wkSoko = Print_JYOGAI_SOKO_T(i) Then
    
            Update_Proc = False
            Exit Function
        
        End If
    
    Next i
    
    
    
'-------------------------------------------------------------  ���o�b�@�z�I���e
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Soko_No, wkSoko)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Retu, wkRetu)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Ren, wkRen)
    
    Call UniCode_Conv(K0_OSAKA_PSTOCK.Dan, wkDan)               '2013.03.23
    
    Call UniCode_Conv(K0_OSAKA_PSTOCK.JGYOBU, wkJGYOBU)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.NAIGAI, wkNaigai)
    Call UniCode_Conv(K0_OSAKA_PSTOCK.HIN_GAI, wkHin_GAI)
    
    sts = BTRV(BtOpGetEqual, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
    Select Case sts
        Case BtNoErr
        
            com = BtOpUpdate
        
        Case BtErrKeyNotFound
        
            com = BtOpInsert
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���o�b�@�z�I���e")
            Exit Function
    End Select
    
    If com = BtOpInsert Then
        
        Call UniCode_Conv(OSAKA_PSTOCKREC.Soko_No, wkSoko)      '�q�ɇ�
        Call UniCode_Conv(OSAKA_PSTOCKREC.Retu, wkRetu)         '�I�ԁ@��
        Call UniCode_Conv(OSAKA_PSTOCKREC.Ren, wkRen)           '�I�ԁ@�A
        Call UniCode_Conv(OSAKA_PSTOCKREC.Dan, wkDan)           '�I�ԁ@�i
        Call UniCode_Conv(OSAKA_PSTOCKREC.JGYOBU, wkJGYOBU)     '���ƕ��敪
        Call UniCode_Conv(OSAKA_PSTOCKREC.NAIGAI, wkNaigai)     '�����O
        Call UniCode_Conv(OSAKA_PSTOCKREC.HIN_GAI, wkHin_GAI)   '�i�ԁi�O���j
                                                                '�v��N��
        Call UniCode_Conv(OSAKA_PSTOCKREC.KEIJYO_YM, Left(Format(Text1(ptxYM).Text & "/01"), 7))
                                                                '�������ɐ�
        Call UniCode_Conv(OSAKA_PSTOCKREC.NYUKO_QTY, "0000000000")
                                                                '�����o�ɐ�
        Call UniCode_Conv(OSAKA_PSTOCKREC.SYUKO_QTY, "0000000000")
                                                                '�����݌Ɏc��
        Call UniCode_Conv(OSAKA_PSTOCKREC.ZAIKO_QTY, "0000000000")
        
        Call UniCode_Conv(OSAKA_PSTOCKREC.FILLER, "")
                                                                '�ް��쐬����
        Call UniCode_Conv(OSAKA_PSTOCKREC.Ins_DateTime, wkDateTime)
    
    
    End If
    
    
    Select Case Mode
        Case 1
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode))
            wkQTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.NYUKO_QTY, Format(wkQTY, "0000000000"))
        Case 2
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode))
            wkQTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.SYUKO_QTY, Format(wkQTY, "0000000000"))
        Case 3
            wkQTY = CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode))
            wkQTY = wkQTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Call UniCode_Conv(OSAKA_PSTOCKREC.ZAIKO_QTY, Format(wkQTY, "0000000000"))
    End Select


    sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrDuplicates
        
            Call File_Error(sts, com, "���o�b�@�z�I���e" & "KEY =" & StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode) & StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
                    
        Case Else
            Call File_Error(sts, com, "���o�b�@�z�I���e")
            Exit Function
    End Select

    Update_Proc = False


End Function
Private Function Output_Proc() As Integer
'-------------------------------------------------------------------
'
'   �I���f�[�^�@�W�v�@���@�f�[�^�o��
'
'-------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean

Dim FileNo          As Integer


    Output_Proc = True


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (F107040CSV) For Output As FileNo


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�N���A�[��", Me.hwnd, 0)
    
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���o�b�@�z�I���e")
                Exit Function
        End Select
    
    
        sts = BTRV(BtOpDelete, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "���o�b�@�z�I���e")
                Exit Function
        End Select
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�W�v��", Me.hwnd, 0)


    wkDateTime = Format(Now, "YYYYMMDDHHMMSS")


'������́u������f�v�ȊO�̃C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N
    Command1.Visible = True
    Command1.Enabled = True


    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time



    PRT_CAN = False

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ���o�ɏ��W�v
    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Print_DATE_S)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "�z�I���\�f�[�^�o�͒��f", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Print_DATE_E Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ɉړ���")
                    Exit Function
            End Select
    
    
            For j = 0 To UBound(Print_Yoin_PLUS_T)
                If Print_Yoin_PLUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(1) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
    
            For j = 0 To UBound(Print_Yoin_MINUS_T)
                If Print_Yoin_MINUS_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    
                    If Update_Proc(2) Then
                        Exit Function
                    End If
                    Exit For
                End If
            Next j
    
            com = BtOpGetNext
    
        Loop

    Next i

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  �݌ɏ��W�v
        
    For i = 0 To UBound(Print_Jgyobu_T)
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, Text1(ptxS_SOKO_No).Text)
        Call UniCode_Conv(K0_ZAIKO.Retu, Text1(ptxS_Retu).Text)
        Call UniCode_Conv(K0_ZAIKO.Ren, Text1(ptxS_Ren).Text)
        Call UniCode_Conv(K0_ZAIKO.Dan, Text1(ptxS_Dan).Text)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Text1(ptxS_Dan).Text)
    
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
    
        com = BtOpGetGreaterEqual
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "�z�I���\�f�[�^�o�͒��f", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Exit Function
            End Select
    
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) = Print_Jgyobu_T(i) Then
                If Update_Proc(3) Then
                    Exit Function
                End If
            End If
    
            com = BtOpGetNext
                
        Loop
                
    Next i
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  ��������J�n
                
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�f�[�^�o�͒�", Me.hwnd, 0)
                
                
    com = BtOpGetFirst
    LCNT = 99
                
                
                
    Do
    
        DoEvents
                                            '������f�v��
        If PRT_CAN Then
            Printer.KillDoc
            Call Input_UnLock               '��ʍ��ڃ��b�N����
            hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "�z�I���\�f�[�^�o�͒��f", Me.hwnd, 0)
            Command1.Visible = False
            Output_Proc = False
            Exit Function
        End If
    
        sts = BTRV(com, OSAKA_PSTOCK_POS, OSAKA_PSTOCKREC, Len(OSAKA_PSTOCKREC), K0_OSAKA_PSTOCK, Len(K0_OSAKA_PSTOCK), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���o�b�@�z�I���e")
                Exit Function
        End Select
    
    
        '�w�b�_�[�R���g���[��
        If LCNT = 99 Then
                
            Write #FileNo, "���ƕ�", "�I    ��", "�i  ��", "�@�@�i             ��", "������", "������", "���݌�"
   
            LCNT = 0
        End If
    
    
        '���ƕ�
        For j = 0 To UBound(JGYOBU_T)
            If JGYOBU_T(j).CODE = StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode) Then
                Write #FileNo, JGYOBU_T(j).NAME,
                Exit For
            End If
        Next j
        '�I��
        Write #FileNo, StrConv(OSAKA_PSTOCKREC.Soko_No, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Retu, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Ren, vbUnicode) & "-" & _
                        StrConv(OSAKA_PSTOCKREC.Dan, vbUnicode),
        '�i��
        Write #FileNo, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode),
        '�i��
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OSAKA_PSTOCKREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OSAKA_PSTOCKREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OSAKA_PSTOCKREC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
                            
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
        Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
        Printer.Print RetBuf;
        '���ɐ�
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.NYUKO_QTY, vbUnicode)), "#,##0"),
        '�o�ɐ�
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.SYUKO_QTY, vbUnicode)), "#,##0"),
        '�݌ɐ�
        Write #FileNo, Format(CLng(StrConv(OSAKA_PSTOCKREC.ZAIKO_QTY, vbUnicode)), "#,##0")
        
        
        
        com = BtOpGetNext
    
    Loop
                
    Close #FileNo
    MsgBox "�u" & F107040CSV & "�v�͐���ɏo�͂���܂����B"
    
    
    If WriteIni(App.EXEName, "LAST_PRINT_DateTime", App.EXEName, Now) Then
        Beep
        MsgBox ("INI�t�@�C���̏������݂Ɏ��s���܂����B" & App.EXEName & "LAST_PRINT_DateTime=")
        Unload Me
    End If
    
    
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�z�I���\�f�[�^�I��", Me.hwnd, 0)

    Output_Proc = False

    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox F107040CSV & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

End Function



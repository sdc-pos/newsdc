VERSION 5.00
Begin VB.Form F1040401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����ؗ��i�A���[�����X�g���"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
      TabIndex        =   12
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
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
      Index           =   6
      Left            =   5640
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ȍ~�s�ړ����̕i�ڂ�ΏۂƂ��܂��B"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   20
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   19
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   18
      Top             =   1920
      Width           =   255
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
      TabIndex        =   17
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ł�"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ŏI�o�ɓ�"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1040401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxLAST_SYU_DT_YY% = 0        '�ŏI�o�ɓ� �N
Private Const ptxLAST_SYU_DT_MM% = 1        '�ŏI�o�ɓ� ��
Private Const ptxLAST_SYU_DT_DD% = 2        '�ŏI�o�ɓ� ��

Private Const Text_Max% = 2                 '��ʍ��ڕʍő���ޯ��

Private Const LMAX% = 46                    '�œ��ő�s��
Private Const MGN_L% = 10                   '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Dim Pdate       As String                   '����J�n���t�iͯ�ް�p�j
Dim Ptime       As String                   '����J�n�����iͯ�ް�p�j
Dim ALARM_DATA  As String                   '�A���[���f�[�^�t���p�X


Dim NormalFont  As New StdFont               '����t�H���g

Dim PRT_CAN     As Boolean                  '����r���L�����Z���v��


Private Function Print_Proc() As Integer

Dim ITEM_com        As Integer
Dim sts             As Integer

Dim LCNT            As Integer

Dim PRINT_OK        As Boolean
Dim ZAIKO_ON        As Boolean

'Dim PRI_NAIGAI      As String * 1
'Dim PRI_HIN_GAI     As String * 13

Dim RetBuf          As String
    
Dim Sumi_Zaiko_Qty  As Long
Dim Mi_Zaiko_Qty    As Long


    Print_Proc = True
    Call Input_Lock           '��ʍ��ڃ��b�N
    Label1.Visible = True


    LCNT = 99
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time
                                            '�i�ڃ}�X�^�ǂݍ��݊J�n
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    ITEM_com = BtOpGetGreaterEqual

    Do
        DoEvents
        
        
        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "�i�ڃ}�X�^")
                Exit Function
        End Select
    
        PRINT_OK = False
        
                
'        If Len(Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))) <> 0 And _
'            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) <> "00000000" Then
            If StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) <= (Text(ptxLAST_SYU_DT_YY).Text & _
                                                            Text(ptxLAST_SYU_DT_MM).Text & _
                                                            Text(ptxLAST_SYU_DT_DD).Text) Then
                PRINT_OK = True
            End If
'        End If
    
        If PRINT_OK Then
                                                            
'            PRI_NAIGAI = ""
'            PRI_HIN_GAI = ""
            ZAIKO_ON = False
        
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If

            If (Sumi_Zaiko_Qty + Mi_Zaiko_Qty) = 0 Then
            Else
                If LCNT > LMAX Then
                    Call Print_Head(LCNT)
'                    PRI_NAIGAI = ""
                End If
        
                ZAIKO_ON = True
        
                Printer.Print Tab(MGN_L);
'               If PRI_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Then
                    Printer.Print Tab(MGN_L);
                    If StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI Then
                        Printer.Print NAIGAI1;
                    Else
                        Printer.Print NAIGAI2;
                    End If
'                    PRI_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
'                   PRI_HIN_GAI = ""
'               End If
            
'                If PRI_HIN_GAI <> StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
            
                    Printer.Print Tab(MGN_L + 10);
                    Printer.Print StrConv(ITEMREC.HIN_GAI, vbUnicode);
                    Printer.Print Tab(MGN_L + 32);
                    Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
                    Printer.Print Tab(MGN_L + 82);
                    
                    If Len(Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))) <> 0 Then
                        Printer.Print Left(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 4) & "/";
                        Printer.Print Mid(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 5, 2) & "/";
                        Printer.Print Right(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 2);
                    End If
'                    PRI_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
'                End If
            
                Printer.Print Tab(MGN_L + 102);
        
                RetBuf = Format((Sumi_Zaiko_Qty + Mi_Zaiko_Qty), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf
                LCNT = LCNT + 1
            
            End If
        
        End If
    
        ITEM_com = BtOpGetNext
    
    Loop

    If LCNT <> 99 Then
        Printer.EndDoc
    End If

    Call Input_UnLock               '��ʍ��ڃ��b�N����
    Label1.Visible = False

    Print_Proc = False

End Function
Private Function OUTPUT_Proc() As Integer

Dim ITEM_com        As Integer
Dim sts             As Integer
Dim Ret             As Integer

Dim PRINT_OK        As Boolean

Dim FileNo          As Integer
Dim fileName        As String
    
Dim Sumi_Zaiko_Qty  As Long
Dim Mi_Zaiko_Qty    As Long

    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock           '��ʍ��ڃ��b�N

    FileNo = FreeFile
    fileName = ALARM_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo

    Write #FileNo, "���O", "�i�ԁi�O�j", "�ŏI���ɓ�", "�ŏI�o�ɓ�", "�݌ɐ�"

                                            '�i�ڃ}�X�^�ǂݍ��݊J�n
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    ITEM_com = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "�i�ڃ}�X�^")
                Exit Function
        End Select
    
        PRINT_OK = False
        
                
'        If Len(Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))) <> 0 And _
'            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) <> "00000000" Then
            If StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) <= (Text(ptxLAST_SYU_DT_YY).Text & _
                                                            Text(ptxLAST_SYU_DT_MM).Text & _
                                                            Text(ptxLAST_SYU_DT_DD).Text) Then
                PRINT_OK = True
            End If
'        End If
    
        If PRINT_OK Then
                                                            
            If Zaiko_Syukei_Proc(Sumi_Zaiko_Qty, _
                                    Mi_Zaiko_Qty, _
                                    StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If

            If (Sumi_Zaiko_Qty + Mi_Zaiko_Qty) = 0 Then
            Else
        
                If StrConv(ITEMREC.NAIGAI, vbUnicode) = NAIGAI_NAI Then
                    Write #FileNo, NAIGAI1,
                Else
                    Write #FileNo, NAIGAI2,
                End If
        
                Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
                
                If Len(Trim(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode))) <> 0 Then
                    Write #FileNo, Left(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode), 4) & "/" _
                            & Mid(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode), 5, 2) & "/" _
                            & Right(StrConv(ITEMREC.LAST_NYU_DT, vbUnicode), 2),
                Else
                    Write #FileNo, ,
                End If
                
                
                
                If Len(Trim(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode))) <> 0 Then
                    Write #FileNo, Left(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 4) & "/" _
                            & Mid(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 5, 2) & "/" _
                            & Right(StrConv(ITEMREC.LAST_SYU_DT, vbUnicode), 2),
                Else
                    Write #FileNo, ,
                End If
                
                            
                Write #FileNo, Format(CLng(Sumi_Zaiko_Qty + Mi_Zaiko_Qty))
            
            End If
        
        End If
    
        ITEM_com = BtOpGetNext
    
    Loop

    Close #FileNo

    Call Input_UnLock               '��ʍ��ڃ��b�N����

    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"
    

    OUTPUT_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If

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
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(26);
    Printer.Print "������  �����ؗ��i�A���[�����X�g  ������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        
                                        '���׈��
    Printer.Print Tab(MGN_L);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 32);
    Printer.Print "�i  ��  ";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "�ŏI�o�ɓ�";
    Printer.Print Tab(MGN_L + 101);
    Printer.Print "�L���݌ɐ�"           '97.07.16
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1040401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040401)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040401)


    F1040401.MousePointer = vbDefault

End Sub


Private Function Err_Chk()
    
Dim i As Integer
    
    Err_Chk = True


    For i = ptxLAST_SYU_DT_YY To ptxLAST_SYU_DT_DD
        If Len(Text(i).Text) = 0 Then
            Select Case i
                Case ptxLAST_SYU_DT_YY
                    Text(i).Text = "0000"
                Case Else
                    Text(i).Text = "00"
            End Select
        Else
            If IsNumeric(Text(i).Text) Then
                Select Case i
                    Case ptxLAST_SYU_DT_YY
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    Case Else
                        Text(i).Text = Format(CInt(Text(i).Text), "00")
                End Select
            End If
        End If
    Next i
    
    Err_Chk = False

End Function

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              '�f�[�^�o��
        
            If Err_Chk() Then
                Exit Sub
            End If
        
            Beep
            ans = MsgBox("�u�����ؗ��i�A���[�����X�g�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxLAST_SYU_DT_YY).SetFocus
        
        Case 8                              '���
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("�u�����ؗ��i�A���[�����X�g�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(ptxLAST_SYU_DT_YY).SetFocus
                    
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
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
        
Dim ALARM_DATE  As String * 8

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
    LOG_F = Trim(c)
                                '�A���[���t�@�C������荞��
    If GetIni("FILE", "ALARM_DATA", "SYS", c) Then
        Beep
        MsgBox "�A���[���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ALARM_DATA = Trim(c)
                                '���t�̃f�t�H���g
    If GetIni(App.EXEName, "ALARM_DATE", "SYS", c) Then
        Beep
        MsgBox "���t�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    If Not IsNumeric(Trim(c)) Then
    Else
        ALARM_DATE = Format(DateAdd("d", -CInt(Trim(c)), Date), "yyyymmdd")
        Text(ptxLAST_SYU_DT_YY).Text = Left(ALARM_DATE, 4)
        Text(ptxLAST_SYU_DT_MM).Text = Mid(ALARM_DATE, 5, 2)
        Text(ptxLAST_SYU_DT_DD).Text = Right(ALARM_DATE, 2)
    End If
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1040401.Caption = "�����ؗ��i�A���[�����X�g����i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1040401.FontName
        .Size = F1040401.FontSize
    End With
    Set Printer.Font = NormalFont


End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1040401 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1040401.Caption = "�����ؗ��i�A���[�����X�g����i" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub



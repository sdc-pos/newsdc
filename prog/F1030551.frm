VERSION 5.00
Begin VB.Form F1030551 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���o�b�����i�ԕʏo�ɕ\���"
   ClientHeight    =   5655
   ClientLeft      =   2325
   ClientTop       =   2715
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3825
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1020
      Width           =   285
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   4725
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2760
      Width           =   480
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   3885
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2760
      Width           =   480
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1620
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1620
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   3840
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1620
      Width           =   615
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4560
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0=���o��/9=�o�Ɋ���"
      Height          =   255
      Index           =   6
      Left            =   4275
      TabIndex        =   34
      Top             =   1140
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������/�Ǎ�����"
      Height          =   255
      Index           =   2
      Left            =   1530
      TabIndex        =   33
      Top             =   3540
      Width           =   2085
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ɋ敪"
      Height          =   255
      Index           =   5
      Left            =   2700
      TabIndex        =   32
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Height          =   315
      Index           =   1
      Left            =   4680
      TabIndex        =   31
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   4
      Left            =   4545
      TabIndex        =   30
      Top             =   3540
      Width           =   120
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      Height          =   315
      Index           =   0
      Left            =   3915
      TabIndex        =   29
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   1
      Left            =   4410
      TabIndex        =   28
      Top             =   2880
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����Ŕ͈�"
      Height          =   255
      Index           =   0
      Left            =   2415
      TabIndex        =   27
      Top             =   2880
      Width           =   1275
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
      TabIndex        =   26
      Top             =   5160
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   25
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   2
      Left            =   4920
      TabIndex        =   24
      Top             =   1740
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   1
      Left            =   4440
      TabIndex        =   23
      Top             =   1740
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����敪"
      Height          =   255
      Index           =   3
      Left            =   2730
      TabIndex        =   22
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ח\���"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   1740
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
Attribute VB_Name = "F1030551"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxPRINT_KBN% = 0             '����敪


Private Const ptxSYUKA_YY% = 1              '�o�ח\����@�N
Private Const ptxSYUKA_MM% = 2              '�o�ח\����@��
Private Const ptxSYUKA_DD% = 3              '�o�ח\����@��

Private Const ptxINS_BIN% = 4               '��

Private Const ptxS_Page% = 5                '����J�n�@�Ő�
Private Const ptxE_Page% = 6                '����I���@�Ő�


Private Const Text_Max% = 6                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbPRINT_KBN% = 0            '����敪


Private Const Print_KBN0$ = "�V�@�K"
Private Const Print_KBN1$ = "�Ĉ��"
Private Const Print_KBN_SIN$ = "0"
Private Const Print_KBN_SAI$ = "1"

Private KASO_NYUKA_SOKO As String * 2       '���z�@���בq�ɔԍ�
Private KASO_SYOHN_SOKO As String * 2       '���z�@���i���q�ɔԍ�
Private KASO_NAI_SOKO As String * 2         '���z�@���E�q�ɔԍ�


Private Const LMAX% = 40                    '�œ��ő�s��
Private Const MGN_L% = 2                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j
'Dim PRT_CAN As Boolean                     '����r���L�����Z���v��


Private NormalFont As New StdFont           '����t�H���g
Private SmallFont As New StdFont            '����t�H���g   '2007.03.29

Private BoldFont As New StdFont             '����t�H���g   '2007.03.29

Private RargeFont As New StdFont             '����t�H���g   '2007.03.29



Private Code39Font As New StdFont           '����t�H���g



'�X�e�[�V������
Private WS_NO       As String * 10


Private Input_Cnt   As Integer
Private Output_Cnt  As Integer



Private P_CNT       As Integer
Private Print_F     As Boolean
Private Const LAST_UPDATE_DAY$ = "[F103055]2010.04.07 11:00"

Private Function Print_Proc() As Integer

Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 1

Dim com             As Integer
Dim sts             As Integer
    
Dim RetBuf          As String
    
Dim Print_F         As Boolean
    
    
    Print_Proc = True

    Call Input_Lock
    
    
    Lcnt = 99
    P_CNT = 0
    Print_F = False
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    Call UniCode_Conv(K1_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text)
    Call UniCode_Conv(K1_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_SOKO, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_RETU, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_REN, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.ST_DAN, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.JGYOBU, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.NAIGAI, "")
    Call UniCode_Conv(K1_Y_SYU_SUM.HIN_NO, "")
    
    
        
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K1_Y_SYU_SUM, Len(K1_Y_SYU_SUM), 1)
    
        Select Case sts
            Case BtNoErr
                
                If StrConv(Y_SYU_SUMREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                    Text(ptxSYUKA_MM).Text & _
                                                                    Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
                
                
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Call Input_UnLock
                Exit Function
        End Select
                                            
        Input_Cnt = Input_Cnt + 1
        Label3(1).Caption = Format(Input_Cnt, "#0")
                                            
                                            
        Print_F = True
                                            
                                            
        Select Case Trim(Text(ptxPRINT_KBN).Text)
            Case ""
            Case "0"
                If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) = 0 Then
                    Print_F = False
                End If
            Case "9"
                If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) <> 0 Then
                    Print_F = False
                End If
        End Select
                                            
                                            
                                            
'        If (CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)) - CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode))) = 0 Then
        
        If Not Print_F Then
        Else
            If Lcnt = 99 Then
                SAVE_SOKO_No = Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1)
            Else
                                                '�q�ɂ̃u���[�N
                If SAVE_SOKO_No <> Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1) Then
                    Lcnt = LMAX + 1
                    SAVE_SOKO_No = Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1)
                End If
            End If

            If Lcnt > LMAX Then                 '�w�b�_�[�R���g���[��
                If Head_Proc(Lcnt) Then
                    Call Input_Lock
                    Exit Function
                End If
            
                '�w��y�[�W��OVER
                If P_CNT > CInt(Text(ptxE_Page).Text) Then
                    If Lcnt <> 99 Then
                        Printer.EndDoc
                    End If
                
                
                
                    Call Input_UnLock
                
                    Print_Proc = False
                
                    Exit Function
                
                End If
            
            
            End If
                                                
                                                            
            If P_CNT < CInt(Text(ptxS_Page).Text) Then
            Else
                                                
                                                
                Output_Cnt = Output_Cnt + 1
                Label3(0).Caption = Format(Output_Cnt, "#0")
                                                
                                                
                '-----------------------------------------------------  '�P�s��
                Printer.Print Tab(MGN_L);
                '�W���I��
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_RETU, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_REN, vbUnicode), 1, 2) & "-";
                Printer.Print Mid(StrConv(Y_SYU_SUMREC.ST_DAN, vbUnicode), 1, 2);
                '�i��(�O)
'                Printer.Print Tab(MGN_L + 14);
                Printer.Print Tab(MGN_L + 18);
                Set Printer.Font = RargeFont
                Printer.Print Left(StrConv(Y_SYU_SUMREC.HIN_NO, vbUnicode), 13);
                Set Printer.Font = NormalFont
    
    '            Printer.Print Tab(MGN_L + 34);
    '            '�ʒu�I����
    '            If Trim(StrConv(Y_SYU_SUMREC.BETU_SOKO, vbUnicode)) = "" Then
    '            Else
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_RETU, vbUnicode), 1, 2) & "-";
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_REN, vbUnicode), 1, 2) & "-";
    '                Printer.Print Mid(StrConv(Y_SYU_SUMREC.BETU_DAN, vbUnicode), 1, 2);
    '            End If
    '            Printer.Print Tab(MGN_L + 46);
    '            '�ʒu���݌ɐ�
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.BETU_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
    '            Printer.Print Tab(MGN_L + 55);
    '            '���i�������E�݌ɐ�
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.SYO_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
    '            '���בq�ɍ݌�
    '            RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.NYU_ZAIKO_QTY, vbUnicode)), "#,##0")
    '            If Len(RetBuf) < 9 Then
    '                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    '            End If
    '            Printer.Print RetBuf;
        
                '�o�ɍς݁^�o�ɗ\�萔
'                Printer.Print Tab(MGN_L + 29);
                Printer.Print Tab(MGN_L + 40);
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.Y_SURYO, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Set Printer.Font = BoldFont
                Printer.Print RetBuf;
                Set Printer.Font = NormalFont
                
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.J_SURYO, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print " (" & RetBuf & ")";
                
                
                
                
                '����t�H���g�ݒ�i�b�������R�X�j
'                Printer.Print Tab(MGN_L + 51);
'                Set Printer.Font = Code39Font
'                                    '�o�[�R�[�h(*�`�[ID*)
'                Printer.Print "*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*";
'                                                        '����t�H���g�ݒ�i�ʏ�j
'                Set Printer.Font = NormalFont
                
'                Printer.Print Tab(MGN_L + 75);
                Printer.Print Tab(MGN_L + 65);
'                Printer.Print "(*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*)";
                Printer.Print StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode);

                '�W���I�@�݌ɐ�
                Printer.Print Tab(MGN_L + 93);
                RetBuf = Format(CLng(StrConv(Y_SYU_SUMREC.ST_ZAIKO_QTY, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Set Printer.Font = SmallFont
                Printer.Print RetBuf;
                Set Printer.Font = NormalFont
                
                '����
                Printer.Print Tab(MGN_L + 105);
                RetBuf = Format(CInt(StrConv(Y_SYU_SUMREC.DATA_CNT, vbUnicode)), "#0")
                If Len(RetBuf) < 4 Then
                    RetBuf = Space(4 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf
                
                
                
                
                '�W���I��BC
                Printer.Print Tab(MGN_L);
                
                
                If Trim(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode)) <> "" Then
                    Set Printer.Font = Code39Font
                    Printer.Print "*/" & StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_RETU, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_REN, vbUnicode) & _
                                            StrConv(Y_SYU_SUMREC.ST_DAN, vbUnicode) & "*";
                    Set Printer.Font = NormalFont
                End If
                Printer.Print Tab(MGN_L + 65);
                                    '�o�[�R�[�h(*�`�[ID*)
                Set Printer.Font = Code39Font
                Printer.Print "*" & StrConv(Y_SYU_SUMREC.SYU_NO, vbUnicode) & "*"
                                                        '����t�H���g�ݒ�i�ʏ�j
                Set Printer.Font = NormalFont
                
                        
                Printer.Print String(115, "��")
'                Printer.Print
            
            End If
            
            
            
            
            
            Lcnt = Lcnt + 3
        
        End If
        
        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If

    Label3(0).Caption = Format(Output_Cnt, "#0")


    Call Input_UnLock

    Print_Proc = False

End Function
                                    
Private Function Head_Proc(Lcnt As Integer) As Integer
Dim i As Integer
Dim sts As Integer

    Head_Proc = True


    P_CNT = P_CNT + 1

    If P_CNT < CInt(Text(ptxS_Page).Text) Or P_CNT > CInt(Text(ptxE_Page).Text) Then
        Lcnt = 8 + MGN_U
        Head_Proc = False
        Exit Function
    End If


    If Print_F Then
        Printer.NewPage
    End If

    Print_F = True



    For i = 1 To MGN_U
        Printer.Print
    Next i

''    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    
    Printer.Print Tab(MGN_L + 31);
    
    '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
    If Trim(Text(ptxINS_BIN).Text) = "" Then
        Printer.Print "�w�֎w��Ȃ��x�o�ɕ\";
    Else
        Printer.Print "�w" + Text(ptxINS_BIN).Text + "�ցx�o�ɕ\";
    End If
    '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
    
    If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SAI Then
        Printer.Print Tab(MGN_L + 63);
        Printer.Print "�y�Ĉ���z";
    End If
    
    Printer.Print Tab(MGN_L + 81);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "  P." & Format(P_CNT, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L + 5);
    Printer.Print "�q�ɁF";
    Printer.Print Left(StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode), 1);
''    Printer.Print Tab(MGN_L + 15);
''    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(Y_SYU_SUMREC.ST_SOKO, vbUnicode))
''    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''    Select Case sts
''        Case BtNoErr
''            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
''        Case BtErrKeyNotFound
''        Case Else
''            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
''            Exit Function
''    End Select
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
'    Printer.Print Tab(MGN_L + 14);
    Printer.Print Tab(MGN_L + 18);
    Printer.Print "�i�ԁi�O���j";
'    Printer.Print Tab(MGN_L + 28);
    Printer.Print Tab(MGN_L + 39);
'    Printer.Print Tab(MGN_L + 35);
'    Printer.Print "�ʒu�I��";
'    Printer.Print Tab(MGN_L + 47);
'    Printer.Print "�ʒu�݌�";
'    Printer.Print Tab(MGN_L + 56);
'    Printer.Print "���i����";
'    Printer.Print Tab(MGN_L + 65);
'    Printer.Print "���בq��";
'    Printer.Print Tab(MGN_L + 40);
    Set Printer.Font = BoldFont
    Printer.Print "�o�׎w����";
    Set Printer.Font = NormalFont
    
'    Printer.Print Tab(MGN_L + 39);
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "(�o�� �ϐ�)";
    
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "�o�ɂh�c";
    
    
    Printer.Print Tab(MGN_L + 92);
    Set Printer.Font = SmallFont
    Printer.Print "�W���I�݌�";
    Set Printer.Font = NormalFont
    Printer.Print Tab(MGN_L + 104);
    Printer.Print "����"
    
    
    Printer.Print String(115, "��")
    
    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYU_HREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Tana_Kensaku = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    
    
    
    
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, StrConv(Y_SYU_HREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYU_HREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYU_HREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.SOKO_NO, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(Y_SYU_HREC.JGYOBU, vbUnicode) Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYU_HREC.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.SOKO_NO, vbUnicode) <> StrConv(ITEMREC.ST_SOKO, vbUnicode) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(ITEMREC.ST_RETU, vbUnicode) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(ITEMREC.ST_REN, vbUnicode) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(ITEMREC.ST_DAN, vbUnicode) Then
                                                '�V�X�e���q�ɂ̔���
                    Call UniCode_Conv(K0_SOKO.SOKO_NO, StrConv(ZAIKOREC.SOKO_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
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
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1030551.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030551)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030551)


    F1030551.MousePointer = vbDefault

End Sub





Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbPRINT_KBN      '����敪
            Text(ptxPRINT_KBN).SetFocus
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim sts     As Integer
Dim com     As Integer
    
Dim c       As String * 128
    
    Select Case Index
        Case 8                              '���
            
            
            If Text(ptxPRINT_KBN).Text <> "0" And Text(ptxPRINT_KBN).Text <> "9" And Trim(Text(ptxPRINT_KBN).Text) <> "" Then
                MsgBox "���͓��e�̓G���[�ł��B"
                Text(ptxPRINT_KBN).SetFocus
                Exit Sub
            End If
            
'>>>>>>>>>>>>>>>>>>>>>  �և��@�`�k�k�w���  2012.12.21
            
'2012.12.21            For i = ptxSYUKA_YY To ptxINS_BIN
            For i = ptxSYUKA_YY To ptxSYUKA_MM
            
'>>>>>>>>>>>>>>>>>>>>>  �և��@�`�k�k�w���  2012.12.21
                
                
                
                
                If Not IsNumeric(Text(i).Text) Then
                    MsgBox "���͓��e�̓G���[�ł��B"
                    Text(i).SetFocus
                    Exit Sub
                End If
            
            
                Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
                            
            
            
            Next i
            
            
'>>>>>>>>>>>>>>>>>>>>>  �և��@�`�k�k�w���  2012.12.21
            If Trim(Text(ptxINS_BIN).Text) = "" Then
            Else
                If Not IsNumeric(Text(ptxINS_BIN).Text) Then
                    MsgBox "���͓��e�̓G���[�ł��B(�և��F���l or ��)"
                    Text(ptxINS_BIN).SetFocus
                    Exit Sub
                End If
            End If
'>>>>>>>>>>>>>>>>>>>>>  �և��@�`�k�k�w���  2012.12.21
            
            
            If Not IsNumeric(Text(ptxS_Page).Text) Then
                Text(ptxS_Page).Text = "001"
            Else
                Text(ptxS_Page).Text = Format(CInt(Text(ptxS_Page).Text), "000")
            End If
            
            If Not IsNumeric(Text(ptxE_Page).Text) Then
                Text(ptxE_Page).Text = "999"
            Else
                Text(ptxE_Page).Text = Format(CInt(Text(ptxE_Page).Text), "000")
            End If
            
            
            
            
            '����ς݂�����
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                Call UniCode_Conv(K2_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                            Text(ptxSYUKA_MM).Text & _
                                                            Text(ptxSYUKA_DD).Text)
                Call UniCode_Conv(K2_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
                
                sts = BTRV(BtOpGetEqual, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                Select Case sts
                    Case BtNoErr
'                        MsgBox "�w����e�́A�o�ɕ\����ς݂ł��B�Ĉ����I�����ĉ������B"
'                        Combo(pcmbPRINT_KBN).SetFocus
'                        Exit Sub
                                            
                        yn = MsgBox("�w����e�́A�o�ɕ\����ς݂ł��B�������p�����܂����H", vbOKCancel + vbDefaultButton2 + vbCritical, "�m�F����")
                        If yn = vbCancel Then
                            Combo(pcmbPRINT_KBN).SetFocus
                            Exit Sub
                        End If
                
                        Call UniCode_Conv(K2_Y_SYU_SUM.SYUKA_YMD, Text(ptxSYUKA_YY).Text & _
                                                                    Text(ptxSYUKA_MM).Text & _
                                                                    Text(ptxSYUKA_DD).Text)
                        Call UniCode_Conv(K2_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
                        com = BtOpGetGreaterEqual
                    
                        Do
                        
                            DoEvents
                        
                            sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                            Select Case sts
                                Case BtNoErr
                                
                                    If StrConv(Y_SYU_SUMREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                                    Text(ptxSYUKA_MM).Text & _
                                                                                    Text(ptxSYUKA_DD).Text) Then
                                        Exit Do
                                    End If
                                
                                    
                                    '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                                    If Trim(Text(ptxINS_BIN).Text) <> "" Then
                                        If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                                            Exit Do
                                        End If
                                    End If
                                    '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                                
                                Case BtErrEOF
                                    Exit Do
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�o�ח\��(���PC�o�ɕ\�p)")
                                    Unload Me
                            End Select
                        
                        
                        
                            sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K2_Y_SYU_SUM, Len(K2_Y_SYU_SUM), 2)
                            If sts Then
                                Call File_Error(sts, BtOpDelete, "�o�ח\��(���PC�o�ɕ\�p)")
                                Unload Me
                            End If
                        
                            com = BtOpGetNext
                        
                        Loop
                    
                                            
                    
                    
                    
                    
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(���PC�o�ɕ\�p)")
                        Unload Me
                End Select
            End If
            
            
            
            
            
            
            
            Beep
            yn = MsgBox("�u�o�ɕ\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                Input_Cnt = 0
                Output_Cnt = 0
                
                
                
                If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                    If Data_Make_Proc() Then
                        Unload Me
                    End If
                Else
                    If RE_Data_Make_Proc() Then
                        Unload Me
                    End If
                End If
                
                If Input_Cnt = 0 Then
                
                    MsgBox "�u�o�ɕ\����v�o�ɑΏۂ��L��܂���ł���� "
                Else
                
                    Input_Cnt = 0
                    Output_Cnt = 0
                    
                    
                                    
                    
                    
                    If Print_Proc() Then
                        Unload Me
                    End If
                
                
                
                    If Output_Cnt = 0 Then
                        MsgBox "�u�o�ɕ\����v�o�ɑΏۂ��L��܂���ł���� "
                    Else
                        MsgBox "�u�o�ɕ\����v����I�����܂����B"
                
                    End If
                
                End If
            
            
            End If
            
            Combo(pcmbPRINT_KBN).SetFocus
                    
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

'
Private Sub Form_Load()

Dim c   As String * 128
Dim i   As Integer
     
Dim sBuffer As String

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        WS_NO = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        WS_NO = "???"
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

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030551.Caption = "���o�b�����i�ԕʏo�ɕ\����i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
                                
                                '���׉��z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", "SYS", c) Then
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
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
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
                                '�o�ח\��(νĲҰ��)�t�@�C���n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��(���PC�o�ɕ\�p)�t�@�C���n�o�d�m
    If Y_SYU_SUM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1030551.FontName
        .Size = 14
        .Bold = False
    End With
                                
                                '����t�H���g(��)�ݒ�   2007.03.29
    With SmallFont
        .NAME = F1030551.FontName
        .Size = 10
    End With
                                
                                '����t�H���g(��)�ݒ�   2007.03.29
    With BoldFont
        .NAME = F1030551.FontName
        .Size = 14
        .Bold = True
    End With
                                
                                
                                
    With RargeFont
        .NAME = F1030551.FontName
        .Size = 18
        .Bold = False
    End With
                                
                                
                                
                                '����t�H���g�ݒ�i�o�[�R�[�h�j
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
                                
                                '��ʏ����ݒ�
    
    
    Combo(pcmbPRINT_KBN).AddItem Print_KBN0 & "   " & Print_KBN_SIN
    Combo(pcmbPRINT_KBN).AddItem Print_KBN1 & "   " & Print_KBN_SAI
    Combo(pcmbPRINT_KBN).ListIndex = 0





    Text(ptxPRINT_KBN).Text = "0"


    Text(ptxSYUKA_YY).Text = Left(Format(Date, "yyyymmdd"), 4)
    Text(ptxSYUKA_MM).Text = Mid(Format(Date, "yyyymmdd"), 5, 2)
    Text(ptxSYUKA_DD).Text = Right(Format(Date, "yyyymmdd"), 2)


    Combo(pcmbPRINT_KBN).SetFocus
    
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
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
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
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030551 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1030551.Caption = "���o�b�����i�ԕʏo�ɕ\����i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
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


Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���o�b�p�o�ɕ\�f�[�^�쐬����(�V�K����)
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim svJGYOBU    As String
Dim svNAIGAI    As String
Dim svHIN_NO    As String
    
Dim Y_SURYO     As Long
Dim J_SURYO     As Long
Dim DATA_CNT    As Integer
    
Dim SKIP_F      As Boolean
    
Dim ID_NO       As String * 12
    
    Data_Make_Proc = True
                                
    Call Input_Lock
                                
                                
                                            '�o�ɕ\�f�[�^�b�k�n�r�d
''    sts = BTRV(BtOpClose, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
''    If sts Then
''        If sts <> BtErrNoOpen Then
''            Call File_Error(sts, BtOpClose, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
''            Call Input_UnLock
''            Exit Function
''        End If
''    End If
''
''                                            '�o�ɕ\�f�[�^�n�o�d�m
''    If Y_SYU_SUM_Open(BtOpenNomal, WS_NO) Then
''        Call Input_UnLock
''        Exit Function
''    End If
                                            
                                            '�O��l�ر�
    
    Call UniCode_Conv(K3_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K3_Y_SYU_SUM.SYU_NO, "")
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
    
        Select Case sts
            Case BtNoErr
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Call Input_UnLock
                Exit Function
        End Select






        sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
            Call Input_UnLock
            Exit Function
        End If
    
        com = BtOpGetNext
    Loop
    
    '-----------------------------------    �o�ח\��νĲҰ���ް��ŏW�v�����J�n
    svJGYOBU = ""
    svNAIGAI = ""
    svHIN_NO = ""
    
    Y_SURYO = 0
    J_SURYO = 0
    DATA_CNT = 0
    
    
    Input_Cnt = 0
    Output_Cnt = 0
    
    Call UniCode_Conv(K5_Y_SYU_H.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K5_Y_SYU_H.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))
    
    Call UniCode_Conv(K5_Y_SYU_H.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K5_Y_SYU_H.NAIGAI, NAIGAI_NAI)
    
    Call UniCode_Conv(K5_Y_SYU_H.HIN_NO, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K5_Y_SYU_H, Len(K5_Y_SYU_H), 5)
    
        Select Case sts
            Case BtNoErr
            
            
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
            
            
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                Text(ptxSYUKA_MM).Text & _
                                                                Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Call Input_UnLock
                Exit Function
        End Select
            
        Input_Cnt = Input_Cnt + 1
    
        SKIP_F = False
        
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_F = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Call Input_UnLock
                Exit Function
        End Select
    
            
        If Not SKIP_F Then
    
            If Trim(svJGYOBU) = "" Then
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                sts = Den_No_Set_Proc(33, Last_JGYOBU, ID_NO)
                If sts Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
            
            End If
        
            If svJGYOBU <> StrConv(Y_SYU_HREC.JGYOBU, vbUnicode) Or _
                svNAIGAI <> StrConv(Y_SYU_HREC.NAIGAI, vbUnicode) Or _
                svHIN_NO <> StrConv(Y_SYU_HREC.HIN_NO, vbUnicode) Then
            
                If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, ID_NO) Then
                    Call Input_UnLock
                    Exit Function
                End If
            
                
                sts = Den_No_Set_Proc(33, Last_JGYOBU, ID_NO)
                If sts Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                Y_SURYO = 0
                J_SURYO = 0
                DATA_CNT = 0
            End If
    
    
        
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) <> "1" Then
            
                Y_SURYO = Y_SURYO + CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
                J_SURYO = J_SURYO + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
                DATA_CNT = DATA_CNT + 1
            
            End If
        
        
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts <> BtNoErr Then
        
                Call File_Error(sts, BtOpUpdate, "�o�ח\��f�[�^")
                Call Input_UnLock
                Exit Function
            End If
        
            Call UniCode_Conv(Y_SYU_HREC.SYU_NO, ID_NO)
            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K5_Y_SYU_H, Len(K5_Y_SYU_H), 5)
            If sts <> BtNoErr Then
        
                Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)�f�[�^")
                Call Input_UnLock
                Exit Function
            End If
        End If
    
        com = BtOpGetNext
    
    
    Loop
    
    If Trim(svJGYOBU) <> "" Then
        If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, ID_NO) Then
            Call Input_UnLock
            Exit Function
        End If
    End If
    
    
    Call Input_UnLock
    
    Data_Make_Proc = False
                                

End Function
Private Function RE_Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���o�b�p�o�ɕ\�f�[�^�쐬����(�Ĉ������)
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
    
    
Dim svID_NO     As String
Dim svJGYOBU    As String
Dim svNAIGAI    As String
Dim svHIN_NO    As String
    
    
Dim Y_SURYO     As Long
Dim J_SURYO     As Long
Dim DATA_CNT    As Integer
    
Dim SKIP_F      As Boolean
    
    RE_Data_Make_Proc = True
                                
    Call Input_Lock
                                
                                
                                            '�o�ɕ\�f�[�^�b�k�n�r�d
''    sts = BTRV(BtOpClose, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
''    If sts Then
''        If sts <> BtErrNoOpen Then
''            Call File_Error(sts, BtOpClose, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
''            Call Input_UnLock
''            Exit Function
''        End If
''    End If
''
''                                            '�o�ɕ\�f�[�^�n�o�d�m
''    If Y_SYU_SUM_Open(BtOpenNomal, WS_NO) Then
''        Call Input_UnLock
''        Exit Function
''    End If
                                            
                                            '�O��l�ر�
    Call UniCode_Conv(K3_Y_SYU_SUM.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K3_Y_SYU_SUM.SYU_NO, "")
    com = BtOpGetGreaterEqual

    Do
        DoEvents
        sts = BTRV(com, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
    
        Select Case sts
            Case BtNoErr
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_SUMREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Call Input_UnLock
                Exit Function
        End Select


        sts = BTRV(BtOpDelete, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K3_Y_SYU_SUM, Len(K3_Y_SYU_SUM), 3)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpDelete, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
            Call Input_UnLock
            Exit Function
        End If
    
        com = BtOpGetNext
    
    Loop
    
    '-----------------------------------    �o�ח\��νĲҰ���ް��ŏW�v�����J�n
    
    svID_NO = ""
    
    Y_SURYO = 0
    J_SURYO = 0
    DATA_CNT = 0
    
    
    Call UniCode_Conv(K7_Y_SYU_H.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))
    Call UniCode_Conv(K7_Y_SYU_H.INS_BIN, Text(ptxINS_BIN).Text)
    Call UniCode_Conv(K7_Y_SYU_H.SYU_NO, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K7_Y_SYU_H, Len(K7_Y_SYU_H), 7)
    
        Select Case sts
            Case BtNoErr
            
                If StrConv(Y_SYU_HREC.SYUKA_YMD, vbUnicode) <> (Text(ptxSYUKA_YY).Text & _
                                                                Text(ptxSYUKA_MM).Text & _
                                                                Text(ptxSYUKA_DD).Text) Then
                    Exit Do
                End If
            
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
                If Trim(Text(ptxINS_BIN).Text) <> "" Then
                    If StrConv(Y_SYU_HREC.INS_BIN, vbUnicode) <> Text(ptxINS_BIN).Text Then
                        Exit Do
                    End If
                End If
                '>>>>>>>>>>>>>>>>>>>    �և��@�`�k�k�w���  2012.12.21
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
                Call Input_UnLock
                Exit Function
        End Select
    
        Input_Cnt = Input_Cnt + 1
    
    
        SKIP_F = False
        
        
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(Y_SYU_HREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                SKIP_F = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��f�[�^")
                Call Input_UnLock
                Exit Function
        End Select
    
            
        If Not SKIP_F Then
    
            If Trim(svID_NO) = "" Then
            
                svID_NO = StrConv(Y_SYU_HREC.SYU_NO, vbUnicode)
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
            
            End If
        
            If svID_NO <> StrConv(Y_SYU_HREC.SYU_NO, vbUnicode) Then
            
                If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, svID_NO) Then
                    Call Input_UnLock
                    Exit Function
                End If
            
            
                svID_NO = StrConv(Y_SYU_HREC.SYU_NO, vbUnicode)
            
            
                svJGYOBU = StrConv(Y_SYU_HREC.JGYOBU, vbUnicode)
                svNAIGAI = StrConv(Y_SYU_HREC.NAIGAI, vbUnicode)
                svHIN_NO = StrConv(Y_SYU_HREC.HIN_NO, vbUnicode)
            
                Y_SURYO = 0
                J_SURYO = 0
                DATA_CNT = 0
            End If
    
    
        
            If StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode) <> "1" Then
            
                Y_SURYO = Y_SURYO + CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
                J_SURYO = J_SURYO + CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
            
                DATA_CNT = DATA_CNT + 1
            
            End If
    
        End If
    
        com = BtOpGetNext
    
    
    Loop
    
    If Trim(svID_NO) <> "" Then
        If Data_Output_Proc(svJGYOBU, svNAIGAI, svHIN_NO, Y_SURYO, J_SURYO, DATA_CNT, svID_NO) Then
            Call Input_UnLock
            Exit Function
        End If
    End If
    
    
    Call Input_UnLock
    
    RE_Data_Make_Proc = False
                                

End Function

Private Function Data_Output_Proc(JGYOBU As String, NAIGAI As String, HIN_NO As String, _
                                    Y_SURYO As Long, J_SURYO As Long, DATA_CNT As Integer, ID_NO As String) As Integer
'----------------------------------------------------------------------------
'                   ���o�b�p�o�ɕ\�f�[�^�o�͏���
'----------------------------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim TEMP_QTY        As Long

Dim Betu_LOCATION   As String * 8





    Data_Output_Proc = True





    Call UniCode_Conv(Y_SYU_SUMREC.SYUKA_YMD, (Text(ptxSYUKA_YY).Text & _
                                                Text(ptxSYUKA_MM).Text & _
                                                Text(ptxSYUKA_DD).Text))

    Call UniCode_Conv(Y_SYU_SUMREC.INS_BIN, Text(ptxINS_BIN).Text)

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_NO)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            
            Data_Output_Proc = False
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Function
    End Select

    Call UniCode_Conv(Y_SYU_SUMREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
    Call UniCode_Conv(Y_SYU_SUMREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))

    Call UniCode_Conv(Y_SYU_SUMREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_SYU_SUMREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_SYU_SUMREC.HIN_NO, HIN_NO)

    Call UniCode_Conv(Y_SYU_SUMREC.Y_SURYO, Format(Y_SURYO, "0000000"))
    Call UniCode_Conv(Y_SYU_SUMREC.J_SURYO, Format(J_SURYO, "0000000"))
        
    Call UniCode_Conv(Y_SYU_SUMREC.SYU_NO, ID_NO)

    Call UniCode_Conv(Y_SYU_SUMREC.DATA_CNT, Format(DATA_CNT, "0000"))

    '�W���I�݌ɐ�
    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        SUMI_QTY = 0
        MI_QTY = 0
    Else
        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode))) Then
            Exit Function
        End If
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.ST_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    '�ʒu���I��
    If Tana_Kensaku(Betu_LOCATION) Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_SOKO, Mid(Betu_LOCATION, 1, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_RETU, Mid(Betu_LOCATION, 3, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_REN, Mid(Betu_LOCATION, 5, 2))
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_DAN, Mid(Betu_LOCATION, 7, 2))
    If Trim(Betu_LOCATION) = "" Then
        SUMI_QTY = 0
        MI_QTY = 0
    Else
                                    '�ʒu�I�@�݌ɐ�
        If Zaiko_Syukei_Proc(SUMI_QTY, _
                                MI_QTY, _
                                StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                Betu_LOCATION) Then
            Exit Function
        End If
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.BETU_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    '���i�����݌ɐ�
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_SYOHN_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    TEMP_QTY = SUMI_QTY + MI_QTY
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_NAI_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    TEMP_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
    Call UniCode_Conv(Y_SYU_SUMREC.SYO_ZAIKO_QTY, Format(TEMP_QTY, "0000000"))
    '���בq�ɍ݌�
    If Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                            KASO_NYUKA_SOKO & "01" & "01" & "01") Then
        Exit Function
    End If
    Call UniCode_Conv(Y_SYU_SUMREC.NYU_ZAIKO_QTY, Format(SUMI_QTY + MI_QTY, "0000000"))
    
    Call UniCode_Conv(Y_SYU_SUMREC.INS_NOW, Format(Now, "YYYYMMDDHHMMSS"))
    
    Call UniCode_Conv(Y_SYU_SUMREC.FILLER, "")

    sts = BTRV(BtOpInsert, Y_SYU_SUM_POS, Y_SYU_SUMREC, Len(Y_SYU_SUMREC), K0_Y_SYU_SUM, Len(K0_Y_SYU_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrDuplicates
            
        Case Else
            Call File_Error(sts, BtOpInsert, "�o�ח\��(���PC�o�ɕ\�p)�f�[�^")
            Exit Function
    End Select




    Data_Output_Proc = False

End Function







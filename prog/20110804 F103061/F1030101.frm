VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1030101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�o�׌��i���x�����s"
   ClientHeight    =   7080
   ClientLeft      =   2025
   ClientTop       =   2655
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
   ScaleHeight     =   7080
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Caption         =   "�����w��Ĉ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3465
      TabIndex        =   19
      Top             =   1320
      Width           =   4845
      Begin VB.ListBox List1 
         Height          =   1020
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��@��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   21.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2310
         TabIndex        =   20
         Top             =   2040
         Width           =   2430
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�`�[�ԍ��w��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3465
      TabIndex        =   16
      Top             =   4440
      Width           =   4845
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   1
         Left            =   2730
         MaxLength       =   7
         TabIndex        =   2
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Index           =   0
         Left            =   1365
         MaxLength       =   7
         TabIndex        =   1
         Top             =   480
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��@��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   21.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2310
         TabIndex        =   3
         Top             =   1080
         Width           =   2430
      End
      Begin VB.Label Label1 
         Caption         =   "�`"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2415
         TabIndex        =   18
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "�`�[�ԍ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�V�K��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6600
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
      Index           =   8
      Left            =   7800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
End
Attribute VB_Name = "F1030101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxS_DEN_NO% = 0
Private Const ptxE_DEN_NO% = 1

Private Const Text_Max% = 1

Private Const plstPrint_Now% = 0


'Private Const LAST_UPDATE_DAY$ = "2013.01.23 08:30"
Private Const LAST_UPDATE_DAY$ = "[F103010] 2016.04.26 09:30"


Dim Pri_Name    As Printer

Private wY_SYU_H_POS    As POSBLK
Private wY_SYU_HREC     As Y_SYU_HREC_Tag
Private wK1_Y_SYU_H     As KEY1_Y_SYU_H
Private wK4_Y_SYU_H     As KEY4_Y_SYU_H

Private Function Print_Proc(Mode As Integer, DATA_CNT As Integer) As Integer
'----------------------------------------------------------------------------
'                   �������
'   mode    0:�V�K����
'           1:�Ĉ��
'----------------------------------------------------------------------------

Dim lPrinterHandl   As Long         '���������ق��擾

Dim sts             As Integer
Dim com             As Integer
Dim wkcom           As Integer
Dim ans             As Integer

Dim sEditWK         As String       '�ҏWܰ�
Dim sJis            As String       '�����ϊ�������
Dim vjis            As String
    
Dim SEQ_NO          As Long
    
Dim DEN_SU          As String
    
Dim SKIP_Flg        As Boolean
    
Dim SV_ID_NO        As String * 7
    
Dim NON_PRINT_Flg   As Boolean
    
Dim PRINT_NOW       As String
    
    
    Print_Proc = True
    
    Call Input_Lock
    
        
    PRINT_NOW = Format(Now, "YYYYMMDDHHMMSS")
    
    
'   ����J�n����
    PrinterDriver_Start "���i���x�����s", lPrinterHandl

    SEQ_NO = 0
    SV_ID_NO = ""

    Select Case Mode
        Case 0
        '-------------------------------------  �V�K����w��
            Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
        Case 1
        '-------------------------------------  �Ĉ���w��
            Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text1(ptxS_DEN_NO).Text)
            Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")
    End Select

    com = BtOpGetGreaterEqual

    Do
    
        DoEvents
        
        SKIP_Flg = False
        
        Select Case Mode
            Case 0
            '-------------------------------------  �V�K����w��
            
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                Select Case sts
                    Case BtNoErr
                    
                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                        Exit Function
                End Select
            
            
            Case 1
            '-------------------------------------  �Ĉ���w��
        
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
                        If Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                            Exit Do
                        End If
                                    
'                       2016.04.26 ����������Ƃ���
'                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
'                            SKIP_Flg = True
'                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                        Exit Function
                End Select
        
        
        End Select
        
        NON_PRINT_Flg = False
        If Trim(SV_ID_NO) = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
            NON_PRINT_Flg = True
        End If
        SV_ID_NO = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7)
        
        
        If Not SKIP_Flg Then
    
            If Not NON_PRINT_Flg Then
    
    
        '       STX�w��
                sEditWK = Chr(&H2)
        '       �ް����M�J�n�w��
                sEditWK = sEditWK & Chr(&H1B) & "A"
    '2006.12.19
                sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+220"
            
                '�`�[�ԍ�
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode)
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                '�^�����
                vjis = Kanji_Conv("H", Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)))
                sEditWK = sEditWK & Chr(&H1B) & "H0160" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                '�A��
                If Mode = 0 Then
                    SEQ_NO = SEQ_NO + 1
                    
                    sEditWK = sEditWK & Chr(&H1B) & "H0330" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                    sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(SEQ_NO, "#0")
                
                
                End If
                '�`�[�ԍ��ް����
                sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode) & "*"
                sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) & "*"
                
                '���Ӑ溰��
    '''            sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0100" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode))
                '���Ӑ於(�����)
                vjis = Kanji_Conv("H", StrConv(Trim(Left(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode), 15)), vbWide))
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
    '''            '�`�[�ԍ��ް����
    '''            sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) & "*"
            
                '�ő�`�[�s���̊l��
                Call UniCode_Conv(wK4_Y_SYU_H.ID_NO, Left(Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)), 7) & "99")
                sts = BTRV(BtOpGetLessEqual, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                    
                        If Left(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 7) <> Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
                            DEN_SU = "01"
                        Else
                            DEN_SU = Right(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 2)
                        End If
                    Case BtErrEOF
                        DEN_SU = "01"
                    Case Else
                        Call File_Error(sts, BtOpGetLessEqual, "�o�ח\��")
                        Exit Function
                End Select
                If Not IsNumeric(DEN_SU) Then
                    DEN_SU = "01"
                End If
                sEditWK = sEditWK & Chr(&H1B) & "H0290" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(DEN_SU, "#0")
                vjis = Kanji_Conv("H", "�_")
                sEditWK = sEditWK & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                
                
            
            
    '''            If PRINT_CNT = DATA_CNT Then
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT1"
    '''            Else
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT0"
    '''            End If
                    
                    
                '�����R�[�h�̊m�F
                Call UniCode_Conv(wK1_Y_SYU_H.PRINT_NOW, StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.INS_NOW, StrConv(Y_SYU_HREC.INS_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.DATA_CNT, StrConv(Y_SYU_HREC.DATA_CNT, vbUnicode))
                                
                wkcom = BtOpGetGreater
                                
                Do
                
                    DoEvents
                
                    sts = BTRV(wkcom, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK1_Y_SYU_H, Len(wK1_Y_SYU_H), 1)
                    Select Case sts
                        Case BtNoErr
                            If Mode = 1 Then
                                If Trim(StrConv(wY_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                                    Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                    Exit Do
                                End If
                            End If
                                                
                            If Mode = 0 Then
                                If Trim(StrConv(wY_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                                    Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                    Exit Do
                                End If
                            End If
                            
                            
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> Left(StrConv(wY_SYU_HREC.ID_NO, vbUnicode), 7) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                            Exit Function
                    End Select
                    
                    wkcom = BtOpGetNext
                    
                    
                Loop
                    
                    
                    
                If Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)) <> "" Then
                    sEditWK = sEditWK & Chr(&H1B) & "CT0"
                Else
                    sEditWK = sEditWK & Chr(&H1B) & "CT1"
                
                End If
            
            
        '       �w�薇��
                sEditWK = sEditWK & Chr(&H1B) & "Q1"
        
            
        '       �ް����M�I���w��
                sEditWK = sEditWK & Chr(&H1B) & "Z"
        
        '       ETX�w��
                sEditWK = sEditWK & Chr(&H3)
            
        '       �ް����M
                PrinterDriver_Write lPrinterHandl, sEditWK
            End If
        
            '����ύX�V
            If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
                
                Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, PRINT_NOW)
                
                Do
                    Select Case Mode
                        Case 0
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                        Case 1
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                    End Select
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "�o�ח\��(νĲҰ��)Ͻ�")
                            Exit Function
                    End Select
                Loop
            
                Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
                Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
                Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
            
                com = BtOpGetGreater
            Else
                com = BtOpGetNext
            
            End If
        Else
            com = BtOpGetNext
        End If
        
    Loop




    '����I������
    
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_Proc = False


End Function
Private Function Print_RE_Proc() As Integer
'----------------------------------------------------------------------------
'                   �ŏI����w�����Ĉ������
'----------------------------------------------------------------------------

Dim lPrinterHandl   As Long         '���������ق��擾

Dim sts             As Integer
Dim com             As Integer
Dim wkcom             As Integer
Dim ans             As Integer

Dim sEditWK         As String       '�ҏWܰ�
Dim sJis            As String       '�����ϊ�������
Dim vjis            As String
    
Dim SEQ_NO          As Long
    
Dim DEN_SU          As String
    
Dim SKIP_Flg        As Boolean
    
Dim SV_ID_NO        As String * 7
    
Dim NON_PRINT_Flg   As Boolean
    
Dim SV_PRINT_NOW    As String
    
    Print_RE_Proc = True
    
    Call Input_Lock
    
        
    
    
'   ����J�n����
    PrinterDriver_Start "���i���x�����s", lPrinterHandl

    SEQ_NO = 0
    SV_ID_NO = ""

    '�ŏI��������̊l��
'''    sts = BTRV(BtOpGetLast, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
'''    Select Case sts
'''        Case BtNoErr
'''
'''            If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
'''                Print_RE_Proc = False
'''                Exit Function
'''            End If
'''
'''        Case BtErrEOF
'''            Print_RE_Proc = False
'''            Exit Function
'''        Case Else
'''            Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
'''            Exit Function
'''    End Select
'''
'''    SV_PRINT_NOW = StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)
    
    SV_PRINT_NOW = Format(List1(plstPrint_Now).List(List1(plstPrint_Now).ListIndex), "YYYYMMDDHHMMSS")
    
    
    Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, SV_PRINT_NOW)
    Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
    Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")

    com = BtOpGetGreaterEqual

    Do
    
        DoEvents
        
        SKIP_Flg = False
        
            
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
        Select Case sts
            Case BtNoErr
            
                If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> SV_PRINT_NOW Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                Exit Function
        End Select
        
        NON_PRINT_Flg = False
        If Trim(SV_ID_NO) = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
            NON_PRINT_Flg = True
        End If
        SV_ID_NO = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7)
        
        
        If Not SKIP_Flg Then
    
            If Not NON_PRINT_Flg Then
    
    
        '       STX�w��
                sEditWK = Chr(&H2)
        '       �ް����M�J�n�w��
                sEditWK = sEditWK & Chr(&H1B) & "A"
    '2006.12.19
                sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+220"
            
                '�`�[�ԍ�
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode)
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                '�^�����
                vjis = Kanji_Conv("H", Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)))
                sEditWK = sEditWK & Chr(&H1B) & "H0160" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                '�A��
                SEQ_NO = SEQ_NO + 1
                
                sEditWK = sEditWK & Chr(&H1B) & "H0330" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(SEQ_NO, "#0")
                
                
                '�`�[�ԍ��ް����
                sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode) & "*"
                sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) & "*"
                
                '���Ӑ溰��
    '''            sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0100" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode))
                '���Ӑ於(�����)
                vjis = Kanji_Conv("H", StrConv(Trim(Left(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode), 15)), vbWide))
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
    '''            '�`�[�ԍ��ް����
    '''            sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) & "*"
            
                '�ő�`�[�s���̊l��
                Call UniCode_Conv(wK4_Y_SYU_H.ID_NO, Left(Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)), 7) & "99")
                sts = BTRV(BtOpGetLessEqual, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                    
                        If Left(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 7) <> Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
                            DEN_SU = "01"
                        Else
                            DEN_SU = Right(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 2)
                        End If
                    Case BtErrEOF
                        DEN_SU = "01"
                    Case Else
                        Call File_Error(sts, BtOpGetLessEqual, "�o�ח\��")
                        Exit Function
                End Select
                If Not IsNumeric(DEN_SU) Then
                    DEN_SU = "01"
                End If
                sEditWK = sEditWK & Chr(&H1B) & "H0290" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(DEN_SU, "#0")
                vjis = Kanji_Conv("H", "�_")
                sEditWK = sEditWK & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                
                
            
            
    '''            If PRINT_CNT = DATA_CNT Then
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT1"
    '''            Else
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT0"
    '''            End If
                    
                    
                '�����R�[�h�̊m�F
                Call UniCode_Conv(wK1_Y_SYU_H.PRINT_NOW, StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.INS_NOW, StrConv(Y_SYU_HREC.INS_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.DATA_CNT, StrConv(Y_SYU_HREC.DATA_CNT, vbUnicode))
                                
                wkcom = BtOpGetGreater
                                
                Do
                
                    DoEvents
                
                    sts = BTRV(wkcom, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK1_Y_SYU_H, Len(wK1_Y_SYU_H), 1)
                    Select Case sts
                        Case BtNoErr
                                            
                            If Trim(StrConv(wY_SYU_HREC.PRINT_NOW, vbUnicode)) <> SV_PRINT_NOW Then
                                Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                Exit Do
                            End If
                            
                            
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> Left(StrConv(wY_SYU_HREC.ID_NO, vbUnicode), 7) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                            Exit Function
                    End Select
                    
                    wkcom = BtOpGetNext
                    
                    
                Loop
                    
                    
                    
                If Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)) <> "" Then
                    sEditWK = sEditWK & Chr(&H1B) & "CT0"
                Else
                    sEditWK = sEditWK & Chr(&H1B) & "CT1"
                
                End If
            
            
        '       �w�薇��
                sEditWK = sEditWK & Chr(&H1B) & "Q1"
        
            
        '       �ް����M�I���w��
                sEditWK = sEditWK & Chr(&H1B) & "Z"
        
        '       ETX�w��
                sEditWK = sEditWK & Chr(&H3)
            
        '       �ް����M
                PrinterDriver_Write lPrinterHandl, sEditWK
            End If
        End If
            
        com = BtOpGetNext
        
            
    Loop




    '����I������
    
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_RE_Proc = False


End Function

Private Sub Command_Click(Index As Integer)

Dim sts         As Integer
Dim i           As Integer
Dim Tana_Cnt    As Long
Dim Yn          As Integer
    
    
    
    Select Case Index
        
        
        
        
        Case 11                             '�u�I���v
            Unload Me
        Case Else
            Beep
    End Select
    
    Exit Sub
    
    
    
End Sub


Private Sub Command1_Click(Index As Integer)
'----------------------------------------------------------------------------
'                   �V�K����̎w��
'
'----------------------------------------------------------------------------



Dim DATA_CNT    As Integer
Dim Yn          As Integer


 '''           DATA_CNT = Print_Cnt_Proc(0)
 '''           If DATA_CNT < 0 Then
 '''               Unload Me
 '''           End If
        
 '''           Yn = MsgBox("���i���x���́u" & StrConv(Format(DATA_CNT, "#,##0"), vbWide) & "�v�����s����܂��B�X�����ł����H", vbYesNo, "�m�F����")
        
            
    Select Case Index
        Case 0
            
            Yn = MsgBox("�u���i���x���v�V�K������s���܂����H", vbYesNo, "�m�F����")
        
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_Proc(0, DATA_CNT) Then
                    Unload Me
                End If
        
        
        
            End If

        Case 1

            If List1(plstPrint_Now).ListIndex < 0 Then
                MsgBox "�w��s��I�����Ă�������"
                
                List1(plstPrint_Now).SetFocus
                List1(plstPrint_Now).ListIndex = 0
                
                Exit Sub
            End If
            
            
            
            'Yn = MsgBox("�u���i���x���v�ŏI����w�����Ĉ�����s���܂����H", vbYesNo, "�m�F����")
            Yn = MsgBox("�u���i���x���v�����w�蕪�@�Ĉ�����s���܂����H", vbYesNo, "�m�F����")          '2012.12.27 �C��    M.T
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_RE_Proc() Then
                    Unload Me
                End If
        
            Else
            
        
            End If

    End Select

ErrHandler:

End Sub

Private Sub Command2_Click()
Dim DATA_CNT    As Integer
Dim Yn          As Integer


'''            DATA_CNT = Print_Cnt_Proc(1)
'''            If DATA_CNT < 0 Then
'''                Unload Me
'''            End If
        
'''            Yn = MsgBox("���i���x���́u" & StrConv(Format(DATA_CNT, "#,##0"), vbWide) & "�v�����s����܂��B�X�����ł����H", vbYesNo, "�m�F����")
            
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.12.27  �ǉ�    M.T
            If Trim(Text1(ptxS_DEN_NO)) = "" And Trim(Text1(ptxE_DEN_NO)) = "" Then
                MsgBox "�J�n�`�[�ԍ��̎w�肪�s���@���@����s�\!", vbExclamation
                Text1(ptxS_DEN_NO).SetFocus
                Call Text1_GotFocus(ptxS_DEN_NO)
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_DEN_NO)) = "" Then
                Text1(ptxE_DEN_NO) = Text1(ptxS_DEN_NO)
            End If
            
            If Trim(Text1(ptxS_DEN_NO)) > Trim(Text1(ptxE_DEN_NO)) Then
                MsgBox "�`�[�ԍ��̎w�肪�s���@���@����s�\�I", vbExclamation
                Text1(ptxE_DEN_NO).SetFocus
                Call Text1_GotFocus(ptxE_DEN_NO)
                Exit Sub
            End If
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>     �����܂�
            
            Yn = MsgBox("�u���i���x���v�Ĉ�����s���܂����H", vbYesNo, "�m�F����")
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_Proc(1, DATA_CNT) Then
                    Unload Me
                End If
                
                
                Text1(ptxE_DEN_NO) = ""
            Else
            
                Text1(ptxS_DEN_NO).SetFocus                     '2012.12.27  �ǉ�    M.T
                Call Text1_GotFocus(ptxS_DEN_NO)                '2012.12.27  �ǉ�    M.T
            
            End If

ErrHandler:

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
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    F1030101.Caption = F1030101.Caption & LAST_UPDATE_DAY           '2016.04.26
    

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    
    
                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    
                                
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
                                '�o�ח\��(νĲҰ��)�n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
                                '�o�ח\��(νĲҰ��)�n�o�d�m
    If wY_SYU_H_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
                                
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Exit For
        End If
    Next
                                    
    '����ς݂���荞��
    If Print_Re_Set_Proc() Then
        Unload Me
    End If
                                
                                
    Command1(0).SetFocus
End Sub

Private Sub Form_Unload(CANCEL As Integer)

Dim sts         As Integer
Dim Wk_Printer  As Printer
                                            
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Pri_Name.DeviceName) Then
            SetWindowsDefaultPrinter Wk_Printer.DeviceName, Wk_Printer.DriverName, Wk_Printer.Port
            Exit For
        End If
    Next
                                            
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��(νĲҰ��)")
        End If
    End If
                                            '�o�ח\��(νĲҰ��)�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��(νĲҰ��)")
        End If
    End If
    
                                            '�o�ח\��(νĲҰ��)�b�k�n�r�d
    sts = BTRV(BtOpClose, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��(νĲҰ��)")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1030101 = Nothing


    End
End Sub
Private Function Print_Cnt_Proc(Mode As Integer) As Long
'----------------------------------------------------------------------------
'                   ��������̃J�E���g
'   mode    0:�V�K�����
'           1:�Ĉ��
'
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim DATA_CNT    As Long


    Print_Cnt_Proc = True

    DATA_CNT = 0



    Select Case Mode
        Case 0
        '---------------------------------------------�V�K��
            Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")


            com = BtOpGetGreaterEqual
        
        
            Do
            
                DoEvents
                
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                Select Case sts
                    Case BtNoErr
                                            
                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                        Exit Function
                End Select
            
                DATA_CNT = DATA_CNT + 1
            
                com = BtOpGetNext
            
            Loop



        Case 1
        '---------------------------------------------�Ĉ��
    
            Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Trim(Text1(ptxS_DEN_NO).Text))
            Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")


            com = BtOpGetGreaterEqual
        
        
            Do
            
                DoEvents
                
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                Select Case sts
                    Case BtNoErr
                                            
                        If Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                        Exit Function
                End Select
            
                DATA_CNT = DATA_CNT + 1
            
                com = BtOpGetNext
            
            Loop
    
    
    
    End Select











    Print_Cnt_Proc = DATA_CNT

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1030101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030101)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030101)


    F1030101.MousePointer = vbDefault

End Sub


Private Function isWindowsNT() As Boolean
  isWindowsNT = IIf(GetVersion() And &H80000000, False, True)
End Function
Private Sub SetWindowsDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal Port As String)
  Dim param As String
  param = DeviceName & "," & DriverName & "," & Port
  WriteProfileString "windows", "device", param
  If isWindowsNT Then
    'Windows NT/2000
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal 0&
  Else
    'Windows 95/98/Me
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal "windows"
  End If
'  Printer.EndDoc
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i       As Integer

Dim W_Str   As String                                               ' 2012.12.27  �ǉ�    M.T

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.12.27  �ǉ�    M.T
    Select Case Index
        Case ptxS_DEN_NO        '�J�n�`�[�ԍ�
            If Not IsNumeric(Text1(Index)) Then
                MsgBox "���l�G���[�I", vbExclamation
                Text1(Index).SetFocus
                Call Text1_GotFocus(Index)
                Exit Sub
            End If
            
            Call Numeric_Check(EDIT_ONLY, Text1(Index).MaxLength, 0, NEGA_DIS, ZSUP_DIS, COMA_DIS, Text1(Index), W_Str)
            Text1(Index) = W_Str
'2013.01.24            If Trim(Text1(ptxE_DEN_NO)) = "" Then
                Text1(ptxE_DEN_NO) = W_Str
'2013.01.24            End If
        Case ptxE_DEN_NO        '�I���`�[�ԍ�
            Call Numeric_Check(EDIT_ONLY, Text1(Index).MaxLength, 0, NEGA_DIS, ZSUP_DIS, COMA_DIS, Text1(Index), W_Str)
            Text1(Index) = W_Str
        
        Case Else
        
        
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>     �����܂�
    Call Tab_Ctrl(Shift)     '�ړ�


End Sub
Function Kanji_Conv(psPara As String, psSiftJis As String) As String
Rem ���JIS���ނ���JIS���ނ֕ϊ�
'   psPara      :   H=HEX   B=Binary
'   psSiftJis   :   ���JIS����

Dim i As Integer    '���������ݺ���
Dim vConv           'ܰ��ϐ�
Dim vHex            '4�޲Ă̼��JIS���ނɕϊ������ݺ���
Dim vUpByte         '���2�޲Ă�1�޲Ăɕϊ������ݺ���
Dim vDownByte       '����2�޲Ă�1�޲Ăɕϊ������ݺ���
    
    vConv = ""                                    'ܰ��ϐ��̏�����
    For i = 1 To Len(psSiftJis)                   '�������J��Ԃ�
        vHex = Hex(Asc(Mid$(psSiftJis, i, 1)))    '�S�޲Ă̼��JIS���ނɕϊ�
        If vHex = "20" Then
           Exit For
        End If
        vUpByte = Val("&h" + Mid$(vHex, 1, 2))    '��ʂQ�޲Ă��P�޲Ăɕϊ�
        vDownByte = Val("&h" + Mid$(vHex, 3, 2))  '���ʂQ�޲Ă��P�޲Ăɕϊ�
        If vUpByte >= &HE0 Then                   '��ʂP�޲Ă��d�Oh�̏ꍇ�̏���
           vUpByte = vUpByte - &H40
        End If
        vUpByte = (vUpByte - &H81) * 2 + &H21
        If vDownByte > &H7F Then                  '���ʂP�޲Ă��W�Oh�ȏ�̏���
           vDownByte = vDownByte - 1
        End If
        If vDownByte > &H9D Then                  '���ʂP�޲Ă��X�dh�ȏ�̏���
           vUpByte = vUpByte + 1
           vDownByte = vDownByte - (&H9E - &H21)
        Else
           vDownByte = vDownByte - (&H40 - &H21)  '���ʂP�޲Ă��X�c�ȉ��̏���
        End If
        Select Case psPara
               Case "H"
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ܰ��ϐ��ɑ�������
               Case "B"
                    vConv = vConv + Chr$(vUpByte) + Chr$(vDownByte)  'ܰ��ϐ��ɑ�������
               Case Else
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ܰ��ϐ��ɑ�������
        End Select
    Next i
    Kanji_Conv = vConv

End Function

Function wY_SYU_H_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              �o�ח\��(νĲҰ��)�f�[�^  �n�o�d�m
'*
'*      ��  ��:Open Mode(Btrieve�Q��)
'*      �߂�l:false ����
'*             true  �ُ�
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    wY_SYU_H_Open = True
                                            '�o�ח\��f�[�^�t���p�X�捞��
    sts = GetIni("FILE", Y_SYU_H_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU_H]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "�o�ח\��(νĲҰ��)�f�[�^")
                Exit Function
        End Select
    Loop
    wY_SYU_H_Open = False
End Function


Private Function Print_Re_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �Ĉ���I��p�̓��t�Z�b�g
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim svPRINT_NOW     As String * 14


    Print_Re_Set_Proc = True

    
    Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "19900101000000")
    Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
    Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
    
    com = BtOpGetGreater

    List1(plstPrint_Now).Clear

    svPRINT_NOW = ""
    
    Do
    
        DoEvents
        
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
        Select Case sts
            Case BtNoErr
                                    
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�o�ח\��(νĲҰ��)")
                Exit Function
        End Select
    
        If Trim(svPRINT_NOW) = "" Then
            svPRINT_NOW = StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)
        
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "zzzzzzzzzzzzzz")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "zzzzz")
        
            
            List1(plstPrint_Now).AddItem Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 7, 2) & " " & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 9, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 11, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 13, 2)

        End If
    
        If svPRINT_NOW <> StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode) Then
        
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "zzzzzzzzzzzzzz")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "zzzzz")
        
            List1(plstPrint_Now).AddItem Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 7, 2) & " " & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 9, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 11, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 13, 2)
        
        End If
    
        com = BtOpGetGreater
    
    Loop

    If List1(plstPrint_Now).ListCount = 0 Then
        List1(plstPrint_Now).AddItem "�Ĉ���Ώۖ���"
        Frame2.Enabled = False
    End If


    Print_Re_Set_Proc = False

End Function

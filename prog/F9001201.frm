VERSION 5.00
Begin VB.Form F9001201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ޗ��ݼޗp���Ɍ��i�[���"
   ClientHeight    =   3465
   ClientLeft      =   2025
   ClientTop       =   2940
   ClientWidth     =   11505
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
   ScaleHeight     =   3465
   ScaleWidth      =   11505
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   375
      Left            =   3720
      OLEDragMode     =   1  '����
      OLEDropMode     =   1  '�蓮
      TabIndex        =   37
      Top             =   240
      Width           =   6975
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   35
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   33
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6510
      MaxLength       =   2
      TabIndex        =   31
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   5985
      MaxLength       =   2
      TabIndex        =   29
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   27
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4725
      MaxLength       =   2
      TabIndex        =   25
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   23
      Top             =   840
      Width           =   405
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2565
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   21
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�p���I��"
      Height          =   975
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   7965
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1320
      Width           =   732
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   0
      Top             =   840
      Width           =   405
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
      Left            =   10425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   8745
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   7905
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   4905
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
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
      Index           =   4
      Left            =   4065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   1905
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
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
      Index           =   1
      Left            =   1065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�V �K"
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
      Left            =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "̧�ٖ�"
      Height          =   255
      Index           =   8
      Left            =   2880
      TabIndex        =   36
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7455
      TabIndex        =   34
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6930
      TabIndex        =   32
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6405
      TabIndex        =   30
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   3
      Left            =   5670
      TabIndex        =   28
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   5145
      TabIndex        =   26
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   24
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   4095
      TabIndex        =   22
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�������v"
      Height          =   255
      Index           =   12
      Left            =   6885
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I�Ԕ͈�"
      Height          =   255
      Index           =   4
      Left            =   2625
      TabIndex        =   15
      Top             =   960
      Width           =   1095
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
      Left            =   225
      TabIndex        =   14
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9001201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '����t�H���g
Dim Code39Font As New StdFont           '����t�H���g


Private Type Print_tbl_tag              '����p�e�[�u��
    NAIGAI          As String * 2
    HIN_GAI         As String * 20
    HIN_NAI         As String * 13
    HIN_NAME        As String
    IRI_QTY         As String * 8
    ST_SOKO         As String * 2
    ST_SOKO_NAME    As String * 5
    ST_RETU         As String * 2
    ST_REN          As String * 2
    ST_DAN          As String * 2
    BIKOU           As String
    GENSAN          As String * 22
    SHIIRE_WORK_CENTER As _
                       String * 8
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag
 


Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer


Dim Err_Log_F   As String


Dim GENSANKOKU  As Boolean

Private Const Update_day$ = "[F900120] 2012.06.04 16:45"


Private Function Print_Proc(Tanaban As String, NAIGAI As String, HIN_GAI As String, Nyuka_DT As String, Qty As String, Gyo As Integer, Retu As Integer, BIKOU As String) As Integer

Dim Maisu       As Integer
Dim sts         As Integer
Dim flg         As Boolean





    Print_Proc = False
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
    flg = False
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            flg = True
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
        
            
            
    If NAIGAI = NAIGAI_NAI Then
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
    Else
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
    End If
    Print_tbl(Gyo, Retu).HIN_GAI = HIN_GAI
    If Not flg Then
        Print_tbl(Gyo, Retu).HIN_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
        Print_tbl(Gyo, Retu).HIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Print_tbl(Gyo, Retu).ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        Print_tbl(Gyo, Retu).ST_RETU = StrConv(ITEMREC.ST_RETU, vbUnicode)
        Print_tbl(Gyo, Retu).ST_REN = StrConv(ITEMREC.ST_REN, vbUnicode)
        Print_tbl(Gyo, Retu).ST_DAN = StrConv(ITEMREC.ST_DAN, vbUnicode)

        Print_tbl(Gyo, Retu).GENSAN = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)
        Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode)

        Call UniCode_Conv(K0_SOKO.Soko_No, Mid(Tanaban, 1, 2))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = Left(StrConv(SOKOREC.SOKO_NAME, vbUnicode), 5)
            Case BtErrKeyNotFound
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                Beep
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                Unload Me
        End Select
    
    
    
    
    Else
        Print_tbl(Gyo, Retu).HIN_NAI = " "
        Print_tbl(Gyo, Retu).HIN_NAME = " "
        Print_tbl(Gyo, Retu).ST_SOKO = " "
        
        Print_tbl(Gyo, Retu).ST_RETU = " "
        Print_tbl(Gyo, Retu).ST_REN = " "
        Print_tbl(Gyo, Retu).ST_DAN = " "
        Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
    
        Print_tbl(Gyo, Retu).GENSAN = ""
        Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = ""
    
    
    End If

    Print_tbl(Gyo, Retu).IRI_QTY = Qty
    Print_tbl(Gyo, Retu).BIKOU = BIKOU


    
    
        
End Function
                                    
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 4
        Text(i).Text = ""
    Next i
    Text(8).Text = ""

    Text(9).Text = "0"
    Text(10).Text = "0"
End Sub





Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim sts             As Integer
Dim i               As Integer




Select Case Index
        
        
        
        
        
        Case 0, 1
        
            If Data_Make_Proc(Index) Then
                Unload Me
            End If
        
        
        
        
        
        
        Case 4
        
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
        
            Beep
            yn = MsgBox("�����������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                sts = Maisu_keisan_Proc()
                
                If sts Then
                    Unload Me
                End If
        
            End If
        
            Text(0).SetFocus
        
        
        Case 8                              '���
            
            
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
            
            
            
            
            
            Beep
            yn = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                
                
                
                
                If Print_Main_Proc() Then
                    Unload Me
                End If
                
                
                Printer.EndDoc
            
            
            End If
            
            Text(0).SetFocus
            
        Case 11                             '�I��
            Unload Me
            
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Form_DblClick()
    PrintForm
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
Dim Pri_Name    As Printer
Dim DEF         As String
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�ޗ��ݼޗp���Ɍ��i�[�������", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
    
    
                                '���O�t�@�C������荞��
    If GetIni(App.EXEName, "ERR_LOG", App.EXEName, c) Then
        Err_Log_F = LOG_F
    Else
        Err_Log_F = Trim(c)
    End If
    
    
    
                                '���Y���}�X�^�X�V�L��
    If GetIni(App.EXEName, "GENSANKOKU", App.EXEName, c) Then
        GENSANKOKU = False
    Else
        If Trim(c) = "1" Then
            GENSANKOKU = True
        Else
            GENSANKOKU = False
        End If
    End If
    
    
    
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        If JGYOBU_T(i).CODE = SHIZAI Then
        Else
            SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
    
            If JGYOBU_T(i).CODE = Last_JGYOBU Then
                F9001201.Caption = "�ޗ��ݼޗp���Ɍ��i�[���(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
                SubMenu(i).Checked = True
                LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
            Else
                SubMenu(i).Checked = False
            End If
        End If
    Next i

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenRead) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '���Y���n�o�d�m
    If GENSAN_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�I�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
                                '�i�ԁ|�I�n�o�d�m
    If ITEM_LOC_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '�f�t�H���g�p���T�C�Y��荞��
    If GetIni(App.EXEName, "DEF", App.EXEName, c) Then
        c = ""
    End If
    DEF = RTrim(c)
                                
                                '����t�H���g�ݒ�
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F9001201.FontName
        .Size = F9001201.FontSize
    End With
    Set Printer.Font = NormalFont
                                
                                '��ʏ����ݒ�
    
    If DEF = Trim(Option1(0).Caption) Then
        Option1(0).Value = True
        Option1(1).Value = False
    Else
        If DEF = Trim(Option1(1).Caption) Then
            Option1(0).Value = False
            Option1(1).Value = True
        Else
            Option1(0).Value = True
            Option1(1).Value = False
        End If
    End If
    
    Combo1.Clear
    For Each Pri_Name In Printers
        If Trim(Pri_Name.DeviceName) = Trim(Printer.DeviceName) Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    Combo1.ListIndex = 0
    
    
    For Each Pri_Name In Printers
        If Trim(Pri_Name.DeviceName) <> Trim(Combo1.Text) Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    
    Combo1.ListIndex = 0
    
    Text(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
                                            '���Y���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���Y���}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If

    Set F9001201 = Nothing


    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F9001201.Caption = "�ޗ��ݼޗp���Ɍ��i�[���(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
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
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim sts_QTY     As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            Select Case Index
                Case 0, 1, 2, 3
                
                
                
                    For i = 0 To 3
                    
                        If Trim(Text(Index).Text) = "" Then
                        Else
                            Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)
                    
                        End If
                    Next

                
                
                    If Trim(Text(Index).Text) = "" Then
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
                
                Case 4, 5, 6, 7
            
                    
                    
                    For i = 4 To 7
                    
                        If Trim(Text(Index).Text) = "" Then
                        Else
                            Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)
                    
                        End If
                    Next
                   
                    If Trim(Text(Index).Text) = "" Then
                        Text(Index).Text = "zz"
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
            
            
            
            End Select
        
        
            For i = Index + 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        
        
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        
        
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF2
            Command(1).Value = True
        Case vbKeyF5
            Command(4).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub


Private Sub Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
Dim wkGENSAN    As String * 15
                                            
    On Error GoTo Err_Proc
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If


'------------------------------------------------   1�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(20);
        Printer.Print "���Ɍ��i�[";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);
        
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(80);
            Printer.Print "���Ɍ��i�[";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If

'------------------------------------------------   2�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 6
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   3�s��   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 14)) + "*";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 14)) + "*"
        End If
'------------------------------------------------   4�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   5�s��   ------------------
       With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "�i��";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print " " & Left(Print_tbl(Gyo, 0).HIN_GAI, 14);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
            Printer.Print "�i��";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print " " & Left(Print_tbl(Gyo, 1).HIN_GAI, 14);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"
        End If
'------------------------------------------------   6�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   7�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "�i��" & " " & LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "�i��" & " " & LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80)
        End If
'------------------------------------------------   8�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   9�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "�@�@����" & ":";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "���ד�" & ":";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print " ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "�@�@����" & ":";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "���ד�" & ":";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print " "
        End If
'------------------------------------------------   10�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   11�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "�W���I��" & ":" & Print_tbl(Gyo, 0).ST_SOKO & "-" & Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
        Printer.Print Tab(30);
        Printer.Print "�@���l" & ":" & RTrim(LeftB(Print_tbl(Gyo, 0).BIKOU, 40));
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Printer.Print "�W���I��" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
            Printer.Print Tab(88);
            Printer.Print "�@���l" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))
        End If
'------------------------------------------------   12�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
        Printer.Print "�@���Y��" & ":" & wkGENSAN;
        
        
        
        Printer.Print Tab(30);
        Printer.Print "�d����" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            
            wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
            Printer.Print "�@���Y��" & ":" & wkGENSAN;
            
            
            Printer.Print Tab(88);
            Printer.Print "�d����" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;
        End If




'------------------------------------------------   13�s��   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 8
        End With
        Set Printer.Font = NormalFont
        
        Printer.Print
        
        If Gyo <> Max_Gyo Then


            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print
            
            



            If Max_Gyo <> 2 Then
            
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                Printer.Print
            Else
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 4
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
            
            
            End If

        End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "�v�����^�[�G���[���������܂����B"
    Else
        MsgBox "���s���G���[�F" & Err.Number
    End If
End Sub


Private Function Maisu_keisan_Proc() As Integer


Dim com         As Integer
Dim sts         As Integer

Dim svTanaban   As String * 8
Dim SvNaigai    As String * 1
Dim SvHin_Gai   As String * 20
Dim Maisu       As Integer



    Maisu_keisan_Proc = True

    Call Input_Lock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�W�v��", Me.hwnd, 0)





    Call UniCode_Conv(K2_ITEM_LOC.SOKO, Text(0).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Retu, Text(1).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Ren, Text(2).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Dan, Text(3).Text)
    
    Call UniCode_Conv(K2_ITEM_LOC.JGYOBU, "")
    Call UniCode_Conv(K2_ITEM_LOC.NAIGAI, "")
    Call UniCode_Conv(K2_ITEM_LOC.HIN_GAI, "")

    com = BtOpGetGreaterEqual

    Maisu = 0

    SvNaigai = ""


    Do
        DoEvents
    
        sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K2_ITEM_LOC, Len(K2_ITEM_LOC), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEM_LOCREC.SOKO, vbUnicode) & StrConv(ITEM_LOCREC.Retu, vbUnicode) & StrConv(ITEM_LOCREC.Ren, vbUnicode) & StrConv(ITEM_LOCREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�i�ԁ|�I��")
                Exit Function
        End Select
    
    
        Maisu = Maisu + Val(StrConv(ITEM_LOCREC.Print_SU, vbUnicode))
    
        com = BtOpGetNext
    
    
    Loop

    Text(8).Text = Format(Maisu, "#0")


    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�W�v�I��", Me.hwnd, 0)

    Maisu_keisan_Proc = False



End Function


Private Function Print_Main_Proc() As Integer


Dim com             As Integer
Dim sts             As Integer


Dim Wk_Printer      As Printer


Dim Gyo             As Integer


Dim Retu            As Integer

Dim wk_LOOP         As Integer

Dim Tanaban         As String

Dim Fsw             As Boolean

    Print_Main_Proc = True

    Call Input_Lock

'�w�蒠�[�p�v�����^���擾
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Combo1.Text) Then
                Set Printer = Wk_Printer
                Exit For
        End If
    Next

    If Option1(0).Value = True Then
        Printer.PaperSize = vbPRPSA5
        Printer.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��
        Max_Gyo = 2
    Else
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��
        Max_Gyo = 5
    End If


 



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@�����", Me.hwnd, 0)


    For Gyo = 0 To UBound(Print_tbl)
        For Retu = 0 To 1
        
            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
        Next Retu
    Next Gyo

    Gyo = 0
    Retu = 0


    For wk_LOOP = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(wk_LOOP).CODE = Last_JGYOBU Then
            JGYOBU_NAME = JGYOBU_T(wk_LOOP).NAME
            Exit For
        End If
    Next wk_LOOP


    Fsw = True

    Call UniCode_Conv(K2_ITEM_LOC.SOKO, Text(0).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Retu, Text(1).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Ren, Text(2).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Dan, Text(3).Text)
    
    Call UniCode_Conv(K2_ITEM_LOC.JGYOBU, "")
    Call UniCode_Conv(K2_ITEM_LOC.NAIGAI, "")
    Call UniCode_Conv(K2_ITEM_LOC.HIN_GAI, "")

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
        sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K2_ITEM_LOC, Len(K2_ITEM_LOC), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEM_LOCREC.SOKO, vbUnicode) & StrConv(ITEM_LOCREC.Retu, vbUnicode) & StrConv(ITEM_LOCREC.Ren, vbUnicode) & StrConv(ITEM_LOCREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    Exit Do
                End If
            
                Fsw = False
            
            Case BtErrEOF
                
                
                
                
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�i�ԁ|�I��")
                Exit Function
        End Select
                
        For wk_LOOP = 1 To Val(StrConv(ITEM_LOCREC.Print_SU, vbUnicode))
            Tanaban = ""
            If Print_Proc(Tanaban, NAIGAI_NAI, StrConv(ITEM_LOCREC.HIN_GAI, vbUnicode), "", "", Gyo, Retu, StrConv(ITEM_LOCREC.BIKOU, vbUnicode)) Then
                Exit Function
            End If
                
                
            Retu = Retu + 1
            If Retu > 1 Then
                Gyo = Gyo + 1
                If Gyo > Max_Gyo Then
                    Call Print_Sub_Proc
                    Printer.NewPage
                    For Gyo = 0 To Max_Gyo
                        For Retu = 0 To 1
        
                            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
                        Next Retu
                    Next Gyo
        
                    Gyo = 0
                End If
                Retu = 0
            End If
                
                
        
        Next wk_LOOP
    
    
        com = BtOpGetNext
    
    
    Loop

    If Not Fsw Then
        Call Print_Sub_Proc
        Printer.NewPage
    End If


    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "��������@����I��", Me.hwnd, 0)

    Print_Main_Proc = False



End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F9001201.MousePointer = vbHourglass

    Call Ctrl_Lock(F9001201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F9001201)


    F9001201.MousePointer = vbDefault

End Sub

Private Function Data_Make_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �u���i���\��t�@�C���v�Ǎ��ݏ���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    
    
Dim fileName        As String
Dim FileNo          As Long
    

Dim wkBuf           As String
Dim wkText          As Variant

Dim wkDATE          As String * 8

Dim Skip_Flg        As Integer


Dim No              As String * 8       '��
Dim HIN_GAI         As String * 20      '�ΊO�i��
Dim IRI_QTY         As String * 8       '������萔
Dim BIKOU           As String * 20      '������l

Dim Tanaban         As String * 8       '�I��

Dim Print_SU        As String * 8       '�������




    Data_Make_Proc = True

    Call Input_Lock

    FileNo = FreeFile
    fileName = Trim(Text1.Text)
    On Error GoTo Error_Proc

    Open fileName For Input As #FileNo

    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�[�t�@�C���@�o�^�����J�n�I�I", Me.hwnd, 0)

                                    '�e�[�u�����Z�b�g
    If Mode = 0 Then
        com = BtOpGetFirst
        
        Do
            DoEvents
        
            sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�i�ځ|�I�}�X�^")
                    Call Input_UnLock
                    Exit Function
            End Select
        
            sts = BTRV(BtOpDelete, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpDelete, "�i�ځ|�I�}�X�^")
                    Call Input_UnLock
                    Exit Function
            End Select
        
        
            com = BtOpGetNext
        Loop
    
    
    
        If GENSANKOKU Then
    
            com = BtOpGetFirst
            
            Do
                DoEvents
            
                sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "���Y���}�X�^")
                        Call Input_UnLock
                        Exit Function
                End Select
            
                sts = BTRV(BtOpDelete, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpDelete, "���Y���}�X�^")
                        Call Input_UnLock
                        Exit Function
                End Select
            
            
                com = BtOpGetNext
            Loop
        
        End If
    
    
        com = BtOpGetFirst
        
        Do
            DoEvents
        
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�i�ڃ}�X�^")
                    Call Input_UnLock
                    Exit Function
            End Select
        
            Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
            Call UniCode_Conv(ITEMREC.ST_RETU, "**")
            Call UniCode_Conv(ITEMREC.ST_REN, "**")
            Call UniCode_Conv(ITEMREC.ST_DAN, "**")
        
            Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
        
        
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                    Call Input_UnLock
                    Exit Function
            End Select
        
        
            com = BtOpGetNext
        Loop
    
    
    
    End If


    Do Until EOF(FileNo)
        
        
        DoEvents
        
        Line Input #FileNo, wkBuf
    
    
    
    
        wkText = Split(wkBuf, vbTab, -1)
    
    
    
    
        Skip_Flg = False
    
        No = ""                         '��
        HIN_GAI = ""                    '�ΊO�i��
        IRI_QTY = ""                    '������萔
        BIKOU = ""                      '������l

        Tanaban = ""                    '�I��

        Print_SU = ""                   '�������
    
    
    
    
        If UBound(wkText) < 0 Then
            Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
            Skip_Flg = True
        End If
    
    
        Select Case UBound(wkText)
            Case 0
                No = wkText(0)
            Case 1
                No = wkText(0)
                HIN_GAI = wkText(1)
            Case 2
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
            Case 3
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
            Case 4
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
                Tanaban = wkText(4)
                Print_SU = "1"
            Case Else
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
                Tanaban = wkText(4)
                Print_SU = wkText(5)
                If Not IsNumeric(Print_SU) Then
                    Print_SU = "0"
                End If
        End Select
    
        If UBound(wkText) < 4 Then
            If Not Skip_Flg Then
                Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
            End If
            Skip_Flg = True
        End If
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
        Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                If Not Skip_Flg Then
                    Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
                End If
                Skip_Flg = True
            Case Else
                   Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                   Exit Function
        End Select
            
                    
                    
                    
                    
                    
                    
        Call UniCode_Conv(K0_TANA.Soko_No, Mid(Tanaban, 1, 2))
        Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
        Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
        Call UniCode_Conv(K0_TANA.Dan, Mid(Tanaban, 7, 2))
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                If Not Skip_Flg Then
                    Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
                End If
                Skip_Flg = True
            Case Else
                   Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                   Exit Function
        End Select
            
            
        If Not Skip_Flg Then
            Call UniCode_Conv(ITEM_LOCREC.No, No)
            Call UniCode_Conv(ITEM_LOCREC.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(ITEM_LOCREC.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(ITEM_LOCREC.HIN_GAI, HIN_GAI)
            Call UniCode_Conv(ITEM_LOCREC.IRI_QTY, IRI_QTY)
            Call UniCode_Conv(ITEM_LOCREC.BIKOU, BIKOU)
        
        
            Call UniCode_Conv(ITEM_LOCREC.SOKO, Mid(Tanaban, 1, 2))
            Call UniCode_Conv(ITEM_LOCREC.Retu, Mid(Tanaban, 3, 2))
            Call UniCode_Conv(ITEM_LOCREC.Ren, Mid(Tanaban, 5, 2))
            Call UniCode_Conv(ITEM_LOCREC.Dan, Mid(Tanaban, 7, 2))
        
        
            Call UniCode_Conv(ITEM_LOCREC.Print_SU, Format(Val(Print_SU), "00000000"))
            Call UniCode_Conv(ITEM_LOCREC.FILLER, "")
        
            sts = BTRV(BtOpInsert, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrDuplicates
                    sts = BTRV(BtOpUpdate, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "�i�ԁ|�I�}�X�^")
                            Exit Function
                    End Select
                Case Else
                       Call File_Error(sts, BtOpInsert, "�i�ԁ|�I�}�X�^")
                       Exit Function
            End Select
        
        
            If StrConv(ITEMREC.ST_SOKO, vbUnicode) = "**" Then
                Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(Tanaban, 1, 2))
                Call UniCode_Conv(ITEMREC.ST_RETU, Mid(Tanaban, 3, 2))
                Call UniCode_Conv(ITEMREC.ST_REN, Mid(Tanaban, 5, 2))
                Call UniCode_Conv(ITEMREC.ST_DAN, Mid(Tanaban, 7, 2))
        
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
        
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                           Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^")
                           Exit Function
                End Select
        
            End If
        
            If GENSANKOKU Then
        
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> "" Then
            
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(GENSANREC.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(GENSANREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(GENSANREC.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                            Call UniCode_Conv(GENSANREC.FILLER, "")
                            
                            Call UniCode_Conv(GENSANREC.INS_TANTO, App.EXEName)
                            Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                        
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                        
                        
                        
                            sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                            Select Case sts
                                Case BtNoErr
                                Case Else
                                       Call File_Error(sts, BtOpInsert, "���Y���}�X�^")
                                       Exit Function
                            End Select
                        
                        
                        
                        Case Else
                               Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
                               Exit Function
                    End Select
                End If
            End If
        
        End If



    Loop




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "���i�[�t�@�C���@�o�^�����I���I�I", Me.hwnd, 0)



    Call Input_UnLock


    Data_Make_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            Data_Make_Proc = False      '





        Case Else
    End Select
    Call Input_UnLock

End Function


Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Trim(Data.Files(1))

End Sub

Public Sub Err_LOG_Proc(No As String, HIN_GAI As String, IRI_QTY As String, BIKOU As String, Tanaban As String, Print_SU As String)


    Call LOG_OUT(Err_Log_F, No & "," & HIN_GAI & "," & IRI_QTY & "," & BIKOU & "," & Tanaban & "," & Print_SU)



End Sub

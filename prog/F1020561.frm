VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1020561 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�o�׌��i���x�����s ([F102056] 2012.03.23 15:30)"
   ClientHeight    =   4740
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
   ScaleHeight     =   4740
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3990
      MaxLength       =   13
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   2
      Left            =   3990
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1920
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3990
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   2535
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4200
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
      Top             =   4200
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
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
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X �V"
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
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "JAN"
      Height          =   255
      Index           =   2
      Left            =   3465
      TabIndex        =   17
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "���@��"
      Height          =   255
      Index           =   1
      Left            =   3150
      TabIndex        =   16
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "�i�@��"
      Height          =   255
      Index           =   0
      Left            =   3150
      TabIndex        =   15
      Top             =   840
      Width           =   750
   End
End
Attribute VB_Name = "F1020561"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxHIN_NO% = 0
Private Const ptxJAN_CODE% = 1
Private Const ptxMAISU% = 2

Private Const Text_Max% = 2

Dim Pri_Name    As Printer
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   �������
'   mode    0:�V�K����
'           1:�Ĉ��
'----------------------------------------------------------------------------

Dim lPrinterHandl   As Long         '���������ق��擾

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim sEditWK         As String       '�ҏWܰ�
Dim sJis            As String       '�����ϊ�������
Dim vjis            As String
    
    
    Print_Proc = True
    
    '�i�`�m���ލX�V
    If Update_Proc() Then
        Exit Function
    End If
    
    Call Input_Lock
    
'   ����J�n����
    PrinterDriver_Start "�i�ڃ��x�����s", lPrinterHandl
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_NO).Text)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            
            Print_Proc = False
            Call Input_UnLock
            Exit Function
            
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Function
    End Select
            
    '       STX�w��
    sEditWK = Chr(&H2)
    '       �ް����M�J�n�w��
    sEditWK = sEditWK & Chr(&H1B) & "A"
    '2006.12.19
    sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+000"
        
            '�i��
'''    sEditWK = sEditWK & Chr(&H1B) & "H0040" & Chr(&H1B) & "V0040" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
'''    sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    '�i���ް����(JAN13)
'''    sEditWK = sEditWK & Chr(&H1B) & "H0040" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
'''    sEditWK = sEditWK & Chr(&H1B) & "B303100" & Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
''''''    '�i��
''''''    vjis = Kanji_Conv("H", Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)))
''''''    sEditWK = sEditWK & Chr(&H1B) & "H0020" & Chr(&H1B) & "V0100" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
''''''    sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
    
    'JAN����
'''    sEditWK = sEditWK & Chr(&H1B) & "H0040" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
'''    sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
    
    
    '�i��
    sEditWK = sEditWK & Chr(&H1B) & "H0240" & Chr(&H1B) & "V0040" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
    
    
    
    
    
    If Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode)) = "" Then
        '�i���ް����(CODE39)
        sEditWK = sEditWK & Chr(&H1B) & "H0240" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "D101100" & "*" & Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & "*"
    Else
        '�i���ް����(JAN13)
        sEditWK = sEditWK & Chr(&H1B) & "H0240" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "B303100" & Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
    End If
    
    'JAN����
    sEditWK = sEditWK & Chr(&H1B) & "H0240" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(ITEMREC.JAN_CODE, vbUnicode))
    
    
    '       �J�b�g�w��
    sEditWK = sEditWK & Chr(&H1B) & "CT" & Text1(ptxMAISU).Text
    '       �w�薇��
    sEditWK = sEditWK & Chr(&H1B) & "Q" & Text1(ptxMAISU).Text
    
        
    '       �ް����M�I���w��
    sEditWK = sEditWK & Chr(&H1B) & "Z"
    
    '       ETX�w��
    sEditWK = sEditWK & Chr(&H3)
        
    '       �ް����M
    PrinterDriver_Write lPrinterHandl, sEditWK
        
        




    '����I������
    
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_Proc = False


End Function

Private Sub Command_Click(Index As Integer)

Dim sts         As Integer
Dim i           As Integer
Dim Yn          As Integer
    
    
    
    Select Case Index
        
        Case 0
            
'''            If Trim(Text1(ptxJAN_CODE).Text) = "" Then
'''            Else
'''                If Not IsNumeric(Trim(Text1(ptxJAN_CODE).Text)) Then
'''                    MsgBox "���͂������ڂ̓G���[�ł��B(���l�̂�)"
'''                    Text1(ptxJAN_CODE).SetFocus
'''                    Exit Sub
'''                Else
'''                    If Len(Trim(Text1(ptxJAN_CODE).Text)) <> 13 Then
'''                        MsgBox "���͂������ڂ̓G���[�ł��B(13���̂�)"
'''                        Text1(ptxJAN_CODE).SetFocus
'''                        Exit Sub
'''                    End If
'''                End If
'''            End If
            
            
            sts = Err_Check_Proc(1)
            
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            
            Yn = MsgBox("�X�V���܂����H", vbYesNo, "�m�F����")
            If Yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
        Case 8
            
'''            If Trim(Text1(ptxJAN_CODE).Text) = "" Then
'''                MsgBox "���͂������ڂ̓G���[�ł��B(�i�`�m���ݒ�)"
'''                Text1(ptxJAN_CODE).SetFocus
'''                Exit Sub
'''            End If
'''
'''            If Not IsNumeric(Text1(ptxMAISU).Text) Then
'''                Text1(ptxMAISU).Text = "1"
'''            End If
'''
'''            Text1(ptxMAISU).Text = Format(CInt(Text1(ptxMAISU).Text), "#0")
            sts = Err_Check_Proc(0)
            
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            
            
            
            
            Yn = MsgBox("�i�����ق̈�����s���܂����H(������JAN���ނ̍X�V���s���܂�)", vbYesNo, "�m�F����")
            If Yn = vbYes Then
            
'''                CommonDialog1.CancelError = True
'''                On Error GoTo ErrHandler
            
'''                CommonDialog1.ShowPrinter
    
    
                If Print_Proc() Then
                    Unload Me
                End If
        
                Text1(ptxHIN_NO).SetFocus
    
            End If

        
        Case 11                             '�u�I���v
            Unload Me
        Case Else
            Beep
    End Select
    
    Exit Sub
    
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
    

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    
                                '�i��Ͻ��n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    
    If JGYOB_TB_Set(0) Then     '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Exit For
        End If
    Next
                                
                                
                                
    
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
                                            
                                            '�i��Ͻ��b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i��Ͻ�")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1020561 = Nothing


    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1020561.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020561)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020561)


    F1020561.MousePointer = vbDefault

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
    
Dim sts As Integer

Dim i   As Integer


    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If


    Select Case Index
        Case ptxHIN_NO
            
            Text1(ptxHIN_NO).Text = StrConv(Text1(ptxHIN_NO).Text, vbUpperCase)
            
            For i = 0 To UBound(JGYOBU_T)
            
                        
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)

            
                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_NO).Text)


                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
                        Text1(ptxJAN_CODE).Text = StrConv(ITEMREC.JAN_CODE, vbUnicode)
                    
                        Exit For
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Sub
                
                
                
                End Select

            Next i


            If i > UBound(JGYOBU_T) Then
                Text1(ptxJAN_CODE).Text = ""
                MsgBox "���͂������ڂ̓G���[�ł��B(�i�Ԗ��o�^)"
                Text1(Index).SetFocus
                Exit Sub
            End If
            Last_JGYOBU = JGYOBU_T(i).CODE
        Case ptxJAN_CODE

            If Trim(Text1(Index).Text) = "" Then
            Else
                If Not IsNumeric(Trim(Text1(Index).Text)) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B(���l�̂�)"
                    Text1(Index).SetFocus
                    Exit Sub
                Else
                    If Len(Trim(Text1(Index).Text)) <> 13 Then
                        MsgBox "���͂������ڂ̓G���[�ł��B(13���̂�)"
                        Text1(Index).SetFocus
                        Exit Sub
                    End If
                End If
            End If

        Case ptxMAISU
            

    End Select
    Call Tab_Ctrl(Shift)

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

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �X�V����
'----------------------------------------------------------------------------


Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

    
    Update_Proc = True
    
    Call Input_Lock
    
    
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_NO).Text)
    
    sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Update_Proc = False
            Call Input_UnLock
            Exit Function
        
        
        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbCancel + vbQuestion, "�m�F����")
            Call Input_UnLock
            Update_Proc = False
            Exit Function
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Function
    End Select
            
    Call UniCode_Conv(ITEMREC.JAN_CODE, Text1(ptxJAN_CODE).Text)

    sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbCancel + vbQuestion, "�m�F����")
            Call Input_UnLock
            Update_Proc = False
            Exit Function
        
        
        Case Else
            Call File_Error(sts, BtOpUpdate, "�i��Ͻ�")
            Exit Function
    End Select






    Call Input_UnLock
    
    Update_Proc = False


End Function
Private Function Err_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer
            
            
    Err_Check_Proc = True
            
    '�i�ԃ`�F�b�N
    Text1(ptxHIN_NO).Text = StrConv(Text1(ptxHIN_NO).Text, vbUpperCase)
    
    For i = 0 To UBound(JGYOBU_T)
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)
        Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_NO).Text)
    
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            
                Text1(ptxJAN_CODE).Text = StrConv(ITEMREC.JAN_CODE, vbUnicode)
                Exit For
            
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Err_Check_Proc = SYS_ERR
                Exit Function
        End Select
    Next i

    If i > UBound(JGYOBU_T) Then
    
                
        Text1(ptxJAN_CODE).Text = ""
        MsgBox "���͂������ڂ̓G���[�ł��B(�i�Ԗ��o�^)"
        Text1(ptxHIN_NO).SetFocus
        Exit Function
    End If


    Last_JGYOBU = JGYOBU_T(i).CODE

    '�i�`�m�R�[�h�`�F�b�N
    If Trim(Text1(ptxJAN_CODE).Text) = "" Then
    Else
        If Not IsNumeric(Trim(Text1(ptxJAN_CODE).Text)) Then
            MsgBox "���͂������ڂ̓G���[�ł��B(���l�̂�)"
            Text1(ptxJAN_CODE).SetFocus
            Exit Function
        Else
            If Len(Trim(Text1(ptxJAN_CODE).Text)) <> 13 Then
                MsgBox "���͂������ڂ̓G���[�ł��B(13���̂�)"
                Text1(ptxJAN_CODE).SetFocus
                Exit Function
            End If
        End If
    End If

    If Mode = 1 Then
        Err_Check_Proc = False
        Exit Function
    End If

    '�����`�F�b�N
    If Not IsNumeric(Text1(ptxMAISU).Text) Then
        Text1(ptxMAISU).Text = "1"
    End If
    
    Text1(ptxMAISU).Text = Format(CInt(Text1(ptxMAISU).Text), "#0")

    Err_Check_Proc = False


End Function

Private Sub Text1_LostFocus(Index As Integer)

    Select Case Index
    
        Case ptxHIN_NO
            Text1(ptxHIN_NO).Text = StrConv(Text1(ptxHIN_NO).Text, vbUpperCase)
    
    End Select


End Sub

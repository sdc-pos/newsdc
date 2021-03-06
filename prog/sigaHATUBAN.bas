Attribute VB_Name = "HATUBN"
Option Explicit
'********************************************************************
'*
'*              ­Τ}X^@t@Cθ`
'*
'********************************************************************
't@Chc
Public Const HATUBAN_ID$ = "HATUBAN"

'y[WTCY
Public Const HATUBAN_PG_SIZ% = 512

'|WVEubN
Public HATUBAN_POS As POSBLK
'********************************************************************
'*
'*                           \’Μθ`
'*
'********************************************************************
'*************************** ΪΌθ` *****************************
'R[hθ`
Type HATUBANREC_Tag
    JGYOBU(0 To 0)          As Byte         'Ζζͺ
    NYK_KBN(0 To 0)         As Byte         'όΧ`[ζͺ
    NYK_DEN_NO(0 To 4)      As Byte         'όΧ`[
    SYK_KBN(0 To 0)         As Byte         'oΧ`[ζͺ
    SYK_DEN_NO(0 To 4)      As Byte         'oΧ`[
    NYK_ID_KBN(0 To 0)      As Byte         'όΧIDζͺ
    NYK_ID_NO(0 To 7)       As Byte         'όΧID
    SYK_ID_KBN(0 To 0)      As Byte         'oΧIDζͺ
    SYK_ID_NO(0 To 10)      As Byte         'oΧID         2006.05.23 7-->11

    OPC_ID_KBN(0 To 0)      As Byte         'εγPCIDζͺ     2006.12.11
    OPC_ID_NO(0 To 5)       As Byte         'εγPCoΧID   2006.12.11

    OPC_DEN_KBN(0 To 0)     As Byte         'εγPC`[ζͺ   2006.12.11
    OPC_DEN_NO(0 To 5)      As Byte         'εγPC`[       2006.12.11



    FILLER(0 To 31)          As Byte         'FILLER            2006.12.11
End Type

'f[^Eobt@
Public HATUBANREC           As HATUBANREC_Tag

'L[θ`
Type KEY0_HATUBAN            'jdxO
    JGYOBU(0 To 0)          As Byte         'Ζζͺ
End Type

'L[Ef[^
Public K0_HATUBAN           As KEY0_HATUBAN

Type HATUBAN_FSpeck
    fs      As BtFileSpeck                  'Μ§²Ω ½Νί―Έ\’Μ
    ks0     As BtKeySpeck                   '·° ½Νί―Έ\’Μ
End Type

Private HATUBAN_Speck As HATUBAN_FSpeck

Private Function HATUBAN_Create() As Integer
'********************************************************************
'*
'*              ­Τ}X^@bqd`sd
'*
'*      ψ  :Θ΅
'*      ίθl:false ³ν
'*             true  Ων
'*
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    HATUBAN_Create = True
                                            '­Τ}X^tpXζέ
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [HATUBAN]ΗέέG[")
        Exit Function
    End If
    
    FullPath = RTrim(c)

    HATUBAN_Speck.fs.recoleng = Len(HATUBANREC)     ' R[h·
    HATUBAN_Speck.fs.PageSize = HATUBAN_PG_SIZ      ' y[WTCY
    HATUBAN_Speck.fs.idexnumb = 1                   ' CfbNX
    HATUBAN_Speck.fs.fileflag = 0                   ' t@CtO
    HATUBAN_Speck.fs.reserve = &H0                  ' \ρΟέ
                                                    ' L[O
    HATUBAN_Speck.ks0.keypos = 1                    ' L[|WV
    HATUBAN_Speck.ks0.keyleng = 1                   ' L[·
    HATUBAN_Speck.ks0.keyflag = BtKfExt             ' L[tO
    HATUBAN_Speck.ks0.keytype = Chr(BtKtString)     ' L[^Cv
    HATUBAN_Speck.ks0.reserve = &H0                 ' \ρΟέ

    sts = BTRV(BtOpCreate, HATUBAN_POS, HATUBAN_Speck, Len(HATUBAN_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "­Τ}X^")
        Exit Function
    End If

    HATUBAN_Create = False

End Function

Public Function HATUBAN_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ­Τ}X^@nodm
'*
'*      ψ  :Open Mode(BtrieveQΖ)
'*      ίθl:false ³ν
'*             true  Ων
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    HATUBAN_Open = True
                                            '­Τ}X^tpXζέ
    sts = GetIni("FILE", HATUBAN_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [HATUBAN]ΗέέG[")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = HATUBAN_Create()        '­Τ}X^μ¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "­Τ}X^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "­Τ}X^")
                Exit Function
        End Select
    Loop

    HATUBAN_Open = False

End Function

Public Function Den_No_Set_Proc(Mode As Integer, JGYOBU As String, DEN_NO As String, Optional MSG As Integer = 1, Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      uoΧ^oΙ όΧ^όΙ €Κv
'*          vζO`[ `[­Τ
'*          εγob`[ΗΑ  2006.12.11
'*
'*  vζOΜ`[Μζέ
'*  (­Τ}X^ΜOPEN/CLOSEΝΔΡ³Ε)
'*  ψF  [hiΘͺsΒ 10:όΧ`[ 11:όΧeLXg@20:oΧ`[ 21:oΧhc 30:εγPCoΧ`[ 31:εγPCoΧIDj
'*          Ζ(ΘͺsΒ)
'*          `[(ΘͺsΒ)
'*          bZ[W\¦(ΘͺΒ@0:\¦³΅@1:\¦Lθ)
'*          gC(gCρ(0`99 0:³ΐ))
'*  ίθl: false       :³ν
'*          true        :Ων
'*          SYS_CANCEL  :XV·¬έΎΩ
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer
Dim wk_No       As Long
Dim W_Cnt       As Integer
    
Dim NYU_KBN     As String * 1
Dim SYU_KBN     As String * 1

Dim NYU_ID_KBN  As String * 1
Dim SYU_ID_KBN  As String * 1

Dim OPC_ID_KBN  As String * 1
Dim OPC_DEN_KBN As String * 1


Dim c           As String * 128

    
    Den_No_Set_Proc = True
    
    DEN_NO = ""
    W_Cnt = 0
    '*------------------------------------------------------'­Τ}X^Ηέέ
    Call UniCode_Conv(K0_HATUBAN.JGYOBU, JGYOBU)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
                        DoEvents
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
                            DoEvents
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("Ό[Εf[^gpΕ·B<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "mFόΝ")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "­Τ}X^", 0)
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    If com = BtOpInsert Then
                                                            'γPΪΜζͺ
        If GetIni("DEN_KBN", "NYU_DEN_KBN", "SYS", c) Then
            Call Log_Out(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_DEN_KBN", "SYS", c) Then
            Call Log_Out(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_DEN_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "NYU_ID_KBN", "SYS", c) Then
            Call Log_Out(LOG_F, "[SYS.INI] [DEN_KBN] [NYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        NYU_ID_KBN = Trim(c)
        
        If GetIni("DEN_KBN", "SYU_ID_KBN", "SYS", c) Then
            Call Log_Out(LOG_F, "[SYS.INI] [DEN_KBN] [SYU_ID_KBN] READ ERROR")
            Den_No_Set_Proc = SYS_ERR
            Exit Function
        End If

        SYU_ID_KBN = Trim(c)
        
        
        'εγob ΗΑ  2006.12.11
        If GetIni("DEN_KBN", "OSAKA_ID_KBN", "SYS", c) Then
            OPC_ID_KBN = ""
        Else
            OPC_ID_KBN = Trim(c)
        End If

        If GetIni("DEN_KBN", "OSAKA_DEN_KBN", "SYS", c) Then
            OPC_DEN_KBN = ""
        Else
            OPC_DEN_KBN = Trim(c)
        End If


        
        
        
        Call UniCode_Conv(HATUBANREC.JGYOBU, JGYOBU)            'Ζ
        Call UniCode_Conv(HATUBANREC.NYK_KBN, NYU_KBN)          'όΧ`[ζͺ
        Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, "00000")       'όΧ`[
        Call UniCode_Conv(HATUBANREC.SYK_KBN, SYU_KBN)          'oΧ`[ζͺ
        Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, "00000")       'oΧ`[
        
        Call UniCode_Conv(HATUBANREC.NYK_ID_KBN, NYU_ID_KBN)    'όΧhcζͺ
        Call UniCode_Conv(HATUBANREC.NYK_ID_NO, "00000000")     'όΧeLXg
        Call UniCode_Conv(HATUBANREC.SYK_ID_KBN, SYU_ID_KBN)    'oΧhcζͺ
        Call UniCode_Conv(HATUBANREC.SYK_ID_NO, "00000000000")  'oΧhc
        
        
        'εγPC 2006.12.17
        If Trim(OPC_ID_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_ID_KBN, OPC_ID_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, "000000")
        
        End If
        
        If Trim(OPC_DEN_KBN) = "" Then
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, "")
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "")
        Else
            Call UniCode_Conv(HATUBANREC.OPC_DEN_KBN, OPC_DEN_KBN)
            Call UniCode_Conv(HATUBANREC.OPC_DEN_NO, "000000")
        
        End If
        
        
        Call UniCode_Conv(HATUBANREC.FILLER, "")
    End If
    
    Select Case Mode
        Case 10
                                    'όΧ`[
            If StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, Format(wk_No, "00000"))
    
        Case 11
                                    'όΧhc
            If StrConv(HATUBANREC.NYK_ID_NO, vbUnicode) = "99999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.NYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.NYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000")
            Call UniCode_Conv(HATUBANREC.NYK_ID_NO, Format(wk_No, "00000000"))
                                
        Case 20
                                'oΧ`[
            If StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode) = "99999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_DEN_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_KBN, vbUnicode) & Format(wk_No, "00000")
            Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, Format(wk_No, "00000"))
        Case 21
                                    'oΧhc
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "99999999999" Then
                wk_No = 1
            Else
                wk_No = CLng(StrConv(HATUBANREC.SYK_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.SYK_ID_KBN, vbUnicode) & Format(wk_No, "00000000000")
            Call UniCode_Conv(HATUBANREC.SYK_ID_NO, Format(wk_No, "00000000000"))
    
        Case 31
                                    'εγhc
            If StrConv(HATUBANREC.SYK_ID_NO, vbUnicode) = "999999" Then
                wk_No = 1
            Else
                wk_No = Val(StrConv(HATUBANREC.OPC_ID_NO, vbUnicode)) + 1
            End If
        
            DEN_NO = StrConv(HATUBANREC.OPC_ID_KBN, vbUnicode) & Format(wk_No, "000000")
            Call UniCode_Conv(HATUBANREC.OPC_ID_NO, Format(wk_No, "000000"))
    
    
    End Select
    '*------------------------------------------------------'­Τ}X^oΝ
    W_Cnt = 0
    Do
        sts = BTRV(com, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                If MSG = 0 Then
                    If RETRY = 0 Then
                        DoEvents
                    Else
                        W_Cnt = W_Cnt + 1
                        If W_Cnt <= RETRY Then
                            DoEvents
                        Else
                            Den_No_Set_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    End If
                Else
                    Beep
                    ans = MsgBox("Ό[Εf[^gpΕ·B<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "mFόΝ")
                    If ans = vbCancel Then
                        Den_No_Set_Proc = SYS_CANCEL
                        Exit Function
                    End If
                End If
            Case Else
                Call File_Error(sts, com, "­Τ}X^")
                Den_No_Set_Proc = SYS_ERR
                Exit Function
        End Select
    Loop

    Den_No_Set_Proc = False          '³νIΉ

End Function



Attribute VB_Name = "PTANA"
Option Explicit
'********************************************************************
'*                                                                  *
'*              Β pIXgσόt@Ciκt@Cj        *
'*                                                                  *
'*          CREATE 2004.04.23                                       *
'********************************************************************
't@Chc
Public Const PTANA_ID = "PTANA"

'y[WTCY
Public Const PTANA_PG_SIZ% = 512

'|WVEubN
Public PTANA_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           \’Μθ`                             *
'*                                                                  *
'********************************************************************
'*************************** ΪΌθ` *****************************
'R[hθ`
Type PTANAREC_Tag
    Packing_No(0 To 3)  As Byte     'Β 
    Rank(0 To 2)        As Byte     'N
    Page_cnt(0 To 0)    As Byte     'y[W(qΙ)
    SEQ(0 To 4)         As Byte     'SEQΤ
    SOKO_NO01(0 To 1)   As Byte     'qΙP
    RETUREN01(0 To 4)   As Byte     'ρEAP
    SOKO_NO02(0 To 1)   As Byte     'qΙQ
    RETUREN02(0 To 4)   As Byte     'ρEAQ
    SOKO_NO03(0 To 1)   As Byte     'qΙR
    RETUREN03(0 To 4)   As Byte     'ρEAR
    SOKO_NO04(0 To 1)   As Byte     'qΙS
    RETUREN04(0 To 4)   As Byte     'ρEAS
    SOKO_NO05(0 To 1)   As Byte     'qΙT
    RETUREN05(0 To 4)   As Byte     'ρEAT
    SOKO_NO06(0 To 1)   As Byte     'qΙU
    RETUREN06(0 To 4)   As Byte     'ρEAU
    SOKO_NO07(0 To 1)   As Byte     'qΙV
    RETUREN07(0 To 4)   As Byte     'ρEAV
    SOKO_NO08(0 To 1)   As Byte     'qΙW
    RETUREN08(0 To 4)   As Byte     'ρEAW
    SOKO_NO09(0 To 1)   As Byte     'qΙX
    RETUREN09(0 To 4)   As Byte     'ρEAX
    SOKO_NO10(0 To 1)   As Byte     'qΙPO
    RETUREN10(0 To 4)   As Byte     'ρEAPO

End Type
'f[^Eobt@
Public PTANAREC         As PTANAREC_Tag


'L[θ`
Type KEY0_PTANA                     'jdxO
    Packing_No(0 To 3)  As Byte     'Β 
    Rank(0 To 2)        As Byte     'N
    Page_cnt(0 To 0)    As Byte     'y[W(qΙ)
    SEQ(0 To 4)         As Byte     'SEQΤ
End Type
    
'L[Ef[^
Public K0_PTANA         As KEY0_PTANA

Private Type PTANA_FSpeck
    fs  As BtFileSpeck              ' Μ§²Ω ½Νί―Έ\’Μ
    ks0 As BtKeySpeck               ' ·° ½Νί―Έ\’Μ
    ks1 As BtKeySpeck               ' ·° ½Νί―Έ\’Μ
    ks2 As BtKeySpeck               ' ·° ½Νί―Έ\’Μ
    ks3 As BtKeySpeck               ' ·° ½Νί―Έ\’Μ
End Type

Private PTANA_Speck    As PTANA_FSpeck
Private Function PTANA_Create() As Integer
'********************************************************************
'*                                                                  *
'*              Β ΚIXgσόt@C  bqd`sd          *
'*                                                                  *
'*      ψ  :Θ΅                                                 *
'*      ίθl:false ³ν                                           *
'*             true  Ων                                           *
'*                                                                  *
'*          CREATE 2004.04.24                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PTANA_Create = True
                                            'Β ΚIXgσόt@CtpXζέ
    sts = GetIni("FILE", PTANA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI[PTANA] ΗέέG[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PTANA_Speck.fs.recoleng = Len(PTANAREC)         ' R[h·
    PTANA_Speck.fs.PageSize = PTANA_PG_SIZ          ' y[WTCY
    PTANA_Speck.fs.idexnumb = 1                     ' CfbNX
    PTANA_Speck.fs.fileflag = 0                     ' t@CtO
    PTANA_Speck.fs.reserve = &H0                    ' \ρΟέ
                                                    
'---------------------------------------------------' L[O
    PTANA_Speck.ks0.keypos = 1                      ' L[|WV
    PTANA_Speck.ks0.keyleng = 4                     ' L[·
    PTANA_Speck.ks0.keyflag = BtKfExt + BtKfSeg     ' L[tO
    PTANA_Speck.ks0.keytype = Chr(BtKtString)       ' L[^Cv
    PTANA_Speck.ks0.reserve = &H0                   ' \ρΟέ

    PTANA_Speck.ks1.keypos = 5                      ' L[|WV
    PTANA_Speck.ks1.keyleng = 3                     ' L[·
    PTANA_Speck.ks1.keyflag = BtKfExt + BtKfSeg     ' L[tO
    PTANA_Speck.ks1.keytype = Chr(BtKtString)       ' L[^Cv
    PTANA_Speck.ks1.reserve = &H0                   ' \ρΟέ

    PTANA_Speck.ks2.keypos = 8                      ' L[|WV
    PTANA_Speck.ks2.keyleng = 1                     ' L[·
    PTANA_Speck.ks2.keyflag = BtKfExt + BtKfSeg     ' L[tO
    PTANA_Speck.ks2.keytype = Chr(BtKtString)       ' L[^Cv
    PTANA_Speck.ks2.reserve = &H0                   ' \ρΟέ

    PTANA_Speck.ks3.keypos = 9                      ' L[|WV
    PTANA_Speck.ks3.keyleng = 5                     ' L[·
    PTANA_Speck.ks3.keyflag = BtKfExt               ' L[tO
    PTANA_Speck.ks3.keytype = Chr(BtKtString)       ' L[^Cv
    PTANA_Speck.ks3.reserve = &H0                   ' \ρΟέ

    sts = BTRV(BtOpCreate, PTANA_POS, PTANA_Speck, Len(PTANA_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Β ΚIXgσόt@C")
        Exit Function
    End If
    
    PTANA_Create = False

End Function

Function PTANA_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              Β ΚIXgσόt@C  nodm              *
'*                                                                  *
'*      ψ  :Open Mode(BtrieveQΖ)                               *
'*      ίθl:false ³ν                                           *
'*             true  Ων                                           *
'*                                                                  *
'*          CREATE 2004.04.24                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PTANA_Open = True
                                            'Β ΚIXgσόt@CtpXζέ
    sts = GetIni("FILE", PTANA_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI ΗέέG[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PTANA_POS, PTANAREC, Len(PTANAREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PTANA_Create()        'Β ΚIXgσόt@Cμ¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PTANA_POS, PTANAREC, Len(PTANAREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Β ΚIXgσόt@C")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Β ΚIXgσόt@C")
                Exit Function
        End Select
    Loop
    PTANA_Open = False
End Function

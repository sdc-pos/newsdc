Attribute VB_Name = "PACKING"
Option Explicit
'********************************************************************
'*                                                                  *
'*              Β }X^  t@Cθ`                          *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
't@Chc
Public Const PACKING_ID = "PACKING"

'y[WTCY
Public Const PACKING_PG_SIZ% = 512

'|WVEubN
Public PACKING_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           \’Μθ`                             *
'*                                                                  *
'********************************************************************
'*************************** ΪΌθ` *****************************
'R[hθ`
Type PACKINGREC_Tag
    PACKING_NO(0 To 3)  As Byte         'Β 
    RANK_A1(0 To 7)     As Byte         'N@`|P
    RANK_A2(0 To 7)     As Byte         'N@`|Q
    RANK_B1(0 To 7)     As Byte         'N@a|P
    RANK_B2(0 To 7)     As Byte         'N@a|Q
    RANK_C1(0 To 7)     As Byte         'N@b|P
    RANK_C2(0 To 7)     As Byte         'N@b|Q
    FILLER(0 To 43)     As Byte         'FILLER
End Type
'f[^Eobt@
Public PACKINGREC       As PACKINGREC_Tag


'L[θ`
Type KEY0_PACKING                       'jdxO
    PACKING_NO(0 To 3)  As Byte         'Β 
End Type
    
'L[Ef[^
Public K0_PACKING       As KEY0_PACKING

Type PACKING_FSpeck
    fs  As BtFileSpeck              ' Μ§²Ω ½Νί―Έ\’Μ
    ks0 As BtKeySpeck               ' ·° ½Νί―Έ\’Μ
End Type

Private PACKING_Speck    As PACKING_FSpeck
Private Function PACKING_Create() As Integer
'********************************************************************
'*                                                                  *
'*              Β }X^  bqd`sd                          *
'*                                                                  *
'*      ψ  :Θ΅                                                 *
'*      ίθl:false ³ν                                           *
'*             true  Ων                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim FullPath    As String
Dim c           As String * 128

    PACKING_Create = True
                                            'Β }X^tpXζέ
    sts = GetIni("FILE", PACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI ΗέέG[")
        Exit Function
    End If

    FullPath = RTrim$(c)

    PACKING_Speck.fs.recoleng = Len(PACKINGREC)     ' R[h·
    PACKING_Speck.fs.PageSize = PACKING_PG_SIZ      ' y[WTCY
    PACKING_Speck.fs.idexnumb = 1                   ' CfbNX
    PACKING_Speck.fs.fileflag = 0                   ' t@CtO
    PACKING_Speck.fs.reserve = &H0                  ' \ρΟέ
                                                    ' L[O
    PACKING_Speck.ks0.keypos = 1                    ' L[|WV
    PACKING_Speck.ks0.keyleng = 4                   ' L[·
    PACKING_Speck.ks0.keyflag = BtKfExt             ' L[tO
    PACKING_Speck.ks0.keytype = Chr(BtKtString)     ' L[^Cv
    PACKING_Speck.ks0.reserve = &H0                 ' \ρΟέ

    sts = BTRV(BtOpCreate, PACKING_POS, PACKING_Speck, Len(PACKING_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Β }X^")
        Exit Function
    End If

    PACKING_Create = False

End Function

Function PACKING_Open(Mode As Integer) As Integer
'********************************************************************
'*                                                                  *
'*              Β }X^  nodm                              *
'*                                                                  *
'*      ψ  :Open Mode(BtrieveQΖ)                               *
'*      ίθl:false ³ν                                           *
'*             true  Ων                                           *
'*                                                                  *
'*          CREATE 2004.02.13                                       *
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    PACKING_Open = True
                                            'Β }X^tpXζέ
    sts = GetIni("FILE", PACKING_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI ΗέέG[")
        Exit Function
    End If
    FullPath = RTrim$(c)

    Do
        sts = BTRV(BtOpOpen, PACKING_POS, PACKINGREC, Len(PACKINGREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = PACKING_Create()        'Β }X^μ¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, PACKING_POS, PACKINGREC, Len(PACKINGREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Β }X^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Β }X^")
                Exit Function
        End Select
    Loop
    PACKING_Open = False
End Function

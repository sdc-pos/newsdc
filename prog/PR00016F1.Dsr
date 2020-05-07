VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00016F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16875
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   29766
   _ExtentY        =   19420
   SectionData     =   "PR00016F1.dsx":0000
End
Attribute VB_Name = "PR00016F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '明細のBtrieve Operation


'計上年月用添字
Private Const ptxKEIJYO_YM% = 0         '対象年月






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "URIAGE_NAME"         '仕入先名称（ｺｰﾄﾞ含む）
    
    Me.Fields.Add "URIAGE01"             '資材
    Me.Fields.Add "URIAGE02"             '工料
    Me.Fields.Add "URIAGE03"             '家賃
    Me.Fields.Add "URIAGE04"             'その他
    Me.Fields.Add "URIAGE05"             '小計
    Me.Fields.Add "URIAGE06"             '派遣工料
    Me.Fields.Add "URIAGE07"             '合計
    


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
    
    
    sts = BTRV(DET_com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K1_P_SHURI_SUM, Len(K1_P_SHURI_SUM), 1)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "資材売上集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    '収支単位
    
    If StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode) = "zzz" Then
        Me.Fields("URIAGE_NAME") = "　　総　合　計　　"
    Else
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Me.Fields("URIAGE_NAME") = StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode) & " " & StrConv(P_CODEREC.C_RNAME, vbUnicode)
            Case BtErrKeyNotFound
                Me.Fields("URIAGE_NAME") = StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードﾏｽﾀ")
                Exit Sub
        End Select
    End If

    Me.Fields("URIAGE01") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("URIAGE02") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("URIAGE03") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("URIAGE04") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")
    
    
        
    TOTAL = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode))

        
    Me.Fields("URIAGE05") = Format(TOTAL, "#,##0")

    Me.Fields("URIAGE06") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")
        
    TOTAL = TOTAL + CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode))

    Me.Fields("URIAGE07") = Format(TOTAL, "#,##0")
        
    


'--------------------------------------------------------------------------

    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim SKIP_Flg    As Boolean
 
Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
 
    Label1.Caption = PR000161.Text1(ptxKEIJYO_YM).Text                  '計上年月
    Label5.Caption = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))   'センター名






    DET_com = BtOpGetFirst


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 0
    Me.PageRightMargin = 0

    Me.documentName = "収支単位別売上集計表："

    DoEvents

End Sub


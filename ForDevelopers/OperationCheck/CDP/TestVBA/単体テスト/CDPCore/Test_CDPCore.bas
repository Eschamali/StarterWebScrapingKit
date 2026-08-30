Attribute VB_Name = "Test_CDPCore"
'==============================================================================================================
'                                   `CDPCore.cls`の単体テスト一式
'==============================================================================================================
Option Explicit

Private Const ThisClassName As String = "Test_CDPCore"




'***************************************************************************************************
'* 機能　　：`StrBufferCheck`がちゃんと`String`変数の上限ぴったりに拡張されるか？
'---------------------------------------------------------------------------------------------------
'* 期待結果：ログを参考に`responseCDP.EndCursor` = `responseCDP.length` = `Len(responseCDP.strBuffer)`となってること
'***************************************************************************************************
Sub バッファー拡張テスト()
    Const FromProcedureName As String = ThisClassName & ".バッファー拡張テスト"
    
    '----- テスト値 ※基本、CDPCore.cls と合わせる-----
    Const MaxVresSize       As Long = 2 ^ 20 + 2 ^ 16
    Const MAX_STR_LEN       As Long = 2 ^ 30 - 1
    '--------------------


    Dim testCDPCore As New CDPCore

    Dim i As Long               '追加回数
    Dim NowEndCursor As Long    '有効末尾位置
    Dim UseResSize As Long      '追記サイズ量

    'バッファー拡張ループ
    Do
        If MAX_STR_LEN - NowEndCursor < MaxVresSize Then UseResSize = MAX_STR_LEN - NowEndCursor Else UseResSize = MaxVresSize
        testCDPCore.printMsg info_, "Add buffer...               : " & UseResSize, FromProcedureName
        NowEndCursor = testCDPCore.Test_StrBufferCheck(UseResSize)
        testCDPCore.printMsg info_, "`responseCDP.EndCursor`     : " & NowEndCursor, FromProcedureName
        i = i + 1
        
        If NowEndCursor = MAX_STR_LEN Then Exit Do
    Loop

    testCDPCore.printMsg info_, "String変数の上限まで満たせました！", FromProcedureName
End Sub

'***************************************************************************************************
'* 機能　　：`CDPCore.cls - IssuanceCommandID`の`RaiseEvent ResetCommandID`テスト
'---------------------------------------------------------------------------------------------------
'* 期待結果：・「上限ぴったりで取り出し」にて、`CDPCore.cls - LimitCommandID`と同じID値としてログに表示されていること
'            ・「上限ぴったりで取り出し」後に結果が格納後、`IssuanceCommandID`を使用すると、定義中のClassに対してリセットイベントが飛んで、結果が取り出せなくなること
'            ・「リセット後:」にて、1 からスタートし始めていること
'***************************************************************************************************
Sub 結果リセットテスト()
    Const FromProcedureName As String = ThisClassName & ".結果リセットテスト"
    Const LimitCommandID    As Long = 2000000000    'CDPコマンド送信時のID上限値　※基本、CDPCore.cls と合わせる


    Dim testCDPCore     As New CDPCore
    Dim testCDPBrowser  As New CDPBrowser: Set testCDPBrowser.ThisCDPCore = testCDPCore
    Dim testCDPContext  As New CDPContext: Set testCDPContext.ThisCDPCore = testCDPCore
    
    'テスト用に結果を送信
    testCDPCore.Test_SendEvent

    '結果を回収
    testCDPCore.printMsg info_, "`CDPBrowser - Result`: " & testCDPBrowser.TakeResultCDP(1001), FromProcedureName
    testCDPCore.printMsg info_, "`CDPContext - Result`: " & testCDPContext.TakeResultCDP(1000), FromProcedureName

    '再度、テスト用に結果を送信
    testCDPCore.Test_SendEvent

    '上限まで愚直にカウントアップ
    Dim NowCount As Long
    Do
        NowCount = testCDPCore.Test_RunCountUPcommandID
        If NowCount Mod 2 ^ 22 = 0 Then DoEvents: testCDPCore.printMsg info_, "Counting... : " & Format(NowCount, "###,#"), FromProcedureName
    Loop While NowCount < LimitCommandID

    '結果を回収　※まだ取れるはず
    testCDPCore.printMsg info_, "-------- 上限ぴったりで取り出し: " & NowCount & " --------", FromProcedureName
    testCDPCore.printMsg info_, "`CDPBrowser - Result`: " & testCDPBrowser.TakeResultCDP(1001), FromProcedureName
    testCDPCore.printMsg info_, "`CDPContext - Result`: " & testCDPContext.TakeResultCDP(1000), FromProcedureName
    testCDPCore.printMsg info_, "`----------------------------------------------------------", FromProcedureName
    
    '再度、テスト用に結果を送信
    testCDPCore.Test_SendEvent

    'リセット誘発
    NowCount = testCDPCore.Test_RunCountUPcommandID
    
    '結果を回収　※取れなくなってるはず
    testCDPCore.printMsg info_, "-------- リセット後: " & NowCount & " --------", FromProcedureName
    testCDPCore.printMsg info_, "`CDPBrowser - Result`: " & testCDPBrowser.TakeResultCDP(1001), FromProcedureName
    testCDPCore.printMsg info_, "`CDPContext - Result`: " & testCDPContext.TakeResultCDP(1000), FromProcedureName
    testCDPCore.printMsg info_, "----------------------------", FromProcedureName
    End
End Sub

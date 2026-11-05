VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDPiframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================
'               `CDPContext.cls`から操作する物のうち、`iframe`系統の処理をまとめたものになります
'                   基本的には内部用となり、`CDPContext.cls`から配線する形となります
'==============================================================================================================
Option Explicit



'***************************************************************************************************
'                                   ■■■ 各種定数 ■■■
'***************************************************************************************************
Private Const ThisClassName As String = "CDPIframe"    'ログのトレース用



'***************************************************************************************************
'                                   ■■■ 各種クラス変数 ■■■
'***************************************************************************************************
'最後に`InitPageGetFrameTree`を実行した際に見つけたiframe情報
Private LastPageFrame   As BiDiCDPJConv
Private LastDeepLevel   As Long         'どこの深さか？

'その他
Private m_LogID         As String       'ログ識別ID



'***************************************************************************************************
'                                 ■■■ 分析処理の入口 ■■■
'***************************************************************************************************
'* 機能　　：検索条件に紐づくiframe情報(Page.Frame)を1つ、グローバル変数に確保します
'---------------------------------------------------------------------------------------------------
'* 返り値　：検索条件にあうiframe情報があれば、`True`を返します
'* 引数　　：frameTree      `Page.getFrameTree`のRETURN OBJECTの中身を渡します
'            iframeName     iframe名    第一優先(<iframe name="[これを指定]")
'            Url            URL         第二優先
'            NestLevel      `childFrames`をどこまで潜るか。0(省略)/1で親ページの直下のみ、nでn階層目まで、負の値で無制限
'            doRetrySecond  指定秒以内に出るまでリトライを繰り返します。0で、1回のみです
'---------------------------------------------------------------------------------------------------
'* 詳細説明：浅い階層から順に探します（幅優先）。同じ確からしさなら、浅い方のiframeを採用します
'* 注意事項：これらの引数を全て省略すると、最初の`Page.Frame`を返却します
'***************************************************************************************************
Public Function InitPageGetFrameTree(frameTree As BiDiCDPJConv, Optional iframeName As String, Optional Url As String, Optional NestLevel As Long) As Boolean
    Const FromProcedureName As String = ThisClassName & ".InitPageGetFrameTree"


    Dim frameTreeNode As BiDiCDPJConv, foundFrame As BiDiCDPJConv
    Dim bestScore As Long, bestDepth As Long

    '1. 各種初期化
    Set LastPageFrame = Nothing
    LastDeepLevel = 0

    '2. `frameTree`を取得して、条件に一番合う`Page.Frame`を探す
    Set frameTreeNode = frameTree.NodeKey("frameTree")
    Set foundFrame = FindBestFrame(frameTreeNode, ResolveMaxDepth(NestLevel), iframeName, Url, bestScore, bestDepth)

    '3. 目的の物っぽいのがなかったら、その旨を出して、抜けます
    If bestScore = -1 Then
        Dim info As String
        info = IIf(LenB(iframeName), iframeName, IIf(LenB(Url), Url, "(no iframe found)"))
        printMsg WARN_, "Unable to find the iframe: " & info, FromProcedureName
        
        Exit Function
    End If

    '4. 目的の物っぽいのがあったのでその記録
    Set LastPageFrame = foundFrame
    LastDeepLevel = bestDepth
    InitPageGetFrameTree = True
    printMsg info_, "Found the iframe with ID: " & LastPageFrame.StringKey("id") & " at level " & LastDeepLevel & ". (Certainty：" & bestScore & "%)", FromProcedureName
End Function

'***************************************************************************************************
'* 機能　　：現時点で開いている全てのiframe情報を、入れ子の深さに応じて字下げしてログに表示します
'---------------------------------------------------------------------------------------------------
'* 引数　　：frameTree      `Page.getFrameTree`のRETURN OBJECTの中身を渡します
'            TabName        タブの名前 ※デバッグ用
'            NestLevel      ※`InitPageGetFrameTree`と同じ
'***************************************************************************************************
Public Sub printChildFrames(frameTree As BiDiCDPJConv, tabName As String, Optional NestLevel As Long)
    Const FromProcedureName As String = ThisClassName & ".printChildFrames"


    Dim frameTreeNode As BiDiCDPJConv
    Set frameTreeNode = frameTree.NodeKey("frameTree")
    If Not frameTreeNode.ExistsKey("childFrames") Then printMsg info_, "'" & tabName & "' does not have an iframe.", FromProcedureName: Exit Sub

    printMsg info_, "---------------------------------------------------", FromProcedureName
    printMsg info_, "childFrames Information", FromProcedureName
    printMsg info_, "---------------------------------------------------", FromProcedureName

    Dim iframeCount As Long
    PrintFrameTree frameTreeNode, 1, ResolveMaxDepth(NestLevel), iframeCount
End Sub

'***************************************************************************************************
'* 機能　　：`NestLevel`引数を、探索する階層の上限に変換します（0は「親ページ直下のみ」=1。負の値は無制限）
'***************************************************************************************************
Private Function ResolveMaxDepth(NestLevel As Long) As Long
    If NestLevel = 0 Then ResolveMaxDepth = 1 Else ResolveMaxDepth = NestLevel
End Function

'***************************************************************************************************
'* 機能　　：`frameTree`を浅い階層から順に辿り（幅優先）、条件に一番合う`Page.Frame`を返します
'---------------------------------------------------------------------------------------------------
'* 返り値　：見つかった`Page.Frame`。無ければ`Nothing`
'* 引数　　：rootTree    `Page.getFrameTree`の`frameTree`
'            maxDepth    探索する階層の上限（負の値で無制限）
'            bestScore   見つかった`Page.Frame`の点数を返します（無ければ-1）
'            bestDepth   見つかった`Page.Frame`の階層を返します（親ページ直下=1）
'---------------------------------------------------------------------------------------------------
'* 詳細説明：浅い階層を全て見てから深い階層へ進むため、同点なら浅い方が残り、満点が出た時点で打ち切れます
'***************************************************************************************************
Private Function FindBestFrame(RootTree As BiDiCDPJConv, maxDepth As Long, iframeName As String, Url As String, _
                               ByRef bestScore As Long, ByRef bestDepth As Long) As BiDiCDPJConv
    Const MaxScore As Long = 100
    Dim pending As New Collection, Item As Variant
    Dim treeNode As BiDiCDPJConv, childList As BiDiCDPJConv, childTree As BiDiCDPJConv
    Dim Depth As Long, token As Long, score As Long

    bestScore = -1
    bestDepth = 0

    '1. 待ち行列に「FrameTreeノード, その子が何階層目か」を積んで、先頭から順に処理する
    pending.Add Array(RootTree, 1)
    Do While pending.Count > 0
        Item = pending(1)
        pending.Remove 1
        Set treeNode = Item(0)
        Depth = Item(1)

        '2. このノードの子iframeを採点し、さらに潜る分を待ち行列の末尾へ積む
        If treeNode.ExistsKey("childFrames") Then
            Set childList = treeNode.NodeKey("childFrames")
            token = childList.FirstChildToken
            Do While token > 0
                Set childTree = childList.NodeFromToken(token)

                score = ScoreFrame(childTree.NodeKey("frame"), iframeName, Url)
                If score > bestScore Then
                    bestScore = score
                    bestDepth = Depth
                    Set FindBestFrame = childTree.NodeKey("frame")
                    If bestScore = MaxScore Then Exit Function
                End If

                If maxDepth < 0 Or Depth < maxDepth Then pending.Add Array(childTree, Depth + 1)
                token = childList.NextToken(token)
            Loop
        End If
    Loop
End Function

'***************************************************************************************************
'* 機能　　：1つの`Page.Frame`が、検索条件にどれだけ合うかを採点します
'---------------------------------------------------------------------------------------------------
'* 返り値　：100(名前完全一致) / 90(URL先頭一致) / 80(条件なし) / 50(名前部分一致) / 40(URL部分一致) / -1(不一致)
'***************************************************************************************************
Private Function ScoreFrame(frame As BiDiCDPJConv, iframeName As String, Url As String) As Long
    ScoreFrame = -1

    Select Case True
        '第一優先：iframe名
        Case StrPtr(iframeName)
            Dim frameName As String: frameName = frame.StringKey("name")
            If frameName = iframeName Then
                ScoreFrame = 100
            ElseIf InStr(1, frameName, iframeName) Then
                ScoreFrame = 50
            End If

        '第二優先：URL
        Case StrPtr(Url)
            Dim StartPos As Long: StartPos = InStr(1, frame.StringKey("url"), Url)
            If StartPos = 1 Then
                ScoreFrame = 90
            ElseIf StartPos Then
                ScoreFrame = 40
            End If

        '第三優先：条件なし（最初に見つかったiframe）
        Case Else
            ScoreFrame = 80
    End Select
End Function

'***************************************************************************************************
'* 機能　　：`printChildFrames`用に、`frameTree`を文書順（深さ優先）で辿りながら表示します
'***************************************************************************************************
Private Sub PrintFrameTree(treeNode As BiDiCDPJConv, Depth As Long, maxDepth As Long, ByRef iframeCount As Long)
    Const FromProcedureName As String = ThisClassName & ".printChildFrames"
    Dim childList As BiDiCDPJConv, childTree As BiDiCDPJConv, token As Long

    If Not treeNode.ExistsKey("childFrames") Then Exit Sub

    Set childList = treeNode.NodeKey("childFrames")
    token = childList.FirstChildToken
    Do While token > 0
        Set childTree = childList.NodeFromToken(token)
        iframeCount = iframeCount + 1

        With childTree.NodeKey("frame")
            printMsg info_, String$((Depth - 1) * 2, " ") & "#" & iframeCount & " (level " & Depth & ")" & vbTab & _
                        "type: " & .StringKey("mimeType") & " | " & _
                        "title: " & .StringKey("name") & " | " & _
                        "url: " & .StringKey("url") & " | " & _
                        "frameID: " & .StringKey("id"), FromProcedureName
        End With

        '上限まで、子孫も続けて表示する（親の直後に子が並ぶよう、ここは再帰）
        If maxDepth < 0 Or Depth < maxDepth Then PrintFrameTree childTree, Depth + 1, maxDepth, iframeCount
        token = childList.NextToken(token)
    Loop
End Sub



'***************************************************************************************************
'                              ■■■ よく使う`Page.Frame`の中身確認 ■■■
'***************************************************************************************************
Property Get id() As String
    id = LastPageFrame.StringKey("id")
End Property

Property Get Url() As String
    Url = LastPageFrame.StringKey("url")
End Property

Property Get Name() As String
    Name = LastPageFrame.StringKey("name")
End Property

Property Get ParentId() As String
    ParentId = LastPageFrame.StringKey("parentId")
End Property

Property Get securityOrigin() As String
    securityOrigin = LastPageFrame.StringKey("securityOrigin")
End Property

Property Get domainAndRegistry() As String
    domainAndRegistry = LastPageFrame.StringKey("domainAndRegistry")
End Property

Property Get DeepLevelIframe() As Long
    DeepLevelIframe = LastDeepLevel
End Property




'***************************************************************************************************
'                           ■■■ CDPHelpers ラッパーメソッド ■■■
'***************************************************************************************************
Private Sub printMsg(LogLevel_ As LogLevelName, strMsg As String, From As String, Optional isHeader As Boolean = False, Optional doRaiseError As Boolean)
    Const LogFileExtension         As String = ".cdp.log"  'ログファイル拡張子


    CDPHelpers.printMsg LogLevel_, strMsg, From, LogFileExtension, m_LogID, isHeader, doRaiseError, CDPHelpers.currentLog.logPath
End Sub



'***************************************************************************************************
'                       ■■■ このClassオブジェクト生成/破棄時のアクション ■■■
'***************************************************************************************************
Private Sub Class_Initialize()
    'Assign an objID for easy debug identification
    m_LogID = "IFID" & Format(Rnd * 1000, "000")
End Sub

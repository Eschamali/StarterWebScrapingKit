Attribute VB_Name = "BiDiPoC"
Option Explicit


Sub TestBiDi_mapperTabJS改()
    Dim DemoBiDi As CDPBrowser: Set DemoBiDi = 設定シートからの起動
    Dim CharConv As New CharacterCodeConversion

    Dim paramsCDP As Dictionary
    Dim ResultCDP As New Dictionary

    Dim BiDiMapperScript() As Byte
    BiDiMapperScript = CharConv.BytesFromSavedFile("C:\Users\XXXX\Downloads", "mapperTab.js")
    Stop

    With DemoBiDi
        .navigate "about:blank#MAPPER_TARGET", isComplete
    
        Set paramsCDP = New Dictionary
        Set ResultCDP = .invokeMethod("Runtime.enable", paramsCDP)

        Dim current_targetID As String: current_targetID = ShSetting01_StartBrowser.deserializeFromTable("Automation Data", 2, "TestBiDi")("targetID")
        paramsCDP.RemoveAll
        paramsCDP.Add "bindingName", "cdp"
        paramsCDP.Add "targetId", current_targetID
        paramsCDP.Add "inheritPermissions", True
        Set ResultCDP = .invokeMethod("Target.exposeDevToolsProtocol", paramsCDP, True)

        ' JSからVBAへメッセージを送るための関数名を "sendBidiResponse" に設定
        paramsCDP.RemoveAll
        paramsCDP.Add "name", "sendBidiResponse"
        .invokeMethod "Runtime.addBinding", paramsCDP

        'mapperTab.js の中身を丸ごと Runtime.evaluate で実行します。
        .jsEval CharConv.BytesToString(BiDiMapperScript)

        'ブートストラップ（初期化）JSの実行
        .jsEval "window.runMapperInstance('" & current_targetID & "')"
        Stop

        .newTab	'あとで、Googleへ遷移するためのやつ

        ' BiDiコマンドをJSの関数経由でMapperに渡す
        Dim bidiCommand As String
        ' BiDiコマンドそのもの（ダブルクォートは2つ重ねてエスケープ）
        bidiCommand = "{""id"": 100, ""method"": ""session.new"", ""params"": { ""capabilities"": {} }}"
        
        paramsCDP.RemoveAll
        ' ★重要：bidiCommand をシングルクォート ' で囲んで「文字列」としてJSに渡す
        paramsCDP.Add "expression", "window.onBidiMessage('" & bidiCommand & "');"
        
        ' 実行！
        Set .BrowserEvents = New Dictionary
        Set ResultCDP = .invokeMethod("Runtime.evaluate", paramsCDP)
        Stop

        ' --- 特定のイベント名が出るまでループ ---
        Const SearchEventName As String = "Runtime.bindingCalled"
        Do
            '非同期イベントを取り出す
            .TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                Dim tmp
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "name               :"; tmp("params")("name")
                    Debug.Print "payload            :"; tmp("params")("payload")
                    Debug.Print "executionContextId :"; tmp("params")("executionContextId") & vbCrLf
                Next

                '見つかったので抜ける
                Exit Do
            End If
            
            .sleep
            DoEvents
        Loop While True
        Set .BrowserEvents = New Dictionary
        Stop
        
        Dim bidiGetTree As String
        ' browsingContext.getTree は現在の全タブ（コンテキスト）を取得するコマンド
        bidiGetTree = "{""id"": 101, ""method"": ""browsingContext.getTree"", ""params"": {}}"
        
        paramsCDP("expression") = "window.onBidiMessage('" & bidiGetTree & "');"
        .invokeMethod "Runtime.evaluate", paramsCDP
        Stop
        


        ' --- 特定のイベント名が出るまでループ ---
        Do
            '非同期イベントを取り出す
            .TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "name               :"; tmp("params")("name")
                    Debug.Print "payload            :"; tmp("params")("payload")
                    Debug.Print "executionContextId :"; tmp("params")("executionContextId") & vbCrLf
                Next

                '見つかったので抜ける
                Exit Do
            End If
            
            .sleep
            DoEvents
        Loop While True
        Set .BrowserEvents = New Dictionary
        Stop

        Dim targetContext As String
        targetContext = "9EF1B3ADE324E61A74C8311ED07D47EC" ' Stopの間に遷移させたいやつを
        
        Dim bidiNavi As String
        bidiNavi = "{""id"": 102, ""method"": ""browsingContext.navigate"", " & _
                   """params"": {""context"": """ & targetContext & """, " & _
                   """url"": ""https://www.google.com"", ""wait"": ""complete""}}"
        
        ' Mapperへ投下！
        paramsCDP("expression") = "window.onBidiMessage('" & bidiNavi & "');"
        .invokeMethod "Runtime.evaluate", paramsCDP
        Stop
        


        ' --- 特定のイベント名が出るまでループ ---
        Do
            '非同期イベントを取り出す
            .TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "name               :"; tmp("params")("name")
                    Debug.Print "payload            :"; tmp("params")("payload")
                    Debug.Print "executionContextId :"; tmp("params")("executionContextId") & vbCrLf
                Next

                '見つかったので抜ける
                Exit Do
            End If
            
            .sleep
            DoEvents
        Loop While True
        Set .BrowserEvents = New Dictionary
        Stop

        ' console.log が発生したら自動的に sendBidiResponse を叩くように予約
        Dim bidiSubscribe As String
        bidiSubscribe = "{""id"": 105, ""method"": ""session.subscribe"", ""params"": {""events"": [""log.entryAdded""]}}"
        
        paramsCDP("expression") = "window.onBidiMessage('" & bidiSubscribe & "');"
        .invokeMethod "Runtime.evaluate", paramsCDP
        
         ' --- 特定のイベント名が出るまでループ ---
        Do
            '非同期イベントを取り出す
            .TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "name               :"; tmp("params")("name")
                    Debug.Print "payload            :"; tmp("params")("payload")
                    Debug.Print "executionContextId :"; tmp("params")("executionContextId") & vbCrLf
                Next

                '見つかったので抜ける
                Exit Do
            End If
            
            .sleep
            DoEvents
        Loop While True
        Set .BrowserEvents = New Dictionary
        Stop    'console.log("VBA見てるー？")  と打ったら、進めてみましょう
      
         ' --- 特定のイベント名が出るまでループ ---
        Do
            '非同期イベントを取り出す
            .TakeEvents

            'イベント名の確認
            If .BrowserEvents("EventMethods").Exists(SearchEventName) Then
                '出ているダイアログの情報の確認
                For Each tmp In .BrowserEvents("EventMethods")(SearchEventName)
                    Debug.Print "name               :"; tmp("params")("name")
                    Debug.Print "payload            :"; tmp("params")("payload")
                    Debug.Print "executionContextId :"; tmp("params")("executionContextId") & vbCrLf
                Next

                '見つかったので抜ける
                Exit Do
            End If
            
            .sleep
            DoEvents
        Loop While True
        Set .BrowserEvents = New Dictionary
        Stop



        
        
        .quit
    End With




End Sub

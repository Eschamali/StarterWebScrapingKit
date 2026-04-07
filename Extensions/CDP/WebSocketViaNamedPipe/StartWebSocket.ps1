# ==========================================
# VBA用 CDP WebSocket 中継器 (Bridge)
# ==========================================

#----------------------------- 1. 初期パラメータ一式 ----------------------------- 
$wsUrl = "ws://127.0.0.1:9222/devtools/page/E3397DFFC3906A51A78CB5B86F8820DA"
$pipeName = "ChromiumWebSocket"
#---------------------------------------------------------------------------------

# 2. WebSocketの準備と接続
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$uri = New-Object System.Uri($wsUrl)
$cts = New-Object System.Threading.CancellationTokenSource
$ws.ConnectAsync($uri, $cts.Token).Wait()
Write-Host "✅ Chrome(WebSocket) に接続しました！"

#------------------------3. Excelへの名前付きパイプ接続処理------------------------
try {
    # サーバーに接続を試みる
    $pipeClient = [System.IO.Pipes.NamedPipeClientStream]::new(".", $pipeName, [System.IO.Pipes.PipeDirection]::InOut)
    Write-Host "VBAからの接続を待っています... 10秒以内に接続して下さい。"
    $pipeClient.Connect(10000) # 10秒間接続を待つ

    $reader = [System.IO.StreamReader]::new($pipeClient)
    $writer = [System.IO.StreamWriter]::new($pipeClient)
    $writer.AutoFlush = $true

    Write-Host "VBAサーバーへの接続完了！コマンド待機中..."
} catch {
    Write-Error "時間内でのExcelからの接続が、確認できませんでした。`n$($_.Exception.Message)"
    exit 1
}
#---------------------------------------------------------------------------------- 

# ==========================================
# 4. 【メインループ】 非同期I/Oを使った双方向通信
# ==========================================
Write-Host "🔄 双方向のデータ中継を開始します..."

try {
    # それぞれの送受信用バッファを用意
    $bufferPipe = New-Object byte[] 8192
    $bufferWs = New-Object byte[] 8192
    
    # 🌟 ArraySegment の作成方法を、C#ライクな安全な書き方に変更！
    $segmentWs = [System.ArraySegment[byte]]::new($bufferWs)

    # パイプとWebSocket、両方の「受信待機タスク」を同時にスタート
    $taskReadPipe = $pipeClient.ReadAsync($bufferPipe, 0, $bufferPipe.Length)
    $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)

    # どちらかが切断されるまで無限ループ
    while ($pipeClient.IsConnected -and $ws.State -eq 'Open') {
        
        # 🌟 ここがキモ！「パイプ」か「WebSocket」、先にデータが来た方から処理を進める
        $idx = [System.Threading.Tasks.Task]::WaitAny($taskReadPipe, $taskReadWs)

        if ($idx -eq 0) {
            # ----------------------------------------------------
            # パイプ(VBA) からデータが来た！ ➡️ WebSocket(Chrome) へ送信
            # ----------------------------------------------------
            $bytesRead = $taskReadPipe.Result
            if ($bytesRead -eq 0) { break } # 切断された

            # 🌟 修正1: VBAから送られてくるデータの末尾の「Null文字(0x00)」を削る
            $realLength = $bytesRead
            # 末尾が 0 (Null文字) なら、送る長さを1バイト短くする
            if ($bufferPipe[$bytesRead - 1] -eq 0) {
                $realLength = $bytesRead - 1
            }

            # ゴミ(Null)を取り除いた長さでChromeへ送信！
            $sendSegment = [System.ArraySegment[byte]]::new($bufferPipe, 0, $realLength)
            $ws.SendAsync($sendSegment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()

            # 次のパイプ受信タスクを再セット
            $taskReadPipe = $pipeClient.ReadAsync($bufferPipe, 0, $bufferPipe.Length)

        } elseif ($idx -eq 1) {
            # ----------------------------------------------------
            # WebSocket(Chrome) からデータが来た！ ➡️ パイプ(VBA) へ送信
            # ----------------------------------------------------
            $result = $taskReadWs.Result
            if ($result.MessageType -eq 'Close') { break }

            # Chromeから受け取った生データをパイプへ書き込む
            $pipeClient.Write($bufferWs, 0, $result.Count)

            # 🌟 修正: 「これで1つのJSONが完全に終わったか？」を確認する！
            if ($result.EndOfMessage) {
                # メッセージの終わりなら、VBAが待っている「Null文字(0x00)」を追記する
                $nullByte = [byte[]]@(0)
                $pipeClient.Write($nullByte, 0, 1)
            }

            # VBAへ即座に送り出す
            $pipeClient.Flush()

            # 次のWebSocket受信タスクを再セット
            $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)
        }
    }
} catch {
    Write-Host "⚠️ 中継中にエラーが発生しました: $($_.Exception.Message)"
}

# 7. お片付け
Write-Host "🛑 通信が終了しました。お片付けします。"
if ($pipeClient -ne $null) { $pipeClient.Dispose() }
if ($ws -ne $null) { $ws.Dispose() }
Write-Host "完了！"

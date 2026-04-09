# ====================================================================================
#	StarterWebScrapingKit - VBA用 CDP WebSocket 中継器 (Bridge)
#	事前に、対象のChromiumに対して、`--remote-debugging-port=9222 --user-data-dir="XXX"` を付与して起動したうえで、実行してください。
#	条件さえ満たせば、「chrome://inspect」から任意のデバイス内のChromium制御も可能になります。
# ====================================================================================



#----------------------------- 1. 初期パラメータ一式 ----------------------------- 
$wsUrl    = "ws://127.0.0.1:9222/devtools/browser/cbb667e3-758f-4cb3-b2a9-85f1b2e3953a"	#`remote-debugging-pipe`相当の`ws`に接続します。`http://127.0.0.1:9222/json/version`にて、確認可能です。
$pipeName = "ChromiumWebSocket"	#Excelから接続する名前付きパイプと一致するようにしてください。
#---------------------------------------------------------------------------------

# 🌟 ミリ秒付きでログを出す便利関数
function Log($msg, $AddArg = @{}) {
    $time = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff")

    # 🌟 @AddArg と書くのが魔法の呪文（スプラッティング）
    # これにより、色などのオプションを外から流し込めるようになります
    Write-Host "[$time] $msg" @AddArg
}

#------------------------2. WebSocketの準備と接続------------------------
try {
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = New-Object System.Uri($wsUrl)
    $cts = New-Object System.Threading.CancellationTokenSource
    $ws.ConnectAsync($uri, $cts.Token).Wait()

    Log "✅ Chromium(WebSocket) に接続しました！" @{ForegroundColor="Green"}
} catch {
    Write-Error "Chromium(WebSocket) に接続できませんでした。`n$($_.Exception.Message)"
    exit 1
}
#---------------------------------------------------------------------------------- 

#------------------------3. Excelへの名前付きパイプ接続処理------------------------
try {
    # 指定のパイプに接続を試みる
    $pipeOptions = [System.IO.Pipes.PipeOptions]::WriteThrough -bor [System.IO.Pipes.PipeOptions]::Asynchronous	# 🌟 バッファリングを絶対許さない「WriteThrough」フラグを追加して接続！
    $pipeClient = [System.IO.Pipes.NamedPipeClientStream]::new(".", $pipeName, [System.IO.Pipes.PipeDirection]::InOut, $pipeOptions)

    Log "VBAからの接続を待っています... 10秒以内に接続して下さい"
    $pipeClient.Connect(10000) # 10秒間接続を待つ

    $reader = [System.IO.StreamReader]::new($pipeClient)
    $writer = [System.IO.StreamWriter]::new($pipeClient)
    $writer.AutoFlush = $true

    Log "VBAサーバーへの接続完了！コマンド待機中..." @{ForegroundColor="Green"}
} catch {
    Write-Error "時間内にExcelからの接続が、確認できませんでした`n$($_.Exception.Message)"
    exit 1
}
#---------------------------------------------------------------------------------- 

# ==========================================
# 4. 【メインループ】 非同期I/Oを使った双方向通信
# ==========================================
Log "🔄 双方向のデータ中継を開始します..." @{ForegroundColor="DarkBlue"}

try {
    # それぞれの送受信用バッファを用意(VBA側の設定`CDPCore`に準拠)
    $bufferPipe = New-Object byte[] 1MB	# 2 ^ 20
    $bufferWs = New-Object byte[] 1MB	# 2 ^ 20

    # 🌟 ArraySegment の作成方法を、C#ライクな安全な書き方に変更！
    $segmentWs = [System.ArraySegment[byte]]::new($bufferWs)
    $cts = New-Object System.Threading.CancellationTokenSource

    # パイプとWebSocket、両方の「受信待機タスク」を同時にスタート
    $taskReadPipe = $pipeClient.ReadAsync($bufferPipe, 0, $bufferPipe.Length)
    $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)

    # VBAからのデータを蓄積するバッファを用意する。パイプ通信の仕様(データを細切れ)に備える
    $vbaReceiveBuffer = New-Object System.IO.MemoryStream

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

            # 今回の破片が「Null文字」で終わっているか判定
            $endsWithNull = ($bufferPipe[$bytesRead - 1] -eq 0)

            # 🌟 ハイブリッド判定ロジック
            if ($vbaReceiveBuffer.Length -eq 0 -and $endsWithNull) {
                # --- 【高速ルート：直通便 🚀】 ---
                # 今まで貯まったものがなく、かつ今回の1回でヌル文字が来た場合
                $realLength = $bytesRead - 1
                $sendSegment = [System.ArraySegment[byte]]::new($bufferPipe, 0, $realLength)

                # 直接 Chrome へ送信！
                $ws.SendAsync($sendSegment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()
                Log "🚀 【直通】 短文JSONを即座に送信しました ($realLength バイト)" @{ForegroundColor="Cyan"}

            } else {
                # --- 【蓄積ルート：慎重便 📦】 ---
                # すでに貯まっている途中があるか、今回のデータがヌル文字で終わっていない場合

                # とりあえずバッファーに貯める
                $vbaReceiveBuffer.Write($bufferPipe, 0, $bytesRead)

                if ($endsWithNull) {
                    # ヌル文字が来た！これでガッチャンコ完了
                    $fullData = $vbaReceiveBuffer.ToArray()
                    $realLength = $fullData.Length - 1 # 最後のヌル文字を除く

                    # 1つの巨大な塊にして Chrome へ送信！
                    $sendSegment = [System.ArraySegment[byte]]::new($fullData, 0, $realLength)
                    $ws.SendAsync($sendSegment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()

                    Log "📦 【合体】 蓄積された長文JSONを送信しました ($realLength バイト)" @{ForegroundColor="DarkCyan"}

                    # 次のためにバッファーを空にする
                    $vbaReceiveBuffer.SetLength(0)

                } else {
                    # まだヌル文字が来ない。次を待つ
                    Log "⏳ 【蓄積中...】 パケットが分割されています (現在 $($vbaReceiveBuffer.Length) バイト)" @{ForegroundColor="Yellow"}
                }
            }

            # 次のパイプ受信タスクを再セット
            $taskReadPipe = $pipeClient.ReadAsync($bufferPipe, 0, $bufferPipe.Length)

        } elseif ($idx -eq 1) {
            # ----------------------------------------------------
            # WebSocket(Chrome) からデータが来た！ ➡️ パイプ(VBA) へ送信
            # ----------------------------------------------------
            $result = $taskReadWs.Result
            if ($result.MessageType -eq 'Close') { break }

            Log "⬅️ 【Chromiumから受信】 $($result.Count) バイト (EndOfMessage: $($result.EndOfMessage))" @{ForegroundColor="Magenta"}

            # Chromeから受け取った生データをパイプへ書き込む
            $pipeClient.Write($bufferWs, 0, $result.Count)

            # 🌟「これで1つのJSONが完全に終わったか？」を確認する！
            if ($result.EndOfMessage) {
                # メッセージの終わりなら、VBAが待っている「Null文字(0x00)」を追記する
                $nullByte = [byte[]]@(0)
                $pipeClient.Write($nullByte, 0, 1)
            }

            # VBAへ即座に送り出す
            $pipeClient.Flush()
            Log "⬅️ 【パイプFlush完了】" @{ForegroundColor="Magenta"}

            # 次のWebSocket受信タスクを再セット
            $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)
        }
    }
} catch {
    Write-Error "⚠️ 中継中にエラーが発生しました`n$($_.Exception.Message)"
}

# 7. お片付け
Log "🛑 通信終了"
if ($pipeClient -ne $null) { $pipeClient.Dispose() }
if ($ws -ne $null) { $ws.Dispose() }
Log "お片付け完了！" @{ForegroundColor="Green"}

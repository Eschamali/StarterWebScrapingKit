# ====================================================================================
#	StarterWebScrapingKit - VBA用 CDP WebSocket 中継器 (Bridge)
#	事前に、対象のChromiumに対して、`--remote-debugging-port=9222 --user-data-dir="XXX"` を付与して起動したうえで、実行してください。
#
#	条件さえ満たせば、「chrome://inspect」から任意のデバイス内のChromium制御も可能になります。
#	→https://developer.chrome.com/docs/devtools/remote-debugging?hl=ja
#
#	【起動順】PowerShell が名前付きパイプのサーバー（待受け）となり、Excel がクライアントとして接続します。
#	先に本スクリプトを実行して待機し、その後 Excel 側で`ConnectNamePipe`を実行してください。
# ====================================================================================



#----------------------------- 1. 初期パラメータ一式(コマンドライン引数対応) -----------------------------
param(
    [string]$wsUrl    = "",			#`remote-debugging-pipe`相当の`ws`に接続します。`http://127.0.0.1:9222/json/version`にて、確認可能です。
    [string]$pipeName = "ChromiumWebSocket",	#Excel側の `ConnectNamePipe` に渡す引数名と一致させてください。
    [uint16]$port     = 9222			#`Remote debugging`でのデフォルトポート番号
)
#---------------------------------------------------------------------------------------------------------

# 🌟 GUI用のライブラリをロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 🌟 ミリ秒付きでログを出す便利関数
function Log($msg, $AddArg = @{}) {
    $time = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff")

    # 🌟 @AddArg と書くのが魔法の呪文（スプラッティング）
    # これにより、色などのオプションを外から流し込めるようになります
    Write-Host "[$time] $msg" @AddArg
}

#------------------------2. WebSocketの準備と接続------------------------
# 2-1. 引数`$wsUrl`が省略されたら、接続リストを取得して、接続リストウィンドウを表示させます
if ([string]::IsNullOrWhiteSpace($wsUrl)) {
    Log "🔎 WebSocketURLが省略されてるため、接続リストを取得します。しばらく経つと、GUI画面が出ますので、接続リスト先を選んでください" @{ForegroundColor="DarkYellow"}

    # --- 接続リストの基となる基本的な場所 ---
    $edgePortFile = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\DevToolsActivePort"
    $baseUrl      = "http://127.0.0.1:$port"


    # 2-1-1. `DevToolsActivePort`があるか？
    if (Test-Path $edgePortFile) {
        try {
            $portInfo = Get-Content $edgePortFile -ErrorAction SilentlyContinue
            if ($portInfo) {
                $directport = $portInfo[0]
                if ($portInfo.Count -gt 1) { $directWsPath = $portInfo[1] }
            }
        } catch {}
    }

    # --- 2-1-2. データの取得試行 ---
    $tabsJson = @()
    $verJson = $null

    # HTTP経由での取得（タブ情報）
    try {
        $tabsJson = Invoke-RestMethod "$baseUrl/json" -ErrorAction Stop -TimeoutSec 2
    } catch {}

    # HTTP経由での取得（バージョン情報）
    try {
        $verJson = Invoke-RestMethod "$baseUrl/json/version" -ErrorAction Stop -TimeoutSec 2
    } catch {}

    # --- 2-1-3. 最終判定：どこにも繋がる気配がない場合 ---
    $hasTabs = ($tabsJson.Count -gt 0)
    $hasVer = ($verJson -and $verJson.webSocketDebuggerUrl)
    $hasDirect = (![string]::IsNullOrEmpty($directWsPath))

    if (!$hasTabs -and !$hasVer -and !$hasDirect) {
        [System.Windows.Forms.MessageBox]::Show(
            "有効な接続先が見つかりませんでした。`n`n" +
            "・ブラウザがデバッグモードで起動しているか`n" +
            "・ポート番号 ($port) が正しいか`n" +
            "を確認してください。", "接続エラー", 0, 16)
        exit 1
    }

    # --- 2-1-4. GUI構築 ---
    $script:sortColumn = -1
    $script:isDescending = $false

    $form = New-Object Windows.Forms.Form
    $form.Text = "Chromium 接続先詳細セレクター (Smart Search)"
    $form.Size = "800, 550"; $form.StartPosition = "CenterScreen"; $form.Font = "Yu Gothic UI, 10"
    $form.MinimumSize = New-Object Drawing.Size(640, 360)

    $tabControl = New-Object Windows.Forms.TabControl
    $tabControl.Dock = "Fill"

    # --- 🌟 無効なタブへの切り替えを物理的にブロックする ---
    $tabControl.add_Selecting({
        param($s, $e)
        if ($e.TabPage -and !$e.TabPage.Enabled) { $e.Cancel = $true }
    })

    # --- タブ1: [タブ情報] ---
    $tab1 = New-Object Windows.Forms.TabPage; $tab1.Text = " タブ情報 "
    $listView = New-Object Windows.Forms.ListView; $listView.View = "Details"; $listView.FullRowSelect = $true; $listView.GridLines = $true; $listView.Dock = "Fill"

    # 🌟 列クリックイベント
    $listView.add_ColumnClick({
        param($sender, $e)
        
        # クリックされた列のフィールド名を取得
        $clickedField = $visibleFields[$e.Column]
        
        # 同じ列をクリックしたら昇順/降順を反転
        if ($script:sortColumn -eq $e.Column) {
            $script:isDescending = !$script:isDescending
        } else {
            $script:sortColumn = $e.Column
            $script:isDescending = $false
        }

        # 🌟 データを並び替える（PowerShellの魔法）
        $tabsJson = $tabsJson | Sort-Object -Property $clickedField -Descending:$script:isDescending
        
        # リストを更新
        &$UpdateList
    })

    # 全フィールド定義
    $allFields = @("type", "title", "url", "id", "description", "devtoolsFrontendUrl", "faviconUrl", "webSocketDebuggerUrl")
    $visibleFields = New-Object System.Collections.Generic.List[string]
    $visibleFields.AddRange([string[]]@("type", "title")) # 固定列

    $UpdateList = {
        $listView.BeginUpdate(); $listView.Columns.Clear(); $listView.Items.Clear()
        foreach ($f in $visibleFields) { $listView.Columns.Add($f, 150) | Out-Null }
        foreach ($t in $tabsJson) {
            $item = New-Object Windows.Forms.ListViewItem([string]$t.$($visibleFields[0]))
            for ($i = 1; $i -lt $visibleFields.Count; $i++) { $item.SubItems.Add([string]$t.$($visibleFields[$i])) | Out-Null }
            $item.Tag = $t.webSocketDebuggerUrl; $listView.Items.Add($item) | Out-Null
        }
        $listView.EndUpdate()
    }

    $menu = New-Object Windows.Forms.ContextMenuStrip
    foreach ($field in $allFields) {
        $m = New-Object Windows.Forms.ToolStripMenuItem($field)
        $m.CheckOnClick = $true; if ($visibleFields.Contains($field)) { $m.Checked = $true }
        if ($field -eq "type" -or $field -eq "title") { $m.Enabled = $false }
        else { $m.Add_Click({ if ($this.Checked) { $visibleFields.Add($this.Text) } else { $visibleFields.Remove($this.Text) }; &$UpdateList }) }
        $menu.Items.Add($m) | Out-Null
    }
    $listView.ContextMenuStrip = $menu
    $tab1.Controls.Add($listView); &$UpdateList

    # --- タブ2: [ブラウザ本体] ---
    $tab2 = New-Object Windows.Forms.TabPage; $tab2.Text = " ブラウザ本体 "
    $radioPanel = New-Object Windows.Forms.Panel; $radioPanel.Dock = "Fill"; $radioPanel.Padding = "30, 30, 30, 30"
    
    $r1 = New-Object Windows.Forms.RadioButton; $r1.Text = "API (json/version) 接続"; $r1.Location = "20, 30"; $r1.AutoSize = $true
    if ($hasVer) { 
        $r1.Tag = $verJson.webSocketDebuggerUrl; $r1.Text += " ($($verJson.Browser))"
        $r1.Checked = $true # 🌟 第一候補
    } else { $r1.Enabled = $false }

    $r2 = New-Object Windows.Forms.RadioButton; $r2.Text = "ActivePort (Direct) 接続"; $r2.Location = "20, 70"; $r2.AutoSize = $true
    if ($hasDirect) { 
        $r2.Tag = "ws://127.0.0.1:$directport$directWsPath"
        # 🌟 r1が使えない場合のみ、r2を初期選択にする
        if (!$r1.Enabled) { $r2.Checked = $true }
    } else { $r2.Enabled = $false }

    $radioPanel.Controls.AddRange(@($r1, $r2))
    $tab2.Controls.Add($radioPanel)

    $tabControl.TabPages.AddRange(@($tab1, $tab2))

    # 🌟 表示・選択不可制御
    if (!$hasTabs) {
        $tab1.Enabled = $false
        $tabControl.SelectedTab = $tab2 # 最初からブラウザ本体タブを表示
    }

    # --- 1. 下部パネルの作成 ---
    $bottom = New-Object Windows.Forms.Panel
    $bottom.Height = 60
    $bottom.Dock = "Bottom"
    # $bottom.BackColor = "LightGray" # 👈 デバッグ用：もしボタンが出なかったらここを有効にしてパネルが見えるか確認

    # --- 2. 接続ボタンの作成 ---
    $btn = New-Object Windows.Forms.Button
    $btn.Text = "接続開始"
    $btn.Size = "120, 35"
    $btn.FlatStyle = "System"
    $btn.DialogResult = [Windows.Forms.DialogResult]::OK

    # 🌟 まずは「左上」の適当な位置に置いて、パネルに追加しちゃう！
    $btn.Location = New-Object Drawing.Point(10, 10) 
    $bottom.Controls.Add($btn)

    # --- 3. フォームにパーツを追加（順番が大事！） ---
    # 先に Bottom を追加して場所を確保！
    $form.Controls.Add($bottom)
    $form.Controls.Add($tabControl)

    # 🌟 追加された「後」で、正しい位置（右端）に移動させる！
    # パネルの幅 ($bottom.Width) を基準に計算
    $btn.Left = $bottom.Width - $btn.Width - 25
    $btn.Top = ($bottom.Height - $btn.Height) / 2

    # 🌟 位置が決まってから「Anchor（いかり）」を下ろす！
    $btn.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

    # --- 最終実行 ---
    if ($form.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        if ($tabControl.SelectedTab -eq $tab1) { 
            if ($listView.SelectedItems.Count -gt 0) { $wsUrl = $listView.SelectedItems[0].Tag } 
        } else { 
            if ($r1.Checked) { $wsUrl = $r1.Tag } elseif ($r2.Checked) { $wsUrl = $r2.Tag } 
        }
    }
    if ([string]::IsNullOrWhiteSpace($wsUrl)) { exit 1 }
}

# --- 2-2. 実際に接続 ---
try {
    Log "🚀 接続先確定 → $wsUrl" @{ForegroundColor="Cyan"}

    Log "📡 Chromium(WebSocket) に接続中です..." @{ForegroundColor="Yellow"}
    $ws = New-Object System.Net.WebSockets.ClientWebSocket
    $uri = New-Object System.Uri($wsUrl)
    $cts = New-Object System.Threading.CancellationTokenSource
    $ws.ConnectAsync($uri, $cts.Token).Wait()

    Log "☑️ Chromium(WebSocket) に接続しました！" @{ForegroundColor="Green"}
} catch {
    Write-Error "❌ Chromium(WebSocket) に接続できませんでした。`n$($_.Exception.Message)"
    exit 1
}
#-----------------------------------------------------------------------------------

#------------------------3. 名前付きパイプ作成と接続待機処理------------------------
try {
    # オプション設定
    $pipeOptions = [System.IO.Pipes.PipeOptions]::WriteThrough -bor [System.IO.Pipes.PipeOptions]::Asynchronous	# 🌟 バッファリングというお節介をなくす「WriteThrough」フラグを追加

    # 名前付きパイプを作成
    # ※指定のパイプ名で、バイト配列による読み書きモードで先着1名、1MB分のバッファーパイプとして用意します
    $pipeServer = [System.IO.Pipes.NamedPipeServerStream]::new(
        $pipeName,
        [System.IO.Pipes.PipeDirection]::InOut,
        1,
        [System.IO.Pipes.PipeTransmissionMode]::Byte,
        $pipeOptions,
        1MB,
        1MB
    )

} catch {
    Write-Error "💫 名前付きパイプ サーバー作成中に、エラーが発生しました。`nコンソールの再起動が必要です`n$($_.Exception.Message)"
    exit 1
}

try {
    Log "📂 名前付きパイプ サーバー:$pipeName を起動しました。Excel からの接続を待機しています..." @{ForegroundColor="Yellow"}
    Log "ℹ️ キャンセルする場合は、このコンソールを閉じてください。" @{ForegroundColor="DarkCyan"}

    # ★ここで Excel が CreateFile するまでブロックします
    $pipeServer.WaitForConnection()

    Log "✅ Excelから接続が来ました！コマンド待機中..." @{ForegroundColor="Green"}


    #---------------------------------------------------------------------------------- 

    # ==========================================
    # 4. 【メインループ】 非同期I/Oを使った双方向通信
    # ==========================================
    Log "🔄 双方向のデータ中継を開始します..." @{ForegroundColor="DarkBlue"}

    # それぞれの送受信用バッファを用意(VBA側の設定`CDPCore`に準拠)
    $bufferPipe = New-Object byte[] 1MB	# 2 ^ 20
    $bufferWs = New-Object byte[] 1MB	# 2 ^ 20

    # 🌟 ArraySegment の作成方法を、C#ライクな安全な書き方に変更！
    $segmentWs = [System.ArraySegment[byte]]::new($bufferWs)
    $cts = New-Object System.Threading.CancellationTokenSource

    # パイプとWebSocket、両方の「受信待機タスク」を同時にスタート
    $taskReadPipe = $pipeServer.ReadAsync($bufferPipe, 0, $bufferPipe.Length)
    $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)

    # VBAからのデータを蓄積するバッファを用意する。パイプ通信の仕様(データを細切れ)に備える
    $vbaReceiveBuffer = New-Object System.IO.MemoryStream

    # どちらかが切断されるまで無限ループ
    while ($pipeServer.IsConnected -and $ws.State -eq 'Open') {

        # 🌟 ここがキモ！「パイプ」か「WebSocket」、先にデータが来た方から処理を進める
        $idx = [System.Threading.Tasks.Task]::WaitAny($taskReadPipe, $taskReadWs)

        if ($idx -eq 0) {
            # ----------------------------------------------------
            # パイプ(VBA) からデータが来た！ ➡️ WebSocket(Chrome) へ送信
            # ----------------------------------------------------
            $bytesRead = $taskReadPipe.Result
            if ($bytesRead -eq 0) { break } # 切断された

            # 読み取った断片データを一旦蓄積バッファに追加
            $vbaReceiveBuffer.Write($bufferPipe, 0, $bytesRead)

            # 現在バッファにあるデータをバイト配列で取得
            $currentData = $vbaReceiveBuffer.ToArray()
            $startIdx = 0
            $foundNull = $false

            # バッファ内のヌル文字 (0x00) を探して切り出し送信を行う
            for ($i = 0; $i -lt $currentData.Length; $i++) {
                if ($currentData[$i] -eq 0) {
                    $msgLength = $i - $startIdx
                    if ($msgLength -gt 0) {
                        # ヌル文字の手前までのメッセージを切り出して送信
                        $sendSegment = [System.ArraySegment[byte]]::new($currentData, $startIdx, $msgLength)
                        $ws.SendAsync($sendSegment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).Wait()
                        Log "🚀 【分割送信】 JSONメッセージを個別に送信しました ($msgLength バイト)" @{ForegroundColor="Cyan"}
                    }
                    $startIdx = $i + 1
                    $foundNull = $true
                }
            }

            # 1つ以上のメッセージが処理された場合、送信済み領域をバッファから取り除く
            if ($foundNull) {
                $vbaReceiveBuffer.SetLength(0)
                $remainingLength = $currentData.Length - $startIdx
                if ($remainingLength -gt 0) {
                    # 送信しきれなかった未完了のデータ（次のJSONの破片など）をバッファに書き戻す
                    $vbaReceiveBuffer.Write($currentData, $startIdx, $remainingLength)
                    Log "⏳ 【蓄積中】 未完了のパケットがあります ($remainingLength バイト)" @{ForegroundColor="Yellow"}
                }
            } else {
                Log "⏳ 【蓄積中...】 パケットが分割されています (現在 $($vbaReceiveBuffer.Length) バイト)" @{ForegroundColor="Yellow"}
            }

            # 次のパイプ受信タスクを再セット
            $taskReadPipe = $pipeServer.ReadAsync($bufferPipe, 0, $bufferPipe.Length)

        } elseif ($idx -eq 1) {
            # ----------------------------------------------------
            # WebSocket(Chrome) からデータが来た！ ➡️ パイプ(VBA) へ送信
            # ----------------------------------------------------
            $result = $taskReadWs.Result
            if ($result.MessageType -eq 'Close') { break }

            Log "📥 【Chromiumから受信】 $($result.Count) バイト (EndOfMessage: $($result.EndOfMessage))" @{ForegroundColor="Magenta"}

            # Chromeから受け取った生データをパイプへ書き込む
            $pipeServer.Write($bufferWs, 0, $result.Count)

            # 🌟「これで1つのJSONが完全に終わったか？」を確認する！
            if ($result.EndOfMessage) {
                # メッセージの終わりなら、VBAが待っている「Null文字(0x00)」を追記する
                $nullByte = [byte[]]@(0)
                $pipeServer.Write($nullByte, 0, 1)
            }

            # VBAへ即座に送り出す
            $pipeServer.Flush()
            Log "📨 【パイプFlush完了】" @{ForegroundColor="Magenta"}

            # 次のWebSocket受信タスクを再セット
            $taskReadWs = $ws.ReceiveAsync($segmentWs, $cts.Token)
        }
    }
} catch {
    Write-Error "⚠️ 中継中にエラーが発生しました`n$($_.Exception.Message)"
} finally {
    # 7. 各種ハンドルをクリーンして、お片付け
    Log "🛑 通信終了"
    if ($pipeServer -ne $null) { $pipeServer.Dispose() }
    if ($ws -ne $null) { $ws.Dispose() }
    Log "🧹 各種ハンドルを解体しました。" @{ForegroundColor="DarkGreen"}
}

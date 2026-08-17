# ====================================================================================
#   WebView2ViaPowerShell - PowerShellがホストするWebView2を、名前付きパイプ経由のCDPで
#   Excel(VBA)から操作できるようにするホストスクリプトです。
#
#   StarterWebScrapingKitの`--remote-debugging-pipe`、WebSocketViaPowerShellの名前付きパイプ
#   サーバーと同じ作法(vbNullChar区切り)で通信します。
#
#   【起動順】本スクリプトが名前付きパイプのサーバー（待受け）となり、Excelがクライアントとして接続します。
#   通常はExcel側の`CDPCoreWebView2Helpers.StartAndConnect`から自動起動されるため、手動実行は不要です。
#
#   【複数タブ(Target)管理について】
#   WebView2の`CoreWebView2`は1インスタンス=1ページであり、Chromium本来の`--remote-debugging-pipe`が
#   提供するブラウザレベルの`Target`ドメイン(複数タブの生成・列挙)はWebView2の.NET APIには存在しません。
#   （なお、1ページ内のiframeセッションについては`Target.setAutoAttach`/`attachToTarget`が実際にネイティブ
#   　対応しており、`CallDevToolsProtocolMethodForSessionAsync`はこのiframeセッション向けの正規APIです。）
#   そのため、複数の独立したWebView2インスタンス(＝タブ相当)をまたぐ`Target.*`コマンドだけは、本スクリプトが
#   横取りして疑似実装します(下記のTargetManager)。
#
#   【重要な実装上の注意】
#   WebView2の非同期API(`...Async`)の完了通知は、生成スレッドのWindowsメッセージループが回っている
#   ことが前提です。そのため、本スクリプトでは`.GetAwaiter().GetResult()`のような同期待機は一切使わず、
#   すべて`.ContinueWith(...)`のコールバック形式で処理します(同期待機するとメッセージポンプが
#   止まり、デッドロックします)。
#
#   さらに、`.ContinueWith(...)`は`TaskScheduler.FromCurrentSynchronizationContext()`を明示的に
#   渡す必要があります(実機検証で判明)。理由：PowerShellプロセスは単一のランスペース(パイプライン)
#   しか持たず、`Application.Run`がそれを占有し続けます。既定のスケジューラ(ThreadPool)で継続処理を
#   実行しようとすると、別スレッドが同じランスペースの使用を試みて事実上デッドロックします。
#   `FromCurrentSynchronizationContext()`を使うことで、継続処理がポンプスレッド上で実行されるため
#   (WebView2自身の`ProcessFailed`等のイベントと同じ扱いになる)、この問題を回避できます。
# ====================================================================================
param(
    [string]$PipeName       = "WebView2CDP",
    [string]$UserDataFolder = "$env:LOCALAPPDATA\WebView2ViaPowerShell\Profile"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Log($msg, $AddArg = @{}) {
    $time = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff")
    Write-Host "[$time] $msg" @AddArg
}

#------------------------ 1. WebView2 DLL読み込み ------------------------
# 追加ダウンロード不要のため、ExcelのPowerQueryアドインに同梱されてるWebView2 SDKを流用します
$wv2Dir = "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin"

if (-not (Test-Path (Join-Path $wv2Dir "WebView2Loader.dll"))) {
    Write-Error "❌ WebView2Loader.dll が見つかりません: $wv2Dir`nOfficeのバージョン/インストール場所を確認してください。"
    exit 1
}

# WebView2Loader.dll(ネイティブ)を確実に見つけさせるため、探索パスに追加しておく
Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class NativeDllPath {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool SetDllDirectory(string lpPathName);
}
"@
[NativeDllPath]::SetDllDirectory($wv2Dir) | Out-Null

# `Microsoft.Web.WebView2.Core.dll`のみで完結します(WinFormsラッパーの`WebView2`コントロールは使わず、
# `CoreWebView2Environment.CreateCoreWebView2ControllerAsync(hwnd)`で素のFormへ直接紐付けるため不要)
Add-Type -Path (Join-Path $wv2Dir "Microsoft.Web.WebView2.Core.dll")

Log "☑️ WebView2 DLLの読み込みが完了しました ($wv2Dir)" @{ForegroundColor = "Green" }

#------------------------ 2. 名前付きパイプサーバー作成 ------------------------
try {
    $pipeOptions = [System.IO.Pipes.PipeOptions]::WriteThrough -bor [System.IO.Pipes.PipeOptions]::Asynchronous
    $pipeServer = [System.IO.Pipes.NamedPipeServerStream]::new(
        $PipeName,
        [System.IO.Pipes.PipeDirection]::InOut,
        1,
        [System.IO.Pipes.PipeTransmissionMode]::Byte,
        $pipeOptions,
        1MB,
        1MB
    )
} catch {
    Write-Error "💫 名前付きパイプ サーバー作成中にエラーが発生しました。`n$($_.Exception.Message)"
    exit 1
}

$script:recvStream = New-Object System.IO.MemoryStream
$writeLock = New-Object object

#------------------------ 3. 送信(フレーミング) ------------------------
# * 機能：CDPレスポンス/イベントのオブジェクトをJson化し、vbNullChar(0x00)区切りでパイプへ書き込みます
# * 注意：WebView2オブジェクトを触らないため、任意のスレッドから呼んでよい。複数の非同期応答が
#         同時に完了しうるので`Monitor`で直列化する(フレーム混線防止)
function Write-FramedMessage($Envelope) {
    $json = $Envelope | ConvertTo-Json -Depth 20 -Compress
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $json.ToCharArray()) {
        [void]$sb.Append($ch)
    }
    $asciiJson = $sb.ToString()

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($asciiJson)
    [System.Threading.Monitor]::Enter($writeLock)
    try {
        if ($pipeServer.IsConnected) {
            $pipeServer.Write($bytes, 0, $bytes.Length)
            $pipeServer.Write([byte[]]@(0), 0, 1)
            $pipeServer.Flush()
        }
    } finally {
        [System.Threading.Monitor]::Exit($writeLock)
    }
}

#------------------------ 4. Target疑似ドメイン管理(TargetManager) ------------------------
$script:Targets           = @{}    # targetId(合成) -> @{ Form; Controller; Core; RegisteredEvents }
$script:SyntheticSessions = @{}    # sessionId(合成、タブ用) -> targetId
$script:DefaultTargetId   = $null
$script:AutoAttachTabs    = $false

function Get-TargetInfo([string]$targetId) {
    $t = $script:Targets[$targetId]
    return @{ targetId = $targetId; type = "page"; title = ""; url = [string]$t.Core.Source; attached = [bool]$t.Attached }
}

function New-SyntheticSession([string]$targetId) {
    $sessId = [Guid]::NewGuid().ToString("N")
    $script:SyntheticSessions[$sessId] = $targetId
    return $sessId
}

# * 機能：指定targetIdに紐づく合成sessionId(アタッチ済みなら)を返します
# * 注意：ネイティブイベントのデリゲート経由で直接`$script:Targets[...]`を参照すると、実機で
#         "null配列にインデックスを付けることはできません"という例外が発生することを確認したため、
#         (`Close-WebView2Target`等と同様に)必ず名前付き関数経由でアクセスする
function Get-TargetSessionId([string]$targetId) {
    $t = $script:Targets[$targetId]
    if (-not $t) { return $null }
    return $t.SessionId
}

# * 機能：`Runtime.executionContextCreated`のうち、「このtargetのメインフレーム」に該当する本物の
#         実フレームIDを確定/記憶し、返します
# * 詳細：広告/トラッキング等を含む実サイトでは、メインフレーム以外にも多数のiframeがそれぞれ
#         「isDefault:true, type:"default"」の実行コンテキストを持つため(実機で確認済み。keisan.siteで
#         1回のナビゲーションにつき10件以上発生)、単純に「defaultなら全部メインフレーム扱い」にすると、
#         頻繁に生成/破棄されるiframeのcontextIdで上書きされてしまい、実行時には既に破棄済みの
#         contextIdを使おうとして`Runtime.callFunctionOn`が失敗する(実機で確認済み)。
#         そのため、このtargetで最初に観測した`frameId`を「本物のメインフレームID」として記憶し、
#         以降はそれと一致するイベントだけをメインフレーム由来として扱う。
function Resolve-MainFrameId([string]$targetId, [string]$observedFrameId) {
    $t = $script:Targets[$targetId]
    if (-not $t) { return $observedFrameId }
    if (-not $t.RealMainFrameId) { $t.RealMainFrameId = $observedFrameId }
    return $t.RealMainFrameId
}

# * 機能：非表示のホストFormを1つ作り、`CreateCoreWebView2ControllerAsync`で紐付けたController/Coreを
#         `$script:Targets`へ登録します。ポップアップ(NewWindowRequested)経由でも共通で使います
function Complete-NewTarget([System.Windows.Forms.Form]$form, $controller, [string]$url, [scriptblock]$onReady) {
    $controller.Bounds = $form.ClientRectangle
    $core = $controller.CoreWebView2

    $targetId = [Guid]::NewGuid().ToString("N")
    $script:Targets[$targetId] = @{ Form = $form; Controller = $controller; Core = $core; RegisteredEvents = @{}; Attached = $false; SessionId = $null; RealMainFrameId = $null }
    if (-not $script:DefaultTargetId) { $script:DefaultTargetId = $targetId }

    $form.Text = "WebView2 - $($targetId.Substring(0, 8))"

    # ウィンドウのリサイズに追従して、WebView2の表示領域も更新する
    $form.Add_ClientSizeChanged({ $controller.Bounds = $form.ClientRectangle }.GetNewClosure()) | Out-Null

    # ウィンドウを手動で閉じられた場合も、TargetManagerの状態を正しく後片付けする(タブを閉じたのと同じ扱い)
    $form.Add_FormClosing({ Close-WebView2Target $targetId }.GetNewClosure()) | Out-Null

    Register-StandardHandlers $targetId
    if ($url) { $core.Navigate($url) }
    if ($onReady) { & $onReady $targetId }

    return $targetId
}

# * 機能：各Targetに共通の`ProcessFailed`/`NewWindowRequested`を配線します
# * 注意：`NewWindowRequested`はポップアップ(window.open等)を、新規の疑似Targetとして捕捉します。
#         `GetDeferral`で完了を保留し、新規Controllerの準備が整ってから`e.NewWindow`をセットします
function Register-StandardHandlers([string]$targetId) {
    $core = $script:Targets[$targetId].Core
    $capturedTargetId = $targetId

    $core.add_ProcessFailed({
        param($sender, $e)
        Log "⚠️ WebView2プロセスがクラッシュしました (targetId=$capturedTargetId): $($e.ProcessFailedKind)" @{ForegroundColor = "Red" }
        Close-WebView2Target $capturedTargetId
    }.GetNewClosure()) | Out-Null

    $core.add_NewWindowRequested({
        param($sender, $e)
        $deferral = $e.GetDeferral()

        $popupForm = New-Object System.Windows.Forms.Form
        $popupForm.ShowInTaskbar = $true
        $popupForm.StartPosition = 'CenterScreen'
        $popupForm.Size = New-Object System.Drawing.Size(1200, 800)
        $popupForm.Show()

        $popCtrlTask = $script:environment.CreateCoreWebView2ControllerAsync($popupForm.Handle)
        $popCtrlTask.ContinueWith([Action[System.Threading.Tasks.Task]]{
            param($t)
            try {
                if ($t.IsFaulted) {
                    Log "❌ ポップアップ用WebView2の作成に失敗しました: $($t.Exception.InnerException.Message)" @{ForegroundColor = "Red" }
                    $popupForm.Dispose()
                    return
                }
                $popTargetId = Complete-NewTarget $popupForm $t.Result $null $null
                $e.NewWindow = $script:Targets[$popTargetId].Core

                Write-FramedMessage (@{ method = "Target.targetCreated"; params = @{ targetInfo = (Get-TargetInfo $popTargetId) } })
                if ($script:AutoAttachTabs) {
                    $sessId = New-SyntheticSession $popTargetId
                    Write-FramedMessage (@{ method = "Target.attachedToTarget"; params = @{ sessionId = $sessId; targetInfo = (Get-TargetInfo $popTargetId); waitingForDebugger = $false } })
                }
            } finally {
                $deferral.Complete()
            }
        }.GetNewClosure(), $script:uiScheduler) | Out-Null
    }.GetNewClosure()) | Out-Null
}

# * 機能：新規のWebView2ターゲット(タブ相当)を非同期に作成します
# * 注意：`CreateCoreWebView2ControllerAsync`の完了を待たず即座に関数を抜けます。準備ができたら
#         `$onReady`コールバックが(ポンプスレッド上で)呼ばれます
function New-WebView2Target([string]$url, [scriptblock]$onReady) {
    $targetForm = New-Object System.Windows.Forms.Form
    $targetForm.ShowInTaskbar = $true
    $targetForm.StartPosition = 'CenterScreen'
    $targetForm.Size = New-Object System.Drawing.Size(1200, 800)
    $targetForm.Show()

    $ctrlTask = $script:environment.CreateCoreWebView2ControllerAsync($targetForm.Handle)
    $ctrlTask.ContinueWith([Action[System.Threading.Tasks.Task]]{
        param($t)
        if ($t.IsFaulted) {
            Log "❌ 新規WebView2ターゲットの作成に失敗しました: $($t.Exception.InnerException.Message)" @{ForegroundColor = "Red" }
            $targetForm.Dispose()
            return
        }
        Complete-NewTarget $targetForm $t.Result $url $onReady | Out-Null
    }.GetNewClosure(), $script:uiScheduler) | Out-Null
}

function Close-WebView2Target([string]$targetId) {
    $t = $script:Targets[$targetId]
    if (-not $t) { return }
    try { $t.Controller.Close() } catch {}
    try { $t.Form.Dispose() } catch {}
    $script:Targets.Remove($targetId)
    if ($script:DefaultTargetId -eq $targetId) {
        $script:DefaultTargetId = if ($script:Targets.Count -gt 0) { @($script:Targets.Keys)[0] } else { $null }
    }
    Write-FramedMessage (@{ method = "Target.targetDestroyed"; params = @{ targetId = $targetId } })
}

function Stop-AllTargets {
    foreach ($tid in @($script:Targets.Keys)) { Close-WebView2Target $tid }
}

#------------------------ 5. イベント登録表(ドメイン→CDPイベント名) ------------------------
# WebView2の`GetDevToolsProtocolEventReceiver`はイベント名ごとの個別登録制のため、対応表を静的に持ち、
# `<Domain>.enable`受信時に遅延登録します(`Target.*`はTargetManagerが自前でemitするため対象外)
$script:DomainEventNames = @{
    "Page"    = @("Page.loadEventFired", "Page.frameNavigated", "Page.javascriptDialogOpening")
    "Network" = @("Network.requestWillBeSent", "Network.responseReceived", "Network.loadingFinished", "Network.loadingFailed")
    "DOM"     = @("DOM.documentUpdated", "DOM.childNodeInserted")
    "Runtime" = @("Runtime.executionContextCreated", "Runtime.consoleAPICalled")
}

function Register-EventsForDomain([string]$targetId, [string]$domain) {
    if (-not $script:DomainEventNames.ContainsKey($domain)) { return }
    $t = $script:Targets[$targetId]
    if (-not $t) { return }

    foreach ($evtName in $script:DomainEventNames[$domain]) {
        if ($t.RegisteredEvents.ContainsKey($evtName)) { continue }
        $receiver = $t.Core.GetDevToolsProtocolEventReceiver($evtName)
        $capturedEvtName = $evtName   # クロージャでの変数捕捉のため、ループ変数をローカルにコピー

        $receiver.add_DevToolsProtocolEventReceived({
            param($sender, $e)
            $paramsObj = $e.ParameterObjectAsJson | ConvertFrom-Json

            # `Runtime.executionContextCreated`のauxData.frameIdには、WebView2内部の実フレームID
            # (例: "5DAD533657C2DE91103ADA378FBA6AD9")がそのまま入っている。しかし`CDPContext.cls`は
            # このframeIdが自分の`targetID`(TargetManagerが発行した合成ID)と一致するかで有効性を
            # 判定しており、実フレームIDのままだと常に不一致でタイムアウトする(実機で確認済み)。
            # ここで合成targetIdに書き換えてから転送する。
            # ただし、広告/トラッキング等のiframeも同じ"isDefault:true"のコンテキストを持つため、
            # このtargetで最初に観測したframeId(=メインフレーム)と一致する場合のみ書き換える。
            # 一致しない(=他のiframe由来の)ものは、実フレームIDのまま転送し、書き換えない
            # (`brTab.targetID`と一致しなくなるので、`CDPContext.cls`側で自然に無視される)。
            if ($capturedEvtName -eq "Runtime.executionContextCreated" -and $paramsObj.context -and $paramsObj.context.auxData) {
                $realFrameId = $paramsObj.context.auxData.frameId
                $mainFrameId = Resolve-MainFrameId $targetId $realFrameId
                if ($realFrameId -eq $mainFrameId) {
                    $paramsObj.context.auxData.frameId = $targetId
                }
            }

            # 本家Chromiumの`--remote-debugging-pipe`は、タブにアタッチ済み(flatセッション)の場合、
            # ドメインイベント(Page.*/Network.*/Runtime.*等)に必ずトップレベルの`sessionId`を付与してくる。
            # `CDPCore.cls`はこの`sessionId`の有無で`CDPBrowserEvent`(ブラウザ向け)か`CDPContextEvent`
            # (タブ向け、`CDPContext.cls`のWithEventsが拾う)かを振り分けており、`sessionId`が無いと
            # タブ向けイベントが`CDPContext.cls`側に一切届かない(実機で確認済み)。WebView2の
            # `GetDevToolsProtocolEventReceiver`はsessionIdを付与してくれないため、ここでTargetManagerが
            # 把握してる合成sessionIdを付与して、本家Chromiumと同じ形にする。
            $envelope = @{ method = $capturedEvtName; params = $paramsObj }
            $currentSessionId = Get-TargetSessionId $targetId
            if ($currentSessionId) { $envelope.sessionId = $currentSessionId }

            Write-FramedMessage ($envelope)
        }.GetNewClosure()) | Out-Null

        $t.RegisteredEvents[$evtName] = $true
    }
}

#------------------------ 6. Target.*コマンドの疑似実装 ------------------------
function Handle-TargetCommand([long]$id, [string]$method, $paramsObj, [string]$sessionId) {
    switch ($method) {
        "Target.getTargets" {
            $infos = @($script:Targets.Keys | ForEach-Object { Get-TargetInfo $_ })
            Write-FramedMessage (@{ id = $id; result = @{ targetInfos = $infos } })
        }
        "Target.getTargetInfo" {
            # `targetId`省略時は、sessionId(合成タブセッション優先、無ければ既定タブ)から解決する
            # ※`CDPContext.refreshTargetInfo`(reattach内の生存確認)が、targetIdなし・sessionIdのみで送ってくる
            $targetId = $null
            if ($paramsObj -and $paramsObj.targetId) {
                $targetId = $paramsObj.targetId
            } elseif ($sessionId -and $script:SyntheticSessions.ContainsKey($sessionId)) {
                $targetId = $script:SyntheticSessions[$sessionId]
            } else {
                $targetId = $script:DefaultTargetId
            }

            if ($targetId -and $script:Targets.ContainsKey($targetId)) {
                Write-FramedMessage (@{ id = $id; result = @{ targetInfo = (Get-TargetInfo $targetId) } })
            } else {
                Write-FramedMessage (@{ id = $id; error = @{ message = "No target found for Target.getTargetInfo" } })
            }
        }
        "Target.createTarget" {
            $url = if ($paramsObj.url) { $paramsObj.url } else { "about:blank" }
            New-WebView2Target $url {
                param($newTargetId)
                Write-FramedMessage (@{ id = $id; result = @{ targetId = $newTargetId } })
            }.GetNewClosure()
        }
        "Target.attachToTarget" {
            $targetId = $paramsObj.targetId
            if (-not $script:Targets.ContainsKey($targetId)) {
                Write-FramedMessage (@{ id = $id; error = @{ message = "No target with id $targetId" } })
                return
            }
            $sessId = New-SyntheticSession $targetId
            $script:Targets[$targetId].Attached = $true
            $script:Targets[$targetId].SessionId = $sessId
            Write-FramedMessage (@{ id = $id; result = @{ sessionId = $sessId } })
            Write-FramedMessage (@{ method = "Target.attachedToTarget"; params = @{ sessionId = $sessId; targetInfo = (Get-TargetInfo $targetId); waitingForDebugger = $false } })
        }
        "Target.closeTarget" {
            Close-WebView2Target $paramsObj.targetId
            Write-FramedMessage (@{ id = $id; result = @{ success = $true } })
        }
        "Target.setAutoAttach" {
            $script:AutoAttachTabs = [bool]$paramsObj.autoAttach
            Write-FramedMessage (@{ id = $id; result = @{} })
        }
        "Target.setDiscoverTargets" {
            Write-FramedMessage (@{ id = $id; result = @{} })
        }
        default {
            Write-FramedMessage (@{ id = $id; error = @{ message = "Unsupported Target method: $method" } })
        }
    }
}

#------------------------ 7. CDPコマンドの解釈・実行(セッションID分岐) ------------------------
# セッションIDには2系統ある:
#   ・PowerShellが合成したもの(タブ＝別Coreインスタンス用) -> $script:SyntheticSessions に載っている
#   ・WebView2自身が発行したもの(iframeセッション用、`Target.attachedToTarget`由来) -> 載っていない
# 後者は本物のWebView2セッションなので、何も合成せずそのまま`CallDevToolsProtocolMethodForSessionAsync`
# に横流しするだけでよい
function Dispatch-CDPCommand([string]$json) {
    $req = $null
    try { $req = $json | ConvertFrom-Json } catch {
        Write-FramedMessage (@{ error = @{ message = "JSON parse error" } })
        return
    }

    $id = $req.id
    $sessionId = $req.sessionId
    $method = $req.method
    $paramsObj = $req.params

    if ($method -like "Target.*") {
        Handle-TargetCommand $id $method $paramsObj $sessionId
        return
    }

    # CDPBrowser.cls(既存クラス)が無条件で送ってくる「ブラウザレベル」コマンドのうち、
    # WebView2に相当する機能が無いものは、ここで最小限のショートカット応答を返す。
    # (`Browser.getVersion`は`.reattach`のgetBrowserInfoから、`Browser.close`は`.quit`から呼ばれる)
    if ($method -eq "Browser.getVersion") {
        $verStr = if ($script:environment) { $script:environment.BrowserVersionString } else { "0.0.0.0" }
        Write-FramedMessage (@{ id = $id; result = @{ protocolVersion = "1.3"; product = "WebView2/$verStr"; revision = "0"; userAgent = "Mozilla/5.0 (WebView2)"; jsVersion = "0" } })
        return
    }
    if ($method -eq "Browser.close") {
        Write-FramedMessage (@{ id = $id; result = @{} })
        Stop-AllTargets
        if ($pipeServer -ne $null) { $pipeServer.Dispose() }
        $script:pumpForm.Close()
        return
    }

    $core = $null
    $isSyntheticSession = $false
    $eventDomainTargetId = $script:DefaultTargetId

    if ($sessionId -and $script:SyntheticSessions.ContainsKey($sessionId)) {
        $isSyntheticSession = $true
        $eventDomainTargetId = $script:SyntheticSessions[$sessionId]
        if ($script:Targets.ContainsKey($eventDomainTargetId)) { $core = $script:Targets[$eventDomainTargetId].Core }
    } elseif ($script:DefaultTargetId -and $script:Targets.ContainsKey($script:DefaultTargetId)) {
        $core = $script:Targets[$script:DefaultTargetId].Core
    }

    if (-not $core) {
        Write-FramedMessage (@{ id = $id; error = @{ message = "No target available" } })
        return
    }

    if ($method -match '^(\w+)\.enable$') {
        Register-EventsForDomain $eventDomainTargetId $Matches[1]
    }

    $paramsJson = if ($paramsObj) { $paramsObj | ConvertTo-Json -Depth 20 -Compress } else { "{}" }

    $task = if ($sessionId -and -not $isSyntheticSession) {
        # WebView2ネイティブのiframeセッション -> ForSession版でそのまま素通し
        $core.CallDevToolsProtocolMethodForSessionAsync($sessionId, $method, $paramsJson)
    } else {
        # ページ本体、あるいは合成セッション(タブ本体) -> 素の呼び出し
        $core.CallDevToolsProtocolMethodAsync($method, $paramsJson)
    }

    $task.ContinueWith([Action[System.Threading.Tasks.Task]]{
        param($t)
        if ($t.IsFaulted) {
            Write-FramedMessage (@{ id = $id; error = @{ message = $t.Exception.InnerException.Message } })
            return
        }
        $resultObj = if ($t.Result) { $t.Result | ConvertFrom-Json } else { @{} }
        $envelope = @{ id = $id; result = $resultObj }
        if ($sessionId) { $envelope.sessionId = $sessionId }
        Write-FramedMessage ($envelope)
    }.GetNewClosure(), $script:uiScheduler) | Out-Null
}

#------------------------ 8. パイプ受信ループ(非同期) ------------------------
function Handle-IncomingBytes([byte[]]$buf, [int]$count) {
    $script:recvStream.Write($buf, 0, $count)
    $all = $script:recvStream.ToArray()
    $start = 0
    for ($i = 0; $i -lt $all.Length; $i++) {
        if ($all[$i] -eq 0) {
            if ($i -gt $start) {
                Dispatch-CDPCommand ([System.Text.Encoding]::UTF8.GetString($all, $start, $i - $start))
            }
            $start = $i + 1
        }
    }
    $script:recvStream.SetLength(0)
    if ($start -lt $all.Length) { $script:recvStream.Write($all, $start, $all.Length - $start) }
}

function Start-PipeAsyncRead {
    $buf = New-Object byte[] (1MB)
    $readTask = $pipeServer.ReadAsync($buf, 0, $buf.Length)
    $readTask.ContinueWith([Action[System.Threading.Tasks.Task]]{
        param($t)
        if ($t.IsFaulted -or $t.Result -eq 0) {
            Log "🔌 パイプが切断されました。終了します。" @{ForegroundColor = "Yellow" }
            Stop-AllTargets
            $script:pumpForm.Close()
            return
        }
        Handle-IncomingBytes $buf $t.Result
        Start-PipeAsyncRead
    }.GetNewClosure(), $script:uiScheduler) | Out-Null
}

function Start-PipeAccept {
    Log "📂 名前付きパイプ サーバー:$PipeName を起動しました。Excelからの接続を待機しています..." @{ForegroundColor = "Yellow" }
    $acceptTask = $pipeServer.WaitForConnectionAsync()
    $acceptTask.ContinueWith([Action[System.Threading.Tasks.Task]]{
        param($t)
        if ($t.IsFaulted) {
            Log "❌ パイプ接続待機中にエラーが発生しました: $($t.Exception.InnerException.Message)" @{ForegroundColor = "Red" }
            $script:pumpForm.Close()
            return
        }
        Log "✅ Excelから接続が来ました！コマンド待機中..." @{ForegroundColor = "Green" }
        Start-PipeAsyncRead
    }, $script:uiScheduler) | Out-Null
}

#------------------------ 9. WebView2環境初期化 + メッセージポンプ ------------------------
# `Application.Run`がこのスレッドをブロックしてメッセージループを回し続けます。WebView2の全ての
# 非同期呼び出しの完了通知は、このループが回っていることが前提です。
$script:pumpForm = New-Object System.Windows.Forms.Form
$script:pumpForm.ShowInTaskbar = $false
$script:pumpForm.WindowState = 'Minimized'
$script:pumpForm.Size = New-Object System.Drawing.Size(1, 1)

$script:pumpForm.Add_Shown({
    # `Application.Run`開始後、Form表示のタイミングで`WindowsFormsSynchronizationContext`が
    # 確立されるため、ここで初めて`FromCurrentSynchronizationContext`が正しく機能します
    $script:uiScheduler = [System.Threading.Tasks.TaskScheduler]::FromCurrentSynchronizationContext()

    Log "🔧 WebView2環境を初期化しています..." @{ForegroundColor = "Yellow" }
    if (-not (Test-Path $UserDataFolder)) { New-Item -ItemType Directory -Path $UserDataFolder -Force | Out-Null }

    # このSDKバージョンの`CoreWebView2EnvironmentOptions`にはパラメーターなしコンストラクターが
    # 存在しないため(実機検証で判明)、カスタマイズ不要な既定設定は`$null`で渡します。
    # 追加のブラウザ起動引数等が必要な場合は、下記のように5引数コンストラクターを使ってください：
    #   New-Object Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions(
    #       "--disable-features=msWebOOUI,msPdfOOUI", $null, $null, $false, $null)
    $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $UserDataFolder, $null)
    $envTask.ContinueWith([Action[System.Threading.Tasks.Task]]{
        param($t)
        if ($t.IsFaulted) {
            Write-Error "❌ CoreWebView2Environment.CreateAsync に失敗しました: $($t.Exception.InnerException.Message)"
            $script:pumpForm.Close()
            return
        }
        $script:environment = $t.Result
        Log "☑️ WebView2環境の初期化が完了しました" @{ForegroundColor = "Green" }

        # 最初のデフォルトタブを1つ用意しておく
        New-WebView2Target "about:blank" {
            param($firstTargetId)
            Log "☑️ 既定のWebView2タブを準備しました (targetId=$firstTargetId)" @{ForegroundColor = "Green" }
            Start-PipeAccept
        }
    }, $script:uiScheduler) | Out-Null
})

try {
    [System.Windows.Forms.Application]::Run($script:pumpForm)
} finally {
    Log "🛑 通信終了" @{ForegroundColor = "DarkGreen" }
    Stop-AllTargets
    if ($pipeServer -ne $null) { $pipeServer.Dispose() }
    Log "🧹 各種ハンドルを解体しました。" @{ForegroundColor = "DarkGreen" }
}

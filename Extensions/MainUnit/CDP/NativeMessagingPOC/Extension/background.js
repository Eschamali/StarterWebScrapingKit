// StarterWebScrapingKit CDP Bridge - background service worker
//
// 役割:
//   ・ツールバーアイコンをクリックしたタブに chrome.debugger でアタッチする
//   ・Native Messaging経由で、Excel(VBA)側の CDPCoreViaNativeMessaging.cls と生CDP-JSON文字列をやり取りする
//   ・VBA -> 拡張機能 : {id, method, params, sessionId?} 形式のコマンド
//   ・拡張機能 -> VBA : {id, result} / {id, error} 形式のコマンド結果、または {method, params, sessionId?} 形式のイベント
//
// 注意:
//   chrome.debugger.sendCommand の結果コールバックは生CDP-JSON文字列を返してくれないため、
//   VBA側が送ってきた `id` をそのままechoして、CDPCore.cls側の既存パース処理(BrowserReceivedDataCheck)が
//   Pipe/WebSocket/WebView2と全く同じ形式として扱えるように、ここで手動合成する。

const HOST_NAME = "com.starterwebscrapingkit.cdpbridge";
const PROTOCOL_VERSION = "1.3";

// tabId -> { port: chrome.runtime.Port }
const sessions = new Map();

console.log("[CDPBridge] background.js loaded. HOST_NAME=" + HOST_NAME);

function log(...args) {
  console.log("[CDPBridge]", ...args);
}

// 接続中のタブのアイコンに「DEV」バッジを表示し、見た目でも接続状態が分かるようにする
function setConnectedBadge(tabId) {
  chrome.action.setBadgeText({ tabId, text: "DEV" });
  chrome.action.setBadgeBackgroundColor({ tabId, color: "#1a73e8" });
  chrome.action.setTitle({ tabId, title: "StarterWebScrapingKit: 接続中(クリックで切断)" });
}

function clearConnectedBadge(tabId) {
  chrome.action.setBadgeText({ tabId, text: "" });
  chrome.action.setTitle({ tabId, title: "StarterWebScrapingKit: このタブをCDPブリッジに接続 / 切断" });
}

function cleanup(tabId) {
  const s = sessions.get(tabId);
  if (!s) return;
  try {
    s.port.disconnect();
  } catch (e) {
    // 既に切断済みの場合は無視
  }
  sessions.delete(tabId);
  clearConnectedBadge(tabId);
  log("cleanup done. tabId=" + tabId);
}

function safeDetach(tabId) {
  chrome.debugger.detach({ tabId }, () => {
    void chrome.runtime.lastError; // 既にdetach済みの場合のエラーは無視
  });
}

// ===== chrome.debuggerに中継せず、拡張機能側だけで処理するCDPコマンド =====
// chrome.debuggerでタブにアタッチしたセッションは、`Browser.*`のようなブラウザ全体スコープの
// コマンドを受け付けない(-32601 Method not found)ため、VBA側(CDPBrowser.getBrowserInfo等)が
// 期待する形の応答を、拡張機能側で入手できる情報から代わりに合成して返す。
// 取れる情報が無ければダミー値で埋めて「成功」扱いにする(VBA側からは失敗として見せない)。
function buildBrowserVersionResult() {
  let ua = "";
  try {
    ua = (typeof navigator !== "undefined" && navigator.userAgent) || "";
  } catch (e) {
    ua = "";
  }

  // "Chrome/123.0.0.0" や "Edg/123.0.0.0" のようなトークンをUser-Agentから拾う
  const match = ua.match(/(Edg|Chrome|Chromium)\/([\d.]+)/);
  let product = "Unknown/0.0.0.0"; // 取れなかった場合のダミー値(VBA側の`Split(br.Type, "/")(1)`が壊れないよう"/"は必ず含める)
  if (match) {
    const name = match[1] === "Edg" ? "Edge" : match[1];
    product = name + "/" + match[2];
  }

  return {
    protocolVersion: PROTOCOL_VERSION,
    product: product,
    revision: "",
    userAgent: ua || "unknown",
    jsVersion: "",
  };
}

const LOCAL_ONLY_METHODS = {
  "Browser.getVersion": buildBrowserVersionResult,
};

// chrome.debugger.sendCommand の `chrome.runtime.lastError.message` は、実際にはCDP側のエラーを
// JSON文字列化したもの(例: `{"code":-32601,"message":"..."}"`)であることが多い。
// そのまま`{code:-32000, message: <そのJSON文字列>}`に包むと二重にネストして読みにくいため、
// パースできる場合は中身の{code, message}をそのまま使う。
function normalizeDebuggerError(lastErrorMessage) {
  if (typeof lastErrorMessage === "string") {
    try {
      const parsed = JSON.parse(lastErrorMessage);
      if (parsed && typeof parsed.code === "number" && typeof parsed.message === "string") {
        return { code: parsed.code, message: parsed.message };
      }
    } catch (e) {
      // JSONとして解釈できなければ、下のフォールバックへ
    }
  }
  return { code: -32000, message: String(lastErrorMessage) };
}

function attachTab(tabId) {
  if (sessions.has(tabId)) {
    log("既にこのタブは接続済みです。tabId=" + tabId + " -> 切断します");
    safeDetach(tabId);
    cleanup(tabId);
    return;
  }

  chrome.debugger.attach({ tabId }, PROTOCOL_VERSION, () => {
    if (chrome.runtime.lastError) {
      log("attach失敗: " + chrome.runtime.lastError.message);
      return;
    }
    log("attach成功。tabId=" + tabId);

    // OOPIF/子ターゲット(iframe, worker等)も、フラットモードのsessionId付きで自動アタッチさせる
    chrome.debugger.sendCommand(
      { tabId },
      "Target.setAutoAttach",
      { autoAttach: true, flatten: true, waitForDebuggerOnStart: false },
      () => {
        if (chrome.runtime.lastError) {
          log("Target.setAutoAttach失敗: " + chrome.runtime.lastError.message);
        }
      }
    );

    let port;
    try {
      port = chrome.runtime.connectNative(HOST_NAME);
    } catch (e) {
      log("connectNative失敗: " + e.message);
      safeDetach(tabId);
      return;
    }

    const s = {
      port,
      rootTargetId: null, // このタブ(ルート)自身のtargetId。`Target.getTargetInfo`で確定させる
      rootSessionId: "root:" + tabId, // ルートタブに割り当てる仮想sessionId({tabId}指定への読み替えに使う)
      targetSessions: new Map(), // 子ターゲットの targetId -> 本物のsessionId (自動アタッチイベントから収集)
      targetInfos: new Map(), // targetId -> targetInfo (`Target.getTargets`をローカル合成するためのキャッシュ)
    };
    sessions.set(tabId, s);

    // ルート自身のtargetIdを控えておく(VBAが`Target.attachToTarget`でルートのsessionIdを要求してきた時に使う)
    chrome.debugger.sendCommand({ tabId }, "Target.getTargetInfo", {}, (info) => {
      if (chrome.runtime.lastError) {
        log("Target.getTargetInfo失敗: " + chrome.runtime.lastError.message);
      } else if (info && info.targetInfo) {
        s.rootTargetId = info.targetInfo.targetId;
        s.targetInfos.set(info.targetInfo.targetId, info.targetInfo);
      }
    });

    // VBA -> 拡張機能 のコマンド中継
    port.onMessage.addListener((msg) => {
      if (!msg || typeof msg.method !== "string") {
        log("不正なメッセージを無視しました: " + JSON.stringify(msg));
        return;
      }

      // chrome.debuggerに中継できないコマンドは、拡張機能側だけで応答を合成する
      if (Object.prototype.hasOwnProperty.call(LOCAL_ONLY_METHODS, msg.method)) {
        log(msg.method + " はローカル応答で処理します(chrome.debuggerには中継しません)");
        port.postMessage({ id: msg.id, sessionId: msg.sessionId, result: LOCAL_ONLY_METHODS[msg.method]() });
        return;
      }

      // `Target.attachToTarget`はchrome.debugger自身がセッション管理をしてる都合上、手動では弾かれる
      // (Not allowed)ため、`Target.setAutoAttach`による自動アタッチで既に把握済みのsessionIdを
      // そのまま返す(ルート自身の場合は、仮想sessionIdを返す)
      if (msg.method === "Target.attachToTarget") {
        const wantedTargetId = msg.params && msg.params.targetId;
        let sid = null;
        if (wantedTargetId && wantedTargetId === s.rootTargetId) {
          sid = s.rootSessionId;
        } else if (wantedTargetId && s.targetSessions.has(wantedTargetId)) {
          sid = s.targetSessions.get(wantedTargetId);
        }

        if (sid) {
          log("Target.attachToTarget をローカル応答で処理します。targetId=" + wantedTargetId + " -> sessionId=" + sid);
          port.postMessage({ id: msg.id, sessionId: msg.sessionId, result: { sessionId: sid } });
        } else {
          log("Target.attachToTarget: 未知のtargetId(自動アタッチ未検出): " + wantedTargetId);
          port.postMessage({
            id: msg.id,
            sessionId: msg.sessionId,
            error: { code: -32000, message: "Unknown targetId (not yet auto-attached): " + wantedTargetId },
          });
        }
        return;
      }

      // ルートの仮想sessionId宛ての`Target.detachFromTarget`は、chrome.debugger的には実在しない
      // セッションなので中継すると`No session with given id`になる。こちらで何もせず成功扱いにする
      // (実際のタブのdetachは、VBA側からの切断操作やタブクローズ等、別経路で行われる)
      if (msg.method === "Target.detachFromTarget" && msg.params && msg.params.sessionId === s.rootSessionId) {
        log("Target.detachFromTarget(ルート仮想session)をローカル応答で処理します。");
        port.postMessage({ id: msg.id, sessionId: msg.sessionId, result: {} });
        return;
      }

      // `Target.getTargets`もブラウザ全体スコープのコマンドのため同様に弾かれる(Not allowed)。
      // このタブ配下で把握済みの範囲(ルート自身+自動アタッチ済みの子ターゲット)だけを合成して返す
      // (`Target.getTargets`のようなブラウザ全体の一覧ではない点に注意)
      if (msg.method === "Target.getTargets") {
        log("Target.getTargets をローカル応答で処理します。known targets=" + s.targetInfos.size);
        port.postMessage({
          id: msg.id,
          sessionId: msg.sessionId,
          result: { targetInfos: Array.from(s.targetInfos.values()) },
        });
        return;
      }

      // ルートタブの仮想sessionId宛てのコマンドは、chrome.debugger的には`{tabId}`指定に読み替える
      let target;
      if (!msg.sessionId || msg.sessionId === s.rootSessionId) {
        target = { tabId };
      } else {
        target = { sessionId: msg.sessionId };
      }

      chrome.debugger.sendCommand(target, msg.method, msg.params || {}, (result) => {
        if (chrome.runtime.lastError) {
          port.postMessage({
            id: msg.id,
            sessionId: msg.sessionId,
            error: normalizeDebuggerError(chrome.runtime.lastError.message),
          });
        } else {
          port.postMessage({ id: msg.id, sessionId: msg.sessionId, result: result || {} });
        }
      });
    });

    port.onDisconnect.addListener(() => {
      log(
        "NativeMessagingホストが切断されました。tabId=" +
          tabId +
          (chrome.runtime.lastError ? " (" + chrome.runtime.lastError.message + ")" : "")
      );
      sessions.delete(tabId);
      clearConnectedBadge(tabId);
      safeDetach(tabId);
    });

    setConnectedBadge(tabId);
    log("Native Messagingホストに接続しました。tabId=" + tabId);
  });
}

// ツールバーアイコンのクリックで、対象タブへの接続/切断をトグルする
chrome.action.onClicked.addListener((tab) => {
  log("action.onClicked発火。tabId=" + tab.id);
  attachTab(tab.id);
});

// ツールバーアイコンが見つからない/クリックできない場合の代替トリガー(右クリックメニュー)
chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create(
    {
      id: "starterwebscrapingkit-cdp-bridge-toggle",
      title: "StarterWebScrapingKit: このタブをCDPブリッジに接続 / 切断",
      contexts: ["page", "action"],
    },
    () => {
      if (chrome.runtime.lastError) {
        log("contextMenus.create失敗: " + chrome.runtime.lastError.message);
      } else {
        log("右クリックメニューを登録しました。");
      }
    }
  );
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  log("contextMenus.onClicked発火。tabId=" + tab.id);
  attachTab(tab.id);
});

// chrome.debugger -> VBA のイベント中継 (ルートセッション、子セッション両方ここに集約される)
chrome.debugger.onEvent.addListener((source, method, params) => {
  const s = sessions.get(source.tabId);
  if (!s) return;

  // 子ターゲットの自動アタッチ/デタッチ/情報更新を検知して、`Target.attachToTarget`/`Target.getTargets`を
  // ローカル応答するためのキャッシュ(targetId -> sessionId / targetInfo)を更新しておく
  if (method === "Target.attachedToTarget" && params && params.targetInfo && params.sessionId) {
    s.targetSessions.set(params.targetInfo.targetId, params.sessionId);
    s.targetInfos.set(params.targetInfo.targetId, params.targetInfo);
  } else if (method === "Target.targetInfoChanged" && params && params.targetInfo) {
    if (s.targetInfos.has(params.targetInfo.targetId)) {
      s.targetInfos.set(params.targetInfo.targetId, params.targetInfo);
    }
  } else if (method === "Target.detachedFromTarget" && params && params.targetId) {
    s.targetSessions.delete(params.targetId);
    s.targetInfos.delete(params.targetId);
  }

  // ルート自身(chrome.debugger的にはsessionId無し)のイベントには、こちらで払い出した仮想sessionIdを
  // 付与し、VBA側で`CDPContext`相当として一貫して扱えるようにする
  const sid = source.sessionId || s.rootSessionId;
  const envelope = { method, params, sessionId: sid };
  s.port.postMessage(envelope);
});

// デバッガのinfobarで「キャンセル」された場合や、タブがクラッシュした場合
chrome.debugger.onDetach.addListener((source, reason) => {
  log("デバッガがdetachされました。tabId=" + source.tabId + ", reason=" + reason);
  cleanup(source.tabId);
});

// タブが閉じられた場合
chrome.tabs.onRemoved.addListener((tabId) => {
  if (sessions.has(tabId)) {
    safeDetach(tabId);
    cleanup(tabId);
  }
});

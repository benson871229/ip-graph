# ip-graph

零依賴的網路流量關聯圖工具。單一 HTML 檔,用瀏覽器打開即可使用,**不需安裝任何套件、離線可用**。

適合在受管制或 air-gapped 的分析環境中,快速把封包或連線紀錄視覺化成 IP 關聯圖。

## 特色

- **零安裝** — 純原生 JavaScript + SVG,不連任何 CDN,只要有瀏覽器就能跑
- **多種輸入** — pcap / pcapng / Zeek `conn.log` / Kibana CSV / Elasticsearch 聚合 JSON
- **N-hop 關聯分析** — 指定聚焦 IP 與跳數(1–3),節點依跳數分層排列,一眼看出關聯層次
- **有向圖** — 箭頭表示連線方向,線條粗細依流量大小
- **資產盤點表對應** — 節點顯示資產名稱,並標記出**不在資產表中的主機**(疑似未列管)
- **威脅情資標記** — 貼上惡意 IP 清單,命中者標紅發光,相關連線同步變紅
- **可匯出 SVG** — 直接放進報告
- 支援滑鼠與觸控(拖曳、縮放、平移)

## 使用方式

1. 用瀏覽器打開 `ip-graph.html`
2. 拖入檔案(pcap 或 Zeek conn.log 等),或展開「改用貼上資料」貼上連線資料
3. 輸入聚焦 IP,拉動 hop 滑桿決定要看幾層關聯

想先看效果:展開「改用貼上資料」按「載入範例」。

### 支援的輸入格式

| 格式 | 說明 |
|------|------|
| pcap / pcapng | 直接在瀏覽器解析,自動彙總對話並依 port 標示服務 |
| Zeek `conn.log` | 支援原始 `#fields` TSV,自動對應 `id.orig_h` / `id.resp_h`,流量取 `orig_ip_bytes` + `resp_ip_bytes` |
| Kibana / SOC 匯出 CSV | 依欄名自動辨識 `source.ip` / `destination.ip` / `network.bytes` |
| Elasticsearch 聚合 JSON | 支援 composite 與巢狀 terms 聚合結果 |
| 自訂 CSV | `來源,目的,權重,標籤`(後兩欄選填) |

### 資產盤點表格式

```
10.20.0.5,PACS-01,影像伺服器
10.20.0.6,HL7-ENGINE,介接引擎
```

不在此清單中的 IP 會以紫色虛線圈標記,可用來發現未列管資產。

### 威脅情資格式

每行一個 IP,支援 `IP # 說明` 或 `IP,說明`,`#` 開頭的行會忽略:

```
45.140.17.3
185.220.101.5 known-c2
```

可從公開情資來源產生,例如:

```bash
curl -s https://raw.githubusercontent.com/stamparm/ipsum/master/ipsum.txt \
  | awk '$2>=3{print $1}' > malicious-ips.txt
```

## 從 API 取得資料(實驗性 · 本分支)

> 這是 `claude/api-so-kibana` 分支的附加功能,主線的零依賴/離線行為完全不變。

展開左側「**從 API 取得資料 · SO / Kibana / ES**」,直接向 Elasticsearch 的 `_search`
送 composite 聚合(來源 × 目的,`sum(bytes)`),回應沿用既有的 ES 聚合解析器繪圖。
支援兩種模式:

| 模式 | 端點 | 說明 |
|------|------|------|
| Elasticsearch 直連 | `https://<host>:9200` | 送 `POST /<index>/_search` |
| Kibana console proxy | `https://<host>:5601` | 送 `POST /api/console/proxy?path=<index>/_search&method=POST`,自動帶 `kbn-xsrf` |

> **Security Onion 使用者注意**:SO 2.x 把所有服務藏在 **443 的反向代理**後面,
> Kibana 掛在網址路徑下(例如 `https://<so>/kibana/app/...`),而 **Elasticsearch 的 9200
> 通常不對分析師開放**。所以請走 **Kibana console proxy 模式**,並把 base path 放進端點:
> 端點填 `https://<so>/kibana`(不是 `:9200`、也不是 `:5601`)。

欄位預設走 ECS(`source.ip` / `destination.ip` / `network.bytes`);Zeek 原始欄位可改成
`id.orig_h` / `id.resp_h` / `orig_ip_bytes`。認證填在「認證」欄:

- API key:`ApiKey <base64>`
- 帳密:`Basic 帳號:密碼`(工具會自動 base64,不必自己編)

超過 5000 對連線會截斷,請縮小時間範圍或加 KQL 過濾。**取回後的 hop / 最小權重 /
資產表 / 威脅情資標記全部照常適用。**

### 推薦:用 `get-so-graph.ps1` 帶帳密查詢(免 CORS)

如果你有 SO/ES 帳密,**最省事的做法是不走瀏覽器**,改用隨附的 PowerShell 抽取腳本:
它用帳密(或 API key)直接向 ES 查詢、在伺服器端彙總,輸出成 CSV 拖進工具。
因為送請求的是 PowerShell 不是瀏覽器,**完全沒有 CORS 問題**。

```powershell
# Security Onion(Kibana 藏在 443 的 /kibana 路徑下)— 最常見的情況
# 省略 -Password 會提示輸入,不留在命令列歷史
# 索引:SO 2.4 用 logs-*;SO 2.3 用 *:so-* 或 so-*。填錯時腳本會自動列出實際存在的索引
.\get-so-graph.ps1 -Server https://10.x.x.x/kibana -Mode kibana -Username analyst `
    -Index "logs-*" -Since now-24h -OutFile graph.csv -SkipCertCheck

# ES 直連(9200 有對你開放時才適用)
.\get-so-graph.ps1 -Server https://so:9200 -Username analyst `
    -Index "logs-*" -Since now-24h -OutFile graph.csv -SkipCertCheck

# 用 API key、Zeek 原始欄位、聚焦某台主機
.\get-so-graph.ps1 -Server https://10.x.x.x/kibana -Mode kibana -ApiKey "AbCd12==" `
    -SrcField id.orig_h -DstField id.resp_h -BytesField orig_ip_bytes `
    -Query 'id.orig_h:10.20.0.30 OR id.resp_h:10.20.0.30' -OutFile ws.csv -SkipCertCheck
```

產出的 `graph.csv` 拖進 `ip-graph.html` 即可生圖,hop / 資產表 / 威脅標記照常適用。
瀏覽器 API 面板適合快速互動;這支腳本適合固定查詢、排程、或不想碰 CORS 的情況。

#### 第一次使用:先跑連通性測試

跑整支腳本前,先用這 6 行確認「連得到 + 認證通過」(把 `10.x.x.x` 換成你的 SO):

```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
Add-Type -TypeDefinition @"
using System.Net; using System.Net.Security; using System.Security.Cryptography.X509Certificates;
public static class TrustAllCerts { public static void Enable() {
  ServicePointManager.ServerCertificateValidationCallback =
    delegate (object s, X509Certificate c, X509Chain ch, SslPolicyErrors e) { return true; }; } }
"@
[TrustAllCerts]::Enable()     # 自簽憑證。不能用 {$true} scriptblock — 5.1 會報 SSL/TLS 信任關係錯誤
$cred=Get-Credential          # 輸入 SO/Kibana 帳密
$u=$cred.UserName;$p=$cred.GetNetworkCredential().Password
$h=@{Authorization="Basic "+[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${u}:${p}"));'kbn-xsrf'='true'}
Invoke-RestMethod -Method Post -Headers $h -Uri "https://10.x.x.x/kibana/api/console/proxy?path=_cluster/health&method=GET"
```

| 結果 | 意義 | 下一步 |
|---|---|---|
| JSON(有 `status`/`cluster_name`) | ✅ 通了 | 直接跑 `get-so-graph.ps1` |
| 401 / 403 | 帳號沒有走 API 的權限 | 在 Kibana 建一把 API key 改用 `-ApiKey` |
| 回傳 HTML(登入頁) | SO 的 SSO 登入代理攔下請求,Basic auth 進不了 443 | 見下方「SO 2.4 SSO 環境」 |
| 逾時 / 連不上 | 網路不通或路徑不對 | 確認瀏覽器開 Kibana 的完整網址,base path 照抄 |

腳本已能自動辨識 HTML 登入頁並印出對應解法(不會再誤判成權限不足)。

#### SO 2.4 SSO 環境:正規解法是開放 9200 直連

SO 2.4 用網頁 SSO(登入頁)保護 443 上的所有服務 — Basic auth 與 API key
都過不了那層,這不是你帳號的問題。官方作法是把 **Elasticsearch 9200** 開放給你的
分析機,然後**直連**(SOC 帳號會同步成 Elasticsearch 帳號,同一組帳密可用):

```bash
# 在 SO 主機的 console 上執行(需要管理員)
sudo so-firewall includehost elasticsearch_rest 你的分析機IP
```

開通後改用 ES 直連模式(去掉 -Mode kibana):

```powershell
.\get-so-graph.ps1 -Server https://SO主機IP:9200 -Username 你的SO帳號 `
    -Index "logs-*" -Since now-24h -OutFile graph.csv -SkipCertCheck
```

### 告警疊圖 · 把 Suricata 告警畫上節點

流量圖只回答「誰跟誰講話」;疊上告警,圖就從流量圖變成**風險圖** —— 節點依告警嚴重度上色,
一眼看出哪些主機該先看,以及它連到誰。步驟:

1. 用 `get-so-alerts.ps1` 帶帳密查 SO 的 Suricata 告警,輸出 `ip,嚴重度,告警名稱`:

   ```powershell
   .\get-so-alerts.ps1 -Server https://10.x.x.x/kibana -Mode kibana -Username analyst `
       -Since now-24h -OutFile alerts.csv -SkipCertCheck
   ```

   預設過濾 `event.dataset:alert`、嚴重度取 `event.severity`、名稱取 `rule.name`;
   不同 SO 版本可用 `-AlertFilter` / `-SevField` / `-SigField` 調整。

2. 把 `alerts.csv` 內容整份貼進工具的「**告警疊圖 · Suricata**」欄位。
3. 勾選「**告警風險模式(依嚴重度上色)**」。

節點會依嚴重度上色(1 紅=最嚴重、2 橘、3 黃,無告警的暗掉),右上角小徽章顯示告警筆數,
critical 節點發光、相關連線變紅,滑過節點可看到告警名稱。左下角列出有告警的主機。
嚴重度沿用 Suricata 慣例(1 最嚴重)。

> 這一步展示了 API 相對「只畫 conn.log 流量」的價值:告警那層資料在 Elasticsearch 裡,
> `zeek-cut | awk` 的 conn.log 沒有,所以只有接 API 才畫得出這張風險圖。

### ⚠️ CORS(瀏覽器 API 面板才會遇到)

瀏覽器從 `file://` 或不同來源打 ES/Kibana 會被 CORS 擋(錯誤訊息通常是 `Failed to fetch`)。
三選一解決:

1. **ES 開 CORS**(推薦):在 `elasticsearch.yml` 加

   ```yaml
   http.cors.enabled: true
   http.cors.allow-origin: "*"            # 或指定來源
   http.cors.allow-headers: Authorization,Content-Type,kbn-xsrf
   http.cors.allow-methods: GET,POST,OPTIONS
   ```

2. **同源部署**:把 `ip-graph.html` 放到與 ES/Kibana 同一台、同一 port 提供,就不受 CORS 限制。
3. **本機小代理**:用隨附的零安裝 PowerShell 代理 `cors-proxy.ps1` 轉發(維持零安裝):

   ```powershell
   # ES 自簽憑證常見,加 -SkipCertCheck;要帳密不進瀏覽器可加 -InjectAuth
   .\cors-proxy.ps1 -Backend https://securityonion:9200 -SkipCertCheck
   ```

   然後在工具的「端點 URL」填 `http://localhost:8080`,索引/欄位照舊。
   代理會補上 CORS 標頭並把 `/<index>/_search` 轉到後端;後端的錯誤狀態也會原樣帶回。

> 安全提醒:認證資訊只存在當下的瀏覽器分頁,不會寫進工作階段檔;但 `allow-origin: "*"`
> 會放寬 ES 的跨來源限制,正式環境請改成指定來源。

## 大流量環境的建議用法

瀏覽器讀 pcap 是整檔載入記憶體,適合數百 MB 以內。流量很大的環境建議**在後端先彙總**,只把結果丟進工具:

```bash
# 從 Zeek conn.log 彙總某台主機的關聯,取前 300 大
cat /nsm/zeek/logs/current/conn.log \
  | zeek-cut id.orig_h id.resp_h orig_ip_bytes resp_ip_bytes \
  | awk -v ip=10.20.0.30 '$1==ip||$2==ip{k=$1","$2; s[k]+=$3+$4} END{for(i in s) print i","s[i]}' \
  | sort -t, -k3 -rn | head -300 > graph.csv
```

輸出通常只有幾 KB,拖進工具即可瞬間生圖。

節點數超過 800 時會自動簡化排版避免瀏覽器卡頓;可透過降低 hop 或提高「最小權重」讓圖更清楚。

## 限制

- 威脅情資為**精確 IP 比對**,不支援 CIDR 網段
- 協定標籤依 port 判斷,非深度封包解析
- 力導向排版為 O(n²),節點過多時效能會下降

## 授權

MIT

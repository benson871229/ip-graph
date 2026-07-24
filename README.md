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

欄位預設走 ECS(`source.ip` / `destination.ip` / `network.bytes`);Zeek 原始欄位可改成
`id.orig_h` / `id.resp_h` / `orig_ip_bytes`。認證填在「認證」欄:

- API key:`ApiKey <base64>`
- 帳密:`Basic 帳號:密碼`(工具會自動 base64,不必自己編)

超過 5000 對連線會截斷,請縮小時間範圍或加 KQL 過濾。**取回後的 hop / 最小權重 /
資產表 / 威脅情資標記全部照常適用。**

### ⚠️ CORS(最常見的卡關點)

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

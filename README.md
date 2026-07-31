# ip-graph

醫療 / SOC 網路分析工具組。**全部零依賴、免安裝、離線可用** —— 單一 HTML 檔用瀏覽器打開,
PowerShell 腳本用 Windows 內建的就能跑。適合受管制或 air-gapped 的分析環境。

## 目錄結構

| 資料夾 | 內容 |
|---|---|
| **`graph/`** | `ip-graph.html` — 網路流量關聯圖工具(主要產出) |
| **`assets/`** | 資產盤點:從防火牆設定擷取 IP↔設備名稱,並合併進既有 Excel 資產表 |
| **`hunting/`** | `threat-hunting-kql.md` — Security Onion / Kibana 的威脅獵捕 KQL 手冊 |
| **`trend/`** | `daily-compare.html` — 多天流量比較:累積基準、首見日、出現天數、成長趨勢 |
| **`intel/`** | `malicious-ip.txt` — 高信度惡意 IP 清單,供關聯圖標記威脅 |

---

# graph/ — 網路流量關聯圖

零依賴的網路流量關聯圖工具。單一 HTML 檔,用瀏覽器打開即可使用。

## 特色

- **零安裝** — 純原生 JavaScript + SVG,不連任何 CDN,只要有瀏覽器就能跑
- **多種輸入** — pcap / pcapng / Zeek `conn.log` / Kibana CSV / Elasticsearch 聚合 JSON
- **N-hop 關聯分析** — 指定聚焦 IP 與跳數(1–3),節點依跳數分層排列,一眼看出關聯層次
- **有向圖** — 箭頭表示連線方向,線條粗細依流量大小
- **資產盤點表對應** — 節點顯示資產名稱,並標記出**不在資產表中的主機**(疑似未列管)
- **威脅情資標記** — 貼上惡意 IP 清單,命中者標紅發光,相關連線同步變紅
- **連線明細** — 單擊節點,右側列出「誰連進來 / 連出去」與各自使用的 port
- **服務識別** — 內建 149 個埠對照(含 DICOM/HL7、企業防毒、OT/ICS、常見惡意用埠),
  未知埠顯示「埠號 unknown」,ICMP 依 type 標示
- **基準比較** — 設為基準後再匯入新資料,新增的節點與流量標成綠色;可只顯示新增
- **可匯出 SVG** — 直接放進報告
- 支援滑鼠與觸控(拖曳、縮放、平移)

## 使用方式

1. 用瀏覽器打開 `graph/ip-graph.html`
2. 拖入檔案(pcap 或 Zeek conn.log 等),或展開「改用貼上資料」貼上連線資料
3. 輸入聚焦 IP,拉動 hop 滑桿決定要看幾層關聯

想先看效果:展開「改用貼上資料」按「載入範例」。

**圖上的操作**

| 動作 | 效果 |
|---|---|
| 單擊節點 | 右側顯示連線明細(入站/出站對象與 port) |
| 雙擊節點 | 設為聚焦 IP 並重繪 |
| 拖曳節點 | 手動調整位置 |
| 滾輪 / 雙指 | 縮放;拖曳空白處平移 |
| 「清空重來」 | 清掉圖與基準,但**保留資產表與威脅情資**,方便重新匯入 |

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

---

# assets/ — 資產盤點

把防火牆設定裡的 IP↔設備名稱擷取出來,並合併進你既有的 Excel 資產表。
產出可直接貼進關聯圖的「資產盤點表」欄位,未列管主機就會在圖上標紫色虛線圈。

| 檔案 | 用途 |
|---|---|
| `get-asset.ps1` | 從 FortiGate 設定檔擷取 IP↔名稱(區段堆疊解析) |
| `get-asset-regex.ps1` | 同上,改用逐行 regex;**自動偵測編碼**,解析不到時輸出診斷 |
| `merge-assets.bas` | **Excel VBA 巨集**,在同一活頁簿內跨工作表合併資產表 |
| `merge-assets.ps1` | 同上的 PowerShell + Excel COM 版本 |

```powershell
# 擷取(解析不到時會印出讀入編碼、config/edit/set 筆數與樣本行)
.\assets\get-asset-regex.ps1 fgt.conf -OutFile assets.csv

# 位址物件多為網段時,加這兩個參數才會有輸出
.\assets\get-asset-regex.ps1 fgt.conf -IncludeSubnets -ExpandRanges -OutFile assets.csv
```

**合併進 Excel 資產表**:`merge-assets.bas` 是巨集版(不受 PowerShell COM 限制,建議優先用)。
Alt+F11 → 插入模組 → 貼上 → 改最上面的工作表名稱設定 → F5。
已有的 IP 補上設備名稱、沒有的新增一列,**所有更動一律紅字**,且不會自動存檔。

---

# hunting/ — 威脅獵捕

`threat-hunting-kql.md`:34 條給 Security Onion / Kibana 用的 KQL 獵捕查詢,
依 資產可見性 / 掃描偵察 / 橫向移動 / C2 / 資料外流 / **醫療專屬** / 惡意檔案 分類。
每條都寫了用途、如何判讀、誤報從哪來。

醫療專屬那一段是一般 SOC 手冊沒有的:明文 DICOM/HL7 盤點、醫療設備對外連線、
非授權主機存取 DICOM、醫療設備上出現 IT 協定。

> 手冊第 0 節先教你確認自己環境的欄位名稱 —— 不同 SO 版本欄位不同,
> 查詢跑出 0 筆時第一個要懷疑的就是欄名。

---

# trend/ — 多天流量比較

`daily-compare.html` 單一檔、零依賴。**一個檔案 = 一天**,拖入多天的 conn.log
(或整個 `/nsm/zeek/logs` 資料夾),日期自動從檔名或資料夾路徑判讀。

解決的是「只能比兩份」的限制 —— 現有工具比得出基準 vs 今天,
但看不出某條連線是偶發還是天天都有、哪一天開始出現、量在不在成長。

一列 = **來源IP → 目的IP : 服務**,四個判定同時算:

| 判定 | 規則 |
|---|---|
| **新增** | 首見日在基準期之後 |
| **量增** | 後段平均比前段成長 ≥3 倍(前段須 >0,否則那叫新增不叫成長) |
| **偶發** | 出現天數 < 觀察天數的 20/30/50%(可調) |
| **例行** | 什麼都沒中 —— 是預設值,不是主動判斷 |

預設**依命中訊號數排序**,可疑的自動浮到最上面。可匯出 CSV(UTF-8 BOM + CRLF)。

### 埠篩選

對外的 80/443 幾乎都是瀏覽雜訊,但**一律排除會把走 443 的外流一起藏掉**。所以:

- **保留有標記的列**(預設開):只藏什麼標記都沒中的純雜訊
- **僅排除對外部**(預設開):內網的 80/443 常是設備管理介面,留著

### 限制

**偶發只看「出現幾天」,看不出規律性** —— 每週固定跑一次的備份,
跟隨機冒出兩次的異常,同樣被標偶發。**標記是用來排序的,不是用來下結論的。**

## 授權

MIT

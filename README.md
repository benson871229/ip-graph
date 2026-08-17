# ip-graph

醫療 / SOC 網路分析工具組。**全部零依賴、免安裝、離線可用** —— 單一 HTML 檔用瀏覽器打開,
PowerShell 腳本用 Windows 內建的就能跑。適合受管制或 air-gapped 的分析環境。

## 目錄結構

| 資料夾 | 內容 |
|---|---|
| **`graph/`** | `ip-graph.html` — 網路流量關聯圖工具(主要產出) |
| **`assets/`** | 資產盤點:從防火牆設定擷取 IP↔設備名稱,並合併進既有 Excel 資產表 |
| **`hunting/`** | `threat-hunting-kql.md` — Security Onion / Kibana 的威脅獵捕 KQL 手冊 |
| **`intel/`** | 惡意 IP 清單,以及批次查 public IP 屬於誰的工具 |

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

# intel/ — IP 情資

| 檔案 | 用途 |
|---|---|
| `malicious-ip.txt` | 高信度惡意 IP 清單,貼進關聯圖的「威脅情資」欄位 |
| `Get-IpOwner.ps1` | 走 RDAP **即時**查一堆 public IP 分別屬於誰,產出 CSV |
| `Build-WhoisDbTable.ps1` | 把公開的 WHOIS-DB repo 爬成一張**離線**的「網段 → 所屬機關」對照表,並可直接比對 IP |

## 批次查 public IP 是誰

圖上的外部 IP 只是一串數字,知道是誰之後意義差很多:
**「醫療設備連到 VPS 供應商」** 比 **「醫療設備連到 203.0.113.7」** 有用得多。

```powershell
# 從檔案讀(可直接餵工具匯出的 CSV,會自動抓每行第一個 IP)
.\intel\Get-IpOwner.ps1 -InFile public-ips.txt -OutFile owners.csv

# 或直接給
.\intel\Get-IpOwner.ps1 -Ip 8.8.8.8,104.131.0.5 -OutFile owners.csv
```

用 **RDAP**(WHOIS 的現代版):HTTPS、免註冊、免 API key,查的是各區域註冊機構的權威資料。
**私有位址會自動跳過**,不送出查詢。

輸出欄位:`IP, 組織, 網段名稱, 國家, 網段, 註冊機構, 分類, 狀態`,
並自動分類成 **雲端/VPS**、**CDN**、**ISP/電信** —— **雲端/VPS 排最前面**,
因為 C2 與惡意基礎設施最常架在那裡。

**要在有網路的機器上跑**,產出的 CSV 帶回離線環境即可。

> 隱私提醒:查詢會把「目的地 IP」送到公開註冊資料庫。查的是對方的登記資料,
> 不會洩漏你的內網位址,但等於告訴外部有人在查這個 IP。這是 SOC 標準做法,
> 若你的規範不允許,請改用離線的 IP-to-ASN 資料集。

## 離線對照表:把 WHOIS-DB 整包爬成一張表

上面那支要連外。如果連查詢本身都不想送出去,`Build-WhoisDbTable.ps1` 把公開的
[WAFLogic/WHOIS-DB](https://github.com/WAFLogic/WHOIS-DB) 整包爬成一張離線對照表:

```powershell
# 有網路的機器:建表(實測 539 個網段)
.\intel\Build-WhoisDbTable.ps1 -Download -OutFile whoisdb-ranges.csv

# 沒網路也行:自己抓 ZIP 解壓後指過去
.\intel\Build-WhoisDbTable.ps1 -RepoPath C:\tmp\WHOIS-DB-main -OutFile whoisdb-ranges.csv

# 內網離線:拿建好的表比對 IP(可直接餵關聯圖匯出的 CSV)
.\intel\Build-WhoisDbTable.ps1 -Table whoisdb-ranges.csv -InFile public-ips.txt -MatchOut who.csv
```

那個 repo 的資料散在好幾層,而且每家註冊機構的欄位名稱都不一樣
(ARIN 用 `NetRange`/`Organization`、RIPE/APNIC 用 `inetnum`/`descr`、LACNIC 用 `owner`/`inetrev`,
APNIC 系甚至把網段藏在 `% Information related to` 註解行),大宗資料還壓在 zip 裡。
這支全部走一遍(zip 不落地解壓),統一成
`網段 / 起始IP / 結束IP / 組織 / 網段名稱 / 國家 / 註冊機構 / 分類`。

比對用**最長前綴**:ARIN 一次會回整條授權鏈,所以表裡同時有 `23.19.0.0/16`(Nobis)
和 `23.19.0.0/19`(Ubiquity),查 `23.19.0.5` 會給你比較精確的後者。私有位址自動跳過。

> **這張表是線索,不是權威。** repo 最後更新是 2023-05,裡面的 WHOIS 檔案時間戳是 2015,
> 全表只有 539 個網段,而且偏 2015 年前後的惡意基礎設施(所以命中率意外地不差,
> 但 `8.8.8.8` 會查成 Level 3 —— 那是它以前的持有者)。
> **命中當成一條可查的方向,查不到很正常**,要權威資料請用上面的 `Get-IpOwner.ps1`。
>
> 該 repo 沒有附授權條款,所以這裡只放腳本、不轉存它的資料;要用請自己抓。

## 授權

MIT

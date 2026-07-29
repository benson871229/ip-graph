# Threat Hunting KQL 手冊(Security Onion / Kibana)

給醫療網路 SOC 用的獵捕查詢集。每一條都說明**這是在找什麼**、**怎麼判讀**、**誤報從哪來**。

貼進 **Kibana → Discover** 的查詢列(或 SO 的 Hunt 介面)即可。語法為 **KQL**,
不是 Lucene —— Kibana 查詢列右側可切換,請確認是 KQL 模式。

---

## 0. 開始之前:先確認你的欄位名稱

**不同 SO 版本欄位不同,先做這一步再往下走。** 任何查詢跑出 0 筆,第一個要懷疑的就是欄名。

在 Discover 選好時間範圍與索引後,左側「Available fields」直接搜尋關鍵字比對;
或用下面這條看有哪些資料集:

```
event.dataset:*
```

然後在左側點 `event.dataset` 看 top values。SO 2.4 常見值:
`conn`、`dns`、`http`、`ssl`、`file`、`notice`、`alert`、`kerberos`、`smb`、`dce_rpc`。

本手冊使用 ECS 欄位(SO 2.x 預設):

| 用途 | 欄位 |
|---|---|
| 來源 / 目的 IP | `source.ip` / `destination.ip` |
| 來源 / 目的 port | `source.port` / `destination.port` |
| 傳輸協定 | `network.transport`(tcp/udp/icmp) |
| 應用協定 | `network.protocol`(zeek 的 service) |
| 流量 | `source.bytes` / `destination.bytes` / `network.bytes` |
| 資料集 | `event.dataset` |
| 告警名稱 / 嚴重度 | `rule.name` / `event.severity` |

> 若你的環境是舊版、欄位是 Zeek 原生名稱(`id.orig_h`、`id.resp_h`、`id.resp_p`),
> 把下面查詢的欄名整批替換即可,邏輯不變。

---

## A. 資產與可見性

### A1. 某台主機的全部連線(pivot 起點)

```
source.ip:10.20.0.30 or destination.ip:10.20.0.30
```

**用途**:從一個 IP 開始展開調查。SOC 告警 pivot 的第一步。
**判讀**:左側點 `destination.port` 看 top values,能一眼看出這台在講哪些服務。
**接下來**:把結果丟進 `ip-graph.html` 畫關聯圖。

### A2. 內網互連(排除對外流量)

```
event.dataset:conn and source.ip:10.0.0.0/8 and destination.ip:10.0.0.0/8
```

**用途**:只看東西向流量,橫向移動都藏在這裡。
**注意**:KQL 支援 CIDR,請換成你自己的網段。

### A3. 內網主機對外連線

```
event.dataset:conn and source.ip:10.0.0.0/8 and not destination.ip:10.0.0.0/8
    and not destination.ip:(172.16.0.0/12 or 192.168.0.0/16)
```

**用途**:醫療設備原則上不該直接對外。這條會直接列出所有違反的。
**誤報**:更新伺服器、NTP、DNS forwarder 是正常的,先建立白名單。

---

## B. 掃描與偵察

### B1. 垂直掃描:一台打同一目標的很多 port

```
event.dataset:conn and source.ip:10.20.0.99
```

**用途**:確認某台是否在做 port 掃描。
**判讀**:**這條必須搭配視覺化才有意義** —— 在 Discover 左側點 `destination.port`,
若 top values 分散且數量極多(幾十上百個不同 port),就是垂直掃描。
**更好的做法**:Visualize → `destination.port` 的 **Unique Count**,依 `source.ip` 分組。

### B2. 水平掃描:一台打很多主機的同一個 port

```
event.dataset:conn and destination.port:445 and source.ip:10.0.0.0/8
```

**用途**:找 SMB 掃描 / 蠕蟲擴散。把 445 換成 3389、22、23 各跑一次。
**判讀**:同樣看 `destination.ip` 的 unique count,一台連上百個目標就是掃描。

### B3. 失敗連線(掃描最可靠的訊號)

```
event.dataset:conn and connection.state:(S0 or REJ or RSTO)
```

**用途**:`S0` = 送了 SYN 沒有回應(目標不存在或 port 關閉),`REJ` = 被拒絕。
大量集中在同一來源 = 掃描。**這是比看 port 數更可靠的訊號**。
**欄位注意**:有些版本是 `zeek.conn.state` 或 `conn_state`,先用第 0 節確認。

### B4. ICMP 掃描(ping sweep)

```
event.dataset:conn and network.transport:icmp
```

**用途**:找 ping sweep 與網路探測。
**判讀**:一個來源對大量 `destination.ip` 送 ICMP = ping sweep。
大量 **目的不可達** 回應則是掃描的副產物。

---

## C. 橫向移動(醫療網路的重點)

### C1. SMB 橫向

```
event.dataset:conn and destination.port:(445 or 139)
    and source.ip:10.0.0.0/8 and destination.ip:10.0.0.0/8
```

**用途**:勒索軟體與 PsExec 類工具的主要途徑。
**判讀**:**工作站 → 工作站** 的 SMB 幾乎都不正常(正常是工作站 → 檔案伺服器)。
用資產表比對來源與目的的角色。

### C2. RDP / 遠端桌面

```
event.dataset:conn and destination.port:3389 and destination.ip:10.0.0.0/8
```

**用途**:攻擊者最愛的移動方式。
**誤報**:IT 維運本來就會用 RDP → 建立「允許的來源」白名單,只看白名單外的。

### C3. WinRM / PowerShell Remoting

```
event.dataset:conn and destination.port:(5985 or 5986)
```

**用途**:無檔案攻擊常用。**在醫療網路裡通常極少見**,一出現就值得查。

### C4. 管理協定從不該有的地方發起

```
event.dataset:conn and destination.port:(22 or 23 or 3389 or 5985 or 5900)
    and not source.ip:10.99.0.0/24
```

**用途**:把 `10.99.0.0/24` 換成你的**維運網段**。這條會列出所有
「非維運網段發起的管理連線」—— 訊號雜訊比極高。

### C5. Kerberos 異常

```
event.dataset:kerberos and kerberos.error_msg:*
```

**用途**:大量 Kerberos 錯誤可能是密碼噴灑或 Kerberoasting。
**注意**:欄名依版本可能是 `zeek.kerberos.error_msg`。

---

## D. C2 與惡意連線

### D1. 命中威脅情資的連線

```
event.dataset:alert and rule.name:(*CNC* or *Trojan* or *Malware* or *Backdoor*)
```

**用途**:Suricata 的 ET 規則命中。
**判讀**:先看 `rule.name` 的 top values,再對 `source.ip` 展開。
**⚠ 優先序陷阱**:KQL 的 `and` 綁得比 `or` 緊。若寫成
`event.dataset:alert and rule.name:*CNC* or rule.name:*Trojan*`,
會變成 `(alert and CNC) or Trojan` —— 第二個條件**沒有被限制在告警資料集內**。
**同層混用 `and` 與 `or` 時一定要加括號。**

### D2. 高嚴重度告警(每日必看)

```
event.dataset:alert and event.severity:1
```

**用途**:Suricata 嚴重度 1 = 最嚴重。**這條應該是你每天上班第一個跑的**。
**接下來**:把命中的來源 IP 記下來,在 ip-graph 的「威脅情資」欄位貼上,關聯圖上會直接標紅。

### D3. 長連線(Beacon / 隧道)

```
event.dataset:conn and event.duration > 3600000000000
```

**用途**:超過 1 小時的連線。C2 的 keepalive、反向 shell 都是長連線。
**單位陷阱**:ECS 的 `event.duration` 是**奈秒**,1 小時 = 3600000000000。
若你的欄位是 Zeek 原生 `duration`(秒),請改成 `duration > 3600`。

### D4. 非標準 port 的加密流量

```
event.dataset:ssl and not destination.port:(443 or 8443 or 993 or 995 or 465 or 636)
```

**用途**:TLS 走在奇怪的 port 上 = C2 常見特徵。

### D5. 自簽憑證

```
event.dataset:ssl and ssl.validation_status:*self*signed*
```

**用途**:C2 伺服器大量使用自簽憑證。
**誤報**:內部設備(印表機、IPMI、醫療設備管理介面)常是自簽 → **限定目的為外部 IP** 可大幅降噪:
```
event.dataset:ssl and ssl.validation_status:*self*signed* and not destination.ip:10.0.0.0/8
```

### D6. 少見的 JA3 指紋

```
event.dataset:ssl and ja3:*
```

**用途**:JA3 是 TLS client 指紋。**正常環境 JA3 種類很少**(瀏覽器、OS 就那幾種)。
**判讀**:看 `ja3` 的 top values,**出現次數極少的**那幾個最可疑。
惡意軟體用自己的 TLS stack → 產生獨特 JA3。

### D7. 已知惡意 / 後門用埠

```
event.dataset:conn and destination.port:(4444 or 1337 or 31337 or 12345 or 5555 or 6666)
```

**用途**:Metasploit(4444)、BackOrifice(31337)、NetBus(12345)等預設埠。
**特性**:誤報極低,一命中幾乎就是事件。

---

## E. 資料外流

### E1. 大量出站流量

```
event.dataset:conn and source.ip:10.0.0.0/8 and not destination.ip:10.0.0.0/8
    and source.bytes > 104857600
```

**用途**:單一連線送出超過 100 MB。醫療影像外流會非常明顯。
**判讀**:重點是**出多於入**。備份/更新是入多於出。

### E2. DNS 隧道:過長的查詢

```
event.dataset:dns and dns.question.name:*
```

**用途**:DNS tunneling 會用很長的子網域夾帶資料。
**判讀**:Discover 左側看 `dns.question.name` —— 找**又長又像亂碼**的域名。
**更精準**:在 Visualize 用 `dns.question.name` 的字串長度做 metric(或先用下面 E3)。

### E3. DNS 隧道:大量 NXDOMAIN

```
event.dataset:dns and dns.response_code:NXDOMAIN
```

**用途**:DNS 隧道與 DGA 惡意軟體會產生大量查不到的域名。
**判讀**:單一來源短時間內數百筆 NXDOMAIN = 高度可疑。

### E4. 連到罕見 TLD

```
event.dataset:dns and dns.question.name:(*.top or *.xyz or *.tk or *.ml or *.cf or *.gq or *.buzz)
```

**用途**:這些 TLD 在正常醫療業務裡幾乎不會出現,但惡意基礎設施大量使用。

---

## F. 醫療場域專屬(你的差異化)

> 這幾條是一般 SOC 手冊不會有的,針對 DICOM / HL7 / FHIR。

### F1. 明文 DICOM(含病患影像)

```
event.dataset:conn and destination.port:(104 or 11112)
```

**用途**:104 / 11112 是**未加密** DICOM,2762 才是 DICOM-TLS。
想反過來確認「有多少是加密的」,查 `destination.port:2762` 比較兩者數量。
**這本身就是一個 finding** —— 病患影像在網路上明文傳輸,是合規問題。
**注意**:這是**現況盤點**,不是攻擊偵測。先盤出來,再談改善。

### F2. 明文 HL7(含病歷資料)

```
event.dataset:conn and destination.port:(2575 or 6661)
```

**用途**:HL7 v2 預設明文,內含 PID(病患識別)段。
**隱私原則**:**只看有沒有這條流量,不要去解析內容**。知道「哪裡有明文 PHI」就夠了。

### F3. 醫療設備對外連線

```
event.dataset:conn and source.ip:(10.20.0.11 or 10.20.0.12 or 10.20.0.5)
    and not destination.ip:10.0.0.0/8
```

**用途**:把來源換成你資產表裡的 **modality / PACS**。
這些設備**不該直接連網際網路**。一有命中就是高優先事件(可能是被入侵,或廠商遠端維護未報備)。

### F4. 非 PACS 的主機存取 DICOM

```
event.dataset:conn and destination.port:(104 or 11112) and not source.ip:(10.20.0.11 or 10.20.0.12)
```

**用途**:把來源白名單換成你**合法的 modality 清單**。
白名單外的主機在存取 DICOM = 未授權存取病患影像。

### F5. 醫療設備上出現 IT 協定

```
event.dataset:conn and destination.ip:(10.20.0.11 or 10.20.0.12)
    and destination.port:(22 or 23 or 445 or 3389 or 5985)
```

**用途**:CT/MRI 這種設備上出現 SSH/SMB/RDP,幾乎都不是正常臨床行為。
**這是我最推薦的一條** —— 醫療設備行為極固定,誤報很低。

---

## G. 惡意檔案與 Web

### G1. 可執行檔下載

```
event.dataset:file and file.mime_type:(*executable* or *msdownload* or *octet-stream*)
```

**用途**:找從網路下載的 EXE/DLL。

### G2. 可疑 User-Agent

```
event.dataset:http and user_agent.original:(*curl* or *wget* or *python* or *powershell* or *Go-http*)
```

**用途**:正常使用者用瀏覽器。這些 UA 出現在**工作站**上通常是腳本或惡意程式。
**誤報**:伺服器上的自動化作業是正常的 → 限定來源為工作站網段。

### G3. 直接用 IP 連 HTTP(沒有域名)

```
event.dataset:http and http.virtual_host:(0* or 1* or 2* or 3* or 4* or 5* or 6* or 7* or 8* or 9*)
```

**用途**:惡意程式常直接連 IP,不經 DNS。Host 標頭是 IP 時開頭必為數字。
**為什麼不用 `not http.virtual_host:*.*`**:那寫法會把所有沒有點的主機名也一起撈進來,
而且 KQL 的萬用字元不是正規表示式,無法精準判斷「這是不是 IP」。用開頭數字比對更可靠。
**誤報**:少數合法網站的主機名以數字開頭(例如 `104.com`),量不大,人工看一眼即可。

---

## H. 自動化

### H1. 能不能在 Kibana 裡排程?

**不能。** SO 預設是 **Elastic Basic 授權,沒有 Alerting / Watcher**(那是付費功能)。
你的 Kibana 也因此沒有 Graph app。所以「在 Kibana 裡設排程告警」這條路是**封死的**。

### H2. 可行的自動化路徑

| 方式 | 可行性 | 說明 |
|---|---|---|
| Kibana Alerting | ❌ | Basic 授權沒有 |
| **SO 內建 Detections** | ✅ **推薦** | SO 2.4 可管理 Suricata / Sigma 規則,適合「持續偵測」 |
| 自寫腳本打 ES `_search` | ✅ | 需要 9200 存取權;本 repo 不附腳本,見下方注意事項 |
| ElastAlert2 | ⚠️ | 要另外部署服務,你的環境不一定允許 |

### H3. 推薦做法:把穩定的獵捕變成 Sigma 規則

本手冊裡**誤報低、可用筆數門檻判斷**的那幾條(見 H4),最適合的歸宿不是排程腳本,
而是寫成 **Sigma 規則**交給 SO 的 Detections 管理 —— 它本來就會持續比對,
命中就進 Alerts,不必另外維護排程與報表。

若你要自己寫腳本排程(PowerShell 或其他),有兩個前置條件:

1. **存取權**:SO 2.4 的 SSO 會擋掉 443 上的 Basic auth,腳本打不進 Kibana。
   需要開放 Elasticsearch 9200 給你的分析機:
   ```bash
   sudo so-firewall includehost elasticsearch_rest <你的分析機IP>
   ```
2. **查詢改寫**:本手冊是 KQL(Kibana 用)。打 ES `_search` 要放進
   `query_string` 並開 `analyze_wildcard`,或改寫成 DSL。

### H4. 哪些適合自動化、哪些不適合

**適合排程**(結果是「有/沒有」,可直接告警):
- D2 高嚴重度告警、D7 惡意用埠、F3 醫療設備對外、F5 設備上的 IT 協定
- B3 失敗連線暴增、E1 大量出站

**不適合排程**(需要人看分布、判斷脈絡):
- B1/B2 掃描(要看 unique count 分布)
- D6 JA3(要看「稀有」而非「有無」)
- E2 DNS 隧道(要看域名長相)

**原則**:能寫成「筆數 > N 就通知」的才適合自動化;需要看分布或需要人判斷「像不像」的,
留給人工獵捕。硬要自動化只會製造大量誤報,然後你就開始忽略它 —— 那比沒有告警更糟。

---

## I. 與 ip-graph 的搭配

```
Kibana 跑 KQL 找到可疑主機
    ↓
在 Discover 匯出結果,或用 Kibana 的表格複製成 CSV
    ↓
拖進 ip-graph.html,設聚焦 IP、看 2 hop
    ↓
貼上資產表 → 未列管主機自動標紫圈
貼上威脅情資 → 命中標紅
單擊節點 → 看它跟誰用什麼 port 通訊
```

**設為基準**功能特別適合搭配獵捕:今天跑完 KQL 存一份基準,明天再比,
**新增的節點與流量會標綠色** —— 這是最省力的「有什麼變了」。

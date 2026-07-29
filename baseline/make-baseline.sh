#!/bin/sh
# ---------------------------------------------------------------------------
#  make-baseline.sh — 依「服務」彙總 Zeek conn.log,產出可在 Excel 審核的流量基準
#
#  為什麼從服務出發:IP 有幾百個,但實際在用的服務埠只有幾十個。
#  不問「每個 IP 用了哪些 port」,改問「每個服務是誰在用」,IP 收斂成網段,
#  數百個 IP 就變成幾十列 —— port 的細節一個都沒少,但不必逐 IP 看。
#
#  用法:
#    ./make-baseline.sh /nsm/zeek/logs/2026-07-*/conn*.log*  > baseline.csv
#    ./make-baseline.sh -c baseline.csv /nsm/zeek/logs/current/conn.log
#
#  選項:
#    -c FILE   比對模式:只印出基準 FILE 裡沒有的列
#    -p N      網段粒度,24(預設)或 16
#    -h        說明
#
#  不需要 zeek-cut(SO 2.4 的 Zeek 在容器裡,主機不一定有):
#  直接讀 conn.log 的 #fields 表頭自己對欄位,欄位順序改版也不受影響。
#  純 sh + awk,支援 .log 與 .log.gz,可一次吃多天。
# ---------------------------------------------------------------------------
set -eu

PREFIX=24
BASE=""

usage() {
    sed -n '2,22p' "$0" | sed 's/^# \{0,1\}//'
    exit 0
}

while getopts "c:p:h" opt 2>/dev/null; do
    case "$opt" in
        c) BASE=$OPTARG ;;
        p) PREFIX=$OPTARG ;;
        h) usage ;;
        *) echo "未知選項。用 -h 看說明。" >&2; exit 2 ;;
    esac
done
shift $((OPTIND - 1))

if [ $# -eq 0 ]; then
    echo "請指定 conn.log(可多個,支援 .gz)。用 -h 看說明。" >&2
    exit 2
fi
if [ "$PREFIX" != "24" ] && [ "$PREFIX" != "16" ]; then
    echo "-p 只支援 24 或 16。" >&2
    exit 2
fi
if [ -n "$BASE" ] && [ ! -f "$BASE" ]; then
    echo "找不到基準檔:$BASE" >&2
    exit 2
fi

# .gz 與純文字混著吃
cat_logs() {
    for f in "$@"; do
        [ -f "$f" ] || continue
        case "$f" in
            *.gz) gzip -dc -- "$f" ;;
            *)    cat -- "$f" ;;
        esac
    done
}

# 比對模式先把基準的鍵(服務/來源網段/目的網段)讀進來
KEYFILE=""
if [ -n "$BASE" ]; then
    KEYFILE=$(mktemp)
    trap 'rm -f "$KEYFILE"' EXIT INT TERM
    # 去掉 BOM 與 CR,跳過表頭,取第 2/3/4 欄當鍵
    sed '1s/^\xef\xbb\xbf//' "$BASE" | tr -d '\r' | awk -F',' '
        NR==1 { next }
        NF>=4 {
            # 還原被引號包起來的欄位(本腳本自己的輸出只會在含逗號時加引號)
            s=$2; d=$3; e=$4
            gsub(/^"|"$/,"",s); gsub(/^"|"$/,"",d); gsub(/^"|"$/,"",e)
            print s "\t" d "\t" e
        }' > "$KEYFILE"
fi

cat_logs "$@" | awk -F'\t' -v prefix="$PREFIX" -v keyfile="$KEYFILE" '
function net(ip,   a, n) {
    # 非內網一律歸成單一組。幾千個外部位址會把表撐爆,而「連了新的外部位址」
    # 這件事在審核時看的是「哪個服務對外」,不是對到哪個 IP。
    if (ip !~ /^10\./ && ip !~ /^192\.168\./ &&
        ip !~ /^172\.(1[6-9]|2[0-9]|3[01])\./ && ip !~ /^127\./) return "外部"
    n = split(ip, a, ".")
    if (n != 4) return "其他"
    if (prefix == 16) return a[1] "." a[2] ".0.0/16"
    return a[1] "." a[2] "." a[3] ".0/24"
}
function addex(k, ip, which,   cur) {
    cur = (which == "s") ? exs[k] : exd[k]
    if (cur == "") cur = ip
    else if (split(cur, _t, ";") < 3) cur = cur ";" ip
    if (which == "s") exs[k] = cur; else exd[k] = cur
}
function csv(s) {
    if (s ~ /[",]/) { gsub(/"/, "\"\"", s); return "\"" s "\"" }
    return s
}
BEGIN {
    # 高風險埠:遠端控制、檔案共享、明文管理、常見後門
    split("21 22 23 69 135 137 138 139 445 512 513 514 1337 3389 4444 5555 " \
          "5800 5900 5901 5902 5938 5985 5986 6666 6667 9001 12345 31337 47001", _r, " ")
    for (i in _r) risky[_r[i]] = 1
    if (keyfile != "") {
        while ((getline line < keyfile) > 0) known[line] = 1
        close(keyfile)
    }
}
# 每個檔案的 #fields 表頭都要重讀:多檔可能來自不同版本
/^#fields/ { for (i = 2; i <= NF; i++) col[$i] = i - 1; hasfields = 1; next }
/^#/ { next }
{
    if (!hasfields) next                       # 沒有表頭就無法安全對欄位
    st = $col["conn_state"]
    # 失敗連線是掃描與探測的產物,不代表服務真的被使用,放進基準只會製造雜訊
    if (st == "S0" || st == "RSTOS0" || st == "SH" || st == "REJ") next
    s = $col["id.orig_h"]; d = $col["id.resp_h"]
    p = $col["id.resp_p"]; sv = $col["service"]; pr = $col["proto"]
    if (s == "" || d == "") next
    svc = (sv == "-" || sv == "") ? p "/" pr : p " " sv
    ns = net(s); nd = net(d)
    k = svc SUBSEP ns SUBSEP nd
    conns[k]++
    bytes[k] += $col["orig_ip_bytes"] + $col["resp_ip_bytes"]
    port[k] = p
    if (!((k SUBSEP s) in seen_s)) { seen_s[k SUBSEP s] = 1; nsrc[k]++; addex(k, s, "s") }
    if (!((k SUBSEP d) in seen_d)) { seen_d[k SUBSEP d] = 1; ndst[k]++; addex(k, d, "d") }
}
END {
    printf "\357\273\277"                       # UTF-8 BOM,否則 Excel 開中文是亂碼
    printf "判定,服務,來源網段,目的網段,來源台數,目的台數,連線數,流量,來源範例,目的範例\r\n"
    n = 0
    for (k in conns) {
        split(k, f, SUBSEP)
        if (keyfile != "" && ((f[1] "\t" f[2] "\t" f[3]) in known)) continue
        flag = ""
        if (port[k] in risky)  flag = "⚠高風險"
        else if (nsrc[k] == 1) flag = "⚠僅1台"
        else if (f[3] == "外部") flag = "對外"
        # 有標記的排前面;其次來源台數少的排前面 —— 可疑的自動浮到最上面
        sortkey = (flag == "" ? "2" : "1") sprintf("%08d", nsrc[k])
        rows[++n] = sortkey SUBSEP flag SUBSEP f[1] SUBSEP f[2] SUBSEP f[3] SUBSEP \
                    nsrc[k] SUBSEP ndst[k] SUBSEP conns[k] SUBSEP bytes[k] SUBSEP \
                    exs[k] SUBSEP exd[k]
    }
    if (!hasfields)
        print "警告:輸入沒有 #fields 表頭,無法對應欄位。請用原始的 conn.log," \
              "不要用 zeek-cut 的輸出。" > "/dev/stderr"
    else if (n == 0 && keyfile == "")
        print "警告:沒有彙總到任何連線。可能整批都是失敗連線(S0/REJ),或檔案是空的。" > "/dev/stderr"
    if (n == 0) { exit 0 }
    for (i = 1; i <= n; i++)
        for (j = i + 1; j <= n; j++)
            if (rows[j] < rows[i]) { t = rows[i]; rows[i] = rows[j]; rows[j] = t }
    for (i = 1; i <= n; i++) {
        split(rows[i], c, SUBSEP)
        printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\r\n",
            csv(c[2]), csv(c[3]), csv(c[4]), csv(c[5]),
            c[6], c[7], c[8], c[9], csv(c[10]), csv(c[11])
    }
}
'

#!/bin/bash
# ─────────────────────────────────────────────
# Portfolio System — 一鍵安裝腳本
# 執行方式：bash setup.sh
# ─────────────────────────────────────────────

set -eo pipefail

# ── 複製時忽略 identical 錯誤 ──
cp_safe() {
    cp "$1" "$2" 2>/dev/null || true
}
GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'

echo ""
echo "📈 Portfolio Daily Report — 安裝精靈"
echo "======================================"
echo ""

# ── 1. 檢查 Python 版本 ──
echo "🔍 檢查 Python..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ 找不到 Python3，請先安裝：https://www.python.org/downloads/${NC}"
    exit 1
fi
PY_VER=$(python3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
echo -e "${GREEN}✅ Python $PY_VER 已安裝${NC}"

# ── 2. 安裝 Python 套件 ──
echo ""
echo "📦 安裝必要套件（約 1 分鐘）..."
pip3 install yfinance reportlab pytz pandas --quiet --break-system-packages 2>/dev/null \
  || pip3 install yfinance reportlab pytz pandas --quiet
echo -e "${GREEN}✅ 套件安裝完成${NC}"

# ── 3. 複製主程式到 Desktop ──
echo ""
echo "📁 複製程式到桌面..."
DESKTOP="$HOME/Desktop"
mkdir -p "$DESKTOP"
cp_safe portfolio_system.py "$DESKTOP/portfolio_system.py"
echo -e "${GREEN}✅ 程式已複製到 $DESKTOP${NC}"

# ── 4. 引導填寫 config.json ──
echo ""
echo "⚙️  設定你的帳號資訊"
echo "────────────────────────────────────────"
echo ""

echo "請輸入你的 Gmail 信箱（用來寄報告）："
read -r GMAIL_SENDER
echo ""

echo "請輸入 Gmail 應用程式密碼（App Password）："
echo -e "${YELLOW}📌 設定方式：Google 帳號 → 安全性 → 兩步驟驗證 → 應用程式密碼${NC}"
read -r -s GMAIL_PASSWORD
echo ""

echo "請輸入收件信箱（可以跟寄件相同）："
read -r GMAIL_RECEIVER
echo ""

echo "請輸入 Anthropic API Key（用來產生 AI 市場摘要）："
echo -e "${YELLOW}📌 申請：https://console.anthropic.com/${NC}"
read -r -s ANTHROPIC_KEY
echo ""

cat > "$DESKTOP/config.json" << EOF
{
  "gmail_sender":      "$GMAIL_SENDER",
  "gmail_password":    "$GMAIL_PASSWORD",
  "gmail_receiver":    "$GMAIL_RECEIVER",
  "anthropic_api_key": "$ANTHROPIC_KEY"
}
EOF
echo -e "${GREEN}✅ config.json 已建立${NC}"

# ── 5. 引導填寫 portfolio.json ──
echo ""
echo "💼 設定你的投資組合"
echo "────────────────────────────────────────"
echo ""
echo "請輸入你的財務目標（台幣，例如 3000000）："
read -r GOAL_TWD
echo ""

cat > "$DESKTOP/portfolio.json" << 'PORTFOLIO_EOF'
{
  "goal_twd": GOAL_PLACEHOLDER,
  "cash_twd": 0,
  "holdings": {
    "AAPL":    {"shares": 10,  "cost_usd": 150.0},
    "QQQ":     {"shares": 5,   "cost_usd": 380.0},
    "0050.TW": {"shares": 100, "cost_twd": 130.0}
  }
}
PORTFOLIO_EOF

# 把 goal 填入
sed -i '' "s/GOAL_PLACEHOLDER/$GOAL_TWD/" "$DESKTOP/portfolio.json" 2>/dev/null \
  || sed -i "s/GOAL_PLACEHOLDER/$GOAL_TWD/" "$DESKTOP/portfolio.json"

echo -e "${GREEN}✅ portfolio.json 已建立（內含範例持倉）${NC}"
echo -e "${YELLOW}📌 請用文字編輯器打開 ~/Desktop/portfolio.json 填入你的實際持股${NC}"

# ── 6. 設定每日自動執行（launchd）──
echo ""
echo "⏰ 設定每日自動執行時間"
echo "────────────────────────────────────────"
echo "1) 每天台灣時間 23:00（美股收盤後）"
echo "2) 每天台灣時間 13:30（台股收盤後）"
echo "3) 兩個時間都要"
echo "4) 暫時不設定"
echo ""
read -r -p "請選擇 (1/2/3/4)：" SCHEDULE_CHOICE

PLIST_DIR="$HOME/Library/LaunchAgents"
mkdir -p "$PLIST_DIR"
PY_PATH=$(which python3)
SCRIPT_PATH="$DESKTOP/portfolio_system.py"

create_plist() {
    local NAME=$1
    local HOUR=$2
    local MINUTE=$3
    cat > "$PLIST_DIR/com.portfolio.$NAME.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.portfolio.$NAME</string>
    <key>ProgramArguments</key>
    <array>
        <string>$PY_PATH</string>
        <string>$SCRIPT_PATH</string>
    </array>
    <key>StartCalendarInterval</key>
    <dict>
        <key>Hour</key>
        <integer>$HOUR</integer>
        <key>Minute</key>
        <integer>$MINUTE</integer>
    </dict>
    <key>StandardOutPath</key>
    <string>$HOME/Desktop/portfolio_log.txt</string>
    <key>StandardErrorPath</key>
    <string>$HOME/Desktop/portfolio_error.txt</string>
    <key>RunAtLoad</key>
    <false/>
</dict>
</plist>
EOF
    launchctl unload "$PLIST_DIR/com.portfolio.$NAME.plist" 2>/dev/null || true
    launchctl load "$PLIST_DIR/com.portfolio.$NAME.plist"
    echo -e "${GREEN}✅ 排程已設定：UTC ${HOUR}:$(printf '%02d' $MINUTE)（台灣時間 $((HOUR+8)):$(printf '%02d' $MINUTE)）${NC}"
}

case $SCHEDULE_CHOICE in
    1) create_plist "night" 15 0 ;;      # UTC 15:00 = 台灣 23:00
    2) create_plist "noon"  5 30 ;;      # UTC 05:30 = 台灣 13:30
    3) create_plist "night" 15 0
       create_plist "noon"  5 30 ;;
    *) echo -e "${YELLOW}⏭️  跳過排程設定，之後可手動執行${NC}" ;;
esac

# ── 7. 試跑一次 ──
echo ""
echo "🚀 要現在試跑一次看看嗎？（會寄一封測試信到你的信箱）"
read -r -p "試跑 (y/n)：" RUN_NOW
if [[ "$RUN_NOW" == "y" || "$RUN_NOW" == "Y" ]]; then
    echo ""
    echo "執行中..."
    python3 "$DESKTOP/portfolio_system.py" && \
        echo -e "${GREEN}✅ 成功！請檢查你的信箱${NC}" || \
        echo -e "${RED}❌ 執行失敗，請確認 portfolio.json 的持倉格式是否正確${NC}"
fi

# ── 完成 ──
echo ""
echo "══════════════════════════════════════"
echo -e "${GREEN}🎉 安裝完成！${NC}"
echo ""
echo "📂 重要檔案位置："
echo "   程式：~/Desktop/portfolio_system.py"
echo "   帳號設定：~/Desktop/config.json"
echo "   持倉設定：~/Desktop/portfolio.json"
echo "   執行日誌：~/Desktop/portfolio_log.txt"
echo ""
echo "✏️  下一步：編輯 portfolio.json 填入你的真實持倉"
echo "   美股格式：\"AAPL\": {\"shares\": 10, \"cost_usd\": 150.0}"
echo "   台股格式：\"0050.TW\": {\"shares\": 100, \"cost_twd\": 130.0}"
echo ""
echo "手動執行：python3 ~/Desktop/portfolio_system.py"
echo "══════════════════════════════════════"

# ── 8. 自動開啟 portfolio.json 讓用戶填寫 ──
echo ""
echo -e "${YELLOW}📝 即將開啟 portfolio.json，請填入你的真實持倉後儲存。${NC}"
echo "   （目前是範例資料，記得改成你自己的股票！）"
echo ""
sleep 2
open -e "$DESKTOP/portfolio.json" 2>/dev/null || \
open "$DESKTOP/portfolio.json" 2>/dev/null || \
echo -e "${YELLOW}請手動開啟：~/Desktop/portfolio.json${NC}"

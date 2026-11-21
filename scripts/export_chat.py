#!/usr/bin/env python3
"""
Script xuất toàn bộ lịch sử cuộc trò chuyện giữa Người dùng và Antigravity
vào thư mục dự án dưới dạng Markdown và file thô JSONL (tự động ẩn các Token/Secrets nhạy cảm).
"""

import os
import sys
import json
import re
from datetime import datetime

CONV_ID = "1598e300-67b6-4e28-b71f-f2059ba01974"
SRC_DIR = f"/Users/nhdbao/.gemini/antigravity/brain/{CONV_ID}/.system_generated/logs"
PROJECT_DIR = "/Volumes/Lexar/Tools/Thesis-Support-Tool-for-Word-macOS-_VN"
DEST_DIR = os.path.join(PROJECT_DIR, "docs", "chat-history")

def redact_secrets(text):
    if not text:
        return ""
    # Mask GitHub Personal Access Tokens
    text = re.sub(r"ghp_[A-Za-z0-9_]{30,}", "ghp_***[REDACTED_TOKEN]***", text)
    # Mask Generic GitHub / Bearer tokens
    text = re.sub(r"github_pat_[A-Za-z0-9_]{30,}", "github_pat_***[REDACTED]***", text)
    return text

def clean_user_text(text):
    if not text:
        return ""
    # Extract user request tag
    m = re.search(r"<USER_REQUEST>(.*?)</USER_REQUEST>", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    # Remove metadata
    text = re.sub(r"<ADDITIONAL_METADATA>.*?</ADDITIONAL_METADATA>", "", text, flags=re.DOTALL)
    # Remove context summary if present
    text = re.sub(r"<CONTEXT_SUMMARY>.*?</CONTEXT_SUMMARY>", "", text, flags=re.DOTALL)
    return redact_secrets(text.strip())

def copy_and_redact_jsonl(src_file, dest_file):
    if not os.path.exists(src_file):
        return
    with open(src_file, "r", encoding="utf-8") as fin, open(dest_file, "w", encoding="utf-8") as fout:
        for line in fin:
            fout.write(redact_secrets(line))

def main():
    os.makedirs(DEST_DIR, exist_ok=True)

    # 1. Copy & Redact raw JSONL files
    src_transcript = os.path.join(SRC_DIR, "transcript.jsonl")
    src_transcript_full = os.path.join(SRC_DIR, "transcript_full.jsonl")

    copy_and_redact_jsonl(src_transcript, os.path.join(DEST_DIR, "transcript.jsonl"))
    copy_and_redact_jsonl(src_transcript_full, os.path.join(DEST_DIR, "transcript_full.jsonl"))

    print(f"Đã sao chép các tệp raw transcript (đã bảo mật) vào: {DEST_DIR}")

    # 2. Parse dialogues
    dialogue = []
    if os.path.exists(src_transcript_full):
        with open(src_transcript_full, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    data = json.loads(line)
                except Exception:
                    continue

                source = data.get("source")
                mtype = data.get("type")
                content = data.get("content", "")
                time_str = data.get("created_at", "")
                tool_calls = data.get("tool_calls", [])

                if source == "USER_EXPLICIT" and mtype == "USER_INPUT":
                    clean = clean_user_text(content)
                    if clean:
                        dialogue.append({
                            "role": "USER",
                            "text": clean,
                            "time": time_str
                        })
                elif source == "MODEL" and mtype == "PLANNER_RESPONSE" and content:
                    if not tool_calls or len(content.strip()) > 30:
                        if "<SYSTEM_MESSAGE>" not in content and "tool_calls" not in content:
                            dialogue.append({
                                "role": "ASSISTANT",
                                "text": redact_secrets(content.strip()),
                                "time": time_str
                            })

    # Merge consecutive assistant messages
    merged = []
    for msg in dialogue:
        if merged and merged[-1]["role"] == msg["role"]:
            merged[-1]["text"] += "\n\n" + msg["text"]
        else:
            merged.append(msg)

    # Write Markdown output
    md_path = os.path.join(DEST_DIR, "LICH_SU_TRO_CHUYEN.md")
    with open(md_path, "w", encoding="utf-8") as out:
        out.write("# 📜 Toàn bộ Nhật ký Cuộc trò chuyện & Tiến trình Dự án\n\n")
        out.write("**Dự án:** Thesis Support Tool for Word (macOS & Windows)\n\n")
        out.write(f"**Conversation ID:** `{CONV_ID}`\n\n")
        out.write(f"**Thời gian cập nhật:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        out.write("---\n\n")
        out.write("## 📌 Tóm tắt các mốc tính năng đã hoàn thành\n\n")
        out.write("1. **Khởi tạo Office Web Add-in:** Kiến trúc Task Pane tiêu chuẩn bằng Office.js API, HTML5, CSS3 Fluent, JavaScript.\n")
        out.write("2. **Quy chuẩn Bảng số liệu:** Table Grid, Times New Roman, AutoFit Window, căn giữa dọc ô (Center), Cột 1 căn trái, Cột 2+ căn phải.\n")
        out.write("3. **Bảng Quốc hiệu - Tiêu ngữ:** Tự động nhận diện nội dung văn bản hành chính Việt Nam, xóa hoàn toàn viền (No border), căn giữa 2 bên, định dạng chuẩn in hoa / đậm / nghiêng.\n")
        out.write("4. **Triển khai Đám mây (GitHub Pages):** Đưa toàn bộ mã nguồn lên GitHub Pages, không phụ thuộc vào `npm start` cục bộ.\n")
        out.write("5. **Cột số liệu không bị nhảy dòng:** Thuật toán tính toán `columnWidth` chia đều cho các cột số, ép vừa khít dòng đơn.\n")
        out.write("6. **Lề ô 4 cạnh (Cell Margin):** Nhập độc lập Trên, Dưới, Trái, Phải (đơn vị cm) kèm các preset thông minh.\n")
        out.write("7. **Định dạng đoạn văn (Paragraph formatting):** Khoảng cách dòng (Line spacing), Before/After, loại bỏ thụt đầu dòng (First Line Indent = 0pt).\n")
        out.write("8. **Bản cài đặt cho Windows:** Tạo file script 1-click `Cai_Dat_Windows.bat`, file nén zip và hướng dẫn sử dụng.\n\n")
        out.write("---\n\n")
        out.write("## 💬 Chi tiết từng lượt trao đổi\n\n")

        turn_idx = 1
        for msg in merged:
            t = msg.get("time", "")
            if msg["role"] == "USER":
                out.write(f"### 👤 Người dùng (Lượt {turn_idx}) - *{t}*\n\n")
                out.write(f"> {msg['text'].replace(chr(10), chr(10) + '> ')}\n\n")
                turn_idx += 1
            else:
                out.write(f"### 🤖 Trợ lý AI (Antigravity)\n\n")
                out.write(f"{msg['text']}\n\n")
                out.write("---\n\n")

    print(f"Đã tạo thành công nhật ký trò chuyện ({len(merged)} lượt trao đổi) tại: {md_path}")

if __name__ == "__main__":
    main()

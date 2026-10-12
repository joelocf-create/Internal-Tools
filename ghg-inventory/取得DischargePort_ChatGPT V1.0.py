# -*- coding: utf-8 -*-
"""
出貨文件 Port of Discharge 擷取工具 (v3)

背景：
    這批出貨文件（B/L / Sea Waybill / FCR）多數是「表格底圖為圖片，
    只有填入的資料是可選取文字」的 PDF——欄位標籤本身（如 PORT OF
    DISCHARGE FROM VESSEL）不在文字層裡，純文字關鍵字比對無法定位。

策略（依序嘗試，抓不到就標記人工確認，不臆測）：
    1. 座標式左右欄比對：適用於「表格底圖為圖片，欄位標籤不在文字層」
       的版型（ATE B/L、OFS/SOM B/L 等）。固定規律：Port of Loading
       行（右欄，含 TAIWAN）的下一行，左欄是 Port of Discharge、
       右欄是 Place of Delivery。
    2. 交錯行還原比對：適用 Maersk Waybill 這類版型——PDF 產生器將
       「標籤行」與「值行」以幾乎相同的 y 座標、逐字交錯的方式排版
       （例如 "Port of Loading" 與 "TAICHUNG" 的每個字母交錯疊在
       同一視覺行），一般的列聚合（含 y 容許誤差）會把兩行文字混在
       一起變成無法辨識的亂碼。用「精確 top 值分組」（不做誤差容許）
       將每個不同 top 值視為獨立一行，即可還原出乾淨的標籤行與值行，
       再用序位/最近距離比對。
    3. 逗號地名 keyword 比對：適用文字層有 "PORT OF DISCHARGE" 標籤
       的版型（Expeditors FCR、COSCO Sea Waybill、Mainfreight B/L、
       Everise B/L），值為「城市, 州/國家」格式。
    4. 空格分隔單字序位比對：適用標籤與值皆為空格分隔單字、值中
       無逗號的版型（Maersk/Damco FCR），依標籤序位對應值序位。
    5. 文字層內容過少（掃描圖檔）：轉圖片 + pytesseract OCR，
       再套用第 3、4 項規則嘗試比對。
    6. SAVANNAH 已知港口關鍵字保底：這批文件的港口目的地實務上幾乎
       全部是 SAVANNAH, GA。部分來源為圖片型 PDF（OCR），OCR 常因
       表格框線、字級或掃描品質造成開頭字母被切掉或誤植（如
       "AVANNAH"、"5AVANNAH"），導致嚴謹的地名 pattern 比對失敗。
       在其他策略都失敗、且限定於 B/L/Waybill 頁面範圍內時，只要
       偵測到「SAVANNAH」（含常見 OCR 開頭字母缺漏的寬鬆比對）
       就直接判定為 "SAVANNAH, GA"，不需要再抓完整欄位版面。
    7. 以上皆無法判斷 → 明確標記 "[未擷取到，請人工確認]"，
       絕不臆測填入猜測值。

使用方式：
    直接執行本程式，只會跳出檔案選擇視窗，選擇要處理的 PDF 檔案
    （可複選）。處理進度會顯示在 PowerShell／命令提示字元視窗。
    處理完成後，結果 Excel 會輸出至所選檔案所在的資料夾。
    OCR 採逐頁處理，並定期暫存；若中途異常結束，下次選取相同檔案可續跑。

輸出：
    <所選檔案資料夾>/port_of_discharge_result.xlsx
"""

import gc
import os
import re
import tkinter as tk
from tkinter import filedialog

import pdfplumber
from pdf2image import convert_from_path, pdfinfo_from_path
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# ---------- 設定 ----------
OCR_DPI = 220
OCR_LANG = "eng"
OCR_TIMEOUT_SECONDS = 60
CHECKPOINT_INTERVAL = 10
MIN_TEXT_LEN_FOR_TEXTLAYER = 30

# 座標式左右欄比對法的欄位分界值（適用 ATE、SOM 等表格式版型）
ATE_COLUMN_SPLIT_X = 150

# 檔名核心代號的判斷 pattern：客戶/廠商代號 + 連字號 + 6 位數日期碼
# + 可選的單一字母版次（如 AS-241231A、TG4-241231F、GDP-240112）。
# 這是這批出貨文件檔名的共同慣例；符合則只取這段當顯示名稱，
# 捨棄後面的尾綴資訊（如 -LOADING、_BL、__277729213_ 等）。
FILENAME_CORE_PATTERN = re.compile(r"^([A-Za-z0-9]+-\d{6}[A-Za-z]?)")


def simplify_filename(filename: str) -> str:
    """
    去除副檔名，並將檔名簡化為「客戶代號-日期碼」核心部分，
    捨棄其後的附加尾綴（發票號、文件類型縮寫等）。
    若檔名不符合預期格式，回傳去除副檔名後的原始檔名，不臆測。
    """
    stem = os.path.splitext(filename)[0]
    m = FILENAME_CORE_PATTERN.match(stem)
    return m.group(1) if m else stem


def extract_words_by_page(path: str):
    """回傳每頁的 words 清單（含座標），與每頁純文字。"""
    pages_words = []
    pages_text = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                pages_words.append(page.extract_words())
                pages_text.append(page.extract_text() or "")
    except Exception as e:
        print(f"  [警告] pdfplumber 讀取失敗: {e}")
    return pages_words, pages_text


def try_extract_port_from_ocr_text(ocr_text: str) -> str:
    """將單頁 OCR 文字套用既有 Port of Discharge 辨識規則。"""
    return (
        try_keyword_layout([ocr_text])
        or try_space_separated_layout([ocr_text])
        or try_known_port_fallback(ocr_text)
    )


def extract_port_via_ocr_pages(
    path: str,
    page_indices=None,
    progress_callback=None,
) -> str:
    """
    逐頁轉圖並執行 OCR，找到 Port of Discharge 後立即停止。

    page_indices：0-indexed 頁碼清單；未提供時處理整份 PDF。
    每次只載入一頁，並在完成後關閉圖片及回收記憶體，避免大量檔案
    批次處理時因高解析度圖片累積而耗盡記憶體。
    """
    def report(message: str):
        if progress_callback:
            progress_callback(message)
        else:
            print(message)

    if page_indices is None:
        try:
            page_count = int(pdfinfo_from_path(path).get("Pages", 0))
        except Exception as e:
            report(f"  [錯誤] 無法取得 PDF 頁數: {e}")
            return ""
        page_indices = list(range(page_count))
    else:
        page_indices = list(dict.fromkeys(page_indices))

    total_pages = len(page_indices)
    for sequence, page_index in enumerate(page_indices, start=1):
        images = []
        page_number = page_index + 1
        try:
            report(f"  OCR 辨識第 {page_number} 頁（{sequence}/{total_pages}）...")
            images = convert_from_path(
                path,
                dpi=OCR_DPI,
                first_page=page_number,
                last_page=page_number,
                grayscale=True,
                thread_count=1,
            )
            if not images:
                continue

            ocr_text = pytesseract.image_to_string(
                images[0],
                lang=OCR_LANG,
                config="--oem 3 --psm 11",
                timeout=OCR_TIMEOUT_SECONDS,
            ).strip()

            port = try_extract_port_from_ocr_text(ocr_text)
            if port:
                report(f"  已於第 {page_number} 頁取得: {port}")
                return port

        except RuntimeError:
            report(
                f"  [警告] 第 {page_number} 頁 OCR 超過 "
                f"{OCR_TIMEOUT_SECONDS} 秒，已跳過。"
            )
        except Exception as e:
            report(f"  [警告] 第 {page_number} 頁 OCR 失敗: {e}")
        finally:
            for image in images:
                try:
                    image.close()
                except Exception:
                    pass
            del images
            gc.collect()

    return ""


def cluster_rows(words: list, y_tolerance: float = 4.0) -> list:
    """將座標相近（同一視覺行，誤差內）的文字合併成一行。
    回傳依 top 排序的行清單，每行為該行所有 words（依 x0 排序）。
    """
    sorted_words = sorted(words, key=lambda w: w["top"])
    clustered = []
    current_row = []
    current_top = None

    for w in sorted_words:
        if current_top is None or abs(w["top"] - current_top) <= y_tolerance:
            current_row.append(w)
            # 用行內第一個字的 top 當代表值，避免逐字漂移累積誤差
            if current_top is None:
                current_top = w["top"]
        else:
            clustered.append(sorted(current_row, key=lambda x: x["x0"]))
            current_row = [w]
            current_top = w["top"]

    if current_row:
        clustered.append(sorted(current_row, key=lambda x: x["x0"]))

    return clustered


def try_coordinate_layout(pages_words: list) -> str:
    """
    座標式左右欄比對法，適用於「表格底圖為圖片，欄位標籤不在文字層」
    的版型（ATE 表格式、OFS/SOM 表格式等）。固定規律：
    Port of Loading 該行（右欄，含 TAIWAN 地名）的下一行，
    就是 Port of Discharge（左欄）與 Place of Delivery（右欄）
    並排的資料行。

    place_pattern 額外支援括號地名（如 "HOCHIMINH (CAT LAI)"）。
    為避免誤判到 shipper/consignee 地址或公司抬頭（如
    "ALL STRONG INDUSTRY (USA) INC." 也可能符合寬鬆地名 pattern），
    額外排除含公司後綴詞（INC、CO.、LTD、CORP 等）的行。
    """
    place_pattern = re.compile(r"^[A-Z][A-Z\s,.\-'()]+$")
    company_suffix_pattern = re.compile(
        r"\b(INC|CO\.|LTD|CORP|LLC|CO,)\b", re.IGNORECASE
    )
    # 已知的表格欄位標籤詞——若 row_b 本身由這些詞組成（而非實際地名），
    # 代表抓到的是標籤行本身，不是資料值行，須排除
    label_word_pattern = re.compile(
        r"^(PORT|PLACE|OF|DISCHARGE|LOADING|DELIVERY|RECEIPT|"
        r"DESTINATION|VESSEL|VOYAGE|CARRIER|CONSIGNEE|NOTIFY|PARTY)$",
        re.IGNORECASE,
    )

    def is_label_row(text: str) -> bool:
        """判斷這段文字是否為欄位標籤本身（每個字都是標籤詞），
        而非實際地名資料。"""
        words_in_text = text.replace(",", " ").split()
        return bool(words_in_text) and all(
            label_word_pattern.match(w) for w in words_in_text
        )

    for words in pages_words:
        rows = cluster_rows(words)

        for i, row_a in enumerate(rows):
            row_a_right = [w for w in row_a if w["x0"] >= ATE_COLUMN_SPLIT_X]
            if not row_a_right:
                continue
            right_a_text = " ".join(w["text"] for w in row_a_right).strip()
            if company_suffix_pattern.search(right_a_text):
                continue
            # 要求含逗號（完整「城市, 國家」地名），排除像單獨一個
            # "TAIWAN" 這種國名，避免誤判到 shipper 地址等雜訊行
            if "," not in right_a_text:
                continue
            if not (place_pattern.match(right_a_text) and "TAIWAN" in right_a_text):
                continue

            # 往後最多找 6 行內的「左右欄皆地名、且不含公司後綴」的行
            for row_b in rows[i + 1: i + 7]:
                row_b_left = [w for w in row_b if w["x0"] < ATE_COLUMN_SPLIT_X]
                row_b_right = [w for w in row_b if w["x0"] >= ATE_COLUMN_SPLIT_X]
                if not row_b_left or not row_b_right:
                    continue

                left_b_text = " ".join(w["text"] for w in row_b_left).strip()
                right_b_text = " ".join(w["text"] for w in row_b_right).strip()

                if company_suffix_pattern.search(left_b_text) or \
                   company_suffix_pattern.search(right_b_text):
                    continue

                # 排除標籤行本身（如 "PORT OF DISCHARGE" / "PLACE OF DELIVERY"）
                if is_label_row(left_b_text) or is_label_row(right_b_text):
                    continue

                if place_pattern.match(left_b_text) and place_pattern.match(right_b_text):
                    return left_b_text.rstrip(",")

    return ""


def cluster_rows_exact_top(words: list) -> dict:
    """
    將文字依「精確 top 值」（四捨五入至 0.1 避免浮點誤差）分組成行，
    不做任何誤差容許（與 cluster_rows 的 y_tolerance 版本不同）。

    用途：部分版型（如 Maersk Waybill）會把「標籤行」與「值行」以
    幾乎相同、但實際上有微小差異的 y 座標排版，兩者的字元在視覺上
    交錯疊在一起（例如 "Port of Loading" 與 "TAICHUNG" 逐字交錯）。
    若用一般的誤差容許聚合（y_tolerance），會把兩行文字混成一行，
    產生無法辨識的亂碼字串。改用精確 top 值分組，可將這兩行文字
    正確拆分開來，各自還原成乾淨的一行。

    回傳：{top值: 該行文字（已依 x0 排序、無空格串接）}
    """
    rows: dict = {}
    for w in words:
        key = round(w["top"], 1)
        rows.setdefault(key, []).append(w)
    return {
        top: "".join(w["text"] for w in sorted(ws, key=lambda x: x["x0"]))
        for top, ws in rows.items()
    }


def try_interleaved_row_layout(pages_words: list) -> str:
    """
    處理「標籤行與值行以逐字交錯方式排版、y 座標幾乎相同」的版型
    （目前已知：Maersk Non-Negotiable Waybill）。

    典型還原後的結果類似（括號為該行第一個字的 x0 座標）：
        (top=298.8, x0=34.9)  "PortofLoading"
        (top=301.1, x0=162.4) "PortofDischarge"
        (top=305.5, x0=34.9)  "TAICHUNG"
        (top=307.8, x0=162.4) "SAVANNAH"

    規律：標籤行與其對應的值行，並非依「文件順序上的下一行」對應
    （因為左欄 PortofLoading/TAICHUNG 與右欄 PortofDischarge/
    SAVANNAH 這兩組各自的 top 值會交錯排列），而是要看「起始 x0
    座標」是否相同——同一欄位的標籤與值會對齊在同一個 x0。因此用
    「PortofDischarge 這個標籤行的起始 x0」去找同一份 words 中，
    其他 top 值裡起始 x0 幾乎相同（誤差 2pt 內）的那一行，即為對應
    的值。
    """
    label_compact_pattern = re.compile(
        r"^(Vessel|VoyageNo|ReceiptNo|PlaceofReceipt|PortofLoading|"
        r"PortofDischarge|PlaceofDelivery|FinalDestination)\.?",
        re.IGNORECASE,
    )

    for words in pages_words:
        rows_by_top = cluster_rows_exact_top(words)
        if not rows_by_top:
            continue

        # 記錄每個 top 值對應行的起始 x0（該行第一個字的 x0）
        row_start_x0 = {
            top: min(w["x0"] for w in words if round(w["top"], 1) == top)
            for top in rows_by_top
        }

        for top, compact in rows_by_top.items():
            if not re.match(r"^PortofDischarge", compact, re.IGNORECASE):
                continue

            label_x0 = row_start_x0[top]

            # 在其他行中找起始 x0 幾乎相同（同一欄位）、且非標籤行的候選
            candidates = []
            for other_top, other_compact in rows_by_top.items():
                if other_top == top:
                    continue
                if abs(row_start_x0[other_top] - label_x0) > 2.0:
                    continue
                if label_compact_pattern.match(other_compact):
                    continue
                # 值應為單純大寫地名（含逗號、括號皆可），不含長句子
                # （若混入大量小寫字母，通常是條款內文而非欄位值）
                cleaned_value = other_compact.rstrip("*").strip()
                if re.match(r"^[A-Z][A-Za-z\s,.\-'()]{1,40}$", cleaned_value) and \
                        not re.search(r"\b(the|of|to|and|or)\b", cleaned_value, re.IGNORECASE):
                    candidates.append((other_top, cleaned_value))

            if candidates:
                # 值通常緊接在標籤行之後（top 值最接近者優先）
                candidates.sort(key=lambda c: abs(c[0] - top))
                return candidates[0][1].strip()

    return ""


# 已知港口目的地保底關鍵字（此批客户/供應鏈實務上出貨港口目的地
# 幾乎全部固定為 SAVANNAH, GA）。OCR 常因表格框線壓在文字上、字級
# 過小、掃描品質等因素，造成開頭 1~2 個字母被切掉或誤判（如
# "AVANNAH"、"5AVANNAH"、"AVANNAH,GA"），嚴謹的地名 pattern 比對
# 因而失敗。這裡用寬鬆比對（允許開頭缺 1~2 碼）在其他策略都失敗時
# 當作最後保底，避免明明看得出是 SAVANNAH 卻仍標記為人工確認。
SAVANNAH_LOOSE_PATTERN = re.compile(r"\b[A5]?[Vv][A4]NN[A4]H\b", re.IGNORECASE)


def try_known_port_fallback(text: str) -> str:
    """
    在其他策略皆失敗時的最後保底：若文字（含 OCR 結果）中偵測到
    「SAVANNAH」（含常見 OCR 開頭字母缺漏/誤植的寬鬆比對），直接
    判定為 "SAVANNAH, GA"。僅在呼叫端已限定於 B/L/Waybill 相關頁面
    範圍時使用，避免誤判到與港口無關但恰好含近似字串的雜訊文字。
    """
    if SAVANNAH_LOOSE_PATTERN.search(text):
        return "SAVANNAH, GA"
    return ""


def normalize_spaced_out_text(text: str) -> str:
    """
    修正部分 PDF（如 Expeditors FCR）因字距渲染造成的「字母間假空格」，
    例如 "PO R T O F D IS C HA R G E" 應還原為 "PORT OF DISCHARGE"。

    做法：逐行處理，若一行內大寫字母＋單一空格的片段比例異常高
    （代表是被字距拆開的假空格），將行內單一空格移除後，
    在已知欄位標籤處重新插入正確空格。
    """
    # 只針對「單一空白字元」且前後都是單一字母的模式做壓縮，
    # 但保留看起來像正常單字間距（多個字母構成的詞）的空格。
    # 簡化策略：偵測到已知欄位標籤的「字母拆開」變體時，直接替換掉。
    known_labels = [
        "PORT OF DISCHARGE",
        "PORT OF LOADING",
        "PLACE OF DELIVERY",
        "PLACE OF RECEIPT",
        "CONSIGNEE",
        "NOTIFY PARTY",
    ]
    result = text
    for label in known_labels:
        # 只允許插入「半形空格」，不可跨行（避免吃掉換行符破壞行結構）
        loose_pattern = ""
        for ch in label:
            if ch == " ":
                loose_pattern += r"[ ]*"
            else:
                loose_pattern += re.escape(ch) + r"[ ]?"
        result = re.sub(loose_pattern, label, result, flags=re.IGNORECASE)
    return result


# 常見「州/國家」尾詞（依這批文件實際出現過的港口目的地整理），
# 用來當作地名組的結尾錨點。多字詞放前面，避免被單字詞提前截斷。
KNOWN_REGION_SUFFIXES = [
    "UNITED STATES", "UNITED KINGDOM",
    "TAIWAN", "CHINA", "VIETNAM", "CANADA",
    "TX", "GA", "CA", "NY", "NC", "SC", "WA", "OR", "IL", "NJ", "PA",
    "QC", "ON", "BC",  # 加拿大省份縮寫
]
# 依詞長排序（長的優先比對，避免 "UNITED STATES" 被 "STATES" 誤截）
KNOWN_REGION_SUFFIXES.sort(key=len, reverse=True)


def split_geo_groups(line: str) -> list:
    """
    將一行內連續出現的多組「城市, 州/國家」地名切分開來。

    做法：以 KNOWN_REGION_SUFFIXES 中的關鍵字出現位置為每組的結尾錨點，
    往前找最近的逗號，逗號與錨點之間即為該組的「城市, 州/國家」文字。
    """
    groups = []
    search_pos = 0
    text = line

    while search_pos < len(text):
        earliest_match = None
        for suffix in KNOWN_REGION_SUFFIXES:
            m = re.search(r"\b" + re.escape(suffix) + r"\b", text[search_pos:])
            if m:
                abs_start = search_pos + m.start()
                abs_end = search_pos + m.end()
                if earliest_match is None or abs_start < earliest_match[0]:
                    earliest_match = (abs_start, abs_end)
        if earliest_match is None:
            break

        suffix_start, suffix_end = earliest_match
        # 往前找最近的逗號，作為城市與州/國家的分界
        comma_pos = text.rfind(",", search_pos, suffix_start)
        if comma_pos == -1:
            # 沒逗號就跳過這個錨點，避免誤判
            search_pos = suffix_end
            continue

        # 城市名：從（上一組結尾或行首）到逗號前，取最後一段大寫詞
        before_comma = text[search_pos:comma_pos]
        city_match = re.search(r"[A-Z][A-Za-z]*(?:\s[A-Z][A-Za-z]*)*$", before_comma)
        if not city_match:
            search_pos = suffix_end
            continue

        city = city_match.group(0).strip()
        region = text[comma_pos + 1: suffix_end].strip()
        groups.append(f"{city}, {region}")
        search_pos = suffix_end

    return groups

def try_space_separated_layout(pages_text: list) -> str:
    """
    處理「標籤與值皆為空格分隔單字、值中無逗號」的版型
    （如 TG8 Maersk FCR：
        標籤行 "PLACEOF RECEIPT PORT OF LOADING PORT OF DISCHARGE PLACEOF DELIVERY"
        值行   "TAICHUNG KAOHSIUNG VANCOUVER BOUCHERVILLE"）
    用標籤在標籤行中的序位，去值行找同序位的單字 token。
    僅在 split_geo_groups 抓不到逗號地名組時使用（避免與其他版型衝突）。
    """
    full_text = "\n".join(pages_text)
    full_text = normalize_spaced_out_text(full_text)
    lines = full_text.splitlines()

    for i, line in enumerate(lines):
        upper_line = line.upper()
        if "PORT OF DISCHARGE" not in upper_line:
            continue

        labels_found = []
        for label in KNOWN_LABELS_IN_ORDER:
            pos = upper_line.find(label)
            if pos != -1:
                labels_found.append((pos, label))
        labels_found.sort(key=lambda x: x[0])
        label_order = [lbl for _, lbl in labels_found]

        if "PORT OF DISCHARGE" not in label_order:
            continue
        discharge_index = label_order.index("PORT OF DISCHARGE")

        for candidate_line in lines[i + 1: i + 3]:
            # 值行必須是「純空格分隔的大寫單字」，且單字數量與標籤數量相符，
            # 否則序位對應不可靠，寧可放棄不猜
            tokens = candidate_line.strip().split()
            if not tokens or not all(re.match(r"^[A-Z][A-Za-z'\-]*$", t) for t in tokens):
                continue
            if len(tokens) != len(label_order):
                continue

            return tokens[discharge_index]

    return ""


# 已知的表頭欄位標籤，依常見出現順序（用於「標籤序列 vs 值序列」位置對應）
KNOWN_LABELS_IN_ORDER = [
    "PLACE OF RECEIPT",
    "PORT OF LOADING",
    "PORT OF DISCHARGE",
    "PLACE OF DELIVERY",
    "FINAL DESTINATION",
]


def try_keyword_layout(pages_text: list) -> str:
    """
    通用關鍵字 fallback，處理兩種常見排版：

    A) 標籤與值分成兩行，且標籤順序與值順序一一對應
       （如 TG8: "...PORT OF LOADING PORT OF DISCHARGE PLACE OF DELIVERY"
             下一行 "TAICHUNG KAOHSIUNG VANCOUVER BOUCHERVILLE"）
       → 用標籤在標籤行中的序位，去值行找同序位的地名 token。

    B) "PORT OF DISCHARGE" 標籤後直接緊接同一段落內連續地名
       （如 VARI: "PORT OF DISCHARGE\\nKAOHSIUNG, TAIWAN LONG BEACH, ..."，
        這裡標籤行只有一個 PORT OF DISCHARGE，但上一行有 PORT OF LOADING，
        值行含兩組地名，需跳過第一組）
       → 抓標籤行內 "PORT OF DISCHARGE" 前面還有幾個已知標籤，
         值行就跳過對應數量的地名組再取值。
    """
    full_text = "\n".join(pages_text)
    full_text = normalize_spaced_out_text(full_text)
    lines = full_text.splitlines()

    for i, line in enumerate(lines):
        upper_line = line.upper()
        if "PORT OF DISCHARGE" not in upper_line:
            continue

        # 統計這一行（標籤行）中，"PORT OF DISCHARGE" 之前還出現了
        # 哪些已知欄位標籤，藉此判斷值行中要跳過幾組地名
        labels_found = []
        for label in KNOWN_LABELS_IN_ORDER:
            pos = upper_line.find(label)
            if pos != -1:
                labels_found.append((pos, label))
        labels_found.sort(key=lambda x: x[0])
        label_order = [lbl for _, lbl in labels_found]

        if "PORT OF DISCHARGE" not in label_order:
            continue
        discharge_index = label_order.index("PORT OF DISCHARGE")

        # 值可能在同一行（標籤與值中間夾雜數字如 "8. PORT OF DISCHARGE"
        # 時仍視為標籤行），也可能在下一行——嘗試接下來 1~2 行找地名 token
        for candidate_line in lines[i + 1: i + 3]:
            # 地名組型態：「城市, 州/國」或單一地名詞（後面緊接大寫新地名時視為下一組）。
            # 用「逗號地名」optionally 後面接單一大寫詞當州/國縮寫，
            # 且用 (?=[A-Z]|$) 確保不會吞掉下一組地名的開頭字母。
            # 地名組拆分：用「州/國家」關鍵字（含常見多字國名）當作
            # 每組的結尾錨點，往前回抓城市名，藉此可靠切分連續多組
            # 地名（如 "KAOHSIUNG, TAIWAN LONG BEACH, UNITED STATES"）。
            geo_groups = split_geo_groups(candidate_line)
            if not geo_groups:
                continue

            if discharge_index < len(geo_groups):
                return geo_groups[discharge_index].strip()
            else:
                return geo_groups[0].strip() + "（標籤/值數量不一致，請人工確認）"

    return ""



# 判斷是否為正式提單（B/L）頁面的強特徵詞。這批文件常是 Commercial
# Invoice + Packing List + Sea Waybill/B/L 合併成一個 PDF，Invoice/
# Packing List 裡也會出現「Port of Discharge」欄位，但值不一定與正式
# 提單一致（提單才是報關/貿易上採用的權威來源），所以優先鎖定含此
# 特徵詞的頁面。"NON-NEGOTIABLE" 是 B/L、Sea Waybill、Waybill 的
# 共同慣用語，Invoice/Packing List 不會出現。
BL_PAGE_MARKER = "NON-NEGOTIABLE"
FCR_PAGE_MARKER = "FORWARDER'S CARGO RECEIPT"


def filter_to_bl_pages(pages_words: list, pages_text: list) -> tuple:
    """
    若文件中有任何頁面含 B/L 強特徵詞，只回傳這些頁面的 words/text
    （及其在原始 PDF 中的頁碼），讓後續擷取邏輯只在正式提單範圍內
    比對，避免被 Invoice/Packing List 裡同名但不同值的欄位干擾。
    若整份文件都沒有這個特徵，回傳原始全部頁面（單頁提單本身也會
    含這個詞，不受影響）。
    """
    # 優先鎖定含 Port of Discharge 欄位的正式運送文件頁面，
    # 避免 Invoice / Packing List 內的商品文字或其他欄位被誤判。
    waybill_indices = [
        i for i, t in enumerate(pages_text)
        if BL_PAGE_MARKER in t.upper() and "PORT OF DISCHARGE" in t.upper()
    ]
    if waybill_indices:
        target_indices = waybill_indices
    else:
        fcr_indices = [
            i for i, t in enumerate(pages_text)
            if FCR_PAGE_MARKER in t.upper() and "PORT OF DISCHARGE" in t.upper()
        ]
        if fcr_indices:
            target_indices = fcr_indices
        else:
            target_indices = [
                i for i, t in enumerate(pages_text) if BL_PAGE_MARKER in t.upper()
            ]

    if not target_indices:
        return pages_words, pages_text, list(range(len(pages_text)))

    return (
        [pages_words[i] for i in target_indices],
        [pages_text[i] for i in target_indices],
        target_indices,
    )


def find_port_of_discharge(pages_words: list, pages_text: list) -> tuple:
    """
    依序嘗試各策略，任一成功即回傳 (結果字串, 已篩選的頁碼清單)：
    0. 若為多頁合併 PDF 且含正式 B/L/Waybill 頁面，先限縮範圍至該頁面
       （避免 Invoice/Packing List 裡同名欄位的不同值造成誤判）。
    1. 座標式左右欄比對（適用 ATE、SOM 等「標籤不在文字層」版型）
    2. 交錯行還原比對（適用 Maersk Waybill 這類標籤/值逐字交錯版型）
    3. 逗號地名 keyword 比對（適用 Expeditors、COSCO、Mainfreight、Everise）
    4. 空格分隔單字序位比對（適用 TG8 Maersk FCR 這類版型）
    5. SAVANNAH 已知港口關鍵字保底（寬鬆比對，容許 OCR 缺字誤植）
    回傳的頁碼清單供呼叫端在文字層失敗時，針對這些頁面做 OCR 補強重試。
    """
    filtered_words, filtered_text, page_indices = filter_to_bl_pages(
        pages_words, pages_text
    )

    result = try_coordinate_layout(filtered_words)
    if result:
        return result, page_indices

    result = try_keyword_layout(filtered_text)
    if result:
        return result, page_indices

    result = try_space_separated_layout(filtered_text)
    if result:
        return result, page_indices

    result = try_interleaved_row_layout(filtered_words)
    if result:
        return result, page_indices

    result = try_known_port_fallback("\n".join(filtered_text))
    if result:
        return result, page_indices

    return "", page_indices


def process_file(path: str, progress_callback=None) -> dict:
    """
    處理單一 PDF 檔案，回傳擷取結果。
    progress_callback(message: str)：若提供，用來回報處理進度
    （GUI 模式下更新畫面用）；未提供時退回 print。
    """
    def report(msg: str):
        if progress_callback:
            progress_callback(msg)
        else:
            print(msg)

    filename = os.path.basename(path)
    display_name = simplify_filename(filename)
    report(f"處理中: {filename}")

    pages_words, pages_text = extract_words_by_page(path)
    full_text = "\n".join(pages_text)
    source = "文字層 (PDF 未 Flatten)"

    if len(full_text) < MIN_TEXT_LEN_FOR_TEXTLAYER:
        report(f"  文字層內容過少（{len(full_text)} 字），改用逐頁 OCR 辨識...")
        source = "OCR (圖片型 PDF，逐頁辨識)"
        port = extract_port_via_ocr_pages(
            path,
            progress_callback=report,
        )
        return {
            "檔案": display_name,
            "完整路徑": path,
            "資料來源": source,
            "Port of Discharge": port or "[未擷取到，請人工確認]",
        }

    port, bl_page_indices = find_port_of_discharge(pages_words, pages_text)

    if not port:
        # 文字層在正式 B/L/Waybill 頁面上比對失敗——可能是該頁渲染
        # 異常（欄位字母被拆散、無法辨識），對這些頁面單獨補強 OCR
        # 後重試，而非放棄。只鎖定 B/L 頁面 OCR，避免整份大型合併
        # PDF（可能數十頁）全頁 OCR 造成效能負擔。
        report(f"  文字層比對未成功，嘗試對相關頁面逐頁補強 OCR...")
        port = extract_port_via_ocr_pages(
            path,
            page_indices=bl_page_indices,
            progress_callback=report,
        )
        if port:
            source = "文字層 + OCR 補強 (逐頁辨識)"

    if not port:
        report(f"  [注意] 未能自動判斷 Port of Discharge，請人工確認。")

    return {
        "檔案": display_name,
        "完整路徑": path,
        "資料來源": source,
        "Port of Discharge": port or "[未擷取到，請人工確認]",
    }


def write_results_to_xlsx(results: list, out_path: str) -> None:
    """
    將處理結果寫入 xlsx，A 欄（檔案）顯示簡化檔名，並建立指向
    原始 PDF 完整路徑的超連結，方便直接點擊開啟檢查。
    """
    FONT_NAME = "Arial"

    wb = Workbook()
    ws = wb.active
    ws.title = "Port of Discharge"

    headers = ["檔案", "資料來源", "Port of Discharge"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(name=FONT_NAME, bold=True)

    hyperlink_font = Font(name=FONT_NAME, color="0000FF", underline="single")
    normal_font = Font(name=FONT_NAME)

    for row in results:
        ws.append([row["檔案"], row["資料來源"], row["Port of Discharge"]])
        row_idx = ws.max_row
        for col_idx in (2, 3):
            ws.cell(row=row_idx, column=col_idx).font = normal_font

        cell_a = ws.cell(row=row_idx, column=1)
        full_path = row.get("完整路徑")
        if full_path:
            # file:// 超連結需要絕對路徑；Windows 路徑含反斜線，
            # openpyxl 的 hyperlink 屬性會原樣寫入，Excel 能正確解析
            cell_a.hyperlink = full_path
            cell_a.font = hyperlink_font
        else:
            cell_a.font = normal_font

    # 自動調整欄寬，方便直接閱讀
    for col_cells in ws.columns:
        max_len = max((len(str(c.value)) for c in col_cells if c.value), default=10)
        col_letter = col_cells[0].column_letter
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

    wb.save(out_path)




def load_checkpoint(checkpoint_path: str, selected_paths: list) -> list:
    """
    載入上次未完成批次的暫存結果，只保留本次仍有選取且路徑完全相同的檔案。
    已成功或已標記人工確認的檔案都可續跑略過，避免 crash 後重新 OCR。
    """
    if not os.path.exists(checkpoint_path):
        return []

    selected_set = {os.path.normcase(os.path.abspath(p)) for p in selected_paths}
    results = []
    try:
        wb = load_workbook(checkpoint_path, read_only=False, data_only=True)
        ws = wb.active
        for row_idx in range(2, ws.max_row + 1):
            cell_a = ws.cell(row=row_idx, column=1)
            full_path = cell_a.hyperlink.target if cell_a.hyperlink else None
            if not full_path:
                continue
            normalized = os.path.normcase(os.path.abspath(full_path))
            if normalized not in selected_set:
                continue
            results.append({
                "檔案": cell_a.value,
                "完整路徑": full_path,
                "資料來源": ws.cell(row=row_idx, column=2).value or "",
                "Port of Discharge": ws.cell(row=row_idx, column=3).value or "",
            })
        wb.close()
    except Exception as e:
        print(f"[警告] 無法讀取續跑暫存檔，將重新處理全部檔案: {e}")
        return []

    return results


def select_pdf_files() -> list:
    """
    僅顯示檔案選擇視窗，不建立完整 GUI 介面。
    使用者可一次選取多個 PDF，選取完成後由 PowerShell／命令提示字元
    顯示處理進度與結果。
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        paths = filedialog.askopenfilenames(
            title="選擇要處理的出貨文件 PDF（可複選）",
            filetypes=[("PDF files", "*.pdf *.PDF"), ("All files", "*.*")],
        )
        return list(paths)
    finally:
        root.destroy()


def wait_before_exit() -> None:
    """讓使用者能在 PowerShell／命令提示字元中查看處理結果。"""
    try:
        input("\n按 Enter 鍵關閉視窗...")
    except (EOFError, KeyboardInterrupt):
        pass


def main():
    selected_paths = select_pdf_files()

    if not selected_paths:
        print("未選擇任何 PDF，程式結束。")
        wait_before_exit()
        return

    total = len(selected_paths)
    output_dir = os.path.dirname(selected_paths[0])
    out_xlsx_path = os.path.join(output_dir, "port_of_discharge_result.xlsx")
    checkpoint_path = os.path.join(output_dir, "port_of_discharge_checkpoint.xlsx")

    results = load_checkpoint(checkpoint_path, selected_paths)
    completed_paths = {
        os.path.normcase(os.path.abspath(row["完整路徑"]))
        for row in results
        if row.get("完整路徑")
    }
    pending_paths = [
        path for path in selected_paths
        if os.path.normcase(os.path.abspath(path)) not in completed_paths
    ]

    print("=" * 72)
    print(f"已選擇 {total} 個 PDF，開始擷取 Port of Discharge。")
    if results:
        print(f"已由暫存檔載入 {len(results)} 個結果，本次剩餘 {len(pending_paths)} 個。")
    print("=" * 72)

    for index, path in enumerate(pending_paths, start=1):
        overall_index = len(results) + 1
        print(f"\n[{overall_index}/{total}] {os.path.basename(path)}")
        try:
            result = process_file(path)
            results.append(result)
            print(f"  結果: {result['Port of Discharge']}")
        except Exception as e:
            print(f"  [錯誤] 處理失敗: {e}")
            results.append({
                "檔案": simplify_filename(os.path.basename(path)),
                "完整路徑": path,
                "資料來源": "處理失敗",
                "Port of Discharge": f"[錯誤: {e}]",
            })

        # 每完成固定數量即覆寫暫存檔，避免長批次中斷後全部重跑
        if index % CHECKPOINT_INTERVAL == 0 or index == len(pending_paths):
            try:
                write_results_to_xlsx(results, checkpoint_path)
                print(f"  [暫存] 已保存 {len(results)}/{total} 個結果。")
            except Exception as e:
                print(f"  [警告] 暫存結果失敗: {e}")

    try:
        write_results_to_xlsx(results, out_xlsx_path)
    except PermissionError:
        print("\n[錯誤] 無法寫入 Excel 結果檔。")
        print("請先關閉已開啟的 port_of_discharge_result.xlsx，再重新執行。")
        wait_before_exit()
        return
    except Exception as e:
        print(f"\n[錯誤] 輸出 Excel 時發生例外: {e}")
        wait_before_exit()
        return

    # 正常完成後移除暫存檔；若中途 crash，暫存檔會保留供下次續跑
    try:
        if os.path.exists(checkpoint_path):
            os.remove(checkpoint_path)
    except Exception as e:
        print(f"[警告] 無法刪除暫存檔: {e}")

    error_count = sum(
        1 for row in results
        if "未擷取到" in row["Port of Discharge"]
        or "錯誤" in row["Port of Discharge"]
    )
    success_count = total - error_count

    print("\n" + "=" * 72)
    print("處理完成")
    print(f"成功擷取: {success_count} 個")
    print(f"需人工確認／失敗: {error_count} 個")
    print(f"結果檔案: {out_xlsx_path}")
    print("=" * 72)

    wait_before_exit()


if __name__ == "__main__":
    main()

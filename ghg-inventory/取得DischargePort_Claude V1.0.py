# -*- coding: utf-8 -*-
"""
出貨文件 Port of Discharge 擷取工具 (v2)

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
    直接執行本程式（例如在 PowerShell / 命令提示字元打
    python extract_port_of_discharge.py，或直接雙擊執行），
    會跳出系統原生的「選擇檔案」對話框，選擇要處理的 PDF 檔案
    （可複選）。選擇後對話框關閉，處理進度與結果會直接顯示在
    當下執行程式的終端機視窗中，不會另外開啟其他視窗。
    處理完成後，結果 Excel 會輸出至「所選檔案所在的資料夾」。

輸出：
    <所選檔案資料夾>/port_of_discharge_result.xlsx
"""

import os
import re
import shutil
import time
import tkinter as tk
from tkinter import filedialog

import pdfplumber
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font

# ---------- 設定 ----------
# OCR 解析度：僅需辨識清楚的印刷體大寫字母（港口/城市名稱），
# 150 DPI 已足夠清晰，且圖片尺寸只有 300 DPI 的約 1/4（面積約
# 1/4），OCR 速度可提升 3~4 倍，大量檔案時效果明顯。若辨識率
# 不理想再視情況調高。
OCR_DPI = 150
OCR_LANG = "eng"
MIN_TEXT_LEN_FOR_TEXTLAYER = 30

# 圖片型 PDF（全篇皆掃描件、無文字層）在「找不到 B/L/FCR 標記頁」時，
# 最多對前幾頁做 OCR 就好，不整份掃到底。這批文件的 Port of
# Discharge 欄位一定出現在單據最前面幾頁（B/L、Waybill、FCR 本身
# 通常就是第 1 頁，後面才是 Invoice/Packing List 明細），超過這個
# 頁數還沒找到，代表這份文件版型特殊，繼續往後 OCR 也多半找不到，
# 只會拖慢整批處理時間，不如及早停下標記人工確認。
OCR_MAX_PAGES_FULL_SCAN = 5

# 單一頁 OCR 的逾時秒數。少數 PDF 頁面可能因為掃描品質、超大尺寸、
# 或損毀等原因導致 tesseract 卡住不動，不設逾時的話，處理近百份
# 檔案時只要卡到一份，整批就會停在那裡動彈不得、看起來像當機。
OCR_TIMEOUT_SECONDS = 30

# Tesseract OCR 執行檔路徑設定。
#
# pytesseract 只是呼叫外部 tesseract.exe 的介面，本身不含 OCR 引擎，
# 所以電腦上必須另外安裝 Tesseract 本體。若安裝時沒有把它加入系統
# PATH（Windows 常見狀況），pytesseract 呼叫時就會報錯
# "tesseract is not installed or it's not in your PATH"。
#
# 這裡依序嘗試：
#   1. 先看系統 PATH 找不找得到 tesseract（shutil.which）——
#      如果找得到，代表安裝時有設定好 PATH，不需要額外指定路徑。
#   2. 找不到的話，依序檢查 Windows 常見的預設安裝路徑
#      （一般安裝在 Program Files，或只裝給目前使用者的 LOCALAPPDATA）。
#   3. 若都找不到，保留 None，稍後在 GUI 啟動時會跳出明確提示，
#      請使用者自行安裝或指定路徑，而不是等處理到一半才報錯。
#
# 若你的 Tesseract 安裝在其他路徑，可以直接把該路徑加進下面的
# CANDIDATE_TESSERACT_PATHS 清單，或者也可以不改程式，直接把
# Tesseract 的資料夾路徑加入系統環境變數 PATH（安裝時勾選相關選項
# 通常就會處理好）。
CANDIDATE_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    "/usr/bin/tesseract",
    "/usr/local/bin/tesseract",
    "/opt/homebrew/bin/tesseract",
]


def configure_tesseract_path() -> str:
    """
    設定 pytesseract 呼叫外部 tesseract 執行檔的路徑。

    回傳：實際設定成功的路徑字串；若完全找不到，回傳空字串
    （呼叫端可依此判斷是否要提示使用者手動安裝/指定路徑）。
    """
    # 1. 系統 PATH 裡找得到，直接用，不需要額外指定
    found_in_path = shutil.which("tesseract")
    if found_in_path:
        pytesseract.pytesseract.tesseract_cmd = found_in_path
        return found_in_path

    # 2. 依序檢查常見安裝路徑
    for candidate in CANDIDATE_TESSERACT_PATHS:
        if candidate and os.path.isfile(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate

    # 3. 都找不到
    return ""


TESSERACT_PATH = configure_tesseract_path()

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


def extract_text_via_ocr(path: str, max_pages: int = None, start_page: int = 1) -> str:
    """
    對整份 PDF（或指定頁面範圍）做 OCR，回傳合併後的文字。

    效能與穩定性考量：
    - start_page / max_pages：只 OCR 第 start_page 頁到第
      (start_page + max_pages - 1) 頁（1-indexed，含頭尾）。用於
      「先掃前幾頁，找不到再補掃剩餘頁面」的兩段式策略時，補掃
      階段可以從 start_page 接續，不必重複 OCR 已經掃過的頁面。
      max_pages 傳 None 時，從 start_page 掃到最後一頁。
    - 逐頁用 convert_from_path 搭配 first_page/last_page 只轉單頁，
      而不是一次把全部頁面轉成圖片塞進記憶體——大量檔案批次處理時，
      一次性把數十頁高解析度圖片都讀進記憶體是常見的記憶體爆量/
      crash 原因，逐頁轉、逐頁丟棄可以把尖峰記憶體壓在單頁範圍內。
    - 每頁 OCR 都設定逾時（OCR_TIMEOUT_SECONDS），避免少數異常頁面
      讓 tesseract 卡住不動，拖垮整批處理進度。
    """
    text_chunks = []

    try:
        with pdfplumber.open(path) as pdf:
            total_pages = len(pdf.pages)
    except Exception as e:
        print(f"  [錯誤] 無法讀取 PDF 頁數: {e}")
        return ""

    last_page = total_pages if max_pages is None else min(total_pages, start_page + max_pages - 1)
    if max_pages is not None and last_page < total_pages:
        print(f"  [提示] 共 {total_pages} 頁，僅 OCR 第 {start_page}-{last_page} 頁"
              f"（Port of Discharge 通常在單據前段，其餘頁面略過以節省時間）")

    for page_num in range(start_page, last_page + 1):
        try:
            images = convert_from_path(
                path, dpi=OCR_DPI, first_page=page_num, last_page=page_num,
            )
        except Exception as e:
            print(f"  [錯誤] 第 {page_num} 頁轉圖片失敗: {e}")
            continue
        if not images:
            continue
        img = images[0]
        try:
            text_chunks.append(
                pytesseract.image_to_string(
                    img, lang=OCR_LANG, timeout=OCR_TIMEOUT_SECONDS
                )
            )
        except RuntimeError:
            print(f"  [警告] 第 {page_num} 頁 OCR 逾時（超過 {OCR_TIMEOUT_SECONDS} 秒），略過")
        except Exception as e:
            print(f"  [警告] 第 {page_num} 頁 OCR 失敗: {e}")
        finally:
            img.close()  # 明確釋放圖片記憶體，避免大量檔案批次處理時累積

    return "\n".join(text_chunks).strip()


def extract_text_via_ocr_single_page(path: str, page_index: int) -> str:
    """
    只對 PDF 中指定的單一頁（0-indexed）做 OCR，不整份轉圖。
    用於文字層本身渲染異常（例如欄位字母被交錯拆散、無法辨識）
    但只有少數頁面需要補強辨識的情況，避免大型多頁 PDF 全頁 OCR
    造成的效能負擔。同樣套用逾時保護與明確釋放圖片記憶體。
    """
    try:
        images = convert_from_path(
            path, dpi=OCR_DPI,
            first_page=page_index + 1, last_page=page_index + 1,
        )
    except Exception as e:
        print(f"  [錯誤] 第 {page_index + 1} 頁轉圖片失敗: {e}")
        return ""
    if not images:
        return ""
    img = images[0]
    try:
        return pytesseract.image_to_string(
            img, lang=OCR_LANG, timeout=OCR_TIMEOUT_SECONDS
        ).strip()
    except RuntimeError:
        print(f"  [警告] 第 {page_index + 1} 頁 OCR 逾時（超過 {OCR_TIMEOUT_SECONDS} 秒），略過")
        return ""
    except Exception as e:
        print(f"  [警告] 第 {page_index + 1} 頁 OCR 失敗: {e}")
        return ""
    finally:
        img.close()


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

    row_a（Port of Loading 該行）必須是「乾淨的城市, 國家」格式
    （逗號分隔片段數 ≤ 2，例如 "KAOHSIUNG, TAIWAN"），藉此排除：
    - Invoice/Packing List 常見的「FOB TAICHUNG, TAIWAN」這種句子
      片段（FOB 開頭，不是欄位表格值）
    - 公司地址區塊裡的「...HSIANG, CHANG-HWA HSIEN, TAIWAN, R.O.C.」
      （逗號分隔片段數多達 3~4 段，且含 HSIEN/R.O.C. 等地址慣用詞）
    這批文件裡，正式表格版型的 Port of Loading 值一定是單純兩段式
    地名，不會是句子或多段地址，這個限制不會誤殺正常案例。
    """
    place_pattern = re.compile(r"^[A-Z][A-Z\s,.\-'()]+$")
    company_suffix_pattern = re.compile(
        r"\b(INC|CO\.|LTD|CORP|LLC|CO,)\b", re.IGNORECASE
    )
    # 常見句子開頭詞／地址慣用詞——若 row_a 含這些，代表是 Invoice
    # 內文或公司地址片段，不是表格欄位值，須排除
    sentence_or_address_noise = re.compile(
        r"\b(FOB|CIF|CFR|EXW|HSIEN|R\.O\.C\.?|SEC\.?|SECTION|ROAD|RD\.?"
        r"|STREET|VILLAGE|HSIANG)\b",
        re.IGNORECASE,
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
            if sentence_or_address_noise.search(right_a_text):
                continue
            # 要求含逗號（完整「城市, 國家」地名），排除像單獨一個
            # "TAIWAN" 這種國名，避免誤判到 shipper 地址等雜訊行
            if "," not in right_a_text:
                continue
            # 表格欄位值應為單純「城市, 國家」兩段式，逗號分隔片段數
            # 超過 2 段（如公司地址常見的三、四段式）就不是欄位值
            if right_a_text.count(",") > 1:
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
        r"^(PlaceofReceipt|PortofLoading|PortofDischarge|PlaceofDelivery"
        r"|FinalDestination)\.?", re.IGNORECASE,
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

            # 在其他行中找起始 x0 幾乎相同（同一欄位）、且非標籤行的候選。
            # 誤差收緊為 1.0pt（原本 2.0pt 太寬，曾誤把 x0 僅相差
            # 0.8pt、但欄位完全不同的 "Voyage No." 標籤行也當作候選）。
            # 同時要求 top 值相近（真正的交錯行，標籤與值是同一視覺行
            # 區塊內的產物，top 差距通常在 15pt 內；曾發生頁面最上方的
            # "RECEIPTNO." 與頁面中段的 "PortofDischarge" 標籤 x0 恰好
            # 只差 0.7pt、但 top 值相差達 250pt 卻仍被誤配對的案例）
            candidates = []
            for other_top, other_compact in rows_by_top.items():
                if other_top == top:
                    continue
                if abs(row_start_x0[other_top] - label_x0) > 1.0:
                    continue
                if abs(other_top - top) > 15.0:
                    continue
                if label_compact_pattern.match(other_compact):
                    continue
                # 值應為單純大寫地名（含逗號、括號、星號皆可——部分港口
                # 名稱後方會標註 * 號附註，如 "VANCOUVER*"），不含長句子
                # （若混入大量小寫字母，通常是條款內文而非欄位值）
                if re.match(r"^[A-Z][A-Za-z\s,.\-'()*]{1,40}$", other_compact) and \
                        not re.search(r"\b(the|of|to|and|or)\b", other_compact, re.IGNORECASE):
                    candidates.append((other_top, other_compact))

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



# 判斷是否為正式提單/收據（B/L、Waybill、FCR）頁面的強特徵詞。這批
# 文件常是 Commercial Invoice + Packing List + Sea Waybill/B/L/FCR
# 合併成一個 PDF，Invoice/Packing List 裡也會出現「Port of Discharge」
# 欄位，但值不一定與正式提單一致（提單/收據才是報關/貿易上採用的
# 權威來源），所以優先鎖定含這些特徵詞的頁面。
#   - "NON-NEGOTIABLE"：B/L、Sea Waybill、Waybill 的共同慣用語
#   - "FORWARDER'S CARGO RECEIPT"：Maersk FCR 專用標記（部分出貨
#     文件只有 FCR、沒有另外的 Waybill 頁，若不認這個標記，
#     filter_to_bl_pages 會 fallback 回傳全部頁面，導致 Invoice/
#     Packing List 頁面也被拿去跑欄位比對，容易誤判）
# Invoice/Packing List 不會出現這些詞。
BL_PAGE_MARKERS = ["NON-NEGOTIABLE", "FORWARDER'S CARGO RECEIPT"]


def filter_to_bl_pages(pages_words: list, pages_text: list) -> tuple:
    """
    若文件中有任何頁面含 B/L/FCR 強特徵詞，只回傳這些頁面的
    words/text（及其在原始 PDF 中的頁碼），讓後續擷取邏輯只在正式
    提單/收據範圍內比對，避免被 Invoice/Packing List 裡同名但不同值
    的欄位干擾。若整份文件都沒有這些特徵，回傳原始全部頁面（單頁
    提單本身也會含這些詞，不受影響）。
    """
    bl_indices = [
        i for i, t in enumerate(pages_text)
        if any(marker in t.upper() for marker in BL_PAGE_MARKERS)
    ]
    if not bl_indices:
        return pages_words, pages_text, list(range(len(pages_text)))

    return (
        [pages_words[i] for i in bl_indices],
        [pages_text[i] for i in bl_indices],
        bl_indices,
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

    result = try_interleaved_row_layout(filtered_words)
    if result:
        return result, page_indices

    result = try_keyword_layout(filtered_text)
    if result:
        return result, page_indices

    result = try_space_separated_layout(filtered_text)
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
        report(f"  文字層內容過少（{len(full_text)} 字），改用 OCR 辨識圖片...")
        source = "OCR (圖片型 PDF)"

        # 先只 OCR 前幾頁（Port of Discharge 通常在單據前段），
        # 找不到才逐步擴大到整份掃完，避免一開始就對整份多頁
        # 掃描檔全部 OCR，浪費時間在跟目的港無關的後段明細頁。
        ocr_text = extract_text_via_ocr(path, max_pages=OCR_MAX_PAGES_FULL_SCAN)
        port = (
            try_keyword_layout([ocr_text])
            or try_space_separated_layout([ocr_text])
            or try_known_port_fallback(ocr_text)
        )
        if not port:
            with pdfplumber.open(path) as pdf:
                total_pages = len(pdf.pages)
            if total_pages > OCR_MAX_PAGES_FULL_SCAN:
                report(f"  前 {OCR_MAX_PAGES_FULL_SCAN} 頁未找到，擴大掃描剩餘頁面...")
                ocr_text_rest = extract_text_via_ocr(
                    path, max_pages=None, start_page=OCR_MAX_PAGES_FULL_SCAN + 1
                )
                port = (
                    try_keyword_layout([ocr_text_rest])
                    or try_space_separated_layout([ocr_text_rest])
                    or try_known_port_fallback(ocr_text_rest)
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
        report(f"  文字層比對未成功，嘗試對相關頁面補強 OCR...")
        for page_idx in bl_page_indices:
            ocr_text = extract_text_via_ocr_single_page(path, page_idx)
            if not ocr_text:
                continue
            port = (
                try_keyword_layout([ocr_text])
                or try_space_separated_layout([ocr_text])
                or try_known_port_fallback(ocr_text)
            )
            if port:
                source = "文字層 + OCR 補強 (B/L 頁面渲染異常)"
                break

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


def select_pdf_files() -> list:
    """
    跳出系統原生的「選擇檔案」對話框，讓使用者選擇要處理的 PDF
    （可複選）。只用來跳出這個對話框，不建立、不顯示任何額外視窗
    ——對話框關閉後（不論選擇或取消），背後這個隱藏的 tkinter 根
    視窗會立刻 destroy 掉，不會留在畫面上。
    """
    root = tk.Tk()
    root.withdraw()  # 不顯示主視窗，只用來跳出檔案選擇對話框
    try:
        paths = filedialog.askopenfilenames(
            title="選擇要處理的出貨文件 PDF（可複選）",
            filetypes=[("PDF files", "*.pdf *.PDF"), ("All files", "*.*")],
        )
    finally:
        root.destroy()
    return list(paths)


def main():
    print("=" * 60)
    print("出貨文件 Port of Discharge 擷取工具")
    print("=" * 60)

    if TESSERACT_PATH:
        print(f"Tesseract OCR：{TESSERACT_PATH}")
    else:
        print(
            "[警告] 找不到 Tesseract OCR，圖片型（掃描）PDF 將無法辨識。\n"
            "       文字型 PDF 不受影響，仍可正常處理。\n"
            "       如需處理掃描 PDF，請安裝 Tesseract："
            "https://github.com/UB-Mannheim/tesseract/wiki\n"
            "       安裝後若仍找不到，可在原始碼的 "
            "CANDIDATE_TESSERACT_PATHS 清單加入實際安裝路徑。"
        )
    print()

    paths = select_pdf_files()
    if not paths:
        print("未選擇任何檔案，程式結束。")
        return

    total = len(paths)
    print(f"已選擇 {total} 個檔案，開始處理...\n")

    output_dir = os.path.dirname(paths[0])
    out_xlsx_path = os.path.join(output_dir, "port_of_discharge_result.xlsx")

    def save_progress(results_so_far: list):
        """存一次目前已完成的結果。中途意外中斷（Ctrl+C、當機、
        關閉終端機）時，前面已經跑完、花了時間的部分不會白費——
        重新執行只需要處理剩下還沒跑過的檔案即可。"""
        try:
            write_results_to_xlsx(results_so_far, out_xlsx_path)
        except Exception as e:
            print(f"\n[錯誤] 寫入 Excel 失敗: {e}")

    results = []
    batch_start_time = time.time()
    try:
        for i, path in enumerate(paths, start=1):
            file_start_time = time.time()
            print(f"[{i}/{total}]", end=" ")
            try:
                result = process_file(path)
                results.append(result)
            except Exception as e:
                print(f"[錯誤] 處理 {os.path.basename(path)} 時發生例外: {e}")
                results.append({
                    "檔案": simplify_filename(os.path.basename(path)),
                    "完整路徑": path,
                    "資料來源": "處理失敗",
                    "Port of Discharge": f"[錯誤: {e}]",
                })

            elapsed = time.time() - file_start_time
            if elapsed > 5:
                # 單一檔案花超過 5 秒才提示耗時，避免正常的文字層
                # 檔案（通常不到 1 秒）也被洗版，只在真的慢（多半是
                # OCR）的檔案上讓使用者知道「還在跑，不是卡死」
                print(f"  （耗時 {elapsed:.1f} 秒）")

            # 每處理 20 個檔案就存一次中繼結果，避免處理到一半發生
            # 未預期狀況（如意外關閉終端機）時，前面已完成的部分
            # 完全遺失、要整批從頭重跑
            if i % 20 == 0 and i < total:
                save_progress(results)
                print(f"  [已自動儲存進度：{i}/{total}]")

    except KeyboardInterrupt:
        print(f"\n\n[中斷] 使用者按下 Ctrl+C，已處理 {len(results)}/{total} 個檔案。")
        print("已完成的部分將先儲存...")

    total_elapsed = time.time() - batch_start_time
    save_progress(results)

    print("\n" + "=" * 60)
    print("處理完成" if len(results) == total else "處理中斷")
    print("=" * 60)
    for r in results:
        print(f"{r['檔案']:<40} | {r['資料來源']:<28} | {r['Port of Discharge']}")

    error_count = sum(
        1 for r in results
        if "未擷取到" in r["Port of Discharge"] or "錯誤" in r["Port of Discharge"]
    )
    success_count = len(results) - error_count
    minutes, seconds = divmod(int(total_elapsed), 60)
    print()
    print(f"共處理 {len(results)}/{total} 個檔案，成功擷取 {success_count} 個，"
          f"需人工確認/失敗 {error_count} 個")
    print(f"總耗時: {minutes} 分 {seconds} 秒")
    print(f"結果已輸出至: {out_xlsx_path}")

    input("\n按 Enter 鍵結束...")


if __name__ == "__main__":
    main()

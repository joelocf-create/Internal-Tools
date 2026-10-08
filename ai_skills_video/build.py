import subprocess, random, math, os, wave, sys
from PIL import Image, ImageDraw, ImageFont

W, H, FPS = 1280, 720, 15
OUT = os.path.dirname(os.path.abspath(__file__))
FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
F = lambda s: ImageFont.truetype(FONT, s)
INK = (35, 40, 60); BLUE = (40, 110, 200); RED = (214, 70, 60); GREEN = (40, 150, 90); ORG = (235, 150, 40)
BG = (250, 250, 246)

# (narration for TTS, subtitle, title)
SCENES = [
    ("AI 很聰明，但每次都要你重新交代做事方法，是不是很累？", "AI 很聰明，但每次都要重新交代做事方法，很累？", "問題"),
    ("AI 的 Skills，就是把一套做事流程，打包成一份可重複使用的說明書。", "AI 的 Skills：把一套做事流程，打包成可重複使用的說明書", "什麼是 Skills"),
    ("一個 Skill 就是一個資料夾，裡面有說明檔、範例，還有腳本。", "一個 Skill＝一個資料夾：說明檔、範例、腳本", "Skill 長怎樣"),
    ("平常，AI 只讀每個 Skill 的名稱和簡介，佔用很少的記憶。", "平常只讀名稱與簡介，幾乎不佔記憶", "第一步：待命"),
    ("當任務符合時，AI 才會自動載入完整內容，照著步驟執行。", "任務符合時，才自動載入完整內容並照步驟執行", "第二步：觸發"),
    ("好處有三：品質穩定、不用重複貼提示詞，團隊還能共用。", "好處：品質穩定、免重複貼提示詞、團隊共用", "為什麼有用"),
    ("就像給新同事一本標準作業手冊，教一次，永遠會做。", "像給新同事一本 SOP：教一次，永遠會做", "比喻"),
    ("現在，你也可以把自己的專長，變成 AI 的 Skill。", "把你的專長，變成 AI 的 Skill", "開始吧"),
]
TARGET = 60.0

# ---------- TTS: pick speed so total ~ TARGET ----------
def tts(speed):
    durs = []
    for i, (n, _, _) in enumerate(SCENES):
        p = f"{OUT}/n{i}.wav"
        subprocess.run(["espeak-ng", "-v", "cmn", "-s", str(speed), "-p", "55", "-w", p, n], check=True)
        with wave.open(p) as w: durs.append(w.getnframes() / w.getframerate())
    return durs
speed = 170
while True:
    durs = tts(speed)
    tot = sum(d + 0.7 for d in durs)
    print("speed", speed, "total", round(tot, 1))
    if tot <= TARGET or speed >= 260: break
    speed += 10
pad = (TARGET - sum(durs)) / len(durs)
pad = max(0.3, min(pad, 2.5))
scene_len = [d + pad for d in durs]
total = sum(scene_len); print("final", round(total, 1), "pad", round(pad, 2))

# ---------- hand-drawn primitives ----------
rnd = random.Random(7)
def jl(d, p0, p1, c=INK, w=4):
    n = max(2, int(math.dist(p0, p1) // 25))
    pts = [(p0[0] + (p1[0] - p0[0]) * k / n + rnd.uniform(-1.5, 1.5), p0[1] + (p1[1] - p0[1]) * k / n + rnd.uniform(-1.5, 1.5)) for k in range(n + 1)]
    d.line(pts, fill=c, width=w, joint="curve")
def rect(d, x, y, w, h, c=INK, fill=None, lw=4):
    if fill: d.rectangle([x, y, x + w, y + h], fill=fill)
    for a, b in [((x, y), (x + w, y)), ((x + w, y), (x + w, y + h)), ((x + w, y + h), (x, y + h)), ((x, y + h), (x, y))]: jl(d, a, b, c, lw)
def circ(d, cx, cy, r, c=INK, fill=None, lw=4):
    if fill: d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill)
    pts = [(cx + (r + rnd.uniform(-1.5, 1.5)) * math.cos(t / 20 * 2 * math.pi), cy + (r + rnd.uniform(-1.5, 1.5)) * math.sin(t / 20 * 2 * math.pi)) for t in range(22)]
    d.line(pts, fill=c, width=lw, joint="curve")
def txt(d, x, y, s, size=34, c=INK, anchor="mm"): d.text((x, y), s, font=F(size), fill=c, anchor=anchor)
def arrow(d, p0, p1, c=INK, w=5):
    jl(d, p0, p1, c, w); a = math.atan2(p1[1] - p0[1], p1[0] - p0[0])
    for s in (-0.5, 0.5): jl(d, p1, (p1[0] - 26 * math.cos(a + s), p1[1] - 26 * math.sin(a + s)), c, w)
def person(d, x, y, c=INK):
    circ(d, x, y - 60, 24, c); jl(d, (x, y - 36), (x, y + 40), c); jl(d, (x - 36, y - 10), (x + 36, y - 10), c)
    jl(d, (x, y + 40), (x - 28, y + 100), c); jl(d, (x, y + 40), (x + 28, y + 100), c)
def robot(d, x, y, c=BLUE):
    rect(d, x - 55, y - 55, 110, 90, c, (225, 238, 255)); rect(d, x - 40, y + 35, 80, 70, c, (225, 238, 255))
    circ(d, x - 22, y - 15, 11, c); circ(d, x + 22, y - 15, 11, c); jl(d, (x - 22, y + 14), (x + 22, y + 14), c)
    jl(d, (x, y - 55), (x, y - 80), c); circ(d, x, y - 88, 8, RED)
def folder(d, x, y, label, c=ORG, w=170, h=110, size=26):
    rect(d, x, y - 18, 70, 18, c, (255, 232, 190)); rect(d, x, y, w, h, c, (255, 236, 200)); txt(d, x + w / 2, y + h / 2, label, size)
def bubble(d, x, y, w, h, s, size=32, c=INK):
    rect(d, x, y, w, h, c, (255, 255, 255)); txt(d, x + w / 2, y + h / 2, s, size, c)
def check(d, x, y, c=GREEN): jl(d, (x - 14, y), (x - 4, y + 12), c, 7); jl(d, (x - 4, y + 12), (x + 16, y - 14), c, 7)

# ---------- each scene = list of (start_sec, dur_sec, drawfn) ----------
def S1():
    def a(d): person(d, 240, 380); txt(d, 240, 520, "你", 32)
    def b(d): robot(d, 1000, 360); txt(d, 1000, 520, "AI", 32, BLUE)
    def c(d):
        for k in range(3):
            rect(d, 430 + k * 28, 200 + k * 22, 200, 130, INK, (255, 255, 255)); [jl(d, (450 + k * 28, 235 + k * 22 + j * 26), (610 + k * 28, 235 + k * 22 + j * 26), (150, 150, 160), 3) for j in range(3)]
        txt(d, 580, 415, "長長的提示詞…", 30, RED)
    def e(d): arrow(d, (660, 300), (900, 330), RED); bubble(d, 700, 150, 230, 70, "又要重貼一次？！", 28, RED)
    return [(0.3, 1.0, a), (1.5, 1.0, b), (2.8, 1.5, c), (4.5, 1.2, e)]
def S2():
    def a(d):
        for k, s in enumerate(["做法 A", "格式 B", "注意 C", "範例 D"]): bubble(d, 120 + (k % 2) * 190, 200 + (k // 2) * 120, 160, 70, s, 28)
    def b(d): arrow(d, (520, 300), (660, 300), INK, 7)
    def c(d):
        rect(d, 700, 170, 330, 260, BLUE, (225, 238, 255), 5); jl(d, (865, 170), (865, 430), BLUE, 5)
        txt(d, 783, 270, "SKILL", 40, BLUE); txt(d, 783, 330, "說明書", 34, BLUE); txt(d, 948, 300, "可重複使用", 26, GREEN)
    def e(d): txt(d, 640, 520, "做事流程  →  一份說明書", 42, RED)
    return [(0.3, 2.0, a), (2.6, 0.7, b), (3.4, 1.8, c), (5.6, 1.3, e)]
def S3():
    def a(d): folder(d, 200, 160, "my-skill/", w=240, h=90, size=34)
    def b(d): jl(d, (230, 250), (230, 560), INK, 4); [jl(d, (230, y), (290, y), INK, 4) for y in (310, 430, 550)]
    def c(d): bubble(d, 300, 280, 330, 60, "SKILL.md  說明檔", 30, BLUE); txt(d, 800, 310, "← 名稱＋簡介＋步驟", 28, BLUE, "lm")
    def e(d): bubble(d, 300, 400, 330, 60, "examples/  範例", 30, GREEN); txt(d, 800, 430, "← 好的成品長怎樣", 28, GREEN, "lm")
    def f(d): bubble(d, 300, 520, 330, 60, "scripts/  腳本", 30, ORG); txt(d, 800, 550, "← 自動執行的工具", 28, ORG, "lm")
    return [(0.3, 1.0, a), (1.3, 0.6, b), (2.0, 1.2, c), (3.4, 1.2, e), (4.8, 1.2, f)]
def S4():
    def a(d): robot(d, 200, 340)
    def b(d):
        for k, s in enumerate(["寫報告", "查關稅", "做簡報", "整理 Excel"]):
            folder(d, 420 + (k % 2) * 260, 190 + (k // 2) * 170, s, w=200, h=100, size=28)
    def c(d):
        for k in range(4): txt(d, 520 + (k % 2) * 260, 205 + (k // 2) * 170 + 125, "名稱＋簡介", 20, GREEN)
        arrow(d, (320, 330), (400, 290), GREEN)
    def e(d):
        txt(d, 640, 585, "記憶用量", 26, INK, "rm"); rect(d, 660, 565, 400, 40, INK); rect(d, 662, 567, 28, 36, GREEN, GREEN); txt(d, 1080, 585, "極少", 26, GREEN, "lm")
    return [(0.3, 0.9, a), (1.3, 1.6, b), (3.2, 1.2, c), (4.7, 1.5, e)]
def S5():
    def a(d): bubble(d, 90, 220, 260, 80, "幫我做關稅摘要", 30, RED)
    def b(d): arrow(d, (360, 260), (450, 260), INK, 6); robot(d, 560, 300); circ(d, 700, 220, 28, GREEN, None, 5); jl(d, (720, 240), (755, 275), GREEN, 7)
    def c(d): folder(d, 790, 180, "關稅 Skill", c=GREEN, w=220, h=110, size=30); arrow(d, (730, 250), (780, 240), GREEN)
    def e(d):
        arrow(d, (900, 310), (900, 360), INK, 6)
        for k, s in enumerate(["1 收集資料", "2 套用格式", "3 輸出結果"]):
            rect(d, 790, 375 + k * 62, 300, 52, INK, (255, 255, 255)); txt(d, 830, 401 + k * 62, s, 28, INK, "lm"); check(d, 1060, 401 + k * 62)
    return [(0.3, 1.0, a), (1.5, 1.5, b), (3.2, 1.3, c), (4.8, 2.2, e)]
def S6():
    out = []
    items = [("品質穩定", GREEN, 200), ("免重複貼提示詞", BLUE, 520), ("團隊共用", ORG, 840)]
    for k, (s, c, x) in enumerate(items):
        def f(d, s=s, c=c, x=x, k=k):
            rect(d, x, 210, 260, 280, c, (255, 255, 255), 5); txt(d, x + 130, 270, str(k + 1), 72, c); txt(d, x + 130, 380, s, 32, INK); check(d, x + 130, 450, c)
        out.append((0.4 + k * 1.6, 1.2, f))
    return out
def S7():
    def a(d): person(d, 260, 380, BLUE); txt(d, 260, 520, "新同事", 30, BLUE)
    def b(d):
        rect(d, 520, 270, 240, 190, ORG, (255, 236, 200), 5); txt(d, 640, 340, "標準作業", 34); txt(d, 640, 390, "手冊 SOP", 34); arrow(d, (500, 360), (400, 360), INK, 6)
    def c(d): robot(d, 1000, 360); arrow(d, (780, 360), (900, 360), INK, 6); txt(d, 1000, 520, "AI + Skill", 30, BLUE)
    def e(d): txt(d, 640, 555, "教一次，永遠會做", 46, RED)
    return [(0.3, 1.0, a), (1.5, 1.4, b), (3.2, 1.3, c), (5.0, 1.2, e)]
def S8():
    def a(d): robot(d, 400, 380); circ(d, 400, 230, 34, ORG, (255, 244, 180), 5); [jl(d, (400 + 50 * math.cos(t), 230 + 50 * math.sin(t)), (400 + 66 * math.cos(t), 230 + 66 * math.sin(t)), ORG, 5) for t in [i * math.pi / 4 for i in range(8)]]
    def b(d): txt(d, 850, 300, "AI Skills", 92, BLUE); jl(d, (650, 350), (1050, 350), RED, 7)
    def c(d): txt(d, 850, 440, "把專長，變成可重複的能力", 36, INK)
    return [(0.3, 1.2, a), (1.8, 1.5, b), (3.6, 1.8, c)]
BUILD = [S1, S2, S3, S4, S5, S6, S7, S8]

def layer(fn):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0)); fn(ImageDraw.Draw(im)); return im
layers = [[(st, du, layer(fn)) for st, du, fn in B()] for B in BUILD]
sub_font = F(40); title_font = F(44)

def wrap(s, d, maxw):
    lines, cur = [], ""
    for ch in s:
        if d.textlength(cur + ch, font=sub_font) > maxw: lines.append(cur); cur = ch
        else: cur += ch
    return lines + [cur]

def frame(si, t):
    im = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(im)
    for x in range(0, W, 40): d.line([(x, 0), (x, H)], fill=(240, 240, 236))
    for y in range(0, H, 40): d.line([(0, y), (W, y)], fill=(240, 240, 236))
    rect(d, 14, 14, W - 28, H - 28, (170, 175, 185), None, 3)
    d.text((50, 50), f"{si+1}/{len(SCENES)}  {SCENES[si][2]}", font=title_font, fill=BLUE, anchor="lm")
    for st, du, lay in layers[si]:
        if t < st: continue
        p = min(1, (t - st) / du); w = int(W * p)
        if w > 0:
            im.paste(lay.crop((0, 0, w, H)), (0, 0), lay.crop((0, 0, w, H)))
            if p < 1: d.ellipse([w - 8, H // 2 - 8, w + 8, H // 2 + 8], fill=(0, 0, 0, 0)) if False else None
    d = ImageDraw.Draw(im)
    lines = wrap(SCENES[si][1], d, W - 160); bh = 30 + 52 * len(lines)
    d.rectangle([0, H - bh - 20, W, H], fill=(20, 24, 40))
    for k, ln in enumerate(lines): d.text((W / 2, H - bh + 20 + 52 * k + 26), ln, font=sub_font, fill=(255, 255, 255), anchor="mm")
    return im

# ---------- audio ----------
concat = []
for i, d_ in enumerate(durs):
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", f"{OUT}/n{i}.wav", "-af", f"adelay=300|300,apad=pad_dur={scene_len[i]}", "-t", str(scene_len[i]), "-ar", "44100", "-ac", "1", f"{OUT}/s{i}.wav"], check=True)
    concat.append(f"file 's{i}.wav'")
open(f"{OUT}/list.txt", "w").write("\n".join(concat))
subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", f"{OUT}/list.txt", "-af", "volume=2.0", f"{OUT}/audio.wav"], check=True)

# ---------- SRT ----------
def ts(t): return f"{int(t//3600):02}:{int(t%3600//60):02}:{int(t%60):02},{int(t*1000)%1000:03}"
t0, srt = 0, []
for i, L in enumerate(scene_len):
    srt.append(f"{i+1}\n{ts(t0+0.3)} --> {ts(t0+0.3+durs[i])}\n{SCENES[i][1]}\n"); t0 += L
open(f"{OUT}/ai_skills.srt", "w", encoding="utf-8").write("\n".join(srt))

# ---------- video ----------
p = subprocess.Popen(["ffmpeg", "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}", "-r", str(FPS), "-i", "-", "-i", f"{OUT}/audio.wav",
                      "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "128k", "-shortest", f"{OUT}/ai_skills.mp4"], stdin=subprocess.PIPE)
for si, L in enumerate(scene_len):
    for f in range(int(round(L * FPS))): p.stdin.write(frame(si, f / FPS).tobytes())
p.stdin.close(); p.wait()
print("done")

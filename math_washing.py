import sys
import os
import tempfile
import warnings
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk, simpledialog
import re
import time
import threading
import ctypes
import hashlib
import json
import base64
import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO

warnings.filterwarnings("ignore", category=FutureWarning)

try:
    import fitz
    import pyautogui
    import pyperclip
    import pygetwindow as gw
    from PIL import Image, ImageGrab, ImageFilter, ImageEnhance
    from google import genai
    from google.genai import types
except ImportError as e:
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("라이브러리 오류",
        f"필수 라이브러리가 없습니다:\n{e}\n\n"
        "pip install google-genai pymupdf pillow pyautogui pyperclip pygetwindow 를 실행하세요.")
    sys.exit(1)

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

pyautogui.PAUSE = 0.05
pyautogui.FAILSAFE = True

# ── win32com 로드 (없으면 pyautogui 전용 모드) ───────────────
try:
    import win32com.client as _w32
    import pythoncom
    import win32clipboard
    _WIN32_AVAILABLE = True
except ImportError:
    _WIN32_AVAILABLE = False
    win32clipboard = None

# ── matplotlib (풀이 뷰어용) ────────────────────────────────
try:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.rcParams['font.family']        = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False
    _MPL_AVAILABLE = True
except ImportError:
    _MPL_AVAILABLE = False

# ============================================================
# [설정]
# ============================================================
GCP_PROJECT  = "project-202ee3d7-34d9-426a-be5"
GCP_LOCATION = "us-central1"
MODEL_PRO   = "gemini-2.5-pro"    # 워싱 / 타이핑-해설있음
MODEL_FLASH = "gemini-2.5-flash"  # 타이핑-해설없음

# 드롭다운용 모델 목록 (앱 시작 후 백그라운드 조회)
_AVAILABLE_MODELS: list[str] = [MODEL_PRO, MODEL_FLASH]
_EXCLUDE_KEYWORDS = ("tts", "image", "embedding", "live", "computer-use", "audio")

def _fetch_available_models() -> list[str]:
    """Vertex AI에서 텍스트 생성 Gemini 모델 목록 조회."""
    try:
        raw = list(_gemini_client.models.list())
        names = []
        for m in raw:
            n = m.name.split("/")[-1]  # publishers/google/models/XXX → XXX
            if not n.startswith("gemini"):
                continue
            if any(kw in n for kw in _EXCLUDE_KEYWORDS):
                continue
            names.append(n)
        return sorted(set(names)) or _AVAILABLE_MODELS
    except Exception:
        return list(_AVAILABLE_MODELS)

def _pick_model(task_mode: str, style: str) -> tuple[str, int]:
    """(model_id, thinking_budget) 반환"""
    if task_mode == "타이핑" and style == "해설없음":
        return MODEL_FLASH, 1024
    return MODEL_PRO, 4096
stop_flag  = False
image_stack = []

# ── 결과 캐시 설정
_CACHE_DIR     = Path.home() / ".math_washing_cache"
_CACHE_MAX_AGE = 7 * 24 * 3600   # 7일
_BATCH_WORKERS = 3                # 동시 API 콜 수

# ── 풀이저장 DB
_HISTORY_DIR   = _CACHE_DIR / "history"
_FINETUNE_DB   = _CACHE_DIR / "finetune_dataset.json"
_AI_CACHE_DB   = _CACHE_DIR / "ai_output_cache.json"
_cache_write_lock = threading.Lock()

# Vertex AI 글로벌 클라이언트 (gcloud auth application-default login 완료 전제)
_gemini_client = genai.Client(
    vertexai=True,
    project=GCP_PROJECT,
    location=GCP_LOCATION,
)

# ============================================================
# ── 이미지 → genai Part 변환 헬퍼
# ============================================================
def pil_to_part(img: Image.Image) -> types.Part:
    buf = BytesIO()
    img.save(buf, format="PNG")
    return types.Part.from_bytes(data=buf.getvalue(), mime_type="image/png")


# ============================================================
# ── 결과 캐시 헬퍼
# ============================================================
def _img_bytes(img: Image.Image) -> bytes:
    buf = BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def _cache_key(img: Image.Image, task_mode: str, q_type: str, target_q: str,
               style: str = "") -> str:
    raw = _img_bytes(img) + f"{task_mode}|{q_type}|{target_q or ''}|{style}".encode()
    return hashlib.md5(raw).hexdigest()

def _cache_get(key: str) -> str | None:
    path = _CACHE_DIR / f"{key}.txt"
    if not path.exists():
        return None
    if time.time() - path.stat().st_mtime > _CACHE_MAX_AGE:
        path.unlink(missing_ok=True)
        return None
    return path.read_text(encoding="utf-8")

def _cache_put(key: str, result: str):
    _CACHE_DIR.mkdir(exist_ok=True)
    (_CACHE_DIR / f"{key}.txt").write_text(result, encoding="utf-8")


# ============================================================
# 1. 한글(HWP) 입력 — COM Dispatch + Visible 방식
# ============================================================
# 표준 방식: 프로그램이 HWP 인스턴스를 직접 띄우고 제어.
# 이미 열려있는 HWP에 붙으려 하지 않고, 자체 인스턴스에서 작업.

_hwp_com            = None   # 메인 스레드에서 생성한 HWP COM 객체
_com_alive          = False  # HWP 인스턴스가 살아있는지
_hwp_marshal_stream = None   # 스레드 간 마샬링용

def com_start_hwp(open_file: str = None) -> tuple:
    """
    HWP 인스턴스를 Dispatch로 생성하고 화면에 표시.
    open_file: 열 파일 경로 (없으면 빈 문서).
    반환: (성공여부, 메시지)
    """
    global _hwp_com, _com_alive, _hwp_marshal_stream

    if not _WIN32_AVAILABLE:
        return False, "pywin32가 설치되지 않음. pip install pywin32"

    # 기존 인스턴스가 살아있으면 재사용
    if _hwp_com is not None:
        try:
            _ = _hwp_com.XHwpDocuments.Count
            _com_alive = True
            return True, "기존 HWP 인스턴스 재사용"
        except Exception:
            _hwp_com = None
            _com_alive = False

    # Dispatch로 새 인스턴스 생성
    last_err = ""
    for cid in ["HWPFrame.HwpObject", "HWPFrame.HwpObject.1", "Hwp.Application"]:
        try:
            obj = _w32.Dispatch(cid)
            # 보안 모듈 등록 (2022 필수)
            try:
                obj.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass

            # 화면에 표시 ★ 핵심
            try:
                obj.XHwpWindows.Item(0).Visible = True
            except Exception:
                pass

            # 파일 열거나 빈 문서로
            if open_file:
                try:
                    fmt = "HWPX" if open_file.lower().endswith(".hwpx") else "HWP"
                    obj.Open(open_file, fmt, "forceopen:true")
                except Exception as e:
                    return False, f"파일 열기 실패: {e}"

            _hwp_com   = obj
            _com_alive = True

            # 마샬 스트림 생성 (워커 스레드에서 사용)
            try:
                _hwp_marshal_stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch, obj)
            except Exception as e:
                _hwp_marshal_stream = None

            _update_com_status(True, f"Dispatch({cid})")
            return True, f"HWP 인스턴스 생성 완료 ({cid})"
        except Exception as e:
            last_err = str(e)
            continue

    _com_alive = False
    _update_com_status(False, "")
    return False, f"HWP Dispatch 실패: {last_err}"


def _update_com_status(ok: bool, cid: str):
    """COM 상태 라벨 업데이트 (GUI 스레드 안전)"""
    try:
        if ok:
            com_status_label.config(
                text=f"🟢 HWP 인스턴스 활성  ({cid})",
                fg="#006600"
            )
        else:
            com_status_label.config(
                text="🔴 HWP 미시작 — [HWP 시작] 버튼을 누르세요",
                fg="#CC0000"
            )
    except Exception:
        pass


# ── pyautogui 폴백 헬퍼 ─────────────────────────────────────
def send_hotkey(*keys, delay=0.3):
    pyautogui.hotkey(*keys)
    time.sleep(delay)

def _close_formula(delay):
    pyautogui.keyDown('shift')
    time.sleep(0.05)
    pyautogui.press('escape')
    pyautogui.keyUp('shift')
    time.sleep(delay)

def paste_text(text, delay=0.1):
    if not text: return
    pyperclip.copy(text)
    time.sleep(0.15)
    send_hotkey('ctrl', 'v', delay=delay)


# ── 메인 입력 루프 ───────────────────────────────────────────
def _split_hwp_problems(text: str) -> list:
    """HWP 입력 텍스트를 문항별로 분리. <미주> 시작 패턴 기준."""
    if '<미주>' not in text:
        return [text] if text.strip() else []
    blocks = re.split(r'(?m)(?=(?:^\d{1,2}\s*<미주>|^<미주>))', text)
    return [b.strip() for b in blocks if b.strip()]


def _insert_page_break(thread_hwp, use_com, WD):
    """강제 쪽 나누기(Ctrl+Shift+Enter) 삽입."""
    if use_com:
        try:
            thread_hwp.HAction.Run("BreakPage")
            return
        except Exception as e:
            _log_com_error("페이지 나누기", e)
    time.sleep(0.1 * WD)
    pyautogui.hotkey('ctrl', 'shift', 'enter')


def _process_one_block(prob_block, thread_hwp, use_com, WD, ID):
    """문항 블록 하나(표·그래프·수식·텍스트)를 HWP에 입력."""
    global stop_flag
    top_parts = re.split(r'(<표>.*?</표>|<그래프>.*?</그래프>)', prob_block, flags=re.DOTALL)

    for top_part in top_parts:
        if stop_flag: break
        if not top_part: continue

        # ── 표 삽입 ──────────────────────────────────────────
        if top_part.startswith('<표>') and top_part.endswith('</표>'):
            inner = top_part[4:-5]
            rows = [r.strip() for r in inner.splitlines() if r.strip()]
            if not rows:
                continue
            if use_com:
                try:
                    # 단일 셀(1행 1열) 표 생성 — 내용 전체를 하나의 박스에
                    hset = thread_hwp.HParameterSet.HTableCreation.HSet
                    thread_hwp.HAction.GetDefault("TableCreate", hset)
                    hset.SetItem("Rows", 1)
                    hset.SetItem("Cols", 1)
                    hset.SetItem("WidthType", 0)
                    thread_hwp.HAction.Execute("TableCreate", hset)
                    # 셀 안에 각 줄을 수식 포함 처리 후 줄바꿈으로 구분
                    for j, row_text in enumerate(rows):
                        # 수식으로 시작하는 셀: HWP COM이 빈 셀에 첫 객체로 수식을 넣으면
                        # 앞 공백이 없어 커서 위치가 어긋나는 문제 방지
                        if row_text.startswith('=='):
                            row_text = ' ' + row_text
                        _insert_content_parts(
                            re.split(r'(==.*?==)', row_text, flags=re.DOTALL),
                            thread_hwp, use_com, WD, ID)
                        if j < len(rows) - 1:
                            thread_hwp.HAction.Run("BreakPara")  # 셀 내 줄바꿈
                    # 표 탈출: MoveTopLevelEnd로 커서를 표 바깥으로
                    thread_hwp.HAction.Run("MoveTopLevelEnd")
                except Exception as e:
                    _log_com_error("표 삽입", e)
                    for row_text in rows:
                        paste_text(row_text, delay=0.1 * ID)
                        send_hotkey('enter', delay=0.15 * ID)
            else:
                for row_text in rows:
                    paste_text(row_text, delay=0.1 * ID)
                    send_hotkey('enter', delay=0.15 * ID)
            continue

        # ── 그래프 삽입 ───────────────────────────────────────
        if top_part.startswith('<그래프>') and top_part.endswith('</그래프>'):
            png_path = top_part[5:-6].strip()
            if png_path and os.path.isfile(png_path):
                try:
                    pil_img = Image.open(png_path)
                    # 흰 배경에 합성 (색반전·투명도 문제 원천 차단)
                    bg = Image.new('RGB', pil_img.size, (255, 255, 255))
                    if pil_img.mode == 'RGBA':
                        bg.paste(pil_img, mask=pil_img.split()[3])
                    else:
                        bg.paste(pil_img.convert('RGB'))
                    # 200px 제한 → 96dpi 기준 약 53mm
                    bg.thumbnail((600, 600), Image.Resampling.LANCZOS)
                    # CF_DIB 형식으로 클립보드에 복사 (InsertPicture 대체)
                    buf = BytesIO()
                    bg.save(buf, 'BMP')
                    dib = buf.getvalue()[14:]  # BMP 파일헤더 14바이트 제거
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib)
                    win32clipboard.CloseClipboard()
                    # HWP에 붙여넣기
                    if use_com:
                        thread_hwp.HAction.Run("Paste")
                    else:
                        time.sleep(0.1)
                        pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.2)
                    root.after(0, lambda: status_label.config(
                        text="✅ 그래프 삽입 완료", fg="#008000"))
                except Exception as e:
                    em = str(e)
                    root.after(0, lambda em=em: messagebox.showerror(
                        "그래프 삽입 실패", f"오류:\n{em}"))
                finally:
                    try:
                        os.unlink(png_path)
                    except Exception:
                        pass
            else:
                root.after(0, lambda: messagebox.showwarning(
                    "그래프 없음", "분석 단계에서 그래프 생성에 실패했습니다."))
            continue

        # ── 기존: 미주·수식·일반 텍스트 처리 ─────────────────
        parts = re.split(r'(<미주>|</미주>|==.*?==)', top_part, flags=re.DOTALL)
        _insert_content_parts(parts, thread_hwp, use_com, WD, ID)


def process_text(raw_text):
    global stop_flag
    WD = window_delay_var.get()
    ID = input_delay_var.get()

    # ── COM 스레드 초기화 ─────────────────────────────────────
    thread_hwp = None
    use_com = False

    if _WIN32_AVAILABLE and _com_alive and _hwp_marshal_stream is not None:
        try:
            pythoncom.CoInitialize()
            thread_hwp = _w32.Dispatch(
                pythoncom.CoGetInterfaceAndReleaseStream(
                    _hwp_marshal_stream, pythoncom.IID_IDispatch))
            use_com = True
        except Exception as e:
            root.after(0, lambda err=str(e): status_label.config(
                text=f"⚠️ 마샬링 실패 → pyautogui 모드: {err[:60]}",
                fg="#FF8800"))

    mode_str = "⚡ COM 풀모드" if use_com else "⌨ pyautogui 모드"
    root.after(0, lambda m=mode_str: status_label.config(
        text=f"⏳ 입력 중… [{m}]", fg="#CC4400"))

    # ── 문항별 분리 → 각 문항 입력 후 페이지 나누기 ──────────
    prob_blocks = _split_hwp_problems(raw_text)
    if not prob_blocks:
        prob_blocks = [raw_text]
    n = len(prob_blocks)

    for i, block in enumerate(prob_blocks):
        if stop_flag: break
        if n > 1:
            root.after(0, lambda x=i+1: status_label.config(
                text=f"⏳ 문항 {x}/{n} 입력 중…", fg="#CC4400"))
        _process_one_block(block, thread_hwp, use_com, WD, ID)
        # 마지막 문항 뒤에는 페이지 나누기 없음
        if not stop_flag and i < n - 1:
            _insert_page_break(thread_hwp, use_com, WD)

    if _WIN32_AVAILABLE:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    root.after(0, finalize_ui, not stop_flag)


def _insert_content_parts(parts, thread_hwp, use_com, WD, ID):
    """미주·수식·일반 텍스트 파트 리스트를 HWP에 순서대로 입력."""
    global stop_flag
    for part in parts:
        if stop_flag: break
        if not part: continue

        # ── 미주 열기 ──────────────────────────────────────
        if part == '<미주>':
            if use_com:
                try:
                    thread_hwp.HAction.Run("InsertEndnote")
                    continue
                except Exception as e:
                    _log_com_error("미주 열기", e)
            activate_hwp()
            send_hotkey('ctrl', 'n', 'e', delay=0.1)
            time.sleep(WD)

        # ── 미주 닫기 ──────────────────────────────────────
        elif part == '</미주>':
            if use_com:
                try:
                    thread_hwp.HAction.Run("MoveTopLevelEnd")
                    continue
                except Exception:
                    try:
                        thread_hwp.HAction.Run("CloseEx")
                        continue
                    except Exception as e:
                        _log_com_error("미주 닫기", e)
            activate_hwp()
            time.sleep(WD * 0.3)
            send_hotkey('shift', 'esc', delay=WD * 0.6)

        # ── 수식 ──────────────────────────────────────────
        elif part.startswith('==') and part.endswith('=='):
            eq_str = part[2:-2].strip()
            if use_com:
                try:
                    # 검증된 방식: SetItem("String") + SetItem("EqString")
                    # → 수식창 없이 백그라운드 삽입, 워커 스레드 안전
                    hset = thread_hwp.HParameterSet.HEqEdit.HSet
                    thread_hwp.HAction.GetDefault("EquationCreate", hset)
                    try:
                        hset.SetItem("String", eq_str)
                    except Exception:
                        pass
                    try:
                        hset.SetItem("EqString", eq_str)
                    except Exception:
                        pass
                    thread_hwp.HAction.Execute("EquationCreate", hset)
                    continue
                except Exception as e:
                    _log_com_error("수식 삽입", e)
            # 폴백: pyautogui
            activate_hwp()
            send_hotkey('ctrl', 'n', 'm', delay=0.1)
            time.sleep(WD)
            paste_text(eq_str, delay=0.2 * ID)
            _close_formula(WD * 0.8)

        # ── 일반 텍스트 ────────────────────────────────────
        else:
            lines = part.split('\n')
            for i, line in enumerate(lines):
                if stop_flag: break
                if line:
                    if use_com:
                        try:
                            thread_hwp.HAction.GetDefault(
                                "InsertText",
                                thread_hwp.HParameterSet.HInsertText.HSet)
                            thread_hwp.HParameterSet.HInsertText.Text = line
                            thread_hwp.HAction.Execute(
                                "InsertText",
                                thread_hwp.HParameterSet.HInsertText.HSet)
                        except Exception as e:
                            _log_com_error("텍스트 입력", e)
                            activate_hwp()
                            paste_text(line, delay=0.1 * ID)
                    else:
                        paste_text(line, delay=0.1 * ID)
                if i < len(lines) - 1:
                    if use_com:
                        try:
                            thread_hwp.HAction.Run("BreakPara")
                        except Exception as e:
                            _log_com_error("엔터", e)
                            activate_hwp()
                            send_hotkey('enter', delay=0.15 * ID)
                    else:
                        send_hotkey('enter', delay=0.15 * ID)


def _log_com_error(action: str, e: Exception):
    msg = f"⚠️ COM [{action}] 실패 → pyautogui 폴백: {type(e).__name__}: {str(e)[:100]}"
    root.after(0, lambda m=msg: status_label.config(text=m, fg="#FF8800"))


def activate_hwp():
    hwp_windows = [w for w in gw.getWindowsWithTitle('한글') if w.visible]
    if not hwp_windows: return False
    try:
        w = hwp_windows[0]
        w.activate()
        if w.isMinimized: w.restore()
        return True
    except Exception: return False


def run_automation():
    global stop_flag
    stop_flag = False
    passed, errors, _ = validate_format(silent=True)
    if not passed:
        if not messagebox.askyesno("형식 오류 감지",
            f"오류 {len(errors)}개가 발견됐습니다.\n\n" +
            "\n".join(f"❌ {e}" for e in errors) +
            "\n\n그래도 입력을 진행할까요?"): return

    # HWP 인스턴스 확인
    if _WIN32_AVAILABLE and not _com_alive:
        # 자동으로 새 인스턴스 시작
        ok, msg = com_start_hwp()
        if not ok:
            if not messagebox.askyesno("HWP 시작 실패",
                f"{msg}\n\npyautogui 모드로 진행할까요? "
                "(기존에 떠있는 HWP 창에 입력)"):
                return
            if not activate_hwp():
                messagebox.showerror("오류", "한글 창을 찾을 수 없습니다!"); return

    # 마샬 스트림 재생성 (스레드 시작 직전)
    if _WIN32_AVAILABLE and _com_alive and _hwp_com is not None:
        global _hwp_marshal_stream
        try:
            _hwp_marshal_stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, _hwp_com)
        except Exception:
            _hwp_marshal_stream = None

    time.sleep(0.3)
    raw_text = input_area.get("1.0", tk.END).strip()
    if not raw_text: return
    threading.Thread(target=process_text, args=(raw_text,), daemon=True).start()

def stop_automation():
    global stop_flag
    stop_flag = True
    status_label.config(text="🛑 작업 중단됨", fg="#FF0000")

def finalize_ui(completed):
    status_label.config(text="✅ 대기 중", fg="#008000")
    if completed:
        messagebox.showinfo("완료", "입력이 완료되었습니다.")

# ============================================================
# 2. 프롬프트
# ============================================================
_SYSTEM_INSTRUCTION = (
    "너는 HWP 수식 데이터 변환기다. 아래 규칙을 절대 어기지 마라.\n\n"

    "▶ 출력 원칙\n"
    "  • 모든 출력은 한국어만. 영어 절대 금지.\n"
    "  • 풀이·사고 과정·중간 계산·재시도 과정 — 절대 출력 금지\n"
    "  • 인사말·부연 설명·마크다운 헤더(#)/불릿(-·*)/코드블록(```) — 절대 출력 금지\n"
    "  • 계산 오류·답 불확실 시: 재시도·재풀이 절대 금지. 최선 결과 1회만 출력.\n"
    "  • 문제 조건·숫자·구간 수정 및 재설계 — 절대 금지. 주어진 문제 원본 그대로.\n"
    "  • '수정:', '다시 시도:', '재설계:', '최종:' 등의 표현 — 절대 출력 금지.\n"
    "  • </미주> 이후 문항 본문 반드시 출력. 본문 없이 끝내기 금지.\n\n"

    "▶ 출력 형식 (한 글자도 변형 금지)\n"
    "문항번호<미주>[정답] 정답값\n"
    "[해설] 핵심수식 및 아이디어만 $해설수식$</미주>본문내용 및 $본문수식$\n\n"

    "  [정답 표기 규칙 — 유형별 다름]\n"
    "    객관식: [정답] ③       ← 선지 원문자를 평문으로. 수식 절대 금지.\n"
    "      ✗ 금지) [정답] $5$  /  [5]  /  [정답] 5\n"
    "      ✓ 유일 형식) [정답] ③\n"
    "    주관식: [정답] $숫자$  ← 반드시 수식으로.\n"
    "      ✓ 예) [정답] $72$\n\n"

    "  [해설 표기 규칙]\n"
    "    수학 강사용. 핵심 계산 흐름과 수식만. 서술어·설명 금지.\n"
    "    ✗ 금지) '식을 정리하면 다음과 같습니다'\n"
    "    ✓ 예) $f'(x) = 3x^{2} - 6x$ → $x = 0, 2$ 대입 → $\\frac{1}{3}$\n\n"

    "  • 문항 여럿이면 블록 반복. 2단 문서: 왼쪽 위→아래 후 오른쪽.\n"
    "  • 이미지에 번호 있으면 그대로, 없으면 01부터.\n\n"

    "▶ 수식 규칙\n"
    "  [표기 방식] 모든 수식은 $...$ 로 감싸고, 안은 표준 LaTeX 문법으로 작성.\n"
    "  [숫자·변수] 문항 번호 제외, 모든 숫자·알파벳은 $...$ 안에.\n"
    "  [객관식 보기] 원문자 ①②③④⑤. (1)~(5) 절대 금지.\n"
    "  [분수] \\frac{분자}{분모} 사용. 슬래시(/) 절대 금지.\n"
    "  [로마체] 기하 점·선분·도형이름·조합 C·순열 P·단위 → \\mathrm{}.\n"
    "    ✓) $\\mathrm{A}$  /  $\\mathrm{AB}$  /  $\\triangle\\mathrm{ABC}$\n"
    "    ✓) $\\binom{5}{2}$  /  $\\binom{5}{2n}$\n"
    "  [기하 라벨 규칙] 문장 안에 점·선분·도형 이름이 나오면 반드시 $...$ 안에.\n"
    "    ✓) '점 $\\mathrm{A}$에서' / '선분 $\\mathrm{AB}$의 길이' / '삼각형 $\\triangle\\mathrm{ABC}$'\n"
    "    ✗ 금지) '점 A에서'($ 밖) / '선분 AB의'($ 밖)\n"
    "  [단위] cm·m·km·mm·s·kg·L 등 단위는 숫자와 같은 $...$ 안에 \\mathrm{}로.\n"
    "    ✓) $5\\,\\mathrm{cm}$  /  $3\\,\\mathrm{m}$\n"
    "    ✗ 금지) $5$ cm  /  $5$ 'cm'\n"
    "  [극한] $\\lim_{조건} 식$\n"
    "  [부등호] \\geq, \\leq 사용.\n"
    "  [곱하기] \\cdot 사용. 별표(*) 절대 금지.\n"
    "  [지수 범위] ^ 뒤 항 반드시 {} 로 묶기.\n"
    "    ✗) $2^x(x-3)$     ✓) $2^{x}(x-3)$\n"
    "  [조건 분기] \\begin{cases}...\\end{cases} 사용.\n"
    "    ✓) $f(x) = \\begin{cases} 3x-1 & (x \\geq 2) \\\\ x+3 & (x < 2) \\end{cases}$\n\n"

    "▶ 표 (박스·테두리 조건)\n"
    "  이미지에 박스나 테두리로 감싸진 조건이 있으면 반드시 <표>...</표> 태그로 감싸라.\n"
    "  각 줄은 줄바꿈으로 구분. 수식($...$)도 허용.\n"
    "  ✓ 예)\n"
    "  <표>\n"
    "  조건 (가): $x \\geq 0$\n"
    "  조건 (나): $y \\leq 3$\n"
    "  </표>\n\n"

    "▶ 그래프·도형 마커\n"
    "  이미지에 좌표계 그래프, 함수 곡선, 도형이 있으면\n"
    "  해당 위치에 반드시 <그래프/> 마커를 삽입해라.\n"
    "  그래프/도형이 있는데 마커를 빠뜨리는 것은 절대 금지.\n\n"

    "▶ 출력 예시\n"
    "04<미주>[정답] ③\n"
    "[해설] $a+2b=10$ 대입 → $\\binom{5}{2n}$ 적용\n"
    "$\\lim_{x \\to 3} f(x) = \\frac{2}{3}$</미주>"
    "$a+2b=10$일 때 값을 구하시오.\n"
    "① $1$  ② $2$  ③ $\\frac{1}{3}$  ④ $4$  ⑤ $5$\n"
)

# 해설없음 전용 system instruction — <미주> 형식 강제 없음
_SYSTEM_INSTRUCTION_NO_ANS = (
    "너는 HWP 수식 데이터 변환기다. 아래 규칙을 절대 어기지 마라.\n\n"

    "▶ 출력 원칙\n"
    "  • 모든 출력은 한국어만. 영어 절대 금지.\n"
    "  • 풀이·사고 과정·중간 계산·재시도 과정 — 절대 출력 금지\n"
    "  • 인사말·부연 설명·마크다운 헤더(#)/불릿(-·*)/코드블록(```) — 절대 금지\n"
    "  • [정답], [해설] 태그 및 내용 있는 <미주>...</미주> — 절대 출력 금지\n"
    "  • <미주></미주> 빈 태그는 반드시 출력 (아래 형식 참조)\n"
    "  • 계산 오류·답 불확실 시: 재시도·재풀이 절대 금지. 최선 결과 1회만 출력.\n"
    "  • 문제 조건·숫자·구간 수정 및 재설계 — 절대 금지. 주어진 문제 원본 그대로.\n\n"

    "▶ 출력 형식\n"
    "문항번호<미주></미주>\n본문내용\n"
    "선지가 있으면 본문 아래에 그대로.\n"
    "  • 문항 번호 바로 뒤 <미주></미주> 삽입, 줄바꿈 후 본문 시작. 공백 삽입 금지.\n"
    "  • 문항 사이: 빈 줄 하나.\n"
    "  • 2단 문서: 왼쪽 위→아래 후 오른쪽.\n"
    "  • 이미지에 번호 있으면 그대로, 없으면 01부터.\n\n"

    "▶ 수식 규칙\n"
    "  [표기 방식] 모든 수식은 $...$ 로 감싸고, 안은 표준 LaTeX 문법으로 작성.\n"
    "  [숫자·변수] 문항 번호 제외, 모든 숫자·알파벳은 $...$ 안에.\n"
    "  [객관식 보기] 원문자 ①②③④⑤. (1)~(5) 절대 금지.\n"
    "  [분수] \\frac{분자}{분모} 사용. 슬래시(/) 절대 금지.\n"
    "  [로마체] 기하 점·선분·도형이름·조합 C·순열 P·단위 → \\mathrm{}.\n"
    "    ✓) $\\mathrm{A}$  /  $\\mathrm{AB}$  /  $\\triangle\\mathrm{ABC}$\n"
    "  [기하 라벨 규칙] 문장 안에 점·선분·도형 이름이 나오면 반드시 $...$ 안에.\n"
    "    ✓) '점 $\\mathrm{A}$에서' / '선분 $\\mathrm{AB}$의 길이'\n"
    "    ✗ 금지) '점 A에서'($ 밖)\n"
    "  [단위] cm·m 등 단위는 숫자와 같은 $...$ 안에 \\mathrm{}로.\n"
    "  [극한] $\\lim_{조건} 식$\n"
    "  [부등호] \\geq, \\leq 사용.\n"
    "  [곱하기] \\cdot 사용. 별표(*) 절대 금지.\n"
    "  [지수 범위] ^ 뒤 항 반드시 {} 로 묶기.\n"
    "  [조건 분기] \\begin{cases}...\\end{cases} 사용.\n\n"

    "▶ 표 (박스·테두리 조건)\n"
    "  이미지에 박스나 테두리로 감싸진 조건이 있으면 반드시 <표>...</표> 태그로 감싸라.\n\n"

    "▶ 그래프·도형 마커\n"
    "  이미지에 좌표계 그래프, 함수 곡선, 도형이 있으면\n"
    "  해당 위치에 반드시 <그래프/> 마커를 삽입해라.\n"
    "  그래프/도형이 있는데 마커를 빠뜨리는 것은 절대 금지.\n\n"

    "▶ 출력 예시\n"
    "04<미주></미주>\n$a+2b=10$일 때 값을 구하시오.\n"
    "① $1$  ② $2$  ③ $\\frac{1}{3}$  ④ $4$  ⑤ $5$\n"
)

_REVIEW_SYSTEM = (
    "너는 수학 문제 풀이 전문가다. 주어진 문제를 풀고 정답만 출력해라.\n"
    "객관식: ①②③④⑤ 중 하나만. 주관식: 숫자만.\n"
    "풀이·설명·이유 절대 금지. 정답 한 글자 또는 숫자만."
)

# ── 검토: 텍스트 결과물 검토용 ──────────────────────────────
_REVIEW_TEXT_SYSTEM = (
    "너는 수학 문제 품질 검토 전문가다. "
    "아래에 주어지는 텍스트는 HWP 수식 형식으로 작성된 수학 문항이다.\n\n"

    "다음 4가지를 독립적으로 판단해라:\n"
    "  ① 정답 정확성 — 제시된 [정답]이 실제로 맞는지 직접 풀어 확인\n"
    "  ② 해설 오류   — [해설]의 수식·논리 흐름에 오류가 있는지\n"
    "  ③ 난이도 적절성 — 의도치 않게 너무 쉽게 풀리는 경로가 있는지\n"
    "     (예: 보기 소거, 특수값 대입으로 정답이 바로 나오는 경우)\n"
    "  ④ 문항 완결성 — 조건 누락, 모순, 불필요 정보 등 논리적 결함\n\n"

    "출력 형식:\n"
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
    "[문항 번호]\n"
    "① 정답: ✅ 정확 / ❌ 오류 → (오류 시 실제 정답과 이유)\n"
    "② 해설: ✅ 정확 / ❌ 오류 → (오류 내용)\n"
    "③ 난이도: ✅ 적절 / ⚠️ 주의 → (쉽게 풀리는 경로 설명)\n"
    "④ 완결성: ✅ 문제없음 / ⚠️ 주의 → (결함 내용)\n"
    "종합: ✅ 통과 / ❌ 수정 필요\n\n"

    "규칙:\n"
    "  • 수식은 == == 안의 HWP 문법 그대로 읽고 해석.\n"
    "  • 모든 출력은 한국어.\n"
    "  • 칭찬·인사·부연 설명 절대 금지. 판정 결과만 출력.\n"
    "  • 문제없는 항목도 반드시 ✅ 표시 후 짧게 확인 결과 명시.\n"
)

# ── 검토: 이미지 직접 검토용 ────────────────────────────────
_REVIEW_IMAGE_SYSTEM = (
    "너는 수학 문제 품질 검토 전문가다. "
    "이미지 속 수학 문항을 직접 읽고 아래 4가지를 판단해라.\n\n"

    "  ① 문항 완결성 — 조건 누락, 모순, 불필요 정보 없는지\n"
    "  ② 정답 존재성 — 직접 풀어서 유일한 정답이 존재하는지 확인. 정답값 명시.\n"
    "  ③ 난이도 적절성 — 의도치 않게 쉽게 풀리는 경로 있는지\n"
    "     (예: 보기 소거, 특수값 대입, 단순 암산 등)\n"
    "  ④ 워싱 적합성 — 숫자 변경 시 풀이 구조가 유지될 수 있는지.\n"
    "     쉽게 풀리게 만드는 구조적 취약점이 있으면 명시.\n\n"

    "출력 형식:\n"
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
    "[문항 번호]\n"
    "① 완결성: ✅ 이상없음 / ⚠️ 주의 → ...\n"
    "② 정답: ✅ 유일 (정답값) / ❌ 오류 → ...\n"
    "③ 난이도: ✅ 적절 / ⚠️ 주의 → ...\n"
    "④ 워싱 적합: ✅ 적합 / ⚠️ 주의 → ...\n"
    "종합: ✅ 통과 / ❌ 수정 필요\n\n"

    "규칙: 한국어만. 칭찬·인사·부연 설명 절대 금지. 판정만 출력.\n"
    "이미지에 문항이 여럿이면 모두 검토.\n"
)


def _build_sys_instr(style: str, number_style: str) -> str:
    """style·number_style 조합에 맞는 시스템 지시문 반환."""
    base = _SYSTEM_INSTRUCTION_NO_ANS if style == "해설없음" else _SYSTEM_INSTRUCTION
    if number_style not in ("순서", "없음"):
        return base
    if style == "해설없음":
        base = (base
            .replace("문항번호<미주></미주>\n본문내용\n",
                     "<미주></미주>\n본문내용\n")
            .replace("  • 문항 번호 바로 뒤 <미주></미주> 삽입, 줄바꿈 후 본문 시작. 공백 삽입 금지.\n",
                     "  • 문항번호(숫자+마침표 등) 절대 출력 금지. <미주></미주>만 출력 후 줄바꿈, 본문 시작.\n")
            .replace("  • 이미지에 번호 있으면 그대로, 없으면 01부터.\n\n",
                     "  • 번호 절대 출력 금지.\n\n")
            .replace("04<미주></미주>\n$a+2b=10$", "<미주></미주>\n$a+2b=10$")
        )
    else:
        base = (base
            .replace("문항번호<미주>", "<미주>")
            .replace("  • 이미지에 번호 있으면 그대로, 없으면 01부터.\n\n",
                     "  • 문항번호(숫자+마침표 등) 절대 출력 금지. <미주> 바로 출력.\n\n")
            .replace("04<미주>", "<미주>")
        )
    return base


def get_prompt(task_mode, q_type, target_q=None, extra_note=None, style=None):
    if task_mode == "타이핑":
        answer_rule = "정답은 원본과 동일한 형식으로."
        format_rule = "원본 문항 형태(객관식/주관식) 그대로 유지."
    elif q_type == "주관식":
        answer_rule = "정답은 1~999 자연수. [정답] == 숫자 == 형식으로 출력."
        format_rule = "주관식으로 출력. 객관식 보기(①~⑤) 생성 절대 금지."
    else:
        answer_rule = "정답은 반드시 선지 원문자(①~⑤)로. [정답] ③ 형식. 수식 금지."
        format_rule = "객관식으로 출력. 본문 마지막에 ①~⑤ 보기 5개 필수."

    if task_mode == "타이핑":
        task_line = (f"작업: 이미지 문제를 100% 그대로 타이핑 변환. "
                     f"[정답]·[해설]은 직접 풀어 채워라. {answer_rule} {format_rule}\n"
                     "⚠️ 타이핑 모드 주의: 이 문항은 원본을 그대로 옮기는 것이므로 "
                     "정답이 분수·무리수·복잡한 식 등 깔끔하지 않은 값일 수 있다. "
                     "계산 결과가 복잡해 보여도 재풀이·재시도 절대 금지 — 나온 값 그대로 출력.")
        solve_line = "내부처리(출력금지): ①풀기 → ②역대입 검산 → ③완료 후 출력."
    else:
        task_line = (
            f"작업: 핵심 개념·구조 유지, 숫자를 바꿔 유사 문항 생성. {answer_rule} {format_rule}\n"
            "워싱 순서(내부처리, 출력금지):\n"
            "  ①원문 완전 풀기(풀이 구조·답의 형태 파악)\n"
            "  ②역산 설계 — 답이 정수·깔끔한 분수가 될 조건을 먼저 파악하고, 그 조건에서 숫자 유도\n"
            "     (숫자를 먼저 바꾸고 답 확인하는 방식 절대 금지)\n"
            "  ③새 숫자로 처음부터 풀기(기존 정답 재활용 금지) → ④역대입 검산\n"
            "  ⑤검산 후에도 답이 무리수·복잡한 식이면 → 구하는 것을 변경\n"
            "     (예: 답이 ∛2이면 'a³ 구하시오'로, 답이 1+√2이면 '(a-1)² 구하시오'로)\n"
            "  ⑥완료 후 출력 — 재시도·재설계 절대 금지, 최선 결과 1회만 출력."
        )
        solve_line = ""

    base = task_line
    if solve_line:
        base += f"\n{solve_line}"

    if target_q:
        base += (f"\n\n[집중 지시] 문항 번호 '{target_q}'에 해당하는 문제 하나만 처리. "
                 "다른 문항 출력 절대 금지.")

    # ── 해설 스타일 오버라이드
    if style == "학생용":
        base += (
            "\n\n[해설 스타일 오버라이드 — 반드시 적용]\n"
            "해설을 학생용으로 작성. 각 핵심 단계에 '~이므로', '~따라서', '~을 이용하면' 등 연결어 사용.\n"
            "해설 길이: 핵심 단계 3~5줄 이내. 부연·반복·재설명 절대 금지.\n"
            "오류 발생 시 재풀이 없이 최선 결과 그대로 출력."
        )
    elif style == "해설없음":
        base += (
            "\n\n[해설 없음 모드 — 반드시 적용]\n"
            "<미주>, [정답], [해설] 태그를 절대 출력하지 말 것.\n"
            "문항 번호 + 문항 본문 + 선지(있으면)만 출력.\n"
            "문항 사이는 빈 줄 하나로 구분. 미주·정답·해설 일체 금지."
        )

    # ── 특이사항 (기존 프롬프트 불변, 맨 끝에 추가)
    if extra_note and extra_note.strip():
        base += f"\n\n[이번 문항 특이사항 — 반드시 반영]\n{extra_note.strip()}"

    return base


# ============================================================
# 3. AI 호출 공통 함수
# ============================================================
def _strip_thinking_leak(text: str) -> str:
    text = re.sub(r'```[\s\S]*?```', '', text)
    text = text.replace("`markdown", "").replace("`html", "").strip()
    # [N점] 보호: $[3점]$ 처럼 AI가 수식으로 감싼 경우 $...$→==...== 변환 전에 처리
    text = re.sub(r'\$+\[(\d+)점\]\$+', r'[==\1==점]', text)
    # AI 수식 표기 통일 → ==...== (Python이 일괄 처리)
    text = re.sub(r'\$\$(.+?)\$\$', r'==\1==', text, flags=re.DOTALL)       # $$...$$
    text = re.sub(r'\$([^$\n]+)\$', r'==\1==', text)                         # $...$
    text = re.sub(r'\\\[(.+?)\\\]', r'==\1==', text, flags=re.DOTALL)        # \[...\]
    text = re.sub(r'\\\((.+?)\\\)', r'==\1==', text)                         # \(...\)
    text = re.sub(r'`(==.*?==)`', r'\1', text, flags=re.DOTALL)              # `==...==` 백틱 벗기기
    # 백틱: LaTeX 패턴 포함 시에만 수식으로 변환 (HWP 변환 전 raw 출력 한정)
    text = re.sub(r'`([^`\n]*(?:\\[a-zA-Z]|[\^_{}])[^`\n]*)`', r'==\1==', text)
    lines = text.split('\n')
    clean = []
    in_block = False
    after_close = False
    for line in lines:
        is_new_block = bool(re.match(r'^\s*[0-9０-９]{1,2}', line)) and '<미주>' in line
        if is_new_block:
            in_block = True
            after_close = False
            clean.append(line)
            if '</미주>' in line:
                after_close = True
            continue
        if not in_block:
            continue
        if '</미주>' in line:
            after_close = True
            clean.append(line)
            continue
        if after_close:
            stripped = line.strip()
            looks_like_body = (
                '==' in line or '$' in line or
                any(c in line for c in '①②③④⑤') or
                stripped == ''
            )
            if not looks_like_body and len(stripped) > 15:
                continue
            clean.append(line)
        else:
            clean.append(line)
    result = '\n'.join(clean).strip() if clean else text.strip()

    # ── </미주> 미닫힘 복구 ──────────────────────────────
    if '<미주>' in result and '</미주>' not in result:
        # AI가 재설계·재시도를 시작한 지점 감지 → 그 앞에서 자름
        redesign = re.search(r'\n' + _LOOP_PAT.pattern, result)
        if redesign:
            result = result[:redesign.start()].rstrip()
        result = result + '\n</미주>[⚠️ 미주 미닫힘 — 본문 없음]'

    return result


def _is_valid_result(text: str, style: str = None) -> bool:
    if not text: return False
    if style == "해설없음":
        return len(text.strip()) > 5
    return text.count('<미주>') > 0 and text.count('<미주>') == text.count('</미주>')

# ============================================================
# ★ 풀이저장 — 과목 인덱스 / DB 함수 / 손글씨→LaTeX / 뷰어 렌더러
# ============================================================
SUBJECT_INDEX = {
    "수능":       {"code": "S",  "units": {"수와 연산":"01","다항식":"02","방정식과 부등식":"03","도형의 방정식":"04","집합과 명제":"05","함수":"06","수열":"07","지수와 로그":"08","삼각함수":"09","미분":"10","적분":"11","확률과 통계":"12","수열의 극한":"13","공간도형과 벡터":"14"}},
    "고1":        {"code": "1",  "units": {"수와 연산":"01","다항식":"02","방정식과 부등식":"03","도형의 방정식":"04","집합과 명제":"05","함수":"06"}},
    "고2":        {"code": "2",  "units": {"지수와 로그":"01","삼각함수":"02","수열":"03","함수의 극한":"04","미분":"05","적분":"06"}},
    "고3/수학1":  {"code": "3A", "units": {"지수함수와 로그함수":"01","삼각함수":"02","수열":"03"}},
    "고3/수학2":  {"code": "3B", "units": {"함수의 극한과 연속":"01","미분":"02","적분":"03"}},
    "확률과 통계":{"code": "PS", "units": {"순열과 조합":"01","확률":"02","통계":"03"}},
    "미적분":     {"code": "CA", "units": {"수열의 극한":"01","미분법":"02","적분법":"03"}},
    "기하":       {"code": "GE", "units": {"이차곡선":"01","평면벡터":"02","공간도형과 벡터":"03"}},
}



def handwriting_to_latex(img: Image.Image) -> str:
    """손글씨 풀이 이미지 → LaTeX (Gemini Flash)."""
    prompt = ("이 손글씨 수학 풀이 이미지를 LaTeX로 변환해라.\n규칙:\n"
              "  • 수식은 $...$로 감싸라 (인라인 수식)\n"
              "  • 여러 줄 수식은 줄마다 따로 $...$\n"
              "  • 한국어 설명 텍스트는 그대로 한국어로\n"
              "  • 분수는 \\frac{분자}{분모}\n"
              "  • 기하 점·선분은 \\mathrm{A}, \\mathrm{AB}\n"
              "  • 설명·인사 없이 변환 결과만 출력\n"
              "  • 마크다운 코드블록(```) 절대 금지")
    resp = _gemini_client.models.generate_content(
        model=MODEL_FLASH,
        contents=[prompt, pil_to_part(img)],
        config=types.GenerateContentConfig(
            thinking_config=types.ThinkingConfig(thinking_budget=1024)))
    return resp.text.strip()


def _wrap_latex_line(line: str, max_chars: int) -> list:
    """긴 줄을 max_chars 기준으로 word-wrap. $...$ 수식은 분리하지 않음."""
    if len(line) <= max_chars:
        return [line]
    # $...$ 토큰과 일반 텍스트 분리
    tokens = re.split(r'(\$[^$]+\$)', line)
    words = []
    for tok in tokens:
        if tok.startswith('$') and tok.endswith('$'):
            words.append(tok)
        else:
            for w in tok.split(' '):
                if w:
                    words.append(w)
    result, current = [], ""
    for w in words:
        sep = " " if current else ""
        if len(current) + len(sep) + len(w) <= max_chars:
            current += sep + w
        elif not current:
            result.append(w)
        else:
            result.append(current)
            current = w
    if current:
        result.append(current)
    return result if result else [line]


def _render_latex_image(latex_text: str, width_px: int, dpi: int = 110) -> "Image.Image":
    """LaTeX 텍스트를 PIL Image로 렌더링 (matplotlib Agg, 스레드 안전)."""
    from PIL import Image as _PILImage
    raw_lines = latex_text.split('\n') if latex_text.strip() else ['(내용 없음)']
    # 가로 폭에 맞게 긴 줄 word-wrap
    chars_per_line = max(20, int(width_px / 14))
    lines = []
    for rl in raw_lines:
        lines.extend(_wrap_latex_line(rl, chars_per_line))
    n = len(lines)
    fig_w  = max(width_px / dpi, 4.0)
    line_h = 0.60
    fig_h  = max(2.5, n * line_h + 0.5)
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=dpi, facecolor='#FAFBFF')
    ax  = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.axis('off'); ax.set_facecolor('#FAFBFF')
    step    = 1.0 / (n + 0.4)
    y_start = 1.0 - step * 0.55
    for i, line in enumerate(lines):
        y = y_start - i * step
        if y < 0.01:
            break
        if not line.strip():
            continue
        try:
            ax.text(0.012, y, line, transform=ax.transAxes, fontsize=19,
                    verticalalignment='center', fontfamily='Malgun Gothic',
                    color='#1a1a2e', clip_on=True)
        except Exception:
            safe = re.sub(r'\$[^$]*\$', '[수식]', line)
            ax.text(0.012, y, safe, transform=ax.transAxes, fontsize=19,
                    verticalalignment='center', color='#AA4400', clip_on=True)
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=dpi, facecolor='#FAFBFF')
    plt.close(fig)
    buf.seek(0)
    return _PILImage.open(buf).copy()


# ── 풀이저장 창 연동: 마지막 결과 보관
_last_latex_result   = [""]   # 편집창에 자동 삽입될 마지막 결과 텍스트
_last_problem_id     = [""]   # 마지막으로 저장된 캐쉬 ID (문항별 중 첫 번째)
_last_problem_text   = [""]
_solution_viewer_ref = [None] # 현재 열린 Toplevel 참조


# ============================================================
# ★ AI 결과 캐쉬 (ai_output_cache.json)
# ============================================================

def _load_ai_cache() -> list:
    if not _AI_CACHE_DB.exists():
        return []
    try:
        return json.loads(_AI_CACHE_DB.read_text(encoding="utf-8"))
    except Exception:
        return []

def _save_ai_cache_db(db: list):
    _CACHE_DIR.mkdir(exist_ok=True)
    tmp = _AI_CACHE_DB.with_suffix('.tmp')
    tmp.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(_AI_CACHE_DB)


def _split_ai_results(result: str, style: str) -> list:
    """AI 결과를 문항별 블록으로 분리. 실패하면 전체를 1개로."""
    result = result.strip()
    if not result:
        return []

    if style == "해설없음":
        # 빈 줄 2개 이상으로 구분
        blocks = re.split(r'\n{2,}', result)
        out = [b.strip() for b in blocks if b.strip()]
        return out if out else [result]

    if '<미주>' not in result:
        return [result]

    # "숫자<미주>" 또는 "<미주>" 패턴을 새 블록 시작점으로 분리
    blocks = re.split(r'(?m)(?=(?:^\d{1,2}\s*<미주>|^<미주>))', result)
    valid  = [b.strip() for b in blocks if b.strip() and '<미주>' in b]
    return valid if valid else [result]

def _extract_answer_from_block(block: str) -> str:
    """문항 블록에서 정답 추출. ==...== 형식도 처리."""
    m = re.search(r'<미주>(.*?)</미주>', block, re.DOTALL)
    src = m.group(1) if m else block
    # 객관식 원문자
    a = re.search(r'\[정답\]\s*([①②③④⑤])', src)
    if a: return a.group(1)
    # 주관식 $숫자$ 또는 ==숫자==
    a = re.search(r'\[정답\]\s*(?:\$\s*|==\s*)(\d+(?:\.\d+)?)(?:\s*\$|\s*==)', src)
    if a: return a.group(1)
    # 숫자만
    a = re.search(r'\[정답\]\s*(\d+)', src)
    if a: return a.group(1)
    return ""

def _extract_ai_output_body(block: str, style: str) -> str:
    """문항 블록에서 본문+선지만 추출 (문항번호·미주 제거)."""
    if style == "해설없음":
        # strip leading "번호<미주></미주>\n" or "<미주></미주>\n"
        body = re.sub(r'^\d{0,2}\s*<미주></미주>\s*\n?', '', block.strip())
        return body.strip()
    # 맨 앞 "숫자<미주>...</미주>" 제거
    body = re.sub(r'^\d{0,2}\s*<미주>.*?</미주>', '', block,
                  count=1, flags=re.DOTALL)
    return body.strip()

def save_ai_output_cache(result: str, style: str = "강사용", raw_text: str = None) -> list:
    """AI 결과를 문항별로 분리해 캐쉬에 저장. 생성된 ID 리스트 반환.
    raw_text: Gemini response.text ($...$ 원본). 있으면 이걸 기준으로 분리·저장.
    없으면(캐시 히트) processed result 기준.
    """
    # 원본 있으면 원본 기준 분리, 없으면 processed fallback
    split_src = raw_text.strip() if raw_text else result
    blocks = _split_ai_results(split_src, style)
    if not blocks and raw_text:
        blocks = _split_ai_results(result, style)
    if not blocks:
        return []

    with _cache_write_lock:
        db    = _load_ai_cache()
        ids   = []
        today = datetime.date.today().strftime("%Y%m%d")
        base_seq = sum(1 for e in db if e.get("id","").startswith(today))
        existing_hashes = {e.get("hash") for e in db if e.get("hash")}

        for i, block in enumerate(blocks):
            import hashlib
            content_hash = hashlib.md5(block.strip().encode()).hexdigest()
            if content_hash in existing_hashes:
                continue  # 중복 스킵
            iid       = f"{today}-{base_seq + len(ids) + 1:03d}"
            ai_output = _extract_ai_output_body(block, style)
            answer    = _extract_answer_from_block(block)
            db.append({
                "id":         iid,
                "date":       today,
                "style":      style,
                "hash":       content_hash,
                "raw_output": block.strip(),
                "ai_output":  ai_output,
                "answer":     answer,
            })
            existing_hashes.add(content_hash)
            ids.append(iid)

        _save_ai_cache_db(db)
    return ids


def build_subject_codes_list(course_var, units_vars) -> list:
    """선택된 과목/중단원 → 코드 리스트. 예: ['S-10', 'S-11']"""
    c    = SUBJECT_INDEX.get(course_var.get(), {})
    code = c.get("code", course_var.get())
    sel  = [f"{code}-{c['units'][u]}" for u, v in units_vars.items()
            if v.get() and u in c.get("units", {})]
    return sel if sel else [code]


def save_finetune_new(item_id: str, subject_list: list, teacher_output: str):
    """파인튜닝 데이터 저장 (신규 형식).
    ai_output / answer 는 ai_output_cache 에서 자동 참조.
    """
    _CACHE_DIR.mkdir(exist_ok=True)
    # AI 캐쉬에서 본문·정답 가져오기
    cache_entry = next((e for e in _load_ai_cache() if e.get("id") == item_id), None)
    # 파인튜닝엔 미처리 원본(raw_output) 사용; 구버전 캐쉬엔 ai_output으로 fallback
    ai_output   = cache_entry.get("raw_output") or cache_entry.get("ai_output", "") if cache_entry else ""
    answer      = cache_entry.get("answer", "")    if cache_entry else ""

    ft_db = []
    if _FINETUNE_DB.exists():
        try:
            ft_db = json.loads(_FINETUNE_DB.read_text(encoding="utf-8"))
        except Exception:
            ft_db = []

    for item in ft_db:
        if item.get("id") == item_id:
            item.update({"subject": subject_list, "raw_output": ai_output,
                         "answer": answer, "teacher_output": teacher_output})
            _FINETUNE_DB.write_text(
                json.dumps(ft_db, ensure_ascii=False, indent=2), encoding="utf-8")
            return

    ft_db.append({"id": item_id, "subject": subject_list,
                  "raw_output": ai_output, "answer": answer,
                  "teacher_output": teacher_output})
    _FINETUNE_DB.write_text(
        json.dumps(ft_db, ensure_ascii=False, indent=2), encoding="utf-8")


# ============================================================
# 3-A. LaTeX → HWP 수식 변환 레이어
# ============================================================

def _find_closing_brace(s: str, open_pos: int) -> int:
    """열린 중괄호에 대응하는 닫힌 중괄호 위치 반환."""
    depth = 0
    for i in range(open_pos, len(s)):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return i
    return len(s) - 1


def _convert_frac(expr: str) -> str:
    """\\frac{A}{B} → {A} over {B} (중첩 분수 재귀 처리)."""
    while r'\frac' in expr:
        match = re.search(r'\\frac\{', expr)
        if not match:
            break
        start = match.start()
        num_open  = match.end() - 1
        num_close = _find_closing_brace(expr, num_open)
        try:
            den_open = expr.index('{', num_close + 1)
        except ValueError:
            break
        den_close = _find_closing_brace(expr, den_open)

        numerator   = _convert_frac(expr[num_open + 1 : num_close])
        denominator = _convert_frac(expr[den_open + 1 : den_close])
        replacement = f'{{{numerator}}} over {{{denominator}}}'
        expr = expr[:start] + replacement + expr[den_close + 1:]
    return expr


def _convert_sqrt_n(expr: str) -> str:
    """\\sqrt[n]{content} → root {n} of {content} (임의 깊이 중첩 지원)."""
    result = ''
    i = 0
    while i < len(expr):
        m = re.search(r'\\sqrt\[([^\]]+)\]\{', expr[i:])
        if not m:
            result += expr[i:]
            break
        result += expr[i : i + m.start()]
        n = m.group(1)
        brace_open = i + m.end() - 1
        brace_close = _find_closing_brace(expr, brace_open)
        inner = expr[brace_open + 1 : brace_close]
        result += f'root {{{n}}} of {{{inner}}}'
        i = brace_close + 1
    return result


def latex_to_hwp(expr: str) -> str:
    """LaTeX 수식 문자열을 HWP 수식 문법으로 변환."""
    expr = expr.strip()

    # 0. 전처리: 변형 frac 통일, 유니코드 그리스 문자 → LaTeX 명령어
    expr = expr.replace('\\dfrac', '\\frac').replace('\\cfrac', '\\frac')
    expr = expr.replace('\\displaystyle', '')
    _UNICODE_GREEK = {
        'α':'\\alpha','β':'\\beta','γ':'\\gamma','δ':'\\delta',
        'ε':'\\epsilon','ζ':'\\zeta','η':'\\eta','θ':'\\theta',
        'ι':'\\iota','κ':'\\kappa','λ':'\\lambda','μ':'\\mu',
        'ν':'\\nu','ξ':'\\xi','π':'\\pi','ρ':'\\rho',
        'σ':'\\sigma','τ':'\\tau','υ':'\\upsilon','φ':'\\phi',
        'χ':'\\chi','ψ':'\\psi','ω':'\\omega',
        'Γ':'\\Gamma','Δ':'\\Delta','Θ':'\\Theta','Λ':'\\Lambda',
        'Π':'\\Pi','Σ':'\\Sigma','Φ':'\\Phi','Ψ':'\\Psi','Ω':'\\Omega',
        '∞':'\\infty','∈':'\\in','∉':'\\notin','∪':'\\cup','∩':'\\cap',
        '≥':'\\geq','≤':'\\leq','≠':'\\neq','≈':'\\approx',
    }
    for uc, cmd in _UNICODE_GREEK.items():
        expr = expr.replace(uc, cmd)

    # 1. 분수 (재귀 처리)
    expr = _convert_frac(expr)

    # 1b. 거듭제곱근 (임의 깊이 중첩 지원)
    expr = _convert_sqrt_n(expr)

    # 2. 단순 1:1 치환 테이블 (긴 패턴 먼저)
    _SIMPLE = [
        # 폰트 — rm 뒤 it 삽입으로 이후 italic 복귀 보장
        (r'\\mathrm\{([^{}]+)\}',            r'rm \1 it'),
        (r'\\mathit\{([^{}]+)\}',            r'it \1'),
        (r'\\mathbf\{([^{}]+)\}',            r'bf \1'),
        (r'\\text\{([^{}]+)\}',              r'rm \1 it'),
        # 도형 기호 — HWP에 없는 사각형 계열은 box{}로 대체
        (r'\\triangle',                       'triangle '),   # 뒤 rm과 자연 연결되도록 공백 포함
        (r'\\square',                         'box{~~}'),
        (r'\\Box\b',                          'box{}'),
        (r'\\angle',                          'angle '),
        # 60분법 각도 — ^\circ / ^{\circ} → DEG
        (r'\^\{?\\circ\}?',                  ' DEG'),
        # 루트·벡터·데코레이션 (\\sqrt[n] 는 1b에서 이미 처리)
        (r'\\sqrt\{([^{}]+)\}',              r'sqrt {\1}'),
        (r'\\overrightarrow\{([^{}]+)\}',    r'overrightarrow {\1}'),
        (r'\\vec\{([^{}]+)\}',               r'vec {\1}'),
        (r'\\overline\{((?:[^{}]|\{[^{}]*\})+)\}', r'bar{\1}'),  # overline → bar{}
        (r'\\bar\{((?:[^{}]|\{[^{}]*\})+)\}',    r'bar{\1}'),  # \bar{} → bar{}
        (r'\\underline\{([^{}]+)\}',         r'underline {\1}'),
        (r'\\hat\{([^{}]+)\}',               r'hat {\1}'),
        (r'\\tilde\{([^{}]+)\}',             r'tilde {\1}'),
        # 화살표
        (r'\\rightarrow|\\to(?![a-zA-Z])',    '`rarrow`'),
        (r'\\leftarrow',                     '`larrow`'),
        (r'\\Rightarrow|\\implies',           '`Rarrow`'),
        (r'\\Leftarrow',                     'Larrow'),
        (r'\\leftrightarrow',                'harrow'),
        (r'\\Leftrightarrow|\\iff',          'Lrarrow'),
        # 관계 연산자
        (r'\\geq|\\ge(?![a-zA-Z])',            ' ge '),
        (r'\\leq|\\le(?![a-zA-Z])',            ' le '),
        (r'\\neq|\\ne\b',                    '`ne`'),
        (r'\\equiv',                         'equiv'),
        (r'\\approx',                        'approx'),
        (r'\\sim\b',                         'sim'),
        # 산술 연산자
        (r'\\cdot',                          'cdot'),
        (r'\\times',                         'times'),
        (r'\\div',                           'div'),
        (r'\\pm',                            '+-'),
        (r'\\mp',                            'mp'),
        (r'\\%',                             '%'),
        # 집합
        (r'\\in\b',                          'in'),
        (r'\\notin',                         'notin'),
        (r'\\subset',                        'subset'),
        (r'\\subseteq',                      'subseteq'),
        (r'\\supset',                        'supset'),
        (r'\\cup',                           'cup'),
        (r'\\cap',                           'cap'),
        (r'\\emptyset',                      'emptyset'),
        # 논리·기타
        (r'\\forall',                        'for all'),
        (r'\\exists',                        'exists'),
        (r'\\because',                       'because'),
        (r'\\therefore',                     'therefore'),
        (r'\\partial',                       'partial'),
        (r'\\nabla',                         'nabla'),
        (r'\\infty',                         'inf'),
        (r'\\ldots|\\cdots|\\dots',          '...'),
        # 그리스 문자
        (r'\\alpha',   'alpha'),  (r'\\beta',    'beta'),
        (r'\\gamma',   'gamma'),  (r'\\delta',   'delta'),
        (r'\\epsilon', 'epsilon'),(r'\\varepsilon','varepsilon'),
        (r'\\zeta',    'zeta'),   (r'\\eta',     'eta'),
        (r'\\theta',   'theta'),  (r'\\vartheta','vartheta'),
        (r'\\iota',    'iota'),   (r'\\kappa',   'kappa'),
        (r'\\lambda',  'lambda'), (r'\\mu',      'mu'),
        (r'\\nu',      'nu'),     (r'\\xi',      'xi'),
        (r'\\pi\b',    ' pi '),    (r'\\varpi',   'varpi'),
        (r'\\rho',     'rho'),    (r'\\sigma',   'sigma'),
        (r'\\tau',     'tau'),    (r'\\upsilon', 'upsilon'),
        (r'\\phi',     'phi'),    (r'\\varphi',  'varphi'),
        (r'\\chi',     'chi'),    (r'\\psi',     'psi'),
        (r'\\omega',   'omega'),
        (r'\\Gamma',   'GAMMA'),  (r'\\Delta',   'DELTA'),
        (r'\\Theta',   'THETA'),  (r'\\Lambda',  'LAMBDA'),
        (r'\\Pi\b',    'PI'),     (r'\\Sigma',   'SIGMA'),
        (r'\\Phi',     'PHI'),    (r'\\Psi',     'PSI'),
        (r'\\Omega',   'OMEGA'),
        # 함수
        (r'\\lim',  'lim'),  (r'\\sin',  'sin'),  (r'\\cos',  'cos'),
        (r'\\tan',  'tan'),  (r'\\cot',  'cot'),  (r'\\sec',  'sec'),
        (r'\\csc',  'csc'),  (r'\\log',  'log'),  (r'\\ln',   'ln'),
        (r'\\exp',  'exp'),  (r'\\max',  'max'),  (r'\\min',  'min'),
        (r'\\gcd',  'gcd'),
        # 적분·합
        (r'\\int',  'int'),  (r'\\iint', 'iint'), (r'\\iiint','iiint'),
        (r'\\oint', 'oint'), (r'\\sum',  'sum'),  (r'\\prod', 'prod'),
        # 집합 중괄호
        (r'\\\{',  ' left{'),
        (r'\\\}',  ' right}'),
        # \sqrt 중첩 중괄호 폴백 — 위 패턴이 실패한 경우 백슬래시만 제거
        (r'\\sqrt\b',                        'sqrt'),
    ]
    for pattern, replacement in _SIMPLE:
        expr = re.sub(pattern, replacement, expr)

    # 2b. 데코레이션 뒤 ^ 처리: bar{...}^n → {bar{...}}^n
    # (선분 제곱 표현 시 bar가 ^에 묶이지 않는 문제 방지)
    expr = re.sub(
        r'((?:bar|hat|tilde|vec)\{(?:[^{}]|\{[^{}]*\})*\})\^',
        r'{\1}^',
        expr
    )

    # 2c. 극한 방향 표기: 0^{+} / 0^{-} → 0+ / 0-
    expr = re.sub(r'([a-zA-Z0-9])\^\{([+-])\}', r'\1\2', expr)

    # 2d. sqrt/root 앞 붙은 문자·숫자 → 공백 삽입 (asqrt → a sqrt)
    expr = re.sub(r'([a-zA-Z0-9])(sqrt\b|root\b)', r'\1 \2', expr)

    # 3. 이항계수 → HWP 조합 표기
    expr = re.sub(
        r'\\binom\{([^{}]+)\}\{([^{}]+)\}',
        r'{}_{{\1}} rm C _{{\2}}', expr)

    # 4. cases 환경
    def _replace_cases(m):
        inner = m.group(1).strip()
        inner = re.sub(r'\\\\', ' # ', inner)  # 줄바꿈
        return f'cases{{{inner}}}'
    expr = re.sub(
        r'\\begin\{cases\}(.*?)\\end\{cases\}',
        _replace_cases, expr, flags=re.DOTALL)

    # 5. LaTeX 공백 기호 → HWP 백틱 공백
    expr = re.sub(r'\\,|\\;|\\!', '`', expr)
    expr = re.sub(r'\\quad|\\qquad', '`', expr)

    # 5b. 프라임(') → `' (함수 표기 간격 확보)
    expr = re.sub(r"'", "`'", expr)

    # 6. 절댓값·조건절 파이프 처리
    # ① \left| / \right| / \mid → 플레이스홀더 (toggle에서 다시 분리되지 않도록)
    _LP = '\x00LP\x00'
    _RP = '\x00RP\x00'
    expr = re.sub(r'\\left\s*\|',  _LP, expr)
    expr = re.sub(r'\\right\s*\|', _RP, expr)
    expr = re.sub(r'\\mid\b',      _LP, expr)
    # ② 나머지 \left / \right 제거
    expr = re.sub(r'\\left\s*|\\right\s*', '', expr)
    # ③ 남은 bare | → 순서대로 left| / right| 토글
    parts = expr.split('|')
    if len(parts) > 1:
        out = [parts[0]]
        for k, part in enumerate(parts[1:], 1):
            out.append(_LP if k % 2 == 1 else _RP)
            out.append(part)
        expr = ''.join(out)
    # ④ 플레이스홀더 → 실제 HWP 표기
    expr = expr.replace(_LP, '`left|`').replace(_RP, '`right|`')
    expr = re.sub(r'\$+', '', expr)

    # 7. rm/it 커맨드가 공백 없이 붙은 경우 보정 (예: itrm → it rm)
    expr = re.sub(r'\bit(?=[a-z])', 'it ', expr)

    # 8. 콤마 뒤 공백 → ,~ (HWP 수식 목록 공백) — 다음이 수식 문자일 때만
    expr = re.sub(r',\s+(?=[a-zA-Z0-9\\({+\-])', ',~', expr)

    # 9. ^ / _ 뒤 단일 영숫자(중괄호 없는 것만) → {} 감싸기
    expr = re.sub(r'\^([a-zA-Z0-9])', r'^{\1}', expr)
    expr = re.sub(r'_([a-zA-Z0-9])',  r'_{\1}', expr)

    # 10. 괄호·대괄호 → left/right (자동 크기 조절)
    #     {}: \{·\}는 step2에서 이미 left{·right}로 처리됨
    #         나머지 {}는 HWP 수식 구조 기호라 변환 금지
    expr = re.sub(r'\(', ' left(', expr)
    expr = re.sub(r'\)', ' right)', expr)
    expr = re.sub(r'\[', ' left[', expr)
    expr = re.sub(r'\]', ' right]', expr)
    expr = re.sub(r'  +', ' ', expr)  # 연속 공백 정리

    return expr.strip()


def _auto_wrap_non_korean(text: str) -> str:
    """
    한글/영문 구분 기반 수식 자동 감싸기.
    전제: 수식에서 한글 없음, 평서문에서 영문·숫자 없음.
    라틴 문자 또는 숫자를 포함하는 비한글 구간 → ==...==
    """
    protected = []

    def _protect(m):
        protected.append(m.group(0))
        # PUA(Private Use Area) 문자로 플레이스홀더 — 라틴/숫자 없어서 _wrap 안 걸림
        return chr(0xE000 + len(protected) - 1)

    for pat in [
        r'(?m)^\d{1,2}(?=[ \t]*<미주>)',    # 행 첫 문항 번호
        r'==.*?==',                           # 이미 감싸진 수식
        r'<그래프>.*?</그래프>',              # 그래프 마커
        r'</?(미주|표|그래프)>',              # 구조 태그
        r'\[정답\]|\[해설\]',                # 정답/해설 태그
        r'[①②③④⑤⑥⑦⑧⑨⑩]',              # 원문자 보기
    ]:
        text = re.sub(pat, _protect, text, flags=re.DOTALL)

    def _wrap(m):
        seg = m.group(0)
        if not re.search(r'[a-zA-Z0-9]', seg):
            return seg
        lpad = seg[:len(seg) - len(seg.lstrip())]
        rpad = seg[len(seg.rstrip()):]
        inner = seg.strip().strip('.,;:!?…')
        if not inner or not re.search(r'[a-zA-Z0-9]', inner):
            return seg
        return f'{lpad}=={inner}=={rpad}'

    # PUA 문자(-)는 플레이스홀더 — 매칭 제외
    text = re.sub(
        r'[^가-힣ㄱ-ㆎᄀ-ᇿ\n-]+',
        _wrap, text
    )

    for i, orig in enumerate(protected):
        text = text.replace(chr(0xE000 + i), orig)

    return text


def _convert_equations_in_result(text: str) -> str:
    """결과 텍스트의 ==...== 안 내용을 LaTeX → HWP 수식 문법으로 변환."""
    # [N점] 보호: ==\[N점]==로 변환된 경우(AI가 수식으로 감쌌을 때) 먼저 처리
    text = re.sub(r'==\[(\d+(?:\.\d+)?)\s*점\]==', r'[==\1==점]', text)
    # [N점] 표기 보호: 일반 텍스트인 경우 (정수/소수 모두)
    text = re.sub(r'\[(\d+(?:\.\d+)?)\s*점\]', r'[==\1==점]', text)
    # 미감싸진 수식 구간 자동 감싸기 (한글/비한글 경계 기반)
    text = _auto_wrap_non_korean(text)
    # 텍스트 괄호가 수식을 감싸는 경우 → 괄호를 수식 안으로 흡수
    text = re.sub(r'\(==(.*?)==\)', r'==(\1)==', text, flags=re.DOTALL)
    text = re.sub(r'\[==(.*?)==\]', r'==[\1]==', text, flags=re.DOTALL)

    def _convert_eq(m):
        latex = m.group(1)
        hwp   = latex_to_hwp(latex)
        return f'=={hwp}=='
    text = re.sub(r'==(.*?)==', _convert_eq, text, flags=re.DOTALL)
    # 수식 밖에 남아있는 LaTeX 잔재 제거
    text = re.sub(r'\\mathrm\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\text\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\mathit\{([^{}]+)\}', r'\1', text)
    return text


# ============================================================
# ★ 검수 함수
# ============================================================

def _extract_wash_answer(raw: str) -> str | None:
    """날것 AI 출력에서 [정답] 값 추출 ($...$, ①~⑤ 형태).
    </미주>가 없어도 [정답] 줄이 있으면 추출."""
    m = re.search(r'<미주>(.*?)</미주>', raw, re.DOTALL)
    search_in = m.group(1) if m else raw   # 미닫힘 시 전체에서 탐색
    # 객관식: [정답] ③
    a = re.search(r'\[정답\]\s*([①②③④⑤])', search_in)
    if a:
        return a.group(1)
    # 주관식: [정답] $72$ 또는 [정답] 72
    a = re.search(r'\[정답\]\s*\$\s*(\d+(?:\.\d+)?)\s*\$', search_in)
    if a:
        return a.group(1)
    a = re.search(r'\[정답\]\s*(\d+)', search_in)
    if a:
        return a.group(1)
    return None

def _extract_problem_for_review(raw: str) -> str:
    """<미주>...</미주> 제거 → 검수 AI에게 넘길 순수 문제 본문 ($수식$ 그대로)"""
    return re.sub(r'<미주>.*?</미주>', '', raw, flags=re.DOTALL).strip()

def _normalize_answer(ans: str) -> str:
    """답안 정규화: 원문자→숫자, 후위 제거, 숫자 파싱"""
    ans = ans.strip()
    for k, v in {'①':'1','②':'2','③':'3','④':'4','⑤':'5'}.items():
        ans = ans.replace(k, v)
    ans = re.sub(r'[번호\.\s]', '', ans)
    try:
        return str(int(float(ans)))
    except ValueError:
        return ans

def _answers_match(a: str, b: str) -> bool:
    return _normalize_answer(a) == _normalize_answer(b)

def _call_review(problem_text: str) -> str:
    """검수 단건 호출 — gemini-2.5-flash, 정답만 반환"""
    try:
        resp = _gemini_client.models.generate_content(
            model=MODEL_FLASH,
            contents=[problem_text],
            config=types.GenerateContentConfig(
                system_instruction=_REVIEW_SYSTEM,
                thinking_config=types.ThinkingConfig(thinking_budget=1024),
                max_output_tokens=16,
            )
        )
        return resp.text.strip()
    except Exception:
        return ""


def _call_gemini(img: Image.Image, task_mode: str, q_type: str,
                 target_q: str = None,
                 use_cache: bool = True, style: str = None,
                 extra_note: str = None,
                 graph_img: Image.Image = None,
                 number_style: str = "원본") -> tuple:
    """(processed_result, raw_text) 반환. 캐시 히트 시 raw_text=None."""
    # style 확정 (None이면 단일 패널 기본값)
    if style is None:
        style = style_var.get()

    # ── 캐시 히트 확인 (재시도 콜은 캐시 스킵)
    cache_key = _cache_key(img, task_mode, q_type, target_q or "", style=style)
    if use_cache:
        cached = _cache_get(cache_key)
        if cached is not None:
            return cached, None

    model_id, thinking_budget = _pick_model(task_mode, style)
    client = _gemini_client
    if extra_note is None:
        extra_raw = single_note_area.get("1.0", tk.END).strip()
        extra_note = extra_raw if extra_raw and extra_raw != _SINGLE_PLACEHOLDER else None
    user_prompt = get_prompt(task_mode, q_type, target_q,
                             extra_note=extra_note, style=style)

    sys_instr = _build_sys_instr(style, number_style)

    if graph_img is not None:
        # 타이핑 모드 전용 — 워싱 모드에선 graph_img가 넘어오지 않음
        user_prompt += (
            "\n\n[그래프 제공] 두 번째 이미지가 이 문항에 포함된 그래프/도형입니다. "
            "원본에서 그래프가 나타나는 위치에 반드시 <그래프/> 를 삽입하라. "
            "<그래프/>를 빠뜨리거나 위치를 틀리는 것은 절대 금지. "
            "bbox JSON·그래프 설명·분석 출력 절대 금지."
        )
        contents = [user_prompt, pil_to_part(img), pil_to_part(graph_img)]
    else:
        contents = [user_prompt, pil_to_part(img)]

    response = client.models.generate_content(
        model=model_id,
        contents=contents,
        config=types.GenerateContentConfig(
            system_instruction=sys_instr,
            thinking_config=types.ThinkingConfig(thinking_budget=thinking_budget),
        )
    )
    raw_text = response.text

    # ── 검수 (워싱 모드만, 날것 raw_text 기준)
    review_mismatch = False
    if task_mode == "워싱":
        wash_ans = _extract_wash_answer(raw_text)
        if wash_ans:
            problem_text = _extract_problem_for_review(raw_text)
            review_ans = _call_review(problem_text)
            if review_ans and not _answers_match(wash_ans, review_ans):
                # 1회 재시도
                retry_resp = client.models.generate_content(
                    model=model_id,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=sys_instr,
                        thinking_config=types.ThinkingConfig(thinking_budget=thinking_budget),
                    )
                )
                retry_raw = retry_resp.text
                retry_wash_ans = _extract_wash_answer(retry_raw)
                retry_review_ans = _call_review(_extract_problem_for_review(retry_raw)) if retry_wash_ans else None
                if retry_wash_ans and retry_review_ans and _answers_match(retry_wash_ans, retry_review_ans):
                    raw_text = retry_raw  # 재시도 성공 → 결과 교체
                else:
                    review_mismatch = True

    result = _strip_thinking_leak(raw_text)

    # <그래프> / <그래프/> 마커 처리
    if graph_img is not None:
        # G 크롭 이미지 — 처리 없이 그대로 저장 후 삽입
        path  = _save_graph_img_pil(graph_img)
        g_tag = f'<그래프>{path}</그래프>' if path else '[그래프/도형 삽입필요]'
        result = re.sub(r'<그래프>.*?</그래프>', g_tag, result, flags=re.DOTALL)
        result = result.replace('<그래프/>', g_tag)
        # AI가 마커를 빠뜨린 경우 결과 끝에 추가
        if path and '<그래프>' not in result:
            result = result.rstrip('\n') + f"\n{g_tag}"
    else:
        # G 크롭 없음 → 위치만 표시
        result = re.sub(r'<그래프>.*?</그래프>', '[그래프/도형 삽입필요]', result, flags=re.DOTALL)
        result = result.replace('<그래프/>', '[그래프/도형 삽입필요]')

    result = _convert_equations_in_result(result)   # ← LaTeX → HWP 변환

    # ── 유효한 결과만 캐시에 저장
    if use_cache and _is_valid_result(result, style=style):
        _cache_put(cache_key, result)

    if review_mismatch:
        return "[검수 불일치]", raw_text

    return result, raw_text




def _save_graph_img_pil(img: Image.Image) -> str | None:
    """수동 크롭된 그래프: 배경 순백 + 원본 품질 유지 PNG 저장."""
    try:
        if img.mode == 'RGBA':
            bg = Image.new('RGB', img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[3])
            img = bg
        else:
            img = img.convert('RGB')

        result = ImageEnhance.Contrast(img).enhance(1.2)
        result = result.filter(ImageFilter.SHARPEN)
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        result.save(tmp.name, format="PNG")
        tmp.close()
        return tmp.name
    except Exception:
        return None

# ============================================================
# 4. Vertex AI 배치 처리
# ============================================================
try:
    from google.cloud import storage as _gcs
    GCS_OK = True
except ImportError:
    GCS_OK = False

GCS_BUCKET    = "mathwash-batch-lkh0117-arch"
_WASH_JOBS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "batch_jobs_wash.json")
_WASH_JOBS_LOCK = threading.Lock()
_SUCCEEDED_STATES  = {"JOB_STATE_SUCCEEDED", "SUCCEEDED"}
_RUNNING_STATES    = {"JOB_STATE_RUNNING", "JOB_STATE_QUEUED", "JOB_STATE_PENDING", "RUNNING"}
_LOOP_PAT = re.compile(
    r'(?:수정[：:]|다시 수정|최종 수정|재설계|다시 시도|다시 풀|재시도|다른 문제)'
)

_STATE_KO = {
    "JOB_STATE_QUEUED":    "대기중",
    "JOB_STATE_PENDING":   "준비중",
    "JOB_STATE_RUNNING":   "처리중",
    "JOB_STATE_SUCCEEDED": "완료 ✅",
    "JOB_STATE_FAILED":    "실패 ❌",
    "JOB_STATE_CANCELLED": "취소됨",
    "JOB_STATE_EXPIRED":   "만료됨",
    "RUNNING":             "처리중",
    "SUCCEEDED":           "완료 ✅",
    "FAILED":              "실패 ❌",
}

def _load_wash_jobs() -> list:
    if os.path.exists(_WASH_JOBS_FILE):
        try:
            with open(_WASH_JOBS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def _save_wash_jobs(jobs: list):
    tmp = _WASH_JOBS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(jobs, f, ensure_ascii=False, indent=2)
    os.replace(tmp, _WASH_JOBS_FILE)

def _gcs_client():
    return _gcs.Client(project=GCP_PROJECT)

def _ensure_bucket():
    client = _gcs_client()
    try:
        return client.get_bucket(GCS_BUCKET)
    except Exception:
        return client.create_bucket(GCS_BUCKET, location="US")

def pil_to_b64(img) -> str:
    buf = BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def _make_wash_request_line(idx, img, task_mode, q_type, target_q, style, extra_note,
                             graph_img=None, number_style: str = "원본") -> str:
    _, thinking_budget = _pick_model(task_mode, style)
    user_prompt = get_prompt(task_mode, q_type, target_q,
                             extra_note=extra_note, style=style)
    parts = [
        {"text": f"<<idx:{idx}>>"},
        {"text": user_prompt},
        {"inlineData": {"mimeType": "image/png", "data": pil_to_b64(img)}}
    ]
    req = {
        "request": {
            "systemInstruction": {"parts": [{"text": _build_sys_instr(style, number_style)}]},
            "contents": [{"role": "user", "parts": parts}],
            "generationConfig": {"thinkingConfig": {"thinkingBudget": thinking_budget}}
        }
    }
    return json.dumps(req, ensure_ascii=False)

def _process_batch_response(text: str) -> str:
    result = _strip_thinking_leak(text)
    result = re.sub(r'<그래프>.*?</그래프>', '[그래프/도형 삽입필요]', result, flags=re.DOTALL)
    result = result.replace('<그래프/>', '[그래프/도형 삽입필요]')
    return _convert_equations_in_result(result)

def submit_wash_batch(selected, task_mode, style, extra_note, status_cb, done_cb, label="",
                      number_style: str = "원본"):
    """selected: [(img, q, q_type) 또는 (img, q, q_type, graph_img), ...]"""
    if not GCS_OK:
        done_cb(None, "google-cloud-storage 미설치\npip install google-cloud-storage")
        return
    try:
        job_id = datetime.datetime.now().strftime("wash_%Y%m%d_%H%M%S")
        total  = len(selected)

        # ── 1. JSONL 빌드
        lines = []
        items_meta = []
        for i, item in enumerate(selected):
            img, q, q_type = item[0], item[1], item[2]
            graph_img = item[3] if len(item) > 3 else None
            status_cb(f"📄 요청 인코딩 {i+1}/{total}…")
            lines.append(_make_wash_request_line(i, img, task_mode, q_type, q, style, extra_note,
                                                  graph_img=graph_img,
                                                  number_style=number_style))
            items_meta.append({"q": q, "q_type": q_type})

        # ── 2. GCS 업로드
        status_cb("☁️ GCS 업로드 중…")
        bucket     = _ensure_bucket()
        input_blob = f"{job_id}/input.jsonl"
        bucket.blob(input_blob).upload_from_string(
            "\n".join(lines).encode("utf-8"), content_type="application/jsonl")
        input_gcs  = f"gs://{GCS_BUCKET}/{input_blob}"
        output_gcs = f"gs://{GCS_BUCKET}/{job_id}/output/"

        # ── 3. 배치 작업 제출
        status_cb("🚀 Vertex AI 배치 작업 제출 중…")
        batch_model, _ = _pick_model(task_mode, style)
        batch = _gemini_client.batches.create(
            model=batch_model,
            src=input_gcs,
            config=types.CreateBatchJobConfig(dest=output_gcs)
        )

        job_info = {
            "id":           job_id,
            "label":        label or job_id,
            "job_name":     batch.name,
            "submitted_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "task_mode":    task_mode,
            "style":        style,
            "item_count":   total,
            "items":        items_meta,
            "gcs_output":   output_gcs,
            "status":       "RUNNING"
        }
        with _WASH_JOBS_LOCK:
            jobs = _load_wash_jobs()
            jobs.append(job_info)
            _save_wash_jobs(jobs)
        done_cb(job_info, None)

    except Exception as e:
        done_cb(None, str(e))

def _wash_batch_status(job_info: dict) -> str:
    try:
        batch = _gemini_client.batches.get(name=job_info["job_name"])
        state = batch.state.name if hasattr(batch.state, "name") else str(batch.state)
    except Exception as e:
        state = f"오류: {e}"
    with _WASH_JOBS_LOCK:
        jobs = _load_wash_jobs()
        for j in jobs:
            if j["id"] == job_info["id"]:
                j["status"] = state
                break
        _save_wash_jobs(jobs)
    return state

def _wash_batch_download(job_info: dict):
    """완료된 배치 GCS 결과 다운로드 + 변환. (idx, text) 리스트 반환."""
    items_meta = job_info.get("items", [])
    client = _gcs_client()
    bucket = client.bucket(GCS_BUCKET)
    prefix = job_info["gcs_output"].replace(f"gs://{GCS_BUCKET}/", "")
    blobs  = [b for b in bucket.list_blobs(prefix=prefix)
              if "prediction" in b.name or b.name.endswith(".jsonl")]
    if not blobs:
        return [], "출력 파일을 찾을 수 없습니다."
    results = []
    wash_answers = {}
    review_problems = {}
    raw_idx = 0
    for blob in sorted(blobs, key=lambda b: b.name):
        for line in blob.download_as_text(encoding="utf-8").strip().split("\n"):
            if not line.strip():
                continue
            try:
                data = json.loads(line)
                candidates = data.get("response", {}).get("candidates", [])
                parts = candidates[0].get("content", {}).get("parts", []) if candidates else []
                # thought=True 파트(내부 추론)를 제외하고 실제 응답 텍스트만 추출
                text = next((p.get("text", "") for p in parts if not p.get("thought", False)), "[응답 없음]")
                # 식별자로 원래 idx 복원 (없으면 raw_idx fallback)
                req_parts = (data.get("request", {})
                                 .get("contents", [{}])[0]
                                 .get("parts", []))
                item_idx = None
                for p in req_parts:
                    m = re.match(r'<<idx:(\d+)>>', p.get("text", ""))
                    if m:
                        item_idx = int(m.group(1))
                        break
                if item_idx is None:
                    item_idx = raw_idx
            except Exception as e:
                text = f"[파싱 오류: {e}]"
                item_idx = raw_idx
            meta = items_meta[item_idx] if item_idx < len(items_meta) else {}
            # 검수용: raw text에서 정답·문제 본문 추출 (처리 전)
            w_ans = _extract_wash_answer(text)
            if w_ans:
                wash_answers[str(item_idx)] = w_ans
                prob = _extract_problem_for_review(text)
                if prob:
                    review_problems[str(item_idx)] = prob
            # 뺑뺑이 감지 → 결과 버리고 불일치로 처리
            is_loop = bool(_LOOP_PAT.search(text)) if text else False
            processed = "[검수 불일치]" if is_loop else _process_batch_response(text)
            raw_ok = (text and not text.startswith(("[응답 없음]", "[파싱 오류")) and not is_loop)
            results.append((item_idx, meta, processed, text if raw_ok else None))
            raw_idx += 1
    # idx 순서로 정렬 후 캐쉬 저장 (제출 순서대로 ID 부여)
    results.sort(key=lambda x: x[0])
    style = job_info.get("style", "강사용")
    for _, _, processed, raw_text in results:
        if raw_text:
            save_ai_output_cache(processed, style, raw_text=raw_text)
    results = [(idx, meta, processed) for idx, meta, processed, _ in results]
    # wash_answers / review_problems job_info에 저장
    if wash_answers:
        with _WASH_JOBS_LOCK:
            jobs = _load_wash_jobs()
            for j in jobs:
                if j["id"] == job_info["id"]:
                    j["wash_answers"]    = wash_answers
                    j["review_problems"] = review_problems
                    break
            _save_wash_jobs(jobs)
        job_info["wash_answers"]    = wash_answers
        job_info["review_problems"] = review_problems
    return results, None

def _submit_review_batch(job_info: dict):
    """워싱 결과에서 추출한 문제 텍스트로 검수 배치 제출."""
    review_problems = job_info.get("review_problems", {})
    if not review_problems:
        return "검수할 문제가 없습니다."
    try:
        job_id = job_info["id"]
        lines = []
        for idx_str, prob_text in sorted(review_problems.items(), key=lambda x: int(x[0])):
            req = {
                "request": {
                    "systemInstruction": {"parts": [{"text": _REVIEW_SYSTEM}]},
                    "contents": [{"role": "user", "parts": [
                        {"text": f"<<idx:{idx_str}>>"},
                        {"text": prob_text}
                    ]}],
                    "generationConfig": {
                        "maxOutputTokens": 16,
                        "thinkingConfig": {"thinkingBudget": 1024}
                    }
                }
            }
            lines.append(json.dumps(req, ensure_ascii=False))

        bucket     = _ensure_bucket()
        input_blob = f"{job_id}/review_input.jsonl"
        bucket.blob(input_blob).upload_from_string(
            "\n".join(lines).encode("utf-8"), content_type="application/jsonl")
        input_gcs  = f"gs://{GCS_BUCKET}/{input_blob}"
        output_gcs = f"gs://{GCS_BUCKET}/{job_id}/review_output/"

        batch = _gemini_client.batches.create(
            model=MODEL_FLASH,
            src=input_gcs,
            config=types.CreateBatchJobConfig(dest=output_gcs)
        )

        with _WASH_JOBS_LOCK:
            jobs = _load_wash_jobs()
            for j in jobs:
                if j["id"] == job_id:
                    j["review_job_name"]  = batch.name
                    j["review_status"]    = "RUNNING"
                    j["review_gcs_output"] = output_gcs
                    break
            _save_wash_jobs(jobs)
        job_info["review_job_name"]   = batch.name
        job_info["review_status"]     = "RUNNING"
        job_info["review_gcs_output"] = output_gcs
        return None
    except Exception as e:
        return str(e)


def _review_batch_status(job_info: dict) -> str:
    try:
        batch = _gemini_client.batches.get(name=job_info["review_job_name"])
        state = batch.state.name if hasattr(batch.state, "name") else str(batch.state)
    except Exception as e:
        state = f"오류: {e}"
    with _WASH_JOBS_LOCK:
        jobs = _load_wash_jobs()
        for j in jobs:
            if j["id"] == job_info["id"]:
                j["review_status"] = state
                break
        _save_wash_jobs(jobs)
    job_info["review_status"] = state
    return state


def _review_batch_download(job_info: dict) -> tuple[dict, str | None]:
    """검수 배치 결과 다운로드 → {idx_str: answer} 반환."""
    client = _gcs_client()
    bucket = client.bucket(GCS_BUCKET)
    prefix = job_info["review_gcs_output"].replace(f"gs://{GCS_BUCKET}/", "")
    blobs  = [b for b in bucket.list_blobs(prefix=prefix)
              if "prediction" in b.name or b.name.endswith(".jsonl")]
    if not blobs:
        return {}, "검수 출력 파일을 찾을 수 없습니다."
    review_answers = {}
    raw_idx = 0
    for blob in sorted(blobs, key=lambda b: b.name):
        for line in blob.download_as_text(encoding="utf-8").strip().split("\n"):
            if not line.strip():
                continue
            try:
                data = json.loads(line)
                candidates = data.get("response", {}).get("candidates", [])
                parts = candidates[0].get("content", {}).get("parts", []) if candidates else []
                ans = next((p.get("text", "") for p in parts if not p.get("thought", False)), "")
                req_parts = (data.get("request", {})
                                 .get("contents", [{}])[0]
                                 .get("parts", []))
                idx_str = str(raw_idx)
                for p in req_parts:
                    m = re.match(r'<<idx:(\d+)>>', p.get("text", ""))
                    if m:
                        idx_str = m.group(1)
                        break
                review_answers[idx_str] = ans.strip()
            except Exception:
                pass
            raw_idx += 1
    with _WASH_JOBS_LOCK:
        jobs = _load_wash_jobs()
        for j in jobs:
            if j["id"] == job_info["id"]:
                j["review_answers"] = review_answers
                j["review_status"]  = "SUCCEEDED"
                break
        _save_wash_jobs(jobs)
    job_info["review_answers"] = review_answers
    job_info["review_status"]  = "SUCCEEDED"
    return review_answers, None


def _submit_retry_batch(job_info: dict, mismatch_indices: list):
    """불일치 문항만 원본 GCS input.jsonl에서 추출해 재시도 배치 제출."""
    try:
        job_id = job_info["id"]
        client = _gcs_client()
        bucket = client.bucket(GCS_BUCKET)

        # 원본 input.jsonl 다운로드
        input_blob = f"{job_id}/input.jsonl"
        original_lines = bucket.blob(input_blob).download_as_text(encoding="utf-8").strip().split("\n")

        # 불일치 idx 행만 필터
        mismatch_set = set(mismatch_indices)
        retry_lines = []
        for line in original_lines:
            if not line.strip():
                continue
            try:
                data = json.loads(line)
                req_parts = (data.get("request", {})
                                 .get("contents", [{}])[0]
                                 .get("parts", []))
                for p in req_parts:
                    m = re.match(r'<<idx:(\d+)>>', p.get("text", ""))
                    if m and int(m.group(1)) in mismatch_set:
                        retry_lines.append(line)
                        break
            except Exception:
                continue

        if not retry_lines:
            return "재시도할 항목이 없습니다."

        retry_input_blob = f"{job_id}/retry_input.jsonl"
        retry_output_gcs = f"gs://{GCS_BUCKET}/{job_id}/retry_output/"
        bucket.blob(retry_input_blob).upload_from_string(
            "\n".join(retry_lines).encode("utf-8"), content_type="application/jsonl")

        batch_model, _ = _pick_model(job_info.get("task_mode", "워싱"),
                                     job_info.get("style", "강사용"))
        batch = _gemini_client.batches.create(
            model=batch_model,
            src=f"gs://{GCS_BUCKET}/{retry_input_blob}",
            config=types.CreateBatchJobConfig(dest=retry_output_gcs)
        )

        with _WASH_JOBS_LOCK:
            jobs = _load_wash_jobs()
            for j in jobs:
                if j["id"] == job_id:
                    j["retry_job_name"]       = batch.name
                    j["retry_status"]         = "RUNNING"
                    j["retry_gcs_output"]     = retry_output_gcs
                    j["retry_mismatch_indices"] = list(mismatch_indices)
                    break
            _save_wash_jobs(jobs)
        job_info["retry_job_name"]         = batch.name
        job_info["retry_status"]           = "RUNNING"
        job_info["retry_gcs_output"]       = retry_output_gcs
        job_info["retry_mismatch_indices"] = list(mismatch_indices)
        return None
    except Exception as e:
        return str(e)


def _retry_batch_status(job_info: dict) -> str:
    try:
        batch = _gemini_client.batches.get(name=job_info["retry_job_name"])
        state = batch.state.name if hasattr(batch.state, "name") else str(batch.state)
    except Exception as e:
        state = f"오류: {e}"
    with _WASH_JOBS_LOCK:
        jobs = _load_wash_jobs()
        for j in jobs:
            if j["id"] == job_info["id"]:
                j["retry_status"] = state
                break
        _save_wash_jobs(jobs)
    job_info["retry_status"] = state
    return state


def _retry_batch_download(job_info: dict) -> dict:
    """재시도 배치 결과 다운로드 → {idx: processed_text} 반환."""
    client = _gcs_client()
    bucket = client.bucket(GCS_BUCKET)
    prefix = job_info["retry_gcs_output"].replace(f"gs://{GCS_BUCKET}/", "")
    blobs  = [b for b in bucket.list_blobs(prefix=prefix)
              if "prediction" in b.name or b.name.endswith(".jsonl")]
    retry_results = {}
    raw_idx = 0
    for blob in sorted(blobs, key=lambda b: b.name):
        for line in blob.download_as_text(encoding="utf-8").strip().split("\n"):
            if not line.strip():
                continue
            try:
                data = json.loads(line)
                candidates = data.get("response", {}).get("candidates", [])
                parts = candidates[0].get("content", {}).get("parts", []) if candidates else []
                text = next((p.get("text", "") for p in parts if not p.get("thought", False)), "[응답 없음]")
                req_parts = (data.get("request", {})
                                 .get("contents", [{}])[0]
                                 .get("parts", []))
                item_idx = raw_idx
                for p in req_parts:
                    m = re.match(r'<<idx:(\d+)>>', p.get("text", ""))
                    if m:
                        item_idx = int(m.group(1))
                        break
                is_loop = bool(_LOOP_PAT.search(text)) if text else False
                retry_results[item_idx] = "[검수 불일치]" if is_loop else _process_batch_response(text)
            except Exception:
                pass
            raw_idx += 1
    with _WASH_JOBS_LOCK:
        jobs = _load_wash_jobs()
        for j in jobs:
            if j["id"] == job_info["id"]:
                j["retry_status"] = "SUCCEEDED"
                break
        _save_wash_jobs(jobs)
    job_info["retry_status"] = "SUCCEEDED"
    return retry_results


def _combined_status(job: dict) -> str:
    """워싱·검수·재시도 상태를 합쳐서 표시 문자열 반환."""
    wash_s   = job.get("status", "")
    review_s = job.get("review_status", "")
    retry_s  = job.get("retry_status", "")
    wash_done   = wash_s in _SUCCEEDED_STATES
    review_done = review_s in _SUCCEEDED_STATES

    if not wash_done:
        return _STATE_KO.get(wash_s, wash_s)

    if not review_s:
        return "워싱완료 / 검수대기"
    if review_s in _RUNNING_STATES:
        return "워싱완료 / 검수 처리중"

    if review_done:
        if not retry_s:
            return "워싱+검수 완료 ✅"
        if retry_s in _RUNNING_STATES:
            return "워싱+검수 완료 / 재시도 처리중"
        if retry_s in _SUCCEEDED_STATES:
            return "최종 완료 ✅"
        return f"워싱+검수 완료 / 재시도:{_STATE_KO.get(retry_s, retry_s)}"

    return f"워싱완료 / 검수:{_STATE_KO.get(review_s, review_s)}"


def open_wash_batch_dialog(result_cb, status_cb):
    dlg = tk.Toplevel()
    dlg.title("📦 배치 작업 관리")
    dlg.geometry("1060x520")
    dlg.resizable(True, True)
    dlg.minsize(800, 380)

    frm = tk.Frame(dlg, padx=10, pady=8)
    frm.pack(fill=tk.BOTH, expand=True)
    frm.columnconfigure(0, weight=1)
    frm.rowconfigure(0, weight=1)
    frm.rowconfigure(1, weight=0)
    frm.rowconfigure(2, weight=0)

    # 트리뷰 + 스크롤바를 전용 서브프레임으로 묶어 레이아웃 분리
    tree_fr = tk.Frame(frm)
    tree_fr.grid(row=0, column=0, sticky="nsew")
    tree_fr.columnconfigure(0, weight=1)
    tree_fr.rowconfigure(0, weight=1)

    _ts = ttk.Style()
    _ts.configure("Batch.Treeview", rowheight=35, font=("맑은 고딕", 10))
    _ts.configure("Batch.Treeview.Heading", font=("맑은 고딕", 10, "bold"))
    cols = ("name", "submitted", "mode", "items", "style", "status")
    tree = ttk.Treeview(tree_fr, columns=cols, show="headings", style="Batch.Treeview")
    tree.heading("name",      text="이름")
    tree.heading("submitted", text="제출 시간")
    tree.heading("mode",      text="모드")
    tree.heading("items",     text="문항수")
    tree.heading("style",     text="해설")
    tree.heading("status",    text="상태")
    tree.column("name",      width=210, minwidth=120, anchor="w",      stretch=True)
    tree.column("submitted", width=155, minwidth=120, anchor="center", stretch=True)
    tree.column("mode",      width=75,  minwidth=60,  anchor="center", stretch=False)
    tree.column("items",     width=65,  minwidth=50,  anchor="center", stretch=False)
    tree.column("style",     width=75,  minwidth=55,  anchor="center", stretch=False)
    tree.column("status",    width=185, minwidth=110, anchor="center", stretch=True)
    sb = ttk.Scrollbar(tree_fr, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")

    dlg_status = tk.StringVar(value="작업을 선택하세요.")
    tk.Label(frm, textvariable=dlg_status, fg="#003388",
             font=("맑은 고딕", 9), anchor="w"
             ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4,2))

    jobs_ref = [_load_wash_jobs()]

    _STYLE_SHORT = {"강사용": "강사", "학생용": "학생", "해설없음": "없음"}

    def _refresh():
        tree.delete(*tree.get_children())
        for j in jobs_ref[0]:
            sk    = _combined_status(j)
            style_s = _STYLE_SHORT.get(j.get("style",""), j.get("style","-"))
            tree.insert("", tk.END, iid=j["id"],
                        values=(j.get("label", j.get("id","-")),
                                j["submitted_at"], j.get("task_mode","-"),
                                j.get("item_count","-"), style_s, sk))

    _refresh()

    def _sel():
        sel = tree.selection()
        if not sel:
            dlg_status.set("먼저 작업을 선택하세요."); return None
        jid = sel[0]
        for j in jobs_ref[0]:
            if j["id"] == jid: return j
        return None

    _busy = [False]  # 상태확인/전체새로고침 중복 실행 방지

    def _set_buttons(enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        for b in _btn_refs:
            try: b.config(state=state)
            except: pass

    def _check():
        if _busy[0]: return
        job = _sel()
        if not job: return
        _busy[0] = True
        dlg_status.set("🔄 상태 확인 중…")
        _set_buttons(False)
        def _r():
            _wash_batch_status(job)
            if job.get("review_job_name") and job.get("review_status") not in _SUCCEEDED_STATES:
                _review_batch_status(job)
            if job.get("retry_job_name") and job.get("retry_status") not in _SUCCEEDED_STATES:
                _retry_batch_status(job)
            jobs_ref[0] = _load_wash_jobs()
            dlg.after(0, lambda: (_refresh(),
                dlg_status.set(f"상태: {_combined_status(job)}"),
                _set_buttons(True)))
            _busy[0] = False
        threading.Thread(target=_r, daemon=True).start()

    def _check_all():
        if _busy[0]: return
        _busy[0] = True
        dlg_status.set("🔄 전체 상태 새로 고침 중…")
        _set_buttons(False)
        def _r():
            jobs = list(jobs_ref[0])
            pending = [j for j in jobs
                       if j.get("status") not in _SUCCEEDED_STATES
                       or (j.get("review_job_name")
                           and j.get("review_status") not in _SUCCEEDED_STATES)
                       or (j.get("retry_job_name")
                           and j.get("retry_status") not in _SUCCEEDED_STATES)]
            for j in pending:
                _wash_batch_status(j)
                if j.get("review_job_name") and j.get("review_status") not in _SUCCEEDED_STATES:
                    _review_batch_status(j)
                if j.get("retry_job_name") and j.get("retry_status") not in _SUCCEEDED_STATES:
                    _retry_batch_status(j)
            jobs_ref[0] = _load_wash_jobs()
            dlg.after(0, lambda: (_refresh(),
                dlg_status.set(f"✅ 전체 새로 고침 완료 ({len(pending)}개 확인)"),
                _set_buttons(True)))
            _busy[0] = False
        threading.Thread(target=_r, daemon=True).start()

    def _download():
        job = _sel()
        if not job: return

        def _r():
            # ── 1단계: 워싱 완료 확인
            wash_state = job.get("status", "")
            if wash_state not in _SUCCEEDED_STATES:
                dlg.after(0, lambda: dlg_status.set("🔄 워싱 상태 확인 중…"))
                wash_state = _wash_batch_status(job)
                if wash_state not in _SUCCEEDED_STATES:
                    dlg.after(0, lambda s=wash_state: (
                        jobs_ref.__setitem__(0, _load_wash_jobs()),
                        _refresh(),
                        dlg_status.set(f"⏳ 워싱 {_STATE_KO.get(s, s)} — 아직 완료되지 않았습니다.")
                    ))
                    return

            # ── 2단계: 검수 배치가 없으면 워싱 다운로드 후 검수 배치 제출
            if not job.get("review_job_name"):
                dlg.after(0, lambda: dlg_status.set("⬇️ 워싱 결과 다운로드 중…"))
                results, err = _wash_batch_download(job)
                if err:
                    dlg.after(0, lambda e=err: dlg_status.set(f"❌ {e}"))
                    return
                dlg.after(0, lambda: dlg_status.set("🚀 검수 배치 제출 중…"))
                err2 = _submit_review_batch(job)
                jobs_ref[0] = _load_wash_jobs()
                if err2:
                    dlg.after(0, lambda e=err2: (_refresh(),
                        dlg_status.set(f"⚠️ 검수 배치 제출 실패: {e}")))
                else:
                    dlg.after(0, lambda: (_refresh(),
                        dlg_status.set("✅ 검수 배치 제출됨 — 나중에 다시 '결과 불러오기'를 누르세요.")))
                return

            # ── 3단계: 검수 배치 진행 중
            review_state = job.get("review_status", "")
            if review_state not in _SUCCEEDED_STATES:
                dlg.after(0, lambda: dlg_status.set("🔄 검수 상태 확인 중…"))
                review_state = _review_batch_status(job)
                jobs_ref[0] = _load_wash_jobs()
                if review_state not in _SUCCEEDED_STATES:
                    dlg.after(0, lambda s=review_state: (_refresh(),
                        dlg_status.set(f"⏳ 검수 {_STATE_KO.get(s, s)} — 아직 완료되지 않았습니다.")))
                    return

            # ── 4단계: 검수 완료 → 다운로드 + 비교 → 불일치 있으면 재시도 배치
            if job.get("review_status") not in _SUCCEEDED_STATES or not job.get("review_answers"):
                dlg.after(0, lambda: dlg_status.set("⬇️ 검수 결과 다운로드 중…"))
                review_answers, err3 = _review_batch_download(job)
                if err3:
                    dlg.after(0, lambda e=err3: dlg_status.set(f"❌ {e}"))
                    return
            else:
                review_answers = job["review_answers"]

            if not job.get("retry_status"):
                wash_answers = job.get("wash_answers", {})
                results, _ = _wash_batch_download(job)
                mismatch_indices = [
                    idx for idx, meta, text in results
                    if (wash_answers.get(str(idx)) and review_answers.get(str(idx)) and
                        not _answers_match(wash_answers[str(idx)], review_answers[str(idx)]))
                ]
                if mismatch_indices:
                    dlg.after(0, lambda n=len(mismatch_indices): dlg_status.set(
                        f"🔁 불일치 {n}문항 재시도 배치 제출 중…"))
                    err4 = _submit_retry_batch(job, mismatch_indices)
                    jobs_ref[0] = _load_wash_jobs()
                    if err4:
                        dlg.after(0, lambda e=err4: (_refresh(),
                            dlg_status.set(f"⚠️ 재시도 배치 제출 실패: {e}")))
                    else:
                        dlg.after(0, lambda: (_refresh(),
                            dlg_status.set("✅ 재시도 배치 제출됨 — 나중에 다시 '결과 불러오기'를 누르세요.")))
                    return
                # 불일치 없으면 바로 최종 출력
                output_parts = [text for _, _, text in results]
                combined = "\n\n".join(p for p in output_parts if p)
                jobs_ref[0] = _load_wash_jobs()
                dlg.after(0, lambda: (
                    _refresh(),
                    result_cb(combined),
                    status_cb(f"✅ 워싱+검수 완료 ({len(results)}문항, 불일치 없음)"),
                    dlg_status.set("✅ 워싱+검수 완료! 결과창에 표시됩니다.")
                ))
                return

            # ── 5단계: 재시도 배치 진행 중
            retry_state = job.get("retry_status", "")
            if retry_state not in _SUCCEEDED_STATES:
                dlg.after(0, lambda: dlg_status.set("🔄 재시도 배치 상태 확인 중…"))
                retry_state = _retry_batch_status(job)
                jobs_ref[0] = _load_wash_jobs()
                if retry_state not in _SUCCEEDED_STATES:
                    dlg.after(0, lambda s=retry_state: (_refresh(),
                        dlg_status.set(f"⏳ 재시도 {_STATE_KO.get(s, s)} — 아직 완료되지 않았습니다.")))
                    return

            # ── 6단계: 재시도 완료 → 병합 후 최종 출력
            dlg.after(0, lambda: dlg_status.set("⬇️ 재시도 결과 다운로드 중…"))
            retry_results = _retry_batch_download(job)
            results, _ = _wash_batch_download(job)
            mismatch_set = set(job.get("retry_mismatch_indices", []))

            output_parts = []
            for idx, meta, text in results:
                if idx in mismatch_set:
                    retry_text = retry_results.get(idx)
                    output_parts.append(retry_text if retry_text else "[검수 불일치]")
                else:
                    output_parts.append(text)

            combined = "\n\n".join(p for p in output_parts if p)
            jobs_ref[0] = _load_wash_jobs()
            dlg.after(0, lambda: (
                _refresh(),
                result_cb(combined),
                status_cb(f"✅ 최종 완료 ({len(results)}문항)"),
                dlg_status.set("✅ 최종 완료! 결과창에 표시됩니다.")
            ))

        threading.Thread(target=_r, daemon=True).start()

    def _delete():
        job = _sel()
        if not job: return
        if not messagebox.askyesno("삭제 확인", "목록에서 삭제할까요?", parent=dlg): return
        jobs_ref[0] = [j for j in jobs_ref[0] if j["id"] != job["id"]]
        _save_wash_jobs(jobs_ref[0]); _refresh()

    btn_frm = tk.Frame(frm)
    btn_frm.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6,0))
    _btn_refs = []
    for txt, cmd, bg in [
        ("🔄 상태 확인",       _check,     "#D8F0FF"),
        ("🔃 전체 새로 고침",   _check_all, "#D0ECFF"),
        ("⬇️ 결과 불러오기",   _download,  "#CCFFCC"),
        ("🗑 삭제",            _delete,    "#FFCCCC"),
        ("닫기",               dlg.destroy,"#EEEEEE"),
    ]:
        b = tk.Button(btn_frm, text=txt, command=cmd, font=("맑은 고딕", 10),
                      bg=bg, relief="raised", bd=2, padx=12, pady=4)
        b.pack(side=tk.LEFT, padx=4)
        if txt != "닫기":
            _btn_refs.append(b)


# ============================================================
# (검토 기능 제거됨)
# ============================================================
# ============================================================
# 6. 단일 분석
# ============================================================
def run_single_clipboard():
    """클립보드 단일 이미지 분석 — 단일 패널 변수 사용."""
    img = ImageGrab.grabclipboard()
    if not isinstance(img, Image.Image):
        messagebox.showwarning("주의", "클립보드에 이미지가 없습니다!\nWin+Shift+S로 먼저 캡처하세요.")
        return
    task_mode = single_mode_var.get()
    q_type    = single_type_var.get() if task_mode == "워싱" else "객관식"
    style     = single_style_var.get()
    extra_raw = single_note_area.get("1.0", tk.END).strip()
    extra     = extra_raw if extra_raw and extra_raw != _SINGLE_PLACEHOLDER else None
    root.after(0, lambda: input_area.delete("1.0", tk.END))
    root.after(0, lambda: status_label.config(
        text=f"🤖 [{task_mode}/해설:{style}] 분석 중...", fg="#0000FF"))
    threading.Thread(target=_single_image_thread,
                     args=(img, task_mode, q_type, style, extra), daemon=True).start()


def _single_image_thread(img, task_mode, q_type, style, extra_note):
    try:
        result, raw_text = _call_gemini(img, task_mode, q_type, style=style, extra_note=extra_note)
        if not _is_valid_result(result, style=style) and result:
            result = "[⚠️ 형식 오류 — 검토 필요]\n" + result
        root.after(0, lambda r=result: input_area.insert(tk.END, r or "[분석 실패]"))
        root.after(0, _update_counter)
        root.after(0, lambda: status_label.config(
            text=f"✅ [{task_mode}] 완료!", fg="#008000"))
        # AI 캐쉬 저장 + 풀이저장 창 연동
        if result:
            ids = save_ai_output_cache(result, style, raw_text=raw_text)
            if ids:
                _last_problem_id[0] = ids[0]
            # 뷰어 미리보기: 본문+보기만 추출 (미주·번호 제거)
            src = raw_text or result
            blocks = _split_ai_results(src, style)
            preview_body = _extract_ai_output_body(blocks[0], style) if blocks else src
            _last_latex_result[0] = preview_body
            root.after(0, _push_result_to_solution_viewer, preview_body)
    except Exception as e:
        err = str(e)
        root.after(0, lambda em=err: messagebox.showerror("AI 오류", em))
        root.after(0, lambda: status_label.config(text="❌ 분석 실패", fg="#FF0000"))


def _push_result_to_solution_viewer(text: str):
    """풀이저장 창이 열려 있으면 문제 미리보기를 업데이트 (편집창은 건드리지 않음)."""
    try:
        win = _solution_viewer_ref[0]
        if win is None:
            return
        update_prev = getattr(win, "_update_preview_fn", None)
        if update_prev:
            update_prev(text)
    except Exception:
        pass


# ============================================================
# 7. PDF 뷰어 (수동 크롭 방식)
# ============================================================

def _detect_crops_by_numbers(doc) -> list:
    """PDF 텍스트에서 문제 번호(01 / 1. / 2 등) 위치로 문항 bbox 추출.
    구조화된 PDF(비스캔)에서만 동작.
    Returns [(pi, rx0, ry0, rx1, ry1), ...]
    """
    NUM_PAT = re.compile(r'^(\d{1,2})[\s\.．·]')
    results = []

    for pi in range(len(doc)):
        page = doc.load_page(pi)
        pw, ph = page.rect.width, page.rect.height
        if pw == 0 or ph == 0:
            continue

        raw = page.get_text("blocks")
        tblks = [b for b in raw if b[6] == 0 and b[4].strip()]
        if not tblks:
            continue

        # 2단 레이아웃 감지 (기존 로직과 동일)
        xs = [(b[0] + b[2]) / 2 for b in tblks if len(b[4].strip()) > 5]
        csplit = pw / 2
        if len(xs) >= 4:
            ln = sum(1 for x in xs if x < csplit)
            rn = sum(1 for x in xs if x >= csplit)
            two_col = ln >= 2 and rn >= 2 and min(ln, rn) / max(ln, rn) > 0.25
        else:
            two_col = False

        col_ranges = [(0, csplit), (csplit, pw)] if two_col else [(0, pw)]

        for cx0, cx1 in col_ranges:
            cw = cx1 - cx0
            col = sorted([b for b in tblks if cx0 <= (b[0] + b[2]) / 2 < cx1],
                         key=lambda b: b[1])

            starts = []  # (y_top, num)
            for b in col:
                txt = b[4].lstrip()
                m = NUM_PAT.match(txt)
                if not m:
                    continue
                n = int(m.group(1))
                # 번호 블록이 컬럼 왼쪽 22% 이내에 있어야 문제 번호로 인정
                if 1 <= n <= 50 and b[0] <= cx0 + cw * 0.22:
                    starts.append((b[1], n))

            if len(starts) < 2:
                continue
            # 최소 1쌍 이상 연속 번호 있어야 신뢰
            nums = [s[1] for s in starts]
            if not any(nums[i + 1] - nums[i] == 1 for i in range(len(nums) - 1)):
                continue

            pad = 3 / ph
            for i, (y0, _) in enumerate(starts):
                y1 = starts[i + 1][0] if i + 1 < len(starts) else ph
                results.append((pi, cx0 / pw, max(0.0, y0 / ph - pad),
                                cx1 / pw, min(1.0, y1 / ph)))

    return results


def _detect_crops_gemini(doc, page_imgs, progress_cb=None) -> list:
    """Gemini 2.5 Flash로 수학 문항 영역 감지 (병렬 처리).
    Returns [(pi, rx0, ry0, rx1, ry1), ...]  (비율 좌표 0~1)
    """
    PROMPT = (
        "이 수학 시험지 이미지에서 각 문항(문제)의 전체 영역을 감지해줘.\n"
        "반드시 JSON 배열만 출력해. 설명 없이.\n"
        "좌표는 이미지 전체 크기 대비 비율(0.0~1.0).\n"
        "두 열 레이아웃이면 왼열 문항들 → 오른열 문항들 순서로.\n\n"
        "중요 규칙:\n"
        "- 문제 번호(1. 2. 등)부터 시작해서 선지(①②③④⑤) 끝까지 반드시 포함\n"
        "- 조건, 보기, 그림, 그래프도 해당 문항 박스에 포함\n"
        "- 해설/풀이/정답 영역은 완전히 제외\n"
        "- 박스가 문항 내용을 조금이라도 잘라내지 않도록 넉넉하게\n\n"
        '형식: [{"x1":0.0,"y1":0.05,"x2":0.5,"y2":0.3}, ...]'
    )
    total = len(doc)
    done  = [0]

    def _process_page(pi):
        base = page_imgs[pi] if page_imgs and pi < len(page_imgs) else None
        if base is None:
            pix  = doc.load_page(pi).get_pixmap(matrix=fitz.Matrix(1.8, 1.8))
            base = Image.open(BytesIO(pix.tobytes("png")))
        try:
            resp = _gemini_client.models.generate_content(
                model=MODEL_FLASH,
                contents=[PROMPT, pil_to_part(base)],
            )
            text = resp.text.strip()
            m = re.search(r'\[.*?\]', text, re.DOTALL)
            if not m:
                return pi, []
            boxes = json.loads(m.group())
            PAD = 0.015  # 1.5% 여백
            page_results = []
            for b in boxes:
                x1 = max(0.0, float(b.get("x1", 0)) - PAD)
                y1 = max(0.0, float(b.get("y1", 0)) - PAD)
                x2 = min(1.0, float(b.get("x2", 1)) + PAD)
                y2 = min(1.0, float(b.get("y2", 1)) + PAD)
                if x2 > x1 and y2 > y1:
                    page_results.append((pi, x1, y1, x2, y2))
            return pi, page_results
        except Exception:
            return pi, []
        finally:
            done[0] += 1
            if progress_cb:
                progress_cb(done[0] - 1)

    page_map = {}
    with ThreadPoolExecutor(max_workers=min(total, 8)) as ex:
        for pi, boxes in ex.map(_process_page, range(total)):
            page_map[pi] = boxes

    results = []
    for pi in range(total):
        results.extend(page_map.get(pi, []))
    return results


def _detect_crops_yolo(doc, page_imgs, progress_cb=None) -> list:
    """DocLayout-YOLO로 레이아웃 감지 후 문항 bbox 추출.
    pip install doclayout-yolo  huggingface_hub  필요.
    Returns [(pi, rx0, ry0, rx1, ry1), ...]
    """
    try:
        from doclayout_yolo import YOLOv10
    except ImportError:
        raise ImportError("doclayout-yolo")

    from pathlib import Path as _Path
    import tempfile as _tmp, os as _os

    cache_dir = _Path.home() / ".math_washing_cache" / "models"
    cache_dir.mkdir(parents=True, exist_ok=True)
    model_file = cache_dir / "doclayout_yolo_docstructbench_imgsz1024.pt"

    if not model_file.exists():
        try:
            from huggingface_hub import hf_hub_download
            hf_hub_download(
                repo_id="juliozhao/DocLayout-YOLO-DocStructBench",
                filename="doclayout_yolo_docstructbench_imgsz1024.pt",
                local_dir=str(cache_dir),
            )
        except Exception as e:
            raise RuntimeError(f"모델 다운로드 실패: {e}")

    model = YOLOv10(str(model_file))
    NUM_PAT = re.compile(r'^(\d{1,2})[\s\.．·]')
    results = []

    for pi, base_img in enumerate(page_imgs):
        if base_img is None:
            continue
        if progress_cb:
            progress_cb(pi)

        page = doc.load_page(pi)
        pw, ph = page.rect.width, page.rect.height
        iw, ih = base_img.size

        # YOLO 예측 (임시 파일 경유)
        tf = _tmp.NamedTemporaryFile(suffix=".jpg", delete=False)
        tf.close()
        try:
            base_img.save(tf.name, "JPEG", quality=88)
            det = model.predict(tf.name, imgsz=1024, conf=0.2,
                                device="cpu", verbose=False)[0]
        finally:
            try: _os.unlink(tf.name)
            except: pass

        boxes_xy  = det.boxes.xyxy.cpu().numpy()   # (N,4) pixel
        classes   = det.boxes.cls.cpu().numpy()    # (N,)
        if len(boxes_xy) == 0:
            continue

        # 텍스트·타이틀 클래스만 (plain text=0, title=1)
        KEEP = {0, 1}
        tboxes = [(boxes_xy[i], boxes_xy[i][1] / ih)
                  for i in range(len(boxes_xy)) if int(classes[i]) in KEEP]
        if not tboxes:
            continue

        # 2단 감지
        raw = page.get_text("blocks")
        tblks = [b for b in raw if b[6] == 0 and b[4].strip()]
        xs = [(b[0] + b[2]) / 2 for b in tblks if len(b[4].strip()) > 5]
        csplit_img = iw / 2
        if len(xs) >= 4:
            ln = sum(1 for x in xs if x < pw / 2)
            rn = sum(1 for x in xs if x >= pw / 2)
            two_col = ln >= 2 and rn >= 2 and min(ln, rn) / max(ln, rn) > 0.25
        else:
            two_col = False

        col_ranges_img = [(0, csplit_img), (csplit_img, iw)] if two_col else [(0, iw)]

        for cx0_i, cx1_i in col_ranges_img:
            col_boxes = sorted(
                [(b, ry) for b, ry in tboxes if cx0_i <= (b[0] + b[2]) / 2 < cx1_i],
                key=lambda x: x[1])
            if len(col_boxes) < 2:
                continue

            rx0_col = cx0_i / iw
            rx1_col = cx1_i / iw

            # PyMuPDF 텍스트로 각 YOLO 박스 안의 문제 번호 확인
            prob_ys = []
            for b, _ in col_boxes:
                clip = fitz.Rect(b[0] * pw / iw, b[1] * ph / ih,
                                 b[2] * pw / iw, b[3] * ph / ih)
                txt  = page.get_text("text", clip=clip).strip()
                first = txt.split('\n')[0].lstrip() if txt else ""
                m = NUM_PAT.match(first)
                if m and 1 <= int(m.group(1)) <= 50:
                    prob_ys.append(b[1] / ih)

            if len(prob_ys) >= 2:
                pad = 0.003
                for i, py in enumerate(prob_ys):
                    py_end = prob_ys[i + 1] if i + 1 < len(prob_ys) else 1.0
                    results.append((pi, rx0_col, max(0.0, py - pad),
                                    rx1_col, min(1.0, py_end)))
            else:
                # 번호 못 찾으면 YOLO 박스들을 공백 기준으로 묶어서 출력
                GAP = 0.02
                merged = []
                cur_y0 = col_boxes[0][0][1] / ih
                cur_y1 = col_boxes[0][0][3] / ih
                for b, _ in col_boxes[1:]:
                    by0, by1 = b[1] / ih, b[3] / ih
                    if by0 - cur_y1 < GAP:
                        cur_y1 = max(cur_y1, by1)
                    else:
                        if cur_y1 - cur_y0 > 0.03:
                            merged.append((cur_y0, cur_y1))
                        cur_y0, cur_y1 = by0, by1
                if cur_y1 - cur_y0 > 0.03:
                    merged.append((cur_y0, cur_y1))
                for ry0, ry1 in merged:
                    results.append((pi, rx0_col, ry0, rx1_col, ry1))

    return results


def _auto_detect_crops(doc, page_imgs=None) -> list:
    """문항 경계 자동 감지.
    1차: 텍스트 문제 번호 위치 기반 (구조화 PDF 전용, 정확)
    2차 fallback: 이미지 수평 공백 밴드 분석 (스캔 PDF 포함)
    Returns [(page_idx, rx1, ry1, rx2, ry2), ...]
    """
    by_num = _detect_crops_by_numbers(doc)
    if len(by_num) >= 2:
        return by_num
    # fallback: 공백 밴드 분석
    results = []

    for pi in range(len(doc)):
        base = page_imgs[pi] if page_imgs and pi < len(page_imgs) else None
        if base is None:
            pix  = doc.load_page(pi).get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
            base = Image.open(BytesIO(pix.tobytes("png")))

        page = doc.load_page(pi)
        pw, ph = page.rect.width, page.rect.height
        if pw == 0 or ph == 0:
            continue

        # ── 분석용 축소 (4배) ────────────────────────────
        SCALE = 4
        aw = max(80, base.width  // SCALE)
        ah = max(80, base.height // SCALE)
        gray = base.convert("L").resize((aw, ah), Image.Resampling.BOX)
        data = gray.tobytes()

        # ── 전체 행 평균으로 적응형 임계값 계산 ──────────
        row_means_all = []
        for y in range(ah):
            off = y * aw
            row_means_all.append(sum(data[off: off + aw]) / aw)
        # 상위 5% (가장 밝은 행들)의 평균 × 0.96 → 공백 기준
        bright = sorted(row_means_all, reverse=True)
        top5   = bright[:max(1, len(bright) // 20)]
        white_thresh = max(210.0, sum(top5) / len(top5) * 0.96)

        # ── 두 열 레이아웃 감지 (텍스트 블록 x좌표 분포) ─
        blocks  = page.get_text("blocks")
        text_xs = [(b[0] + b[2]) / 2
                   for b in blocks if b[6] == 0 and len(b[4].strip()) > 8]
        col_split = pw / 2
        if len(text_xs) >= 4:
            left_n  = sum(1 for x in text_xs if x < col_split)
            right_n = sum(1 for x in text_xs if x >= col_split)
            is_two_col = (left_n >= 2 and right_n >= 2 and
                          min(left_n, right_n) / max(left_n, right_n) > 0.25)
        else:
            is_two_col = False

        col_ranges = ([(0, pw)] if not is_two_col
                      else [(0, col_split), (col_split, pw)])

        for col_x0, col_x1 in col_ranges:
            ax0   = int(col_x0 * aw / pw)
            ax1   = int(col_x1 * aw / pw)
            if ax1 <= ax0:
                continue
            col_w = ax1 - ax0

            MIN_GAP_ROWS = max(3, ah // 100)   # 최소 공백 높이 (~1% 페이지)
            MIN_SEG_FRAC = 0.035               # 최소 문항 높이 (3.5%)
            MARGIN_FRAC  = 0.025               # 머리말/꼬리말 무시 범위

            # 이 열의 행별 공백 여부
            row_white = []
            for y in range(ah):
                off = y * aw
                s   = sum(data[off + ax0: off + ax1])
                row_white.append(s / col_w >= white_thresh)

            # 연속 공백 밴드 수집
            gaps: list = []
            i = 0
            while i < ah:
                if row_white[i]:
                    j = i
                    while j < ah and row_white[j]:
                        j += 1
                    if j - i >= MIN_GAP_ROWS:
                        gaps.append((i / ah, j / ah))
                    i = j
                else:
                    i += 1

            # 공백 사이 콘텐츠 세그먼트
            segments: list = []
            prev = 0.0
            for g0, g1 in gaps:
                if g0 - prev >= MIN_SEG_FRAC:
                    segments.append((prev, g0))
                prev = g1
            if 1.0 - prev >= MIN_SEG_FRAC:
                segments.append((prev, 1.0))

            rx0 = max(0.0, col_x0 / pw - 0.005)
            rx1 = min(1.0, col_x1 / pw + 0.005)

            for ry0, ry1 in segments:
                h = ry1 - ry0
                if h < MIN_SEG_FRAC:
                    continue
                # 머리말/꼬리말 제외: 매우 작고 상단/하단 가장자리에 있는 것
                if h < MIN_SEG_FRAC * 2 and (ry0 < MARGIN_FRAC or ry1 > 1 - MARGIN_FRAC):
                    continue
                results.append((pi, rx0, ry0, rx1, ry1))

    return results


def _open_pdf_viewer(filepath: str):
    """PDF를 팝업 뷰어로 열어 사용자가 직접 문항·그래프를 크롭."""
    from PIL import ImageTk
    task_mode = multi_mode_var.get()
    style     = multi_style_var.get()
    extra_raw = multi_note_area.get("1.0", tk.END).strip()
    extra     = extra_raw if extra_raw and extra_raw != _MULTI_PLACEHOLDER else None

    try:
        doc = fitz.open(filepath)
    except Exception as e:
        messagebox.showerror("PDF 오류", str(e)); return

    # ── 팝업 창 ───────────────────────────────────────
    win = tk.Toplevel(root)
    win.title(f"📄 문항 크롭 — {os.path.basename(filepath)}")
    win.geometry("1500x900")
    win.configure(bg="#3C3C3C")
    win.minsize(1000, 600)

    # ── 상태 ───────────────────────────────────────────
    state = {
        "scale":      1.0,
        "crop_mode":  None,      # None | "question" | "graph"
        "drag_start": None,      # (canvas_x, canvas_y)
        "drag_rect":  None,      # temp canvas item id
        "selected":   None,      # index into crops (question only)
        "resize":     None,      # None | {"crop": c, "handle": "nw"|..., "pi": pi}
    }
    crops          = []   # [{"kind","idx","page_rect","crop_img","graph","q_type_var","list_frame"}, ...]
    page_size_cache = []  # [(base_w, base_h), ...] — populated first from page.rect (no pixel render)
    page_base_imgs  = []  # [PIL.Image | None, ...] — filled progressively after size scan
    page_photos    = []   # PhotoImage refs — only for visible pages
    page_offsets   = []   # [(canvas_x, canvas_y, disp_w, disp_h), ...]
    _redraw_timer  = [None]
    PAGE_GAP       = 12
    BASE_SCALE     = 2.5

    # ── 레이아웃 ────────────────────────────────────────
    hint_bar = tk.Frame(win, bg="#D0E4FF", pady=5, padx=10)
    hint_bar.pack(fill=tk.X, side=tk.TOP)
    tk.Label(hint_bar,
             text="C: 문항 크롭  |  G: 그래프 크롭  |  Esc: 취소  |  Ctrl/Cmd+휠: 줌  |  박스 클릭: 선택 후 핸들 드래그로 크기조정",
             font=("맑은 고딕", 10), bg="#D0E4FF", fg="#003399").pack(side=tk.LEFT)
    mode_lbl = tk.Label(hint_bar, text="",
                        font=("맑은 고딕", 11, "bold"), bg="#D0E4FF", fg="#CC3300")
    mode_lbl.pack(side=tk.LEFT, padx=20)
    # 줌 버튼
    tk.Button(hint_bar, text="＋", font=("맑은 고딕", 11, "bold"), bg="#D0E4FF",
              relief="flat", padx=6,
              command=lambda: (state.__setitem__("scale", min(6.0, state["scale"] * 1.2)), _redraw())
              ).pack(side=tk.RIGHT, padx=2)
    tk.Button(hint_bar, text="－", font=("맑은 고딕", 11, "bold"), bg="#D0E4FF",
              relief="flat", padx=6,
              command=lambda: (state.__setitem__("scale", max(0.2, state["scale"] / 1.2)), _redraw())
              ).pack(side=tk.RIGHT, padx=2)
    tk.Label(hint_bar, text="줌:", font=("맑은 고딕", 9), bg="#D0E4FF").pack(side=tk.RIGHT, padx=(8, 0))
    page_lbl = tk.Label(hint_bar, text="PDF 로딩 중…",
                        font=("맑은 고딕", 9), bg="#D0E4FF", fg="#555")
    page_lbl.pack(side=tk.RIGHT, padx=(0, 12))

    body = tk.PanedWindow(win, orient=tk.HORIZONTAL, sashwidth=6,
                          sashrelief="raised", bg="#3C3C3C")
    body.pack(fill=tk.BOTH, expand=True)

    # ── 좌: PDF Canvas ────────────────────────────────
    left_fr = tk.Frame(body, bg="#2A2A2A")
    body.add(left_fr, stretch="always", minsize=300)

    v_scroll = ttk.Scrollbar(left_fr, orient="vertical")
    h_scroll = ttk.Scrollbar(left_fr, orient="horizontal")
    pdf_canvas = tk.Canvas(left_fr, bg="#555", cursor="crosshair",
                            yscrollcommand=v_scroll.set,
                            xscrollcommand=h_scroll.set,
                            highlightthickness=0)
    v_scroll.config(command=pdf_canvas.yview)
    h_scroll.config(command=pdf_canvas.xview)
    v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
    pdf_canvas.pack(fill=tk.BOTH, expand=True)

    # ── 우: 크롭 목록 ─────────────────────────────────
    right_fr = tk.Frame(body, bg="#F0F4FF")
    body.add(right_fr, stretch="never", minsize=220)
    right_fr.grid_rowconfigure(1, weight=1)
    right_fr.grid_columnconfigure(0, weight=1)

    tk.Label(right_fr, text="📋 크롭된 문항 목록",
             font=("맑은 고딕", 11, "bold"),
             bg="#F0F4FF", fg="#003399").grid(row=0, column=0, columnspan=2, pady=(8, 4))

    list_canvas = tk.Canvas(right_fr, bg="#F0F4FF", highlightthickness=0)
    list_vs = ttk.Scrollbar(right_fr, orient="vertical", command=list_canvas.yview)
    list_canvas.config(yscrollcommand=list_vs.set)
    list_vs.grid(row=1, column=1, sticky="ns")
    list_canvas.grid(row=1, column=0, sticky="nsew", padx=(6, 0))

    list_inner = tk.Frame(list_canvas, bg="#F0F4FF")
    list_inner_id = list_canvas.create_window((0, 0), window=list_inner, anchor="nw")
    list_inner.bind("<Configure>",
                    lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all")))
    list_canvas.bind("<Configure>",
                     lambda e: list_canvas.itemconfig(list_inner_id, width=e.width))

    btn_bar = tk.Frame(right_fr, bg="#F0F4FF", pady=8, padx=6)
    btn_bar.grid(row=2, column=0, columnspan=2, sticky="ew")

    # ── PDF 렌더링 (스레드) ───────────────────────────
    def _render_pages():
        nonlocal page_size_cache, page_base_imgs
        try:
            total_p = len(doc)
            # Phase 1: read page sizes instantly (no pixel rendering) → show layout immediately
            sizes = []
            for i in range(total_p):
                r = doc.load_page(i).rect
                sizes.append((int(r.width * BASE_SCALE), int(r.height * BASE_SCALE)))
            page_size_cache = sizes
            page_base_imgs  = [None] * total_p
            root.after(0, _first_draw)
            # Phase 2: render pixels progressively, one page at a time
            for i in range(total_p):
                root.after(0, lambda n=i+1, t=total_p:
                            page_lbl.config(text=f"렌더링 {n}/{t}p…"))
                pix = doc.load_page(i).get_pixmap(
                    matrix=fitz.Matrix(BASE_SCALE, BASE_SCALE))
                page_base_imgs[i] = Image.open(BytesIO(pix.tobytes("png")))
                root.after(0, _schedule_redraw)
            root.after(0, lambda: page_lbl.config(
                text=f"총 {total_p}p  ·  Ctrl+휠: 줌"))
        except Exception as e:
            root.after(0, lambda em=str(e): messagebox.showerror("렌더링 오류", em))

    def _first_draw():
        page_lbl.config(text=f"총 {len(page_size_cache)}p  로딩 중…")
        _redraw()

    # ── Canvas 재그리기 ───────────────────────────────
    def _schedule_redraw():
        if _redraw_timer[0]:
            root.after_cancel(_redraw_timer[0])
        _redraw_timer[0] = root.after(80, _redraw)

    def _redraw():
        if not page_size_cache:
            return
        pdf_canvas.delete("all")
        page_photos.clear()
        page_offsets.clear()

        cw = max(pdf_canvas.winfo_width(), 600)
        y  = PAGE_GAP

        # Layout pass: use pre-calculated sizes (no image loading needed)
        for base_w, base_h in page_size_cache:
            disp_w = max(200, int((cw - 20) * state["scale"]))
            disp_h = int(base_h * disp_w / base_w)
            cx     = max(0, (cw - disp_w) // 2)
            page_offsets.append((cx, y, disp_w, disp_h))
            y += disp_h + PAGE_GAP

        total_w = max(cw, max(ox + ow for ox, _, ow, _ in page_offsets) + PAGE_GAP)
        pdf_canvas.configure(scrollregion=(0, 0, total_w, y))

        # Viewport culling: only create PhotoImage for visible pages
        vt  = pdf_canvas.canvasy(0)
        vb  = pdf_canvas.canvasy(max(pdf_canvas.winfo_height(), 100))
        BUF = 300  # pre-load 300px beyond visible area

        for i, (ox, oy, ow, oh) in enumerate(page_offsets):
            if oy + oh < vt - BUF or oy > vb + BUF:
                pdf_canvas.create_rectangle(ox, oy, ox + ow, oy + oh,
                                            fill="#666666", outline="", tags="page")
            else:
                base_img = page_base_imgs[i] if page_base_imgs else None
                if base_img is None:
                    pdf_canvas.create_rectangle(ox, oy, ox + ow, oy + oh,
                                                fill="#666666", outline="", tags="page")
                    pdf_canvas.create_text(ox + ow // 2, oy + oh // 2,
                                           text=f"{i+1}p", fill="#AAAAAA",
                                           font=("맑은 고딕", 12), tags="page")
                else:
                    resized = base_img.resize((ow, oh), Image.Resampling.LANCZOS)
                    photo   = ImageTk.PhotoImage(resized)
                    page_photos.append(photo)
                    pdf_canvas.create_image(ox, oy, anchor="nw", image=photo, tags="page")

        _draw_all_boxes()

    # ── 크롭 박스 오버레이 ────────────────────────────
    def _draw_all_boxes():
        pdf_canvas.delete("cropbox")
        sel_crop = None
        if state["selected"] is not None and state["selected"] < len(crops):
            sel_crop = crops[state["selected"]]
        for c in crops:
            _draw_one_box(c, is_selected=(c is sel_crop))

    def _draw_one_box(c, is_selected=False):
        pi, rx1, ry1, rx2, ry2 = c["page_rect"]
        if pi >= len(page_offsets):
            return
        ox, oy, ow, oh = page_offsets[pi]
        x1 = ox + int(rx1 * ow)
        y1 = oy + int(ry1 * oh)
        x2 = ox + int(rx2 * ow)
        y2 = oy + int(ry2 * oh)

        if c["kind"] == "question":
            color = "#00AA00" if is_selected else "#0066FF"
            lbl, lw = str(c["idx"]), 3
        else:
            color, lbl, lw = "#FF6600", f"G{c['parent_idx']}", 2

        pdf_canvas.create_rectangle(x1, y1, x2, y2,
                                    outline=color, width=lw, tags="cropbox")
        pdf_canvas.create_text(x1 + 5, y1 + 5, anchor="nw",
                               text=lbl, fill=color,
                               font=("맑은 고딕", 11, "bold"), tags="cropbox")
        # X 버튼 (우측 상단)
        bx, by = x2 - 2, y1 + 2
        ov = pdf_canvas.create_oval(bx - 13, by, bx + 1, by + 14,
                                    fill="#FF3333", outline="white", width=1,
                                    tags="cropbox")
        tx = pdf_canvas.create_text(bx - 6, by + 7, text="✕",
                                    fill="white", font=("맑은 고딕", 8, "bold"),
                                    tags="cropbox")
        pdf_canvas.tag_bind(ov, "<Button-1>", lambda e, crop=c: _remove_crop(crop))
        pdf_canvas.tag_bind(tx, "<Button-1>", lambda e, crop=c: _remove_crop(crop))

        # 리사이즈 핸들 (선택된 문항 박스만)
        if is_selected and c["kind"] == "question":
            mx, my = (x1 + x2) // 2, (y1 + y2) // 2
            hs = 5
            for hdir, hx, hy in [
                ("nw", x1, y1), ("n", mx, y1), ("ne", x2, y1),
                ("w",  x1, my),               ("e",  x2, my),
                ("sw", x1, y2), ("s", mx, y2), ("se", x2, y2),
            ]:
                pdf_canvas.create_rectangle(
                    hx - hs, hy - hs, hx + hs, hy + hs,
                    fill="white", outline="#00AA00", width=2,
                    tags=("cropbox", f"handle:{hdir}"))

    # ── 좌표 변환 ─────────────────────────────────────
    def _abs_xy(event):
        return pdf_canvas.canvasx(event.x), pdf_canvas.canvasy(event.y)

    def _page_rel(ax, ay):
        for i, (ox, oy, ow, oh) in enumerate(page_offsets):
            if ox <= ax < ox + ow and oy <= ay < oy + oh:
                return i, (ax - ox) / ow, (ay - oy) / oh
        return None, None, None

    # ── 크롭 확정 ─────────────────────────────────────
    def _finalize_crop(sx, sy, ex, ey):
        pi1, rx1, ry1 = _page_rel(sx, sy)
        pi2, rx2, ry2 = _page_rel(ex, ey)
        if pi1 is None or pi1 != pi2:
            return
        rx1, rx2 = sorted([rx1, rx2])
        ry1, ry2 = sorted([ry1, ry2])
        if (rx2 - rx1) < 0.005 or (ry2 - ry1) < 0.005:
            return
        base = page_base_imgs[pi1] if page_base_imgs else None
        if base is None:
            messagebox.showwarning("로딩 중", f"{pi1+1}페이지가 아직 로딩 중입니다. 잠시 후 다시 시도하세요.")
            return
        bw, bh = base.size
        crop_img = base.crop((int(rx1*bw), int(ry1*bh), int(rx2*bw), int(ry2*bh)))
        if state["crop_mode"] == "question":
            _do_q_crop(pi1, rx1, ry1, rx2, ry2, crop_img)
        elif state["crop_mode"] == "graph":
            _do_g_crop(pi1, rx1, ry1, rx2, ry2, crop_img)

    def _do_q_crop(pi, rx1, ry1, rx2, ry2, crop_img):
        idx = sum(1 for c in crops if c["kind"] == "question") + 1
        c = {
            "kind":       "question",
            "idx":        idx,
            "page_rect":  (pi, rx1, ry1, rx2, ry2),
            "crop_img":   crop_img,
            "graph":      None,
            "q_type_var": tk.StringVar(value="객관식"),
            "list_frame": None,
        }
        crops.append(c)
        state["selected"] = len(crops) - 1
        _build_list_item(c)
        _draw_all_boxes()
        _highlight_sel()

    def _do_g_crop(pi, rx1, ry1, rx2, ry2, crop_img):
        sel_idx = state["selected"]
        if sel_idx is None:
            messagebox.showwarning("그래프 크롭", "먼저 문항을 크롭하거나 목록에서 선택하세요.")
            return
        parent = crops[sel_idx]
        if parent["kind"] != "question":
            return
        parent["graph"] = {"page_rect": (pi, rx1, ry1, rx2, ry2), "crop_img": crop_img}
        # 기존 graph box 제거 후 추가
        crops[:] = [c for c in crops if not (c["kind"] == "graph"
                                              and c.get("parent_idx") == parent["idx"])]
        crops.append({
            "kind":       "graph",
            "parent_idx": parent["idx"],
            "page_rect":  (pi, rx1, ry1, rx2, ry2),
        })
        _refresh_list_item(parent)
        _draw_all_boxes()

    # ── 크롭 삭제 ─────────────────────────────────────
    def _remove_crop(c):
        if c["kind"] == "question":
            crops[:] = [x for x in crops
                        if not (x.get("parent_idx") == c["idx"] and x["kind"] == "graph")]
            if c["list_frame"]:
                c["list_frame"].destroy()
            crops.remove(c)
            # 번호 재정렬 + graph parent_idx 갱신
            old_to_new = {}
            qi = 1
            for x in crops:
                if x["kind"] == "question":
                    old_to_new[x["idx"]] = qi
                    x["idx"] = qi; qi += 1
            for x in crops:
                if x["kind"] == "graph" and x.get("parent_idx") in old_to_new:
                    x["parent_idx"] = old_to_new[x["parent_idx"]]
            _rebuild_list()
        elif c["kind"] == "graph":
            parent = next((x for x in crops if x["kind"] == "question"
                           and x["idx"] == c.get("parent_idx")), None)
            if parent:
                parent["graph"] = None
                _refresh_list_item(parent)
            crops.remove(c)
        if state["selected"] is not None and state["selected"] >= len(
                [x for x in crops if x["kind"] == "question"]):
            state["selected"] = None
        _draw_all_boxes()

    # ── 우측 목록 UI ──────────────────────────────────
    def _bind_scroll(widget):
        """자식 위젯까지 마우스 휠 이벤트를 list_canvas 스크롤로 연결."""
        widget.bind("<MouseWheel>", _on_list_wheel)
        for child in widget.winfo_children():
            _bind_scroll(child)

    def _build_list_item(c):
        is_typing = (task_mode == "타이핑")
        fr = tk.Frame(list_inner, bg="#FFFFFF", bd=1, relief="groove")
        fr.pack(fill=tk.X, padx=4, pady=3)
        c["list_frame"] = fr

        hdr = tk.Frame(fr, bg="#FFFFFF")
        hdr.pack(fill=tk.X, padx=4, pady=(4, 0))
        # 드래그 핸들
        dh = tk.Label(hdr, text="⠿", font=("맑은 고딕", 13),
                      bg="#FFFFFF", fg="#BBBBBB", cursor="fleur")
        dh.pack(side=tk.LEFT, padx=(0, 4))
        dh.bind("<ButtonPress-1>",   lambda e, crop=c: _dnd_start(e, crop))
        dh.bind("<B1-Motion>",       lambda e, crop=c: _dnd_motion(e, crop))
        dh.bind("<ButtonRelease-1>", lambda e, crop=c: _dnd_end(e, crop))
        tk.Label(hdr, text=f"문항 {c['idx']}",
                 font=("맑은 고딕", 10, "bold"), bg="#FFFFFF",
                 fg="#003399").pack(side=tk.LEFT)
        tk.Button(hdr, text="✕", font=("맑은 고딕", 8, "bold"),
                  bg="#FF4444", fg="white", relief="flat", bd=0,
                  padx=3, pady=0,
                  command=lambda crop=c: _remove_crop(crop)).pack(side=tk.RIGHT)

        thumb = c["crop_img"].copy()
        thumb.thumbnail((280, 180), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(thumb)
        img_lbl = tk.Label(fr, image=photo, bg="#FFFFFF")
        img_lbl.image = photo
        img_lbl.pack(padx=4, pady=4)

        if c.get("graph"):
            gth = c["graph"]["crop_img"].copy()
            gth.thumbnail((280, 140), Image.Resampling.LANCZOS)
            gp = ImageTk.PhotoImage(gth)
            gl = tk.Label(fr, image=gp, bg="#FFF0E0", bd=1, relief="solid")
            gl.image = gp
            gl.pack(padx=4, pady=(0, 2))
            tk.Label(fr, text="그래프 포함", font=("맑은 고딕", 8),
                     bg="#FFFFFF", fg="#FF6600").pack()

        if not is_typing:
            rf = tk.Frame(fr, bg="#FFFFFF")
            rf.pack(padx=4, pady=(0, 6))
            ttk.Radiobutton(rf, text="객관식",
                            variable=c["q_type_var"], value="객관식").pack(side=tk.LEFT, padx=4)
            ttk.Radiobutton(rf, text="주관식",
                            variable=c["q_type_var"], value="주관식").pack(side=tk.LEFT, padx=4)

        def _select_item(event, crop=c):
            try:
                state["selected"] = next(i for i, x in enumerate(crops) if x is crop)
            except StopIteration:
                pass
            _highlight_sel()

        for w in [fr] + fr.winfo_children():
            try: w.bind("<Button-1>", _select_item)
            except Exception: pass

        _bind_scroll(fr)
        return fr

    def _refresh_list_item(c):
        if c["list_frame"]:
            c["list_frame"].destroy()
            c["list_frame"] = None
        _rebuild_list()

    def _rebuild_list():
        for w in list_inner.winfo_children():
            w.destroy()
        for c in crops:
            if c["kind"] == "question":
                _build_list_item(c)
        _highlight_sel()

    def _highlight_sel():
        q_crops = [c for c in crops if c["kind"] == "question"]
        for i, c in enumerate(q_crops):
            if c["list_frame"] and c["list_frame"].winfo_exists():
                bg = "#E0ECFF" if crops.index(c) == state["selected"] else "#FFFFFF"
                try:
                    c["list_frame"].config(bg=bg)
                    for w in c["list_frame"].winfo_children():
                        w.config(bg=bg)
                except Exception:
                    pass

    # ── 드래그앤드롭 순서 변경 ────────────────────────
    _dnd = {"src": None, "indicator": None}

    def _dnd_indicator():
        if _dnd["indicator"] is None or not _dnd["indicator"].winfo_exists():
            _dnd["indicator"] = tk.Frame(list_inner, bg="#2266CC", height=3)
        return _dnd["indicator"]

    def _find_drop_idx(abs_screen_y):
        """절대 화면 y → q_crops 삽입 위치(0-based) 반환."""
        q_crops = [c for c in crops if c["kind"] == "question"]
        for i, c in enumerate(q_crops):
            if not (c["list_frame"] and c["list_frame"].winfo_exists()):
                continue
            fy = c["list_frame"].winfo_rooty()
            fh = c["list_frame"].winfo_height()
            if abs_screen_y < fy + fh // 2:
                return i
        return len(q_crops)

    def _dnd_start(event, crop):
        _dnd["src"] = crop

    def _dnd_motion(event, crop):
        if _dnd["src"] is None:
            return
        abs_y = event.widget.winfo_rooty() + event.y
        drop_i = _find_drop_idx(abs_y)
        q_crops = [c for c in crops if c["kind"] == "question"]
        ind = _dnd_indicator()
        try:
            ind.pack_forget()
            if drop_i == 0:
                ind.pack(fill=tk.X, padx=4, pady=0, before=q_crops[0]["list_frame"])
            elif drop_i >= len(q_crops):
                ind.pack(fill=tk.X, padx=4, pady=0, after=q_crops[-1]["list_frame"])
            else:
                ind.pack(fill=tk.X, padx=4, pady=0, before=q_crops[drop_i]["list_frame"])
        except Exception:
            pass

    def _dnd_end(event, crop):
        src = _dnd["src"]
        _dnd["src"] = None
        if _dnd["indicator"] and _dnd["indicator"].winfo_exists():
            _dnd["indicator"].pack_forget()
        if src is None:
            return
        abs_y = event.widget.winfo_rooty() + event.y
        drop_i = _find_drop_idx(abs_y)
        q_crops = [c for c in crops if c["kind"] == "question"]
        src_i = q_crops.index(src)
        # 같은 자리거나 바로 다음이면 변화 없음
        if drop_i == src_i or drop_i == src_i + 1:
            return
        q_crops.pop(src_i)
        if drop_i > src_i:
            drop_i -= 1
        q_crops.insert(drop_i, src)

        # crops 리스트 재구성 (그래프는 부모 뒤에 유지)
        new_crops = []
        for qc in q_crops:
            new_crops.append(qc)
            g = next((c for c in crops if c["kind"] == "graph"
                      and c.get("parent_idx") == qc["idx"]), None)
            if g:
                new_crops.append(g)
        crops[:] = new_crops

        # 번호 재정렬
        old_to_new: dict = {}
        qi = 1
        for c in crops:
            if c["kind"] == "question":
                old_to_new[c["idx"]] = qi
                c["idx"] = qi
                qi += 1
        for c in crops:
            if c["kind"] == "graph" and c.get("parent_idx") in old_to_new:
                c["parent_idx"] = old_to_new[c["parent_idx"]]

        _rebuild_list()
        _draw_all_boxes()

    # ── 키보드 단축키 ─────────────────────────────────
    def _on_key(event):
        k = event.keysym.lower()
        if k == "c":
            state["crop_mode"] = "question"
            mode_lbl.config(text="[ 📐 문항 크롭 — 드래그 ]", fg="#0055CC")
        elif k == "g":
            q_list = [c for c in crops if c["kind"] == "question"]
            if not q_list:
                messagebox.showwarning("그래프 크롭", "먼저 문항을 크롭하세요.")
                return
            if state["selected"] is None:
                state["selected"] = crops.index(q_list[-1])
            state["crop_mode"] = "graph"
            mode_lbl.config(text="[ 📊 그래프 크롭 — 드래그 ]", fg="#CC4400")
        elif k == "escape":
            state["crop_mode"] = None
            state["drag_start"] = None
            if state["drag_rect"]:
                pdf_canvas.delete(state["drag_rect"])
                state["drag_rect"] = None
            mode_lbl.config(text="")

    # ── 마우스 드래그 ─────────────────────────────────
    _HANDLE_HIT = 8  # 핸들 감지 픽셀 반경

    def _hit_handle(ax, ay):
        """선택된 박스의 핸들 방향 반환. 없으면 None."""
        si = state["selected"]
        if si is None or si >= len(crops):
            return None
        c = crops[si]
        if c["kind"] != "question":
            return None
        pi, rx1, ry1, rx2, ry2 = c["page_rect"]
        if pi >= len(page_offsets):
            return None
        ox, oy, ow, oh = page_offsets[pi]
        cx1, cy1 = ox + rx1 * ow, oy + ry1 * oh
        cx2, cy2 = ox + rx2 * ow, oy + ry2 * oh
        mx, my = (cx1 + cx2) / 2, (cy1 + cy2) / 2
        for hdir, hx, hy in [
            ("nw", cx1, cy1), ("n", mx, cy1), ("ne", cx2, cy1),
            ("w",  cx1, my),                  ("e",  cx2, my),
            ("sw", cx1, cy2), ("s", mx, cy2),  ("se", cx2, cy2),
        ]:
            if abs(ax - hx) <= _HANDLE_HIT and abs(ay - hy) <= _HANDLE_HIT:
                return hdir
        return None

    def _hit_crop(ax, ay):
        """클릭 좌표 안에 있는 question crop의 crops[] 인덱스 반환. 없으면 None."""
        for i, c in enumerate(crops):
            if c["kind"] != "question":
                continue
            pi, rx1, ry1, rx2, ry2 = c["page_rect"]
            if pi >= len(page_offsets):
                continue
            ox, oy, ow, oh = page_offsets[pi]
            if (ox + rx1 * ow <= ax <= ox + rx2 * ow and
                    oy + ry1 * oh <= ay <= oy + ry2 * oh):
                return i
        return None

    def _on_press(event):
        ax, ay = _abs_xy(event)

        # ① 핸들 클릭 → 리사이즈 모드
        hdir = _hit_handle(ax, ay)
        if hdir:
            si = state["selected"]
            c = crops[si]
            pi = c["page_rect"][0]
            state["resize"] = {"crop": c, "handle": hdir, "pi": pi}
            return

        # ② 박스 클릭 → 선택 (크롭 모드가 없을 때만)
        if state["crop_mode"] is None:
            ci = _hit_crop(ax, ay)
            if ci is not None:
                state["selected"] = ci
                _draw_all_boxes()
                _highlight_sel()
                return

        # ③ 새 크롭 드래그 시작
        if state["crop_mode"] is not None:
            state["drag_start"] = (ax, ay)

    def _on_motion(event):
        ax, ay = _abs_xy(event)

        # 리사이즈 중
        if state["resize"]:
            r = state["resize"]
            c = r["crop"]
            pi, rx1, ry1, rx2, ry2 = c["page_rect"]
            if pi >= len(page_offsets):
                return
            ox, oy, ow, oh = page_offsets[pi]
            new_rx = max(0.0, min(1.0, (ax - ox) / ow))
            new_ry = max(0.0, min(1.0, (ay - oy) / oh))
            hdir = r["handle"]
            MIN = 0.02
            if "w" in hdir:
                rx1 = min(new_rx, rx2 - MIN)
            if "e" in hdir:
                rx2 = max(new_rx, rx1 + MIN)
            if "n" in hdir:
                ry1 = min(new_ry, ry2 - MIN)
            if "s" in hdir:
                ry2 = max(new_ry, ry1 + MIN)
            c["page_rect"] = (pi, rx1, ry1, rx2, ry2)
            _draw_all_boxes()
            return

        # 새 크롭 드래그 중
        if state["drag_start"] is None:
            return
        sx, sy = state["drag_start"]
        if state["drag_rect"]:
            pdf_canvas.delete(state["drag_rect"])
        color = "#0066FF" if state["crop_mode"] == "question" else "#FF6600"
        state["drag_rect"] = pdf_canvas.create_rectangle(
            sx, sy, ax, ay, outline=color, width=2, dash=(4, 4))

    def _on_release(event):
        ax, ay = _abs_xy(event)

        # 리사이즈 완료 → crop_img 갱신
        if state["resize"]:
            r = state["resize"]
            c = r["crop"]
            pi, rx1, ry1, rx2, ry2 = c["page_rect"]
            if pi < len(page_base_imgs) and page_base_imgs[pi]:
                base = page_base_imgs[pi]
                bw, bh = base.size
                c["crop_img"] = base.crop((
                    int(rx1 * bw), int(ry1 * bh),
                    int(rx2 * bw), int(ry2 * bh)))
                _refresh_list_item(c)
            state["resize"] = None
            _draw_all_boxes()
            return

        # 새 크롭 드래그 완료
        if state["drag_start"] is None:
            return
        if state["drag_rect"]:
            pdf_canvas.delete(state["drag_rect"])
            state["drag_rect"] = None
        sx, sy = state["drag_start"]
        state["drag_start"] = None
        _finalize_crop(sx, sy, ax, ay)
        state["crop_mode"] = None
        mode_lbl.config(text="")

    # ── 스크롤 / 줌 ───────────────────────────────────
    def _on_mousewheel(event):
        if event.state & 0x4:  # Ctrl → 줌
            factor = 1.12 if event.delta > 0 else (1.0 / 1.12)
            state["scale"] = max(0.2, min(6.0, state["scale"] * factor))
            _redraw()
        else:
            pdf_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            _schedule_redraw()

    def _on_list_wheel(event):
        list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ── 이벤트 바인딩 ─────────────────────────────────
    win.bind("<Key>", _on_key)
    pdf_canvas.bind("<ButtonPress-1>",   _on_press)
    pdf_canvas.bind("<B1-Motion>",       _on_motion)
    pdf_canvas.bind("<ButtonRelease-1>", _on_release)
    pdf_canvas.bind("<MouseWheel>",      _on_mousewheel)
    v_scroll.config(command=lambda *a: (pdf_canvas.yview(*a), _schedule_redraw()))
    pdf_canvas.bind("<Configure>",
                    lambda e: _redraw() if page_size_cache else None)
    body.bind("<ButtonRelease-1>",
              lambda e: _redraw() if page_size_cache else None)
    list_canvas.bind("<MouseWheel>", _on_list_wheel)
    list_inner.bind("<MouseWheel>",  _on_list_wheel)

    # ── 하단 버튼 ─────────────────────────────────────
    def _collect():
        result = []
        for c in crops:
            if c["kind"] != "question":
                continue
            g_img  = c["graph"]["crop_img"] if c["graph"] else None
            q_type = c["q_type_var"].get() if task_mode == "워싱" else "객관식"
            result.append((c["crop_img"], str(c["idx"]), q_type, g_img))
        return result

    def _start_immediate():
        items = _collect()
        if not items:
            messagebox.showwarning("없음", "크롭된 문항이 없습니다."); return
        win.destroy()
        threading.Thread(
            target=_run_selected_analysis,
            args=(items, task_mode, False),
            kwargs={"style": style, "extra_note": extra,
                    "number_style": number_style_var.get()},
            daemon=True).start()

    def _start_batch():
        items = _collect()
        if not items:
            messagebox.showwarning("없음", "크롭된 문항이 없습니다."); return
        label = simpledialog.askstring(
            "배치 이름 설정",
            f"이 배치 작업의 이름을 입력하세요.\n({len(items)}문항 / {task_mode} / {style or '강사용'})\n"
            "(취소하면 제출이 취소됩니다)",
            parent=win)
        if label is None:
            return
        win.destroy()
        def _sub():
            def _st(msg): root.after(0, lambda m=msg: status_label.config(text=m, fg="#FF8800"))
            def _dn(job_info, err):
                if err:
                    root.after(0, lambda em=err: messagebox.showerror("배치 오류", em))
                    root.after(0, lambda: status_label.config(text="❌ 배치 제출 실패", fg="#FF0000"))
                else:
                    root.after(0, lambda: status_label.config(
                        text=f"✅ 배치 제출 완료 ({len(items)}문항) — 나중에 결과 불러오기",
                        fg="#008000"))
            submit_wash_batch(items, task_mode, style or "강사용", extra or "", _st, _dn,
                              label=label.strip() or "이름없음",
                              number_style=number_style_var.get())
        threading.Thread(target=_sub, daemon=True).start()

    _crop_busy = [False]  # 자동/AI 크롭 중복 실행 방지

    def _apply_detected(detected, label):
        added = 0
        for pi, rx1, ry1, rx2, ry2 in detected:
            if pi >= len(page_base_imgs) or page_base_imgs[pi] is None:
                continue
            base = page_base_imgs[pi]
            bw, bh = base.size
            crop_img = base.crop((int(rx1*bw), int(ry1*bh),
                                  int(rx2*bw), int(ry2*bh)))
            _do_q_crop(pi, rx1, ry1, rx2, ry2, crop_img)
            added += 1
        mode_lbl.config(text=f"✅ {label} {added}개 감지 완료", fg="#006600")
        win.after(3000, lambda: mode_lbl.config(text=""))

    def _run_auto_detect():
        if _crop_busy[0]: return
        total_pages = len(doc)
        not_loaded  = (len(page_base_imgs) < total_pages or
                       any(img is None for img in page_base_imgs))
        if not_loaded:
            messagebox.showwarning("로딩 중",
                "아직 렌더링 중인 페이지가 있습니다.\n"
                "잠시 후 다시 시도하세요.", parent=win)
            return
        _crop_busy[0] = True
        mode_lbl.config(text="⏳ 자동 감지 중…", fg="#884400")
        win.update_idletasks()

        try:
            detected = _auto_detect_crops(doc, page_imgs=page_base_imgs)
            if not detected:
                mode_lbl.config(text="")
                messagebox.showinfo("자동 감지",
                    "감지된 문항이 없습니다.\n수동으로 크롭하거나 AI 크롭을 사용하세요.", parent=win)
                return
            _apply_detected(detected, "자동")
        finally:
            _crop_busy[0] = False

    def _run_yolo_detect():
        if _crop_busy[0]: return
        total_pages = len(doc)
        not_loaded  = (len(page_base_imgs) < total_pages or
                       any(img is None for img in page_base_imgs))
        if not_loaded:
            messagebox.showwarning("로딩 중",
                "아직 렌더링 중인 페이지가 있습니다.\n"
                "잠시 후 다시 시도하세요.", parent=win)
            return

        def _run():
            _crop_busy[0] = True
            win.after(0, lambda: mode_lbl.config(text="⏳ Gemini 감지 중…", fg="#440088"))
            try:
                def _prog(pi):
                    win.after(0, lambda p=pi: mode_lbl.config(
                        text=f"⏳ Gemini 감지 중… {p+1}/{len(doc)}페이지", fg="#440088"))
                detected = _detect_crops_gemini(doc, page_base_imgs, progress_cb=_prog)
                if not detected:
                    win.after(0, lambda: (
                        mode_lbl.config(text=""),
                        messagebox.showinfo("AI 크롭", "감지 결과가 없습니다.", parent=win)))
                    return
                win.after(0, lambda d=detected: _apply_detected(d, "AI"))
            except Exception as e:
                em = str(e)
                win.after(0, lambda m=em: (
                    mode_lbl.config(text="❌ AI 오류", fg="#CC0000"),
                    messagebox.showerror("AI 크롭 오류", m, parent=win)))
            finally:
                _crop_busy[0] = False

        threading.Thread(target=_run, daemon=True).start()

    tk.Button(btn_bar, text="🤖 AI 크롭",
              command=_run_yolo_detect,
              font=("맑은 고딕", 10, "bold"), bg="#F0E8FF",
              relief="raised", bd=2, padx=10, pady=4).pack(side=tk.LEFT, padx=4)

    if task_mode == "워싱":
        def _set_all_type(qtype):
            for c in crops:
                if c["kind"] == "question":
                    c["q_type_var"].set(qtype)
        tk.Button(btn_bar, text="전체 객관식",
                  command=lambda: _set_all_type("객관식"),
                  font=("맑은 고딕", 9), bg="#E0E8FF",
                  relief="raised", bd=2, padx=8, pady=3).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_bar, text="전체 주관식",
                  command=lambda: _set_all_type("주관식"),
                  font=("맑은 고딕", 9), bg="#E0E8FF",
                  relief="raised", bd=2, padx=8, pady=3).pack(side=tk.LEFT, padx=3)

    tk.Button(btn_bar, text="📦 배치 제출 (50% 할인)",
              command=_start_batch,
              font=("맑은 고딕", 10, "bold"), bg="#FFF0B0",
              relief="raised", bd=2, padx=10, pady=4).pack(side=tk.RIGHT, padx=4)
    tk.Button(btn_bar, text="🚀 즉시 분석",
              command=_start_immediate,
              font=("맑은 고딕", 11, "bold"), bg="#A0C9FF",
              relief="raised", bd=3, padx=14, pady=6).pack(side=tk.RIGHT, padx=4)

    threading.Thread(target=_render_pages, daemon=True).start()


# ============================================================
# 8. 이미지 장바구니
# ============================================================
def add_to_stack():
    img = ImageGrab.grabclipboard()
    if not isinstance(img, Image.Image):
        messagebox.showwarning("주의", "클립보드에 이미지가 없습니다!\nWin+Shift+S 캡처 후 다시 누르세요.")
        return
    image_stack.append(img.copy())
    _refresh_stack_label()
    status_label.config(text=f"✅ {len(image_stack)}번째 이미지 추가됨", fg="#008000")


def analyze_stack():
    """장바구니 이미지들에서 문항 감지 후 썸네일 다이얼로그 호출."""
    global stop_flag
    if not image_stack:
        messagebox.showwarning("주의", "장바구니가 비어 있습니다!"); return
    task_mode = multi_mode_var.get()
    style     = multi_style_var.get()
    extra_raw = multi_note_area.get("1.0", tk.END).strip()
    extra     = extra_raw if extra_raw and extra_raw != _MULTI_PLACEHOLDER else None
    stop_flag = False
    images    = list(image_stack)
    threading.Thread(target=_detect_and_show_thumbnails,
                     args=(images, task_mode, style, extra, True), daemon=True).start()


# ============================================================
# ★ 썸네일 선택 다이얼로그 (PDF/장바구니 공통)
# ============================================================
def _detect_and_show_thumbnails(images: list, task_mode: str,
                                 style: str = None, extra_note: str = None,
                                 clear_stack_after: bool = False):
    items = [(img, img, f"문제{idx+1}") for idx, img in enumerate(images)]
    root.after(0, lambda its=items, tm=task_mode, st=style, en=extra_note, cs=clear_stack_after:
               _show_thumbnail_dialog(its, tm, st, en, cs))


def _show_thumbnail_dialog(items: list, task_mode: str,
                            style: str = None, extra_note: str = None,
                            clear_stack_after: bool = False):
    """
    items: [(full_img, crop_img, question_num), ...]
    각 항목마다 크롭된 썸네일 + 체크박스 + 객관식/주관식 라디오 표시.
    """
    win = tk.Toplevel(root)
    win.title("🎯 문항별 설정")
    win.geometry("1400x900")
    win.configure(bg="#F8F9FB")
    win.minsize(1400, 800)

    # ── 상단: 안내 + 일괄 선택 ────────────────────────────
    top_bar = tk.Frame(win, bg="#E8F0FF", pady=10, padx=12)
    top_bar.pack(fill=tk.X)
    tk.Label(top_bar,
             text=f"총 {len(items)}개 문항 감지  |  모드: {task_mode}",
             font=("맑은 고딕", 12, "bold"), bg="#E8F0FF",
             fg="#003399").pack(side=tk.LEFT)

    # 일괄 버튼
    bulk_fr = tk.Frame(win, bg="#F8F9FB", pady=8)
    bulk_fr.pack(fill=tk.X, padx=12)
    tk.Label(bulk_fr, text="일괄 적용:", font=("맑은 고딕", 10, "bold"),
             bg="#F8F9FB").pack(side=tk.LEFT, padx=(0, 8))

    # 각 항목의 상태 저장 (checkbox_var, type_var)
    states = []  # [(included_var, type_var), ...]
    graph_imgs = [None] * len(items)  # 문항별 그래프 크롭 이미지

    def _open_graph_crop_win(src_img, idx, preview_lbl):
        """src_img 위에서 드래그 크롭 → graph_imgs[idx] 저장 + preview_lbl 업데이트."""
        cwin = tk.Toplevel(win)
        cwin.title("📐 그래프 영역 드래그 선택")
        cwin.grab_set()

        MAX_W, MAX_H = 1400, 900
        sw, sh = src_img.size
        scale = min(MAX_W / sw, MAX_H / sh, 1.0)
        disp_w, disp_h = int(sw * scale), int(sh * scale)
        disp_img = src_img.resize((disp_w, disp_h), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(disp_img)

        cwin.geometry(f"{disp_w + 20}x{disp_h + 110}")
        info_lbl = tk.Label(cwin, text="드래그로 그래프 영역을 선택하세요. 재드래그로 수정 가능.",
                 font=("맑은 고딕", 10), pady=4)
        info_lbl.pack()

        cv = tk.Canvas(cwin, width=disp_w, height=disp_h, cursor="crosshair",
                       highlightthickness=0)
        cv.pack()
        cv.create_image(0, 0, anchor="nw", image=photo)
        cv._photo = photo  # GC 방지

        rect_id = [None]
        start = [0, 0]
        pending = [None]  # 드래그 후 저장 전 임시 이미지

        def _press(e):
            start[0], start[1] = e.x, e.y
            if rect_id[0]:
                cv.delete(rect_id[0])
                rect_id[0] = None

        def _drag(e):
            if rect_id[0]:
                cv.delete(rect_id[0])
            rect_id[0] = cv.create_rectangle(
                start[0], start[1], e.x, e.y,
                outline="#FF6600", width=2, dash=(4, 2))

        def _release(e):
            x0 = min(start[0], e.x); y0 = min(start[1], e.y)
            x1 = max(start[0], e.x); y1 = max(start[1], e.y)
            if x1 - x0 < 5 or y1 - y0 < 5:
                return
            ox0, oy0 = int(x0 / scale), int(y0 / scale)
            ox1, oy1 = int(x1 / scale), int(y1 / scale)
            raw = src_img.crop((ox0, oy0, ox1, oy1))
            pending[0] = _process_graph_image(raw) if task_mode == "타이핑" else raw
            info_lbl.config(
                text=f"선택됨 ({ox1-ox0}×{oy1-oy0}px). 재드래그로 수정, [저장]으로 확정.")
            save_btn.config(state=tk.NORMAL)

        def _save():
            if pending[0] is None:
                return
            graph_imgs[idx] = pending[0]
            pw, ph = 160, 100
            prev = pending[0].copy()
            prev.thumbnail((pw, ph), Image.Resampling.LANCZOS)
            p = ImageTk.PhotoImage(prev)
            preview_lbl.config(image=p, text="", relief="solid", bd=1)
            preview_lbl._photo = p
            thumb_photos.append(p)
            cwin.destroy()

        def _clear():
            graph_imgs[idx] = None
            preview_lbl.config(image="", text="없음", relief="flat")
            preview_lbl._photo = None
            cwin.destroy()

        cv.bind("<ButtonPress-1>", _press)
        cv.bind("<B1-Motion>", _drag)
        cv.bind("<ButtonRelease-1>", _release)
        btn_fr = tk.Frame(cwin)
        btn_fr.pack(pady=6)
        save_btn = tk.Button(btn_fr, text="저장", command=_save,
                  font=("맑은 고딕", 10), bg="#CCFFCC",
                  relief="raised", bd=2, padx=10, pady=4, state=tk.DISABLED)
        save_btn.pack(side=tk.LEFT, padx=4)
        tk.Button(btn_fr, text="그래프 없음 (제거)", command=_clear,
                  font=("맑은 고딕", 10), bg="#FFCCCC",
                  relief="raised", bd=2, padx=8, pady=4).pack(side=tk.LEFT, padx=4)

    def _bulk(include=None, qtype=None):
        for iv, tv in states:
            if include is not None: iv.set(include)
            if qtype is not None:   tv.set(qtype)

    is_typing = (task_mode == "타이핑")

    if not is_typing:
        tk.Button(bulk_fr, text="전체 객관식", font=("맑은 고딕", 10),
                  bg="#DDEEFF", relief="raised", bd=2, padx=8, pady=2,
                  command=lambda: _bulk(qtype="객관식")).pack(side=tk.LEFT, padx=3)
        tk.Button(bulk_fr, text="전체 주관식", font=("맑은 고딕", 10),
                  bg="#FFE8DD", relief="raised", bd=2, padx=8, pady=2,
                  command=lambda: _bulk(qtype="주관식")).pack(side=tk.LEFT, padx=3)
        tk.Frame(bulk_fr, width=20, bg="#F8F9FB").pack(side=tk.LEFT)
    tk.Button(bulk_fr, text="전체 포함", font=("맑은 고딕", 10),
              bg="#DDFFDD", relief="raised", bd=2, padx=8, pady=2,
              command=lambda: _bulk(include=True)).pack(side=tk.LEFT, padx=3)
    tk.Button(bulk_fr, text="전체 제외", font=("맑은 고딕", 10),
              bg="#FFDDDD", relief="raised", bd=2, padx=8, pady=2,
              command=lambda: _bulk(include=False)).pack(side=tk.LEFT, padx=3)

    # ── 중앙: 스크롤 가능한 썸네일 리스트 ─────────────────
    canvas_fr = tk.Frame(win, bg="#F8F9FB")
    canvas_fr.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

    canvas = tk.Canvas(canvas_fr, bg="#FFFFFF", highlightthickness=1,
                        highlightbackground="#CCCCCC")
    scrollbar = ttk.Scrollbar(canvas_fr, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    list_fr = tk.Frame(canvas, bg="#FFFFFF")
    list_fr_id = canvas.create_window((0, 0), window=list_fr, anchor="nw")

    def _on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    list_fr.bind("<Configure>", _on_frame_configure)

    def _on_canvas_configure(event):
        canvas.itemconfig(list_fr_id, width=event.width)
    canvas.bind("<Configure>", _on_canvas_configure)

    # 마우스 휠 스크롤
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # ── 문항별 행 생성 (세로 구조: 썸네일 위 / 컨트롤 아래) ──────────
    # 썸네일은 전체 페이지 이미지(full_img) — AI 전송은 여전히 crop_img
    thumb_photos   = []
    img_label_data = []  # [(label, full_img), ...] — 창 확정 후 리사이즈용
    win.thumb_photos = thumb_photos

    from PIL import ImageTk

    _s = ttk.Style()
    for _name, _bg in [("Even", "#FFFFFF"), ("Odd", "#F5F8FC")]:
        _s.configure(f"{_name}Check.TCheckbutton", background=_bg, font=("맑은 고딕", 12), indicatorsize=20)
        _s.configure(f"{_name}Radio.TRadiobutton", background=_bg, font=("맑은 고딕", 12), indicatorsize=20)
        _s.map(f"{_name}Check.TCheckbutton", background=[("active", _bg)])
        _s.map(f"{_name}Radio.TRadiobutton", background=[("active", _bg)])

    for i, (full_img, crop_img, q) in enumerate(items):
        row_bg = "#FFFFFF" if i % 2 == 0 else "#F5F8FC"

        # 구분선 (첫 항목 제외)
        if i > 0:
            tk.Frame(list_fr, bg="#BBBBBB", height=2).pack(fill=tk.X)

        row = tk.Frame(list_fr, bg=row_bg, pady=12, padx=20)
        row.pack(fill=tk.X)

        # ① 문항 번호 레이블
        q_label = f"문항 {q}" if q else f"{i+1}번"
        tk.Label(row, text=q_label,
                 font=("맑은 고딕", 14, "bold"),
                 bg=row_bg, fg="#003399").pack(anchor="w", pady=(0, 6))

        # ② 썸네일 — 전체 페이지 이미지, 초기 1360px로 렌더링 후 창 크기 확정 시 업데이트
        src_w, src_h = full_img.size
        init_w = 1360
        init_h = int(src_h * init_w / src_w)
        init_thumb = full_img.resize((init_w, init_h), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(init_thumb)
        thumb_photos.append(photo)

        img_lbl = tk.Label(row, image=photo, bg=row_bg,
                           relief="solid", bd=1, anchor="w")
        img_lbl.pack(fill=tk.X, pady=(0, 10))
        img_label_data.append((img_lbl, full_img))

        # ③ 컨트롤 영역 (가로 나열)
        ctrl_fr = tk.Frame(row, bg=row_bg)
        ctrl_fr.pack(fill=tk.X)

        # 포함 체크박스
        included = tk.BooleanVar(value=True)
        _prefix = "Even" if i % 2 == 0 else "Odd"
        ttk.Checkbutton(ctrl_fr, text="✔ 분석 포함",
                        variable=included,
                        style=f"{_prefix}Check.TCheckbutton").pack(side=tk.LEFT, padx=(0, 32))

        # 유형 라디오 (타이핑 모드에서는 숨김)
        q_type = tk.StringVar(value="객관식")
        if not is_typing:
            ttk.Radiobutton(ctrl_fr, text="객관식", variable=q_type, value="객관식",
                            style=f"{_prefix}Radio.TRadiobutton").pack(side=tk.LEFT, padx=(0, 16))
            ttk.Radiobutton(ctrl_fr, text="주관식", variable=q_type, value="주관식",
                            style=f"{_prefix}Radio.TRadiobutton").pack(side=tk.LEFT)

        # 그래프 크롭 버튼 + 미리보기 (타이핑/워싱 공통)
        graph_sep = tk.Frame(ctrl_fr, width=30, bg=row_bg)
        graph_sep.pack(side=tk.LEFT)
        preview_lbl = tk.Label(ctrl_fr, text="없음", font=("맑은 고딕", 9),
                               bg=row_bg, fg="#888888", width=12, height=4,
                               anchor="center")
        preview_lbl.pack(side=tk.LEFT, padx=(0, 6))
        _idx = i  # 클로저 캡처용
        tk.Button(ctrl_fr, text="📐 그래프 크롭",
                  font=("맑은 고딕", 10), bg="#FFF0CC",
                  relief="raised", bd=2, padx=8, pady=2,
                  command=lambda si=full_img, ii=_idx, pl=preview_lbl:
                      _open_graph_crop_win(si, ii, pl)).pack(side=tk.LEFT)

        states.append((included, q_type))

    # 창이 그려진 뒤 실제 창 너비로 리사이즈
    def _resize_to_win():
        win_w = win.winfo_width()
        if win_w <= 1:
            win.after(100, _resize_to_win)
            return
        new_w = win_w - 40
        for lbl, img in img_label_data:
            sw, sh = img.size
            new_h = int(sh * new_w / sw)
            t = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            p = ImageTk.PhotoImage(t)
            thumb_photos.append(p)
            lbl.config(image=p)
            lbl._photo = p

    win.after(150, _resize_to_win)

    # ── 하단: 실행 버튼 ───────────────────────────────────
    btn_bar = tk.Frame(win, bg="#F8F9FB", pady=12)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _get_selected():
        sel = []
        for i, ((full_img, crop_img, q), (iv, tv)) in enumerate(zip(items, states)):
            if iv.get():
                g = graph_imgs[i]
                if g is not None:
                    sel.append((crop_img, q, tv.get(), g))
                else:
                    sel.append((crop_img, q, tv.get()))
        return sel

    def _start():
        selected = _get_selected()
        if not selected:
            messagebox.showwarning("선택 없음", "포함된 문항이 없습니다."); return
        canvas.unbind_all("<MouseWheel>")
        win.destroy()
        root.after(0, lambda t=len(selected): status_label.config(
            text=f"🤖 분석 중… (0/{t})", fg="#0000FF"))
        threading.Thread(target=_run_selected_analysis,
                         args=(selected, task_mode, clear_stack_after),
                         kwargs={"style": style, "extra_note": extra_note,
                                 "number_style": number_style_var.get()},
                         daemon=True).start()

    def _batch():
        selected = _get_selected()
        if not selected:
            messagebox.showwarning("선택 없음", "포함된 문항이 없습니다."); return
        label = simpledialog.askstring(
            "배치 이름 설정",
            f"이 배치 작업의 이름을 입력하세요.\n({len(selected)}문항 / {task_mode} / {style or '강사용'})\n"
            "(취소하면 제출이 취소됩니다)",
            parent=win)
        if label is None:
            return
        canvas.unbind_all("<MouseWheel>")
        win.destroy()
        def _submit():
            def _status(msg):
                root.after(0, lambda m=msg: status_label.config(text=m, fg="#FF8800"))
            def _done(job_info, err):
                if err:
                    root.after(0, lambda em=err: messagebox.showerror("배치 오류", em))
                    root.after(0, lambda: status_label.config(text="❌ 배치 제출 실패", fg="#FF0000"))
                else:
                    root.after(0, lambda: status_label.config(
                        text=f"✅ 배치 제출 완료 ({len(selected)}문항) — 나중에 결과 불러오기 가능",
                        fg="#008000"))
            submit_wash_batch(selected, task_mode, style or "강사용",
                              extra_note or "", _status, _done,
                              label=label.strip() or "이름없음",
                              number_style=number_style_var.get())
        threading.Thread(target=_submit, daemon=True).start()

    def _cancel():
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    tk.Button(btn_bar, text="취소", command=_cancel,
              font=("맑은 고딕", 10), bg="#FFCCCC",
              relief="raised", bd=2, padx=16, pady=5).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_bar, text="📦 배치 제출 (50% 할인)",
              command=_batch, font=("맑은 고딕", 11, "bold"),
              bg="#FFF0B0", relief="raised", bd=2,
              padx=20, pady=7).pack(side=tk.RIGHT, padx=6)
    tk.Button(btn_bar, text="🚀 즉시 분석",
              command=_start, font=("맑은 고딕", 12, "bold"),
              bg="#A0C9FF", relief="raised", bd=3,
              padx=24, pady=8).pack(side=tk.RIGHT, padx=6)


def _run_selected_analysis(selected: list, task_mode: str,
                            clear_stack_after: bool,
                            style: str = None, extra_note: str = None,
                            number_style: str = "원본"):
    """selected: [(img, q, q_type) 또는 (img, q, q_type, graph_img), ...]"""
    global stop_flag
    stop_flag = False
    root.after(0, lambda: input_area.delete("1.0", tk.END))
    total = len(selected)
    done_count = [0]
    results = {}
    next_out = [1]
    lock = threading.Lock()

    def _flush_ordered():
        while next_out[0] in results:
            r, raw_text = results.pop(next_out[0])
            if r:
                save_ai_output_cache(r, style, raw_text=raw_text)
                root.after(0, lambda rv=r: input_area.insert(tk.END, rv + "\n\n"))
                root.after(0, _update_counter)
            next_out[0] += 1

    def _process_one(idx: int, item):
        if stop_flag:
            return idx, None, None
        img, q, q_type = item[0], item[1], item[2]
        graph_img = item[3] if len(item) > 3 else None
        q_label = q if q else f"{idx}번"
        try:
            result, raw_text = _call_gemini(img, task_mode, q_type, target_q=q,
                                            style=style, extra_note=extra_note,
                                            graph_img=graph_img, number_style=number_style)
            if not _is_valid_result(result, style=style):
                if result:
                    r = "[⚠️ 형식 오류 — 검토 필요]\n" + result
                    if number_style == "순서":
                        r = f"{idx}. " + r
                    return idx, r, raw_text
                return idx, None, None
            if number_style == "순서" and result:
                result = f"{idx}. " + result.strip()
            if number_style == "없음" and style != "해설없음" and result:
                result = result.replace("</미주>", "</미주>\n")
            return idx, result, raw_text
        except Exception as e:
            err = str(e)
            root.after(0, lambda lb=q_label, em=err: input_area.insert(
                tk.END, f"[{lb} 실패: {em}]\n\n"))
            return idx, None, None

    with ThreadPoolExecutor(max_workers=_BATCH_WORKERS) as executor:
        futures = {
            executor.submit(_process_one, i, item): i
            for i, item in enumerate(selected, 1)
        }
        for future in as_completed(futures):
            idx, result, raw_text = future.result()
            with lock:
                done_count[0] += 1
                d = done_count[0]
                results[idx] = (result, raw_text)
                root.after(0, lambda d=d, t=total: status_label.config(
                    text=f"🤖 배치 분석 중 ({d}/{t} 완료)…", fg="#0000FF"))
                _flush_ordered()

    if clear_stack_after and not stop_flag:
        image_stack.clear()
        root.after(0, _refresh_stack_label)

    if stop_flag:
        root.after(0, lambda: status_label.config(text="🛑 분석 중단됨", fg="#FF0000"))
    else:
        root.after(0, lambda t=total: status_label.config(
            text=f"✅ 분석 완료! ({t}문항)", fg="#008000"))


def clear_stack():
    image_stack.clear()
    _refresh_stack_label()


def _refresh_stack_label():
    n = len(image_stack)
    stack_label.config(text=f"📦 {n}장", fg="#CC00CC" if n > 0 else "#777777")


# ============================================================
# 9. 형식 자동검사
# ============================================================

def _show_latex_leak_dialog(msg: str, latex_leaks: list):
    """LaTeX 잔재 발견 시 — 검사 결과 + 'Claude에게 복사' 버튼 다이얼로그."""
    win = tk.Toplevel(root)
    win.title("⚠️ 형식 검사 — LaTeX 잔재 발견")
    win.geometry("620x460")
    win.configure(bg="#FFFEF0")
    win.grab_set()

    tk.Label(win, text="⚠️ 변환 누락된 LaTeX 패턴이 있습니다.",
             font=("맑은 고딕", 11, "bold"), bg="#FFFEF0", fg="#CC4400").pack(pady=(14, 4))

    area = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("맑은 고딕", 10),
                                     bg="#FFFFF8", relief="sunken", bd=1, height=12)
    area.pack(fill=tk.BOTH, expand=True, padx=14, pady=6)
    area.insert(tk.END, msg)
    area.config(state=tk.DISABLED)

    def _copy_for_claude():
        # 누락 패턴 목록 정리
        leak_lines = "\n".join(
            f"  수식 {i}: {' / '.join(p)}" for i, p in latex_leaks)
        clipboard_text = (
            "[v28 LaTeX 변환 누락 보고]\n\n"
            f"누락된 LaTeX 패턴 ({len(latex_leaks)}건):\n"
            f"{leak_lines}\n\n"
            "위 패턴들을 latex_to_hwp() 변환 테이블에 추가해주세요.\n"
            "각 패턴의 HWP 수식 문법 대응값을 알려주시면 바로 적용하겠습니다."
        )
        pyperclip.copy(clipboard_text)
        copy_btn.config(text="✅ 복사됨!", bg="#AAFFAA")
        win.after(2000, lambda: copy_btn.config(text="📋 Claude에게 복사", bg="#FFE8A0"))

    btn_row = tk.Frame(win, bg="#FFFEF0")
    btn_row.pack(pady=8)

    copy_btn = tk.Button(btn_row, text="📋 Claude에게 복사",
                         command=_copy_for_claude,
                         font=("맑은 고딕", 10, "bold"),
                         bg="#FFE8A0", relief="raised", bd=2, padx=14, pady=5)
    copy_btn.pack(side=tk.LEFT, padx=6)

    tk.Button(btn_row, text="닫기", command=win.destroy,
              font=("맑은 고딕", 10), bg="#FFCCCC",
              relief="raised", bd=2, padx=14, pady=5).pack(side=tk.LEFT, padx=6)


def validate_format(silent=False):
    text = input_area.get("1.0", tk.END).strip()
    if not text:
        if not silent: messagebox.showwarning("검사", "검사할 내용이 없습니다.")
        return False, [], []

    errors, warnings_list = [], []

    open_cnt  = text.count("<미주>")
    close_cnt = text.count("</미주>")
    if open_cnt != close_cnt:
        errors.append(f"<미주> 열기 {open_cnt}개 vs 닫기 {close_cnt}개 — 불균형")

    pyo_open  = text.count("<표>")
    pyo_close = text.count("</표>")
    if pyo_open != pyo_close:
        errors.append(f"<표> 열기 {pyo_open}개 vs 닫기 {pyo_close}개 — 불균형")

    eq_cnt = len(re.findall(r'==', text))
    if eq_cnt % 2 != 0:
        block_pattern = re.split(r'(?=\n\s*\d{1,2}<미주>)', text)
        cumulative, culprit = 0, None
        for seg in block_pattern:
            seg_eq = len(re.findall(r'==', seg))
            if (cumulative + seg_eq) % 2 != 0 and culprit is None:
                m = re.search(r'(\d{1,2})<미주>', seg)
                culprit = f"문항 {m.group(1)}" if m else "첫 번째 블록"
            cumulative += seg_eq
        hint = f" (최초 홀수 전환: {culprit})" if culprit else ""
        errors.append(f"== 태그 홀수({eq_cnt}개) — 수식 짝 안 맞음{hint}")

    minju_blocks = re.findall(r'<미주>(.*?)</미주>', text, re.DOTALL)
    for i, block in enumerate(minju_blocks, 1):
        if not block.strip():  # 해설없음 모드의 빈 미주 자리표시자 — 검사 생략
            continue
        if '[정답]' not in block: errors.append(f"{i}번째 미주: [정답] 없음")
        if '[해설]' not in block: warnings_list.append(f"{i}번째 미주: [해설] 없음")

    if single_type_var.get() == "주관식":
        answers = re.findall(r'\[정답\]\s*==\s*(.*?)\s*==', text)
        for i, ans in enumerate(answers, 1):
            clean = ans.strip().replace('`', '').replace(' ', '')
            try:
                val = int(clean)
                if not (1 <= val <= 999):
                    errors.append(f"{i}번 정답 '{clean}': 1~999 범위 초과")
            except ValueError:
                warnings_list.append(f"{i}번 정답 '{clean}': 자연수 불명확 (수동 확인)")

    text_outside_eq = re.sub(r'==.*?==', '', text, flags=re.DOTALL)
    bt_cnt = text_outside_eq.count('`')
    if bt_cnt > 0:
        warnings_list.append(f"수식(==) 바깥에 백틱(`) {bt_cnt}개 발견")

    # LaTeX 잔재 감지 (변환 누락 패턴)
    latex_leaks = []
    eq_contents = re.findall(r'==(.*?)==', text, re.DOTALL)
    for i, eq in enumerate(eq_contents, 1):
        found = re.findall(r'\\[a-zA-Z]+', eq)
        if found:
            latex_leaks.append((i, list(dict.fromkeys(found))))  # 중복 제거, 순서 유지
    if latex_leaks:
        leak_summary = ", ".join(
            f"수식{i}({' '.join(p)})" for i, p in latex_leaks[:5])
        if len(latex_leaks) > 5:
            leak_summary += f" 외 {len(latex_leaks)-5}건"
        warnings_list.append(f"LaTeX 잔재 {len(latex_leaks)}개 수식 — 변환 누락: {leak_summary}")

    if not silent:
        lines = [f"📋 문항 수: {open_cnt}개\n"]
        if errors:
            lines.append("【오류】\n")
            lines += [f"  ❌ {e}" for e in errors]
        if warnings_list:
            lines.append("\n【경고】\n")
            lines += [f"  ⚠️ {w}" for w in warnings_list]
        if not errors and not warnings_list:
            lines.append("✅ 이상 없음! 바로 입력 가능합니다.")
        msg = "\n".join(lines)

        # LaTeX 잔재가 있으면 복사 버튼 포함 커스텀 다이얼로그
        if latex_leaks:
            _show_latex_leak_dialog(msg, latex_leaks)
        elif errors:
            messagebox.showwarning("형식 검사 결과", msg)
        else:
            messagebox.showinfo("형식 검사 결과", msg)

    return len(errors) == 0, errors, warnings_list


# ============================================================
# 10. 카운터 & 저장
# ============================================================
def _update_counter():
    count = input_area.get("1.0", tk.END).count("<미주>")
    counter_label.config(
        text=f"📝 {count}문항",
        fg="#0055AA" if count > 0 else "#888888"
    )


def save_result_txt():
    text = input_area.get("1.0", tk.END).strip()
    if not text:
        messagebox.showwarning("주의", "저장할 내용이 없습니다."); return
    filepath = filedialog.asksaveasfilename(
        title="결과 저장", defaultextension=".txt",
        filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")])
    if not filepath: return
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(text)
        messagebox.showinfo("저장 완료", f"저장되었습니다:\n{filepath}")
    except Exception as e:
        messagebox.showerror("저장 오류", str(e))


# ============================================================
# ★ 풀이 저장 / 뷰어 Toplevel
# ============================================================
def open_solution_viewer():
    """[📝 풀이 저장 / 뷰어 열기] 버튼 핸들러."""
    global _solution_viewer_ref

    # 이미 열려있으면 앞으로
    if _solution_viewer_ref[0] is not None:
        try:
            _solution_viewer_ref[0].lift()
            _solution_viewer_ref[0].focus_set()
            # 최신 결과로 편집창 업데이트
            if _last_latex_result[0]:
                _push_result_to_solution_viewer(_last_latex_result[0])
            return
        except tk.TclError:
            _solution_viewer_ref[0] = None

    from PIL import ImageTk as _ITk

    win = tk.Toplevel(root)
    _solution_viewer_ref[0] = win
    win.title("📝 풀이 저장  v1.3  |  LaTeX 편집 → HWP / 파인튜닝")

    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    ww, wh = min(1300, sw - 60), min(920, sh - 60)
    win.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")
    win.minsize(900, 650)
    win.configure(bg="#F5F7FF")
    win.grid_rowconfigure(0, weight=1)
    win.grid_rowconfigure(1, weight=0)
    win.grid_columnconfigure(0, weight=1)

    vert_pw = tk.PanedWindow(win, orient=tk.VERTICAL,
                             bg="#AAAAAA", sashwidth=6, sashrelief="raised")
    vert_pw.grid(row=0, column=0, sticky="nsew", padx=10, pady=(4, 0))
    win.after(100, lambda: vert_pw.sash_place(0, 0, 280))

    def _on_close():
        _solution_viewer_ref[0] = None
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", _on_close)

    # ── 뷰어 내부 상태 ───────────────────────────────────────
    _vphoto   = [None]
    _vpending = [None]
    _vrunning = [False]
    _vdebid   = [None]
    _curr_id  = tk.StringVar(value=_last_problem_id[0])
    _curr_prob = [_last_problem_text[0]]

    # ── 상단: 문제 불러오기 ──────────────────────────────────
    TOP_BG = "#EEF2FF"
    top_fr = tk.LabelFrame(vert_pw, text="  📂  문제 불러오기  ",
                            font=("맑은 고딕", 10, "bold"),
                            bg=TOP_BG, padx=10, pady=6, relief="groove", bd=2)
    top_fr.grid_columnconfigure(1, weight=1)
    top_fr.grid_rowconfigure(1, weight=1)
    vert_pw.add(top_fr, minsize=150)

    tk.Label(top_fr, text="ID:", bg=TOP_BG,
             font=("맑은 고딕", 9, "bold")).grid(row=0, column=0, padx=(0, 4), sticky="w")
    tk.Entry(top_fr, textvariable=_curr_id, font=("맑은 고딕", 10),
             width=20, relief="sunken", bd=1).grid(row=0, column=1, sticky="w", padx=(0, 6))

    prev_wrap = tk.Frame(top_fr, bg="#F8F8F8", relief="sunken", bd=1)
    prev_wrap.grid(row=1, column=1, columnspan=4, sticky="nsew", pady=(6, 0))
    prev_wrap.grid_columnconfigure(0, weight=1)
    prev_wrap.grid_rowconfigure(0, weight=1)
    prev_sb = ttk.Scrollbar(prev_wrap, orient=tk.VERTICAL)
    prev_sb.grid(row=0, column=1, sticky="ns")
    prob_preview = tk.Canvas(prev_wrap, bg="#FEFEFE", highlightthickness=0,
                              yscrollcommand=prev_sb.set)
    prob_preview.grid(row=0, column=0, sticky="nsew")
    prev_sb.config(command=prob_preview.yview)
    prob_preview.bind("<MouseWheel>", lambda e: prob_preview.yview_scroll(
        int(-1*(e.delta/120)), "units"))
    _prev_photo = [None]

    def _refresh_prob_preview():
        text = _curr_prob[0]
        prob_preview.delete("all")
        if not text:
            return
        if not _MPL_AVAILABLE:
            prob_preview.create_text(8, 8, anchor="nw", text=text,
                                      font=("맑은 고딕", 11), fill="#333")
            return
        def _render():
            try:
                w = max(prob_preview.winfo_width() or 500, 500)
                img = _render_latex_image(text, width_px=w, dpi=110)
                def _apply(i=img):
                    photo = _ITk.PhotoImage(i)
                    _prev_photo[0] = photo
                    prob_preview.delete("all")
                    prob_preview.create_image(4, 4, anchor="nw", image=photo)
                    prob_preview.configure(scrollregion=(0, 0, i.width+8, i.height+8))
                root.after(0, _apply)
            except Exception as ex:
                em = str(ex)
                root.after(0, lambda: (
                    prob_preview.delete("all"),
                    prob_preview.create_text(8, 8, anchor="nw", text=em,
                        font=("맑은 고딕", 9), fill="#CC0000")))
        threading.Thread(target=_render, daemon=True).start()

    def _load_by_id():
        iid   = _curr_id.get().strip()
        entry = next((e for e in _load_ai_cache() if e.get("id") == iid), None)
        if not entry:
            messagebox.showwarning("없음", f"ID '{iid}' 없음", parent=win)
            return
        # 뷰어 캔버스: 본문+보기만 표시 (미주·번호 제거된 ai_output 우선)
        ao = entry.get("ai_output", "").strip()
        if ao:
            _curr_prob[0] = ao
        else:
            raw = entry.get("raw_output", "")
            _curr_prob[0] = (_extract_ai_output_body(raw, entry.get("style", "강사용"))
                             if raw else "")
        _refresh_prob_preview()
        ans = entry.get("answer", "")
        viewer_status.config(
            text=f"✅ 로드: {iid}" + (f"  정답: {ans}" if ans else ""),
            fg="#006600")

    def _open_history_picker():
        items = _load_ai_cache()
        if not items:
            messagebox.showinfo("캐쉬 없음", "저장된 분석 캐쉬가 없습니다.", parent=win)
            return

        hw = tk.Toplevel(win); hw.title("AI 캐쉬 선택"); hw.geometry("1300x680")
        hw.grab_set()
        hw.grid_rowconfigure(0, weight=1)
        hw.grid_rowconfigure(1, weight=0)
        hw.grid_columnconfigure(0, weight=1)

        _cs = ttk.Style()
        _cs.configure("Cache.Treeview", rowheight=80, font=("맑은 고딕", 10))
        _cs.configure("Cache.Treeview.Heading", font=("맑은 고딕", 9, "bold"))

        cols = ("ID", "날짜", "스타일", "정답", "문제 본문 미리보기")
        main_pw = tk.PanedWindow(hw, orient=tk.HORIZONTAL,
                                 bg="#AAAAAA", sashwidth=6, sashrelief="raised")
        main_pw.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))
        tree_fr = tk.Frame(main_pw)
        tree_fr.grid_rowconfigure(0, weight=1)
        tree_fr.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(tree_fr, columns=cols, show="headings",
                            style="Cache.Treeview", selectmode="browse")
        tree.heading("ID",           text="ID")
        tree.heading("날짜",         text="날짜")
        tree.heading("스타일",       text="스타일")
        tree.heading("정답",         text="정답")
        tree.heading("문제 본문 미리보기", text="문제 본문 미리보기")
        tree.column("ID",           width=130, stretch=False)
        tree.column("날짜",         width=80,  stretch=False)
        tree.column("스타일",       width=72,  stretch=False)
        tree.column("정답",         width=58,  stretch=False)
        tree.column("문제 본문 미리보기", width=500, stretch=True)

        def _wrap(text, w=62, lines=3):
            text = text.replace('\n', ' ').strip()
            out = []
            while text and len(out) < lines:
                if len(text) <= w:
                    out.append(text); break
                cut = text[:w].rfind(' ')
                if cut < 8: cut = w
                out.append(text[:cut]); text = text[cut:].lstrip()
            return '\n'.join(out)

        def _entry_preview(it: dict) -> str:
            """캐쉬 엔트리에서 태그 없는 본문 미리보기 반환."""
            ao = it.get("ai_output", "").strip()
            if ao:
                return _wrap(ao)
            # 구버전 엔트리: raw_output에서 미주 제거
            raw = it.get("raw_output", "")
            if raw:
                clean = _extract_ai_output_body(raw, it.get("style", "강사용"))
                return _wrap(clean or raw)
            return ""

        for it in reversed(items):
            tree.insert("", tk.END, iid=it.get("id",""), values=(
                it.get("id",""), it.get("date",""),
                it.get("style",""), it.get("answer",""), _entry_preview(it)))

        sb = ttk.Scrollbar(tree_fr, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        main_pw.add(tree_fr, minsize=350)

        # ── 라텍스 미리보기 패널 ─────────────────────────────
        prev_lf = tk.LabelFrame(main_pw, text="  수식 미리보기 (선택 시 자동 렌더링)  ",
                                font=("맑은 고딕", 9), bg="#F4F6FF", bd=1, relief="groove")
        prev_lf.grid_rowconfigure(0, weight=1)
        prev_lf.grid_columnconfigure(0, weight=1)
        _pk_canvas = tk.Canvas(prev_lf, bg="#FAFBFF", highlightthickness=0)
        _pk_canvas.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        _pk_photo  = [None]
        main_pw.add(prev_lf, minsize=300)

        def _render_pk_preview(text: str):
            if not _MPL_AVAILABLE or not text.strip():
                return
            def _r():
                try:
                    hw.update_idletasks()
                    w = max(_pk_canvas.winfo_width() or 860, 300)
                    img = _render_latex_image(text, width_px=w, dpi=100)
                    def _apply(i=img):
                        photo = _ITk.PhotoImage(i)
                        _pk_photo[0] = photo
                        _pk_canvas.delete("all")
                        _pk_canvas.create_image(0, 0, anchor="nw", image=photo)
                    root.after(0, _apply)
                except Exception:
                    pass
            threading.Thread(target=_r, daemon=True).start()

        def _on_tree_select(event):
            sel = tree.selection()
            if not sel:
                return
            iid = str(tree.item(sel[0])["values"][0])
            entry = next((e for e in items if e.get("id") == iid), None)
            if not entry:
                return
            ao = entry.get("ai_output", "").strip()
            if not ao:
                raw = entry.get("raw_output", "")
                ao = _extract_ai_output_body(raw, entry.get("style", "강사용")) if raw else ""
            if ao:
                _render_pk_preview(ao)

        tree.bind("<<TreeviewSelect>>", _on_tree_select)

        btn_fr = tk.Frame(hw)
        btn_fr.grid(row=1, column=0, pady=(6, 8))

        def _pick():
            sel = tree.selection()
            if not sel: return
            _curr_id.set(str(tree.item(sel[0])["values"][0]))
            _load_by_id(); hw.destroy()

        def _delete():
            sel = tree.selection()
            if not sel: return
            item_id = str(tree.item(sel[0])["values"][0])
            if not messagebox.askyesno("삭제 확인", f"'{item_id}' 항목을 캐쉬에서 삭제할까요?",
                                       parent=hw):
                return
            with _cache_write_lock:
                db = _load_ai_cache()
                db = [e for e in db if e.get("id") != item_id]
                _save_ai_cache_db(db)
            tree.delete(sel[0])

        tk.Button(btn_fr, text="✅ 선택", command=_pick,
                  font=("맑은 고딕", 10, "bold"),
                  bg="#4F6EF7", fg="white", relief="flat", padx=16, pady=4).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_fr, text="🗑 삭제", command=_delete,
                  font=("맑은 고딕", 10),
                  bg="#FFCCCC", relief="flat", padx=16, pady=4).pack(side=tk.LEFT, padx=6)

    tk.Button(top_fr, text="🔍 ID로 불러오기", command=_load_by_id,
              font=("맑은 고딕", 9), bg="#D8E8FF", relief="raised", bd=2,
              padx=8).grid(row=0, column=2, padx=3)
    tk.Button(top_fr, text="📋 캐쉬 목록", command=_open_history_picker,
              font=("맑은 고딕", 9), bg="#DCF0DC", relief="raised", bd=2,
              padx=8).grid(row=0, column=3, padx=3)

    # 최신 ID가 있으면 자동 로드
    if _curr_id.get():
        win.after(150, _load_by_id)
    tk.Label(top_fr, text="문제 미리보기:", bg=TOP_BG,
             font=("맑은 고딕", 9, "bold")).grid(row=1, column=0, sticky="nw", pady=(6, 0))

    # ── 중단: 손글씨 업로드 행 ───────────────────────────────
    MID_BG = "#FAFBFF"
    mid_fr = tk.Frame(vert_pw, bg=MID_BG)
    vert_pw.add(mid_fr, minsize=200)
    mid_fr.grid_rowconfigure(1, weight=1)
    mid_fr.grid_columnconfigure(0, weight=1)

    upload_row = tk.Frame(mid_fr, bg=MID_BG, relief="groove", bd=1)
    upload_row.grid(row=0, column=0, sticky="ew", pady=(0, 4))

    tk.Label(upload_row, text="✋ 손글씨 이미지:", bg=MID_BG,
             font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=10)
    img_label_v = tk.Label(upload_row, text="(이미지 없음)", bg=MID_BG,
                            fg="#999", font=("맑은 고딕", 9))
    img_label_v.pack(side=tk.LEFT, padx=6)

    _cur_img = [None]

    def _load_file_v():
        fp = filedialog.askopenfilename(parent=win, title="손글씨 이미지 선택",
                filetypes=[("이미지","*.png *.jpg *.jpeg *.bmp"),("모든 파일","*.*")])
        if not fp: return
        _cur_img[0] = Image.open(fp)
        img_label_v.config(text=f"✅ {Path(fp).name}", fg="#006600")

    def _load_clip_v():
        try:
            img = ImageGrab.grabclipboard()
            if img is None:
                messagebox.showwarning("클립보드", "클립보드에 이미지가 없습니다.", parent=win)
                return
            _cur_img[0] = img
            img_label_v.config(text="✅ 클립보드 이미지", fg="#006600")
        except Exception as e:
            messagebox.showerror("오류", str(e), parent=win)

    def _run_flash_v():
        if _cur_img[0] is None:
            messagebox.showwarning("이미지 없음", "먼저 이미지를 선택하세요.", parent=win)
            return
        viewer_status.config(text="⏳ Flash 변환 중…", fg="#CC4400")
        conv_btn_v.config(state="disabled")
        def _w():
            try:
                latex = handwriting_to_latex(_cur_img[0])
                root.after(0, lambda lt=latex: _set_editor(lt))
            except Exception as e:
                em = str(e)
                root.after(0, lambda em=em: (
                    messagebox.showerror("변환 오류", em, parent=win),
                    viewer_status.config(text="❌ 변환 실패", fg="#CC0000")))
            finally:
                root.after(0, lambda: conv_btn_v.config(state="normal"))
        threading.Thread(target=_w, daemon=True).start()

    tk.Button(upload_row, text="📁 파일 선택", command=_load_file_v,
              font=("맑은 고딕", 9), bg="#DDEEFF", relief="raised", bd=2,
              padx=8).pack(side=tk.LEFT, padx=3, pady=6)
    tk.Button(upload_row, text="📋 클립보드", command=_load_clip_v,
              font=("맑은 고딕", 9), bg="#DDEEFF", relief="raised", bd=2,
              padx=8).pack(side=tk.LEFT, padx=3, pady=6)
    conv_btn_v = tk.Button(upload_row, text="🔄 Flash 변환", command=_run_flash_v,
                           font=("맑은 고딕", 10, "bold"), bg="#A0C9FF",
                           relief="raised", bd=2, padx=12, pady=2)
    conv_btn_v.pack(side=tk.LEFT, padx=(8, 3), pady=6)

    # ── 좌우 PanedWindow ─────────────────────────────────────
    from tkinter import scrolledtext as _stmod
    lr = tk.PanedWindow(mid_fr, orient=tk.HORIZONTAL,
                        bg="#CCCCCC", sashwidth=6, sashrelief="raised")
    lr.grid(row=1, column=0, sticky="nsew")

    ed_fr = tk.LabelFrame(lr, text="  ✏️  LaTeX 편집창  ",
                           font=("맑은 고딕", 10, "bold"), bg=MID_BG,
                           relief="groove", bd=2)
    latex_ed = _stmod.ScrolledText(ed_fr, wrap=tk.WORD,
                                   font=("Consolas", 16), bg="#FEFFFE", relief="flat")
    latex_ed.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
    lr.add(ed_fr, minsize=300)

    vw_fr = tk.LabelFrame(lr, text="  📐  수식 뷰어  (실시간 렌더링)  ",
                           font=("맑은 고딕", 10, "bold"), bg=MID_BG,
                           relief="groove", bd=2)
    vw_sb = ttk.Scrollbar(vw_fr, orient=tk.VERTICAL)
    vw_sb.pack(side=tk.RIGHT, fill=tk.Y)
    vw_canvas = tk.Canvas(vw_fr, bg="#FAFBFF", yscrollcommand=vw_sb.set,
                           highlightthickness=0)
    vw_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    vw_sb.config(command=vw_canvas.yview)
    vw_canvas.bind("<MouseWheel>", lambda e: vw_canvas.yview_scroll(
        int(-1 * (e.delta / 120)), "units"))
    if not _MPL_AVAILABLE:
        vw_canvas.create_text(10, 10, anchor='nw',
            text="⚠️ matplotlib 미설치\npip install matplotlib",
            fill="#AA4400", font=("맑은 고딕", 11))
    lr.add(vw_fr, minsize=300)

    # ── 뷰어 업데이트 함수 (클로저) ──────────────────────────
    def _apply_img(img):
        photo = _ITk.PhotoImage(img)
        _vphoto[0] = photo
        vw_canvas.delete("all")
        vw_canvas.create_image(0, 0, anchor='nw', image=photo)
        vw_canvas.configure(scrollregion=(0, 0, img.width, img.height))

    def _render_loop_v():
        while True:
            text = _vpending[0]
            if text is None:
                _vrunning[0] = False
                return
            _vpending[0] = None
            try:
                w = max(vw_canvas.winfo_width(), 400)
                img = _render_latex_image(text, width_px=w)
                root.after(0, lambda i=img: _apply_img(i))
            except Exception as e:
                em = str(e)
                root.after(0, lambda em=em: (
                    vw_canvas.delete("all"),
                    vw_canvas.create_text(10, 10, anchor='nw',
                        text=f"렌더링 오류:\n{em}", fill="#CC0000",
                        font=("맑은 고딕", 9))
                ))
            time.sleep(0.05)

    def _update_viewer_v(latex_text: str):
        if not _MPL_AVAILABLE:
            return
        _vpending[0] = latex_text
        if not _vrunning[0]:
            _vrunning[0] = True
            threading.Thread(target=_render_loop_v, daemon=True).start()

    def _set_editor(text: str):
        latex_ed.delete("1.0", tk.END)
        latex_ed.insert(tk.END, text)
        _update_viewer_v(text)
        viewer_status.config(text="✅ 변환 완료 — 우측 뷰어 확인 후 수정하세요", fg="#006600")

    def _on_editor_change(event=None):
        if _vdebid[0]: root.after_cancel(_vdebid[0])
        _vdebid[0] = root.after(500, lambda: _update_viewer_v(
            latex_ed.get("1.0", tk.END).rstrip("\n")))

    latex_ed.bind("<KeyRelease>", _on_editor_change)

    # 창 노출 후 초기 렌더링
    def _initial_v(ev=None):
        win.unbind("<Map>")
        _update_viewer_v("$f(x) = \\dfrac{1}{2}x^{2} + 3x - 1$")
        if _last_latex_result[0]:
            _curr_prob[0] = _last_latex_result[0]
            win.after(300, _refresh_prob_preview)
    win.bind("<Map>", _initial_v)

    def _update_preview_fn(text: str):
        _curr_prob[0] = text
        _refresh_prob_preview()

    # 창 속성에 참조 저장 (push 연동용)
    win._latex_editor      = latex_ed
    win._update_viewer_fn  = _update_viewer_v
    win._update_preview_fn = _update_preview_fn

    # ── 하단: 과목 선택 + 버튼 ──────────────────────────────
    BOT_BG = "#F0FFF4"
    bot_fr = tk.Frame(win, bg=BOT_BG, relief="groove", bd=1)
    bot_fr.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 8))

    course_row_v = tk.Frame(bot_fr, bg=BOT_BG)
    course_row_v.pack(fill=tk.X, padx=10, pady=(8, 4))
    tk.Label(course_row_v, text="과목:", bg=BOT_BG,
             font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))

    course_var_v = tk.StringVar(value="수능")
    ttk.Combobox(course_row_v, textvariable=course_var_v,
                 values=list(SUBJECT_INDEX.keys()), width=14,
                 state="readonly").pack(side=tk.LEFT, padx=(0, 10))

    unit_frame_v = tk.Frame(course_row_v, bg=BOT_BG)
    unit_frame_v.pack(side=tk.LEFT, fill=tk.X, expand=True)
    _unit_vars_v = {}; _unit_cbs_v = []

    def _refresh_units_v(*_):
        for w in _unit_cbs_v: w.destroy()
        _unit_cbs_v.clear(); _unit_vars_v.clear()
        course = course_var_v.get()
        if course not in SUBJECT_INDEX: return
        for uname in SUBJECT_INDEX[course]["units"]:
            var = tk.BooleanVar(value=False)
            cb  = tk.Checkbutton(unit_frame_v, text=uname, variable=var,
                                  bg=BOT_BG, font=("맑은 고딕", 8), anchor="w")
            cb.pack(side=tk.LEFT, padx=2)
            _unit_vars_v[uname] = var; _unit_cbs_v.append(cb)

    subj_preview_v = tk.Label(course_row_v, text="S", bg=BOT_BG,
                               fg="#4F6EF7", font=("맑은 고딕", 9, "bold"))
    subj_preview_v.pack(side=tk.LEFT, padx=(8, 0))

    def _upd_preview_v(*_):
        codes = build_subject_codes_list(course_var_v, _unit_vars_v)
        subj_preview_v.config(text=f"코드: {', '.join(codes)}")

    def _refresh_and_trace(*a):
        _refresh_units_v(*a)
        for var in _unit_vars_v.values():
            var.trace_add("write", lambda *_: root.after(50, _upd_preview_v))
        _upd_preview_v()

    course_var_v.trace_add("write", _refresh_units_v)
    course_var_v.trace_add("write", _refresh_and_trace)
    _refresh_and_trace()

    # HWP 상태 행
    hwp_row_v = tk.Frame(bot_fr, bg="#F0F0F0", relief="groove", bd=1)
    hwp_row_v.pack(fill=tk.X, padx=10, pady=(0, 4))
    hwp_st_lbl = tk.Label(hwp_row_v,
        text="🔴 pywin32 없음" if not _WIN32_AVAILABLE else
             ("🟢 HWP 연결됨" if _com_alive else "🔴 HWP 미연결"),
        font=("맑은 고딕", 9), bg="#F0F0F0",
        fg="#006600" if (_WIN32_AVAILABLE and _com_alive) else "#CC0000")
    hwp_st_lbl.pack(side=tk.LEFT, padx=10, pady=4)

    def _btn_hwp_start_v():
        if not _WIN32_AVAILABLE:
            messagebox.showwarning("pywin32 없음", "pip install pywin32", parent=win)
            return
        ok, msg = com_start_hwp()
        hwp_st_lbl.config(
            text="🟢 HWP 연결됨" if ok else "🔴 HWP 미연결",
            fg="#006600" if ok else "#CC0000")
        if not ok:
            messagebox.showwarning("실패", msg, parent=win)

    tk.Button(hwp_row_v, text="📄 HWP 시작", command=_btn_hwp_start_v,
              font=("맑은 고딕", 9, "bold"), bg="#C8F0C8",
              relief="raised", bd=2, padx=10).pack(side=tk.RIGHT, padx=6, pady=3)

    # 버튼 행
    viewer_btn_row = tk.Frame(bot_fr, bg=BOT_BG)
    viewer_btn_row.pack(pady=(2, 4))
    viewer_status = tk.Label(bot_fr, text="준비 완료",
                             font=("맑은 고딕", 9), fg="#0033AA", bg=BOT_BG)
    viewer_status.pack(pady=(0, 6))

    def _save_finetune_v():
        sol = latex_ed.get("1.0", tk.END).strip()
        if not sol:
            messagebox.showwarning("없음", "풀이가 없습니다.", parent=win); return
        iid = _curr_id.get().strip()
        if not iid:
            messagebox.showwarning("ID 없음",
                "분석 후 자동으로 ID가 설정됩니다.\n"
                "또는 캐쉬 목록에서 문항을 선택하세요.", parent=win); return
        subj_list = build_subject_codes_list(course_var_v, _unit_vars_v)
        save_finetune_new(iid, subj_list, sol)
        viewer_status.config(
            text=f"💾 저장 완료 — ID:{iid}  subject:{subj_list}", fg="#006600")
        messagebox.showinfo("저장 완료",
            f"ID: {iid}\nSubject: {subj_list}\n파일: {_FINETUNE_DB}", parent=win)

    def _hwp_input_v():
        global stop_flag, _hwp_marshal_stream
        sol = latex_ed.get("1.0", tk.END).strip()
        if not sol:
            messagebox.showwarning("없음", "입력할 내용이 없습니다.", parent=win); return
        if not _WIN32_AVAILABLE:
            messagebox.showwarning("pywin32 없음", "HWP 입력은 pywin32 필요", parent=win)
            return
        if not _com_alive:
            ok, msg = com_start_hwp()
            if not ok:
                if not messagebox.askyesno("HWP 없음", f"{msg}\n\npyautogui 모드로 진행?",
                                           parent=win):
                    return
                if not activate_hwp():
                    messagebox.showerror("오류", "한글 창을 찾을 수 없습니다.", parent=win)
                    return
        if _WIN32_AVAILABLE and _com_alive and _hwp_com is not None:
            try:
                _hwp_marshal_stream = pythoncom.CoMarshalInterThreadInterfaceInStream(
                    pythoncom.IID_IDispatch, _hwp_com)
            except Exception:
                _hwp_marshal_stream = None
        stop_flag = False
        viewer_status.config(text="⏳ HWP 입력 중…", fg="#CC4400")
        converted = _convert_equations_in_result(sol)
        threading.Thread(target=process_text, args=(converted,), daemon=True).start()

    tk.Button(viewer_btn_row, text="💾 파인튜닝 데이터 저장",
              command=_save_finetune_v,
              font=("맑은 고딕", 11, "bold"), bg="#FFE066",
              relief="raised", bd=2, height=2, padx=20).pack(side=tk.LEFT, padx=6)
    tk.Button(viewer_btn_row, text="  📝  HWP 변환 및 입력  ",
              command=_hwp_input_v,
              font=("맑은 고딕", 11, "bold"), bg="#4F6EF7", fg="white",
              relief="raised", bd=2, height=2, padx=20).pack(side=tk.LEFT, padx=6)
    tk.Button(viewer_btn_row, text="■ 중단",
              command=lambda: globals().__setitem__('stop_flag', True),
              font=("맑은 고딕", 10), bg="#FFCCCC",
              relief="raised", bd=2, height=2, width=7).pack(side=tk.LEFT, padx=6)


# ============================================================
# 11. GUI
# ============================================================
BG        = "#F8F9FB"
BG_SINGLE = "#EEF4FF"   # 클립보드 단일 패널
BG_MULTI  = "#F0FFF4"   # PDF/장바구니 패널
BG_NOTE   = "#FFFDF0"

# HWP 입력 속도 상수 (UI 슬라이더 제거, 고정값)
window_delay_var = type('_V', (), {'get': lambda s: 0.8})()
input_delay_var  = type('_V', (), {'get': lambda s: 1.0})()

root = tk.Tk()
root.title(f"수학워싱 v28  |  {MODEL_PRO} / {MODEL_FLASH}  |  LaTeX 변환")

screen_w = root.winfo_screenwidth()
screen_h = root.winfo_screenheight()
win_w = min(1280, screen_w - 60)
win_h = min(920, screen_h - 60)
pos_x = (screen_w - win_w) // 2
pos_y = max(10, (screen_h - win_h) // 2 - 20)
root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
root.configure(bg=BG)
root.minsize(1100, 780)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1)

_ttk_style = ttk.Style()
_ttk_style.theme_use("clam")

# ── 라벨 헬퍼
def _lbl(parent, text, bg, fg="#333", **kw):
    return tk.Label(parent, text=text, bg=bg, fg=fg,
                    font=("맑은 고딕", 9, "bold"), **kw)

# ── placeholder 헬퍼
def _make_placeholder(widget, text, fg_normal="#000", fg_ph="#AAA"):
    widget.insert("1.0", text)
    widget.config(fg=fg_ph)
    def _in(e):
        if widget.get("1.0", tk.END).strip() == text:
            widget.delete("1.0", tk.END); widget.config(fg=fg_normal)
    def _out(e):
        if not widget.get("1.0", tk.END).strip():
            widget.insert("1.0", text); widget.config(fg=fg_ph)
    widget.bind("<FocusIn>", _in)
    widget.bind("<FocusOut>", _out)
    return text   # returns placeholder string

# ══════════════════════════════════════════════════════════════
# 모델 선택 행
# ══════════════════════════════════════════════════════════════
model_row = tk.Frame(root, bg=BG)
model_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 2))

tk.Label(model_row, text="⚙ 모델 설정", bg=BG,
         font=("맑은 고딕", 9, "bold"), fg="#444").pack(side=tk.LEFT, padx=(0, 10))

tk.Label(model_row, text="PRO (워싱/해설있음):", bg=BG,
         font=("맑은 고딕", 9), fg="#333").pack(side=tk.LEFT)
_pro_var = tk.StringVar(value=MODEL_PRO)
_pro_cb  = ttk.Combobox(model_row, textvariable=_pro_var,
                         values=[MODEL_PRO], width=26, state="readonly",
                         font=("맑은 고딕", 9))
_pro_cb.pack(side=tk.LEFT, padx=(4, 14))

tk.Label(model_row, text="FLASH (해설없음/손글씨/검수):", bg=BG,
         font=("맑은 고딕", 9), fg="#333").pack(side=tk.LEFT)
_flash_var = tk.StringVar(value=MODEL_FLASH)
_flash_cb  = ttk.Combobox(model_row, textvariable=_flash_var,
                           values=[MODEL_FLASH], width=26, state="readonly",
                           font=("맑은 고딕", 9))
_flash_cb.pack(side=tk.LEFT, padx=(4, 14))

_model_status = tk.Label(model_row, text="모델 목록 조회 중…", bg=BG,
                          font=("맑은 고딕", 8), fg="#888")
_model_status.pack(side=tk.LEFT)

def _on_pro_change(*_):
    global MODEL_PRO
    MODEL_PRO = _pro_var.get()

def _on_flash_change(*_):
    global MODEL_FLASH
    MODEL_FLASH = _flash_var.get()

_pro_var.trace_add("write", _on_pro_change)
_flash_var.trace_add("write", _on_flash_change)

def _load_models_bg():
    models = _fetch_available_models()
    def _update():
        global _AVAILABLE_MODELS
        _AVAILABLE_MODELS = models
        _pro_cb.configure(values=models)
        _flash_cb.configure(values=models)
        _model_status.config(text=f"모델 {len(models)}개", fg="#444")
    root.after(0, _update)

threading.Thread(target=_load_models_bg, daemon=True).start()

# ══════════════════════════════════════════════════════════════
# 상단: 클립보드 단일(좌) | PDF/장바구니(우)
# ══════════════════════════════════════════════════════════════
top_frame = tk.Frame(root, bg=BG)
top_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 4))
top_frame.grid_columnconfigure(0, weight=1)
top_frame.grid_columnconfigure(1, weight=1)

# ──────────────────────────────────────────────
# 왼쪽: 클립보드 단일
# ──────────────────────────────────────────────
single_box = tk.LabelFrame(top_frame, text=" 📋 클립보드 단일 ",
                             font=("맑은 고딕", 10, "bold"),
                             bg=BG_SINGLE, padx=10, pady=8, relief="groove", bd=2)
single_box.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

# 모드
r1 = tk.Frame(single_box, bg=BG_SINGLE)
r1.pack(fill=tk.X, pady=(0, 4))
_lbl(r1, "모드:", BG_SINGLE, "#0044CC").pack(side=tk.LEFT, padx=(0, 4))
single_mode_var = tk.StringVar(value="워싱")
for v in ["타이핑", "워싱"]:
    ttk.Radiobutton(r1, text=v, variable=single_mode_var, value=v).pack(side=tk.LEFT, padx=3)

# 유형 (워싱일 때만 활성)
r2 = tk.Frame(single_box, bg=BG_SINGLE)
r2.pack(fill=tk.X, pady=(0, 4))
_lbl(r2, "유형:", BG_SINGLE).pack(side=tk.LEFT, padx=(0, 4))
single_type_var = tk.StringVar(value="객관식")
_type_radios = []
for v in ["객관식", "주관식"]:
    rb = ttk.Radiobutton(r2, text=v, variable=single_type_var, value=v)
    rb.pack(side=tk.LEFT, padx=3)
    _type_radios.append(rb)
_lbl(r2, "(타이핑 시 자동감지)", BG_SINGLE, "#999").pack(side=tk.LEFT, padx=4)

def _on_single_mode(*_):
    s = "disabled" if single_mode_var.get() == "타이핑" else "!disabled"
    for rb in _type_radios:
        rb.state([s])
single_mode_var.trace_add("write", _on_single_mode)
_on_single_mode()

# 해설
r3 = tk.Frame(single_box, bg=BG_SINGLE)
r3.pack(fill=tk.X, pady=(0, 6))
_lbl(r3, "해설:", BG_SINGLE).pack(side=tk.LEFT, padx=(0, 4))
single_style_var = tk.StringVar(value="강사용")
for v in ["강사용", "학생용", "해설없음"]:
    ttk.Radiobutton(r3, text=v, variable=single_style_var, value=v).pack(side=tk.LEFT, padx=3)

# 특이사항
nf1 = tk.Frame(single_box, bg=BG_NOTE, padx=6, pady=4, relief="groove", bd=1)
nf1.pack(fill=tk.X)
nh1 = tk.Frame(nf1, bg=BG_NOTE); nh1.pack(fill=tk.X)
_lbl(nh1, "특이사항", BG_NOTE, "#774400").pack(side=tk.LEFT)
single_note_area = tk.Text(nf1, height=2, font=("맑은 고딕", 9),
                             bg="#FFFFF8", relief="sunken", bd=1, wrap=tk.WORD)
single_note_area.pack(fill=tk.X, pady=(3, 0))
_SINGLE_PLACEHOLDER = _make_placeholder(
    single_note_area, "예) 박스 조건 (가)(나) 절대 생략 금지")
tk.Button(nh1, text="지우기",
          command=lambda: (single_note_area.delete("1.0", tk.END),
                           single_note_area.event_generate("<FocusOut>")),
          font=("맑은 고딕", 8), bg="#FFCCCC", relief="flat", bd=1, padx=4
          ).pack(side=tk.RIGHT)

# 분석 버튼
tk.Button(single_box, text="📋 클립보드 분석 시작",
          command=run_single_clipboard,
          font=("맑은 고딕", 11, "bold"), bg="#A0C9FF",
          relief="raised", bd=2, pady=6
          ).pack(fill=tk.X, pady=(8, 0))

# ──────────────────────────────────────────────
# 오른쪽: PDF / 클립보드 장바구니
# ──────────────────────────────────────────────
multi_box = tk.LabelFrame(top_frame, text=" 📁 PDF / 클립보드 장바구니 ",
                            font=("맑은 고딕", 10, "bold"),
                            bg=BG_MULTI, padx=10, pady=8, relief="groove", bd=2)
multi_box.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

# 모드 + 해설 한 줄
r4 = tk.Frame(multi_box, bg=BG_MULTI)
r4.pack(fill=tk.X, pady=(0, 4))
_lbl(r4, "모드:", BG_MULTI, "#0044CC").pack(side=tk.LEFT, padx=(0, 4))
multi_mode_var = tk.StringVar(value="워싱")
for v in ["타이핑", "워싱"]:
    ttk.Radiobutton(r4, text=v, variable=multi_mode_var, value=v).pack(side=tk.LEFT, padx=3)
tk.Frame(r4, width=12, bg=BG_MULTI).pack(side=tk.LEFT)
_lbl(r4, "해설:", BG_MULTI).pack(side=tk.LEFT, padx=(0, 4))
multi_style_var = tk.StringVar(value="강사용")
for v in ["강사용", "학생용", "해설없음"]:
    ttk.Radiobutton(r4, text=v, variable=multi_style_var, value=v).pack(side=tk.LEFT, padx=3)

# 번호 스타일 행
r4b = tk.Frame(multi_box, bg=BG_MULTI)
r4b.pack(fill=tk.X, pady=(0, 4))
_lbl(r4b, "번호:", BG_MULTI, "#0044CC").pack(side=tk.LEFT, padx=(0, 4))
number_style_var = tk.StringVar(value="원본")
for v in ["원본", "순서", "없음"]:
    ttk.Radiobutton(r4b, text=v, variable=number_style_var, value=v).pack(side=tk.LEFT, padx=3)

# PDF 행
r5 = tk.Frame(multi_box, bg=BG_MULTI)
r5.pack(fill=tk.X, pady=(0, 4))
def _open_pdf():
    fp = filedialog.askopenfilename(
        title="PDF 선택", filetypes=[("PDF", "*.pdf"), ("모든 파일", "*.*")])
    if fp:
        _open_pdf_viewer(fp)
tk.Button(r5, text="📁 PDF 불러오기", command=_open_pdf,
          font=("맑은 고딕", 9, "bold"), bg="#DDEEFF", relief="raised", bd=2,
          padx=8).pack(side=tk.LEFT)

# 장바구니 행
r6 = tk.Frame(multi_box, bg=BG_MULTI)
r6.pack(fill=tk.X, pady=(0, 6))
stack_label = tk.Label(r6, text="📦 0장", font=("맑은 고딕", 9, "bold"),
                        bg=BG_MULTI, fg="#777")
stack_label.pack(side=tk.LEFT, padx=(0, 6))
tk.Button(r6, text="➕ 추가", command=add_to_stack,
          font=("맑은 고딕", 9), bg="#FFFACD", relief="raised", bd=2, padx=6
          ).pack(side=tk.LEFT, padx=2)
tk.Button(r6, text="🎯 썸네일 선택", command=analyze_stack,
          font=("맑은 고딕", 9, "bold"), bg="#C8F7C5", relief="raised", bd=2, padx=6
          ).pack(side=tk.LEFT, padx=2)
tk.Button(r6, text="🗑 비우기", command=clear_stack,
          font=("맑은 고딕", 9), bg="#FFCCCC", relief="raised", bd=2, padx=6
          ).pack(side=tk.LEFT, padx=2)

# 특이사항
nf2 = tk.Frame(multi_box, bg=BG_NOTE, padx=6, pady=4, relief="groove", bd=1)
nf2.pack(fill=tk.X)
nh2 = tk.Frame(nf2, bg=BG_NOTE); nh2.pack(fill=tk.X)
_lbl(nh2, "특이사항 (전체 적용)", BG_NOTE, "#774400").pack(side=tk.LEFT)
multi_note_area = tk.Text(nf2, height=2, font=("맑은 고딕", 9),
                           bg="#FFFFF8", relief="sunken", bd=1, wrap=tk.WORD)
multi_note_area.pack(fill=tk.X, pady=(3, 0))
_MULTI_PLACEHOLDER = _make_placeholder(
    multi_note_area, "예) 이차방정식이 인수분해 되지 않는 형태")
tk.Button(nh2, text="지우기",
          command=lambda: (multi_note_area.delete("1.0", tk.END),
                           multi_note_area.event_generate("<FocusOut>")),
          font=("맑은 고딕", 8), bg="#FFCCCC", relief="flat", bd=1, padx=4
          ).pack(side=tk.RIGHT)

_lbl(multi_box,
     "▲ PDF: 뷰어에서 직접 크롭(C/G) → 즉시분석 or 배치제출  |  장바구니: 썸네일 선택",
     BG_MULTI, "#336600").pack(anchor="w", pady=(6, 0))

# ══════════════════════════════════════════════════════════════
# 중단: AI 분석 결과
# ══════════════════════════════════════════════════════════════
mid_frame = ttk.LabelFrame(root, text="  🤖  AI 분석 결과  (검수 및 수정)  ", padding=4)
mid_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=3)
input_area = scrolledtext.ScrolledText(
    mid_frame, wrap=tk.WORD, font=("맑은 고딕", 11), bg="#FAFCFF", relief="flat")
input_area.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

# ══════════════════════════════════════════════════════════════
# 하단: HWP 제어 + 버튼
# ══════════════════════════════════════════════════════════════
bot_frame = tk.Frame(root, bg=BG)
bot_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 8))

# HWP 행
com_row = tk.Frame(bot_frame, bg="#F0F0F0", relief="groove", bd=1)
com_row.pack(fill=tk.X, pady=(0, 4))

com_status_label = tk.Label(
    com_row,
    text="🔴 pywin32 없음 — pip install pywin32 후 재시작" if not _WIN32_AVAILABLE
         else "🔴 HWP 미시작",
    font=("맑은 고딕", 9), bg="#F0F0F0", fg="#CC0000")
com_status_label.pack(side=tk.LEFT, padx=10, pady=4)

def _btn_hwp_start():
    if not _WIN32_AVAILABLE:
        messagebox.showwarning("pywin32 없음",
            "pip install pywin32 실행 후 프로그램을 재시작하세요."); return
    ok, msg = com_start_hwp()
    if not ok:
        messagebox.showwarning("HWP 시작 실패",
            f"{msg}\n\nHWP 2022: 도구 → 매크로 보안 → 낮음 설정 후 재시도")

def _btn_hwp_open_file():
    if not _WIN32_AVAILABLE:
        messagebox.showwarning("pywin32 없음", "pip install pywin32 필요"); return
    filepath = filedialog.askopenfilename(
        title="HWP 파일 열기",
        filetypes=[("한글 문서", "*.hwp *.hwpx"), ("모든 파일", "*.*")])
    if not filepath: return
    ok, msg = com_start_hwp(open_file=filepath)
    if not ok:
        messagebox.showwarning("파일 열기 실패", msg)

btn_fr = tk.Frame(com_row, bg="#F0F0F0")
btn_fr.pack(side=tk.RIGHT, padx=6)
tk.Button(btn_fr, text="📄 빈 HWP 시작", command=_btn_hwp_start,
          font=("맑은 고딕", 9, "bold"), bg="#C8F0C8", relief="raised", bd=2,
          width=13).pack(side=tk.LEFT, padx=2, pady=3)
tk.Button(btn_fr, text="📂 HWP 파일 열기", command=_btn_hwp_open_file,
          font=("맑은 고딕", 9), bg="#D8E8F8", relief="raised", bd=2,
          width=14).pack(side=tk.LEFT, padx=2, pady=3)


# 상태
status_label = tk.Label(bot_frame,
    text="준비 완료. 클립보드 캡처 후 분석하거나 PDF를 불러오세요.",
    font=("맑은 고딕", 10), fg="#0033AA", bg=BG)
status_label.pack(pady=(0, 4))

# 버튼 행
btn_row = tk.Frame(bot_frame, bg=BG)
btn_row.pack()

counter_label = tk.Label(btn_row, text="📝 0문항",
                          font=("맑은 고딕", 10, "bold"), fg="#888", bg=BG)
counter_label.pack(side=tk.LEFT, padx=(0, 10))

tk.Button(btn_row, text="💾 TXT 저장", command=save_result_txt,
          font=("맑은 고딕", 10), bg="#DCF0FF", relief="raised", bd=2,
          width=10).pack(side=tk.LEFT, padx=3)
tk.Button(btn_row, text="🔍 형식 검사", command=validate_format,
          font=("맑은 고딕", 10), bg="#FFF0CC", relief="raised", bd=2,
          width=10).pack(side=tk.LEFT, padx=3)
tk.Button(btn_row, text="📝 풀이 저장 / 뷰어 열기", command=open_solution_viewer,
          font=("맑은 고딕", 10), bg="#E8FFE8", relief="raised", bd=2,
          width=20).pack(side=tk.LEFT, padx=3)
tk.Button(btn_row, text="📥 배치 결과 불러오기",
          command=lambda: open_wash_batch_dialog(
              lambda t: (input_area.delete("1.0", tk.END), input_area.insert(tk.END, t), _update_counter()),
              lambda m: status_label.config(text=m, fg="#008000")),
          font=("맑은 고딕", 10), bg="#FFFFCC", relief="raised", bd=2,
          width=16).pack(side=tk.LEFT, padx=3)

tk.Frame(btn_row, width=12, bg=BG).pack(side=tk.LEFT)

tk.Button(btn_row, text="  ▶  한글(HWP) 입력 시작  ",
          command=run_automation,
          height=2, font=("맑은 고딕", 11, "bold"),
          bg="#A0C9FF", relief="raised", bd=3).pack(side=tk.LEFT, padx=3)
tk.Button(btn_row, text="■ 중단", command=stop_automation,
          height=2, font=("맑은 고딕", 10),
          bg="#FFCCCC", relief="raised", bd=2, width=7).pack(side=tk.LEFT, padx=3)

root.mainloop()

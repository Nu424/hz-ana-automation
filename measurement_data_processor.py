"""
実験装置の測定データ処理自動化ツール

このツールは、HZ-ANAソフトウェアによる測定データのCSV出力作業を自動化するツールです。

"""

import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import pyautogui
import psutil
import pyperclip
from pathlib import Path
from typing import List, Optional

# Windows API関連のインポート
try:
    import win32gui
    import win32con
    import win32api
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    print("警告: pywin32がインストールされていません。基本的なウィンドウアクティブ化のみ使用されます。")

# 設定ファイルをインポート
try:
    from config import (
        PYAUTOGUI_SETTINGS,
        WINDOW_SETTINGS,
        FILE_SETTINGS,
        GUI_OPERATIONS,
        DEBUG_SETTINGS,
    )
except ImportError:
    # 設定ファイルが見つからない場合のデフォルト設定
    PYAUTOGUI_SETTINGS = {"PAUSE": 0.5, "FAILSAFE": True}
    WINDOW_SETTINGS = {"TARGET_WINDOW_NAME": "HZ-ANA", "ACTIVATION_WAIT_TIME": 1.0}
    FILE_SETTINGS = {
        "TARGET_EXTENSIONS": [".dat", ".txt", ".csv", ".tsv", ".data"],
        "DIALOG_WAIT_TIME": 1.0,
        "FILE_OPEN_WAIT_TIME": 2.0,
        "EXPORT_WAIT_TIME": 2.0,
        "PROCESS_INTERVAL": 1.0,
        "DATA_RESET_WAIT_TIME": 0.5,
    }
    GUI_OPERATIONS = {
        "EXPORT_TAB_COUNT": 3,
        "EXPORT_SELECTION_COUNT": 3,
        "EXPORT_BUTTON_TAB_COUNT": 10,
        "KEY_PRESS_INTERVAL": 0.2,
    }
    DEBUG_SETTINGS = {
        "VERBOSE_OUTPUT": True,
        "DRY_RUN": False,
        "MAX_FILES_PER_SESSION": None,
    }


class MeasurementDataProcessor:
    """測定データ処理自動化クラス"""

    def __init__(self):
        """初期化処理"""
        # PyAutoGUIの設定
        pyautogui.FAILSAFE = PYAUTOGUI_SETTINGS["FAILSAFE"]
        pyautogui.PAUSE = PYAUTOGUI_SETTINGS["PAUSE"]

        # ウィンドウ名
        self.target_window_name = WINDOW_SETTINGS["TARGET_WINDOW_NAME"]

        # 処理対象ファイルリスト
        self.target_files = []

        # デバッグモード
        self.dry_run = DEBUG_SETTINGS["DRY_RUN"]
        self.verbose = DEBUG_SETTINGS["VERBOSE_OUTPUT"]

    def _log(self, message: str):
        """ログ出力（詳細モードの場合のみ）"""
        if self.verbose:
            print(message)

    def _input_text_via_clipboard(self, text: str) -> bool:
        """
        クリップボード経由でテキストを入力する

        Args:
            text (str): 入力するテキスト

        Returns:
            bool: 入力に成功した場合True
        """
        try:
            # 現在のクリップボード内容を保存
            original_clipboard = ""
            try:
                original_clipboard = pyperclip.paste()
            except Exception:
                pass  # クリップボードが空の場合など

            # テキストをクリップボードにコピー
            pyperclip.copy(text)
            time.sleep(0.05)  # クリップボード操作の安定化 - 短縮

            # Ctrl+V で貼り付け
            pyautogui.hotkey("ctrl", "v")

            # 元のクリップボード内容を復元
            try:
                pyperclip.copy(original_clipboard)
            except Exception:
                pass  # 復元に失敗しても処理は続行

            self._log(f"クリップボード経由でテキストを入力: {text[:50]}...")
            return True

        except Exception as e:
            print(f"クリップボード経由のテキスト入力でエラーが発生しました: {e}")
            return False

    def check_hz_ana_window_exists(self) -> bool:
        """
        HZ-ANAウィンドウが存在するかを確認する

        Returns:
            bool: ウィンドウが存在する場合True、存在しない場合False
        """
        try:
            # 実行中のプロセスを確認
            for proc in psutil.process_iter(["pid", "name", "cmdline"]):
                try:
                    # プロセス名やコマンドラインにHZ-ANAが含まれているかチェック
                    if proc.info["name"] and "hz-ana" in proc.info["name"].lower():
                        self._log(f"プロセス名でHZ-ANAを検出: {proc.info['name']}")
                        return True
                    if proc.info["cmdline"]:
                        cmdline_str = " ".join(proc.info["cmdline"]).lower()
                        if "hz-ana" in cmdline_str:
                            self._log(f"コマンドラインでHZ-ANAを検出: {cmdline_str}")
                            return True
                except (
                    psutil.NoSuchProcess,
                    psutil.AccessDenied,
                    psutil.ZombieProcess,
                ):
                    continue

            # PyAutoGUIでウィンドウを検索
            try:
                windows = pyautogui.getWindowsWithTitle(self.target_window_name)
                if len(windows) > 0:
                    self._log(f"ウィンドウタイトルでHZ-ANAを検出: {len(windows)}個")
                    return True
            except Exception as e:
                self._log(f"ウィンドウ検索エラー: {e}")

            return False

        except Exception as e:
            print(f"ウィンドウ確認中にエラーが発生しました: {e}")
            return False

    def get_target_directories(self) -> List[str]:
        """
        処理対象のディレクトリを複数選択する

        Returns:
            List[str]: 選択されたディレクトリのパスリスト
        """
        directories = []

        # Tkinterのルートウィンドウを非表示で作成
        root = tk.Tk()
        root.withdraw()

        messagebox.showinfo(
            "フォルダ選択",
            "処理対象のファイルが入っているフォルダを選択してください。\n"
            "複数のフォルダを選択する場合は、一つずつ選択してください。\n"
            "すべて選択し終わったら「キャンセル」を押してください。",
        )

        while True:
            directory = filedialog.askdirectory(
                title=f"処理対象フォルダを選択 (現在{len(directories)}個選択済み)"
            )

            if not directory:  # キャンセルが押された場合
                break

            if directory not in directories:
                directories.append(directory)
                print(f"フォルダが追加されました: {directory}")
            else:
                messagebox.showwarning("重複", "このフォルダは既に選択されています。")

        root.destroy()

        if not directories:
            print("フォルダが選択されませんでした。プログラムを終了します。")
            return []

        print(f"選択されたフォルダ数: {len(directories)}")
        return directories

    def get_target_files_directly(self) -> List[str]:
        """
        処理対象のファイルを直接複数選択する

        Returns:
            List[str]: 選択されたファイルのパスリスト
        """
        files = []

        # Tkinterのルートウィンドウを非表示で作成
        root = tk.Tk()
        root.withdraw()

        # 対象拡張子リストを作成（TARGET_EXTENSIONS + csv）
        target_extensions = FILE_SETTINGS["TARGET_EXTENSIONS"].copy()
        if ".csv" not in target_extensions:
            target_extensions.append(".csv")

        # ファイルタイプの設定
        filetypes = [
            ("対応ファイル", " ".join([f"*{ext}" for ext in target_extensions])),
            ("すべてのファイル", "*.*"),
        ]

        messagebox.showinfo(
            "ファイル選択",
            "処理対象のファイルを選択してください。\n"
            f"対応形式: {', '.join(target_extensions)}\n"
            "複数のファイルを一度に選択できます。\n"
            "Ctrlキーを押しながらクリックで複数選択、Shiftキーで範囲選択が可能です。\n"
            "すべて選択し終わったら「キャンセル」を押してください。",
        )

        while True:
            selected_files = filedialog.askopenfilenames(
                title=f"処理対象ファイルを選択 (現在{len(files)}個選択済み)",
                filetypes=filetypes,
            )

            if not selected_files:  # キャンセルが押された場合
                break

            # 選択されたファイルを処理
            new_files_count = 0
            for file_path in selected_files:
                # 拡張子チェック
                file_ext = Path(file_path).suffix.lower()
                if file_ext not in target_extensions:
                    messagebox.showwarning(
                        "非対応ファイル",
                        f"このファイル形式は対応していません: {os.path.basename(file_path)} ({file_ext})\n"
                        f"対応形式: {', '.join(target_extensions)}\n"
                        "このファイルはスキップされます。",
                    )
                    continue

                if file_path not in files:
                    files.append(file_path)
                    new_files_count += 1
                    print(f"ファイルが追加されました: {os.path.basename(file_path)}")
                else:
                    print(f"重複ファイルをスキップ: {os.path.basename(file_path)}")

            if new_files_count > 0:
                print(f"{new_files_count}個の新しいファイルが追加されました")

        root.destroy()

        if not files:
            print("ファイルが選択されませんでした。プログラムを終了します。")
            return []

        print(f"選択されたファイル数: {len(files)}")
        return files

    def collect_target_files(self, directories: List[str]) -> List[str]:
        """
        指定されたディレクトリから処理対象ファイルを収集する

        Args:
            directories (List[str]): 対象ディレクトリのリスト

        Returns:
            List[str]: 処理対象ファイルのパスリスト
        """
        all_files = []
        target_extensions = FILE_SETTINGS["TARGET_EXTENSIONS"]

        for directory in directories:
            try:
                directory_path = Path(directory)
                if not directory_path.exists():
                    print(f"警告: ディレクトリが存在しません: {directory}")
                    continue

                # ディレクトリ内のファイルを取得
                files_in_dir = []
                for file_path in directory_path.iterdir():
                    if file_path.is_file():
                        # 拡張子をチェック
                        if file_path.suffix.lower() in target_extensions:
                            files_in_dir.append(str(file_path))

                all_files.extend(files_in_dir)
                print(
                    f"ディレクトリ '{directory}' から {len(files_in_dir)} 個のファイルを検出"
                )

            except Exception as e:
                print(f"ディレクトリ '{directory}' の処理中にエラーが発生しました: {e}")

        # ファイル数制限の適用
        max_files = DEBUG_SETTINGS["MAX_FILES_PER_SESSION"]
        if max_files and len(all_files) > max_files:
            print(
                f"警告: ファイル数が制限を超えています。最初の{max_files}件のみ処理します。"
            )
            all_files = all_files[:max_files]

        print(f"合計 {len(all_files)} 個のファイルが処理対象です")
        return all_files

    def _find_window_by_title(self, title: str) -> Optional[int]:
        """
        ウィンドウタイトルでウィンドウハンドルを検索する（pywin32使用）
        
        Args:
            title (str): 検索するウィンドウタイトル
            
        Returns:
            Optional[int]: 見つかったウィンドウハンドル、見つからない場合はNone
        """
        if not PYWIN32_AVAILABLE:
            return None
            
        def enum_windows_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title.lower() in window_title.lower():
                    windows.append(hwnd)
            return True
        
        windows = []
        try:
            win32gui.EnumWindows(enum_windows_callback, windows)
            return windows[0] if windows else None
        except Exception as e:
            self._log(f"ウィンドウ検索エラー (pywin32): {e}")
            return None

    def _activate_window_with_pywin32(self, hwnd: int) -> bool:
        """
        pywin32を使用してウィンドウを強制的にアクティブ化する
        
        Args:
            hwnd (int): ウィンドウハンドル
            
        Returns:
            bool: アクティブ化に成功した場合True
        """
        if not PYWIN32_AVAILABLE:
            return False
            
        try:
            # ウィンドウが最小化されている場合は復元
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
            
            # ウィンドウを前面に表示
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            
            # フォーカスを設定（複数の方法を試行）
            try:
                # 方法1: SetForegroundWindow
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                try:
                    pass
                    # # 方法2: より強力な方法
                    # # 現在のスレッドと対象ウィンドウのスレッドを関連付け
                    # current_thread = win32api.GetCurrentThreadId()
                    # target_thread = win32gui.GetWindowThreadProcessId(hwnd)[0]
                    
                    # if current_thread != target_thread:
                    #     win32gui.AttachThreadInput(current_thread, target_thread, True)
                    #     win32gui.SetForegroundWindow(hwnd)
                    #     win32gui.AttachThreadInput(current_thread, target_thread, False)
                    # else:
                    #     win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    # 方法3: BringWindowToTop
                    win32gui.BringWindowToTop(hwnd)
            
            # 最終確認
            time.sleep(0.3)
            foreground_hwnd = win32gui.GetForegroundWindow()
            success = foreground_hwnd == hwnd
            
            if success:
                self._log(f"pywin32でウィンドウアクティブ化成功: {hwnd}")
            else:
                self._log(f"pywin32でウィンドウアクティブ化失敗: 期待={hwnd}, 実際={foreground_hwnd}")
                
            return success
            
        except Exception as e:
            self._log(f"pywin32ウィンドウアクティブ化エラー: {e}")
            return False

    def activate_hz_ana_window(self) -> bool:
        """
        HZ-ANAウィンドウをアクティブ化する（pywin32による強化版）

        Returns:
            bool: アクティブ化に成功した場合True
        """
        if self.dry_run:
            print("[DRY RUN] ウィンドウのアクティブ化をスキップ")
            return True

        print("HZ-ANAウィンドウのアクティブ化を開始...")
        
        # リトライ機能付きでアクティブ化を試行
        max_retries = WINDOW_SETTINGS.get("ACTIVATION_RETRY_COUNT", 2)
        retry_interval = WINDOW_SETTINGS.get("ACTIVATION_RETRY_INTERVAL", 0.5)
        
        for attempt in range(max_retries + 1):
            if attempt > 0:
                self._log(f"アクティブ化再試行 {attempt}/{max_retries}")
                time.sleep(retry_interval)

            # 方法1: pywin32を使用した強力なアクティブ化
            if PYWIN32_AVAILABLE and WINDOW_SETTINGS.get("USE_PYWIN32_FIRST", True):
                self._log("pywin32を使用したウィンドウアクティブ化を試行")
                hwnd = self._find_window_by_title(self.target_window_name)
                if hwnd:
                    if self._activate_window_with_pywin32(hwnd):
                        print("HZ-ANAウィンドウをアクティブ化しました (pywin32)")
                        return True
                    else:
                        self._log("pywin32でのアクティブ化に失敗、PyAutoGUIにフォールバック")
                else:
                    self._log("pywin32でウィンドウが見つかりません、PyAutoGUIにフォールバック")

            # 方法2: PyAutoGUIによるフォールバック
            try:
                self._log("PyAutoGUIを使用したウィンドウアクティブ化を試行")
                windows = pyautogui.getWindowsWithTitle(self.target_window_name)

                if not windows:
                    self._log("PyAutoGUIでもウィンドウが見つかりません")
                    continue

                # 最初に見つかったウィンドウをアクティブ化
                target_window = windows[0]
                target_window.activate()
                time.sleep(WINDOW_SETTINGS["ACTIVATION_WAIT_TIME"])

                print("HZ-ANAウィンドウをアクティブ化しました (PyAutoGUI)")
                return True

            except Exception as e:
                self._log(f"PyAutoGUIでのアクティブ化エラー: {e}")
                continue
        
        # すべての試行が失敗した場合
        print(f"ウィンドウのアクティブ化に失敗しました（{max_retries + 1}回試行）")
        return False

    def reset_display_data(self) -> bool:
        """
        表示しているデータをリセットする（Alt+D→下キー1回→Enter）

        Returns:
            bool: 操作に成功した場合True
        """
        if self.dry_run:
            print("[DRY RUN] データリセットをスキップ")
            return True

        try:
            # Alt+D でデータメニューを開く
            pyautogui.hotkey("alt", "d")
            # 下キー1回でリセット項目を選択
            pyautogui.press("down")
            # Enterでリセット実行
            pyautogui.press("enter")
            time.sleep(FILE_SETTINGS["DATA_RESET_WAIT_TIME"])

            self._log("データをリセットしました")
            return True

        except Exception as e:
            print(f"データリセット処理中にエラーが発生しました: {e}")
            return False

    def open_file_dialog(self) -> bool:
        """
        ファイルを開く画面を表示する

        Returns:
            bool: 操作に成功した場合True
        """
        if self.dry_run:
            print("[DRY RUN] ファイルを開く画面の表示をスキップ")
            return True

        try:
            # Alt+F → Enter でファイルを開く画面を表示
            pyautogui.hotkey("alt", "f")
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
            pyautogui.press("enter")
            time.sleep(FILE_SETTINGS["DIALOG_WAIT_TIME"])

            self._log("ファイルを開く画面を表示しました")
            return True

        except Exception as e:
            print(f"ファイルを開く画面の表示中にエラーが発生しました: {e}")
            return False

    def input_file_path_and_open(self, file_path: str) -> bool:
        """
        ファイルパスを入力してファイルを開く
        
        Args:
            file_path (str): 開くファイルのパス
            
        Returns:
            bool: 操作に成功した場合True
        """
        if self.dry_run:
            print(
                f"[DRY RUN] ファイルオープンをスキップ: {os.path.basename(file_path)}"
            )
            return True

        try:
            # Windows形式のパスに変換（/を\に変換）
            windows_path = file_path.replace("/", "\\")
            
            # ファイルパスを入力
            if not self._input_text_via_clipboard(windows_path):
                return False
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
            pyautogui.press("enter")
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])

            # Tab → Space で選択したファイルを開く
            pyautogui.press("tab")
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
            pyautogui.press("space")
            time.sleep(FILE_SETTINGS["FILE_OPEN_WAIT_TIME"])

            print(f"ファイルを開きました: {os.path.basename(file_path)}")
            return True

        except Exception as e:
            print(f"ファイルを開く処理中にエラーが発生しました: {e}")
            return False

    def export_to_csv(self) -> bool:
        """
        解析ファイルをCSV形式で書き出す

        Returns:
            bool: 操作に成功した場合True
        """
        if self.dry_run:
            print("[DRY RUN] CSV書き出しをスキップ")
            return True

        try:
            # Alt+F → 上キー×2 → Enter でファイル出力画面を表示
            pyautogui.hotkey("alt", "f")
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
            pyautogui.press("up")
            pyautogui.press("up")
            time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
            pyautogui.press("enter")
            time.sleep(FILE_SETTINGS["DIALOG_WAIT_TIME"])

            # Tab×？で出力形式選択フォームに移動
            for _ in range(GUI_OPERATIONS["EXPORT_TAB_COUNT"]):
                pyautogui.press("tab")
                time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])

            # Space → 下キー を繰り返し、出力形式をすべて選択
            for _ in range(GUI_OPERATIONS["EXPORT_SELECTION_COUNT"]):
                pyautogui.press("space")
                time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])
                pyautogui.press("down")
                time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])

            # Tab×10で書き出しボタンに移動
            for _ in range(GUI_OPERATIONS["EXPORT_BUTTON_TAB_COUNT"]):
                pyautogui.press("tab")
                time.sleep(GUI_OPERATIONS["KEY_PRESS_INTERVAL"])

            # Spaceで書き出しを完了
            pyautogui.press("space")
            time.sleep(FILE_SETTINGS["EXPORT_WAIT_TIME"])

            # 上書き確認が表示された場合に備えて「Y」を押す
            # （既存ファイルがある場合の上書き確認対応）
            pyautogui.press("y")
            time.sleep(0.5)  # 上書き処理の完了を待つ

            print("CSV書き出しを完了しました")
            return True

        except Exception as e:
            print(f"CSV書き出し処理中にエラーが発生しました: {e}")
            return False

    def process_single_file(self, file_path: str) -> bool:
        """
        単一ファイルの処理（データリセット → ファイルオープン → CSV書き出し）

        Args:
            file_path (str): 処理するファイルのパス

        Returns:
            bool: 処理に成功した場合True
        """
        print(f"\n処理開始: {os.path.basename(file_path)}")

        # データリセット
        if not self.reset_display_data():
            return False

        # ファイルを開く画面を表示
        if not self.open_file_dialog():
            return False

        # ファイルパスを入力してファイルを開く
        if not self.input_file_path_and_open(file_path):
            return False

        # 再度HZ-ANAをアクティブ化する
        if not self.activate_hz_ana_window():
            if not self.dry_run:
                messagebox.showerror(
                    "エラー", "HZ-ANAウィンドウのアクティブ化に失敗しました。"
                )
                return False

        # CSV書き出し
        if not self.export_to_csv():
            return False

        print(f"処理完了: {os.path.basename(file_path)}")
        return True

    def run(self):
        """メイン処理を実行する"""
        print("=== 測定データ処理自動化ツール ===")
        if self.dry_run:
            print("*** DRY RUN モード - 実際のGUI操作は行いません ***")
        print("開始します...\n")

        try:
            # 1. HZ-ANAウィンドウの存在確認
            print("1. HZ-ANAウィンドウの確認中...")
            if not self.check_hz_ana_window_exists():
                if not self.dry_run:
                    messagebox.showerror(
                        "エラー",
                        "HZ-ANAウィンドウが見つかりません。\n"
                        "HZ-ANAソフトウェアを起動してから再実行してください。",
                    )
                    return
                else:
                    print("[DRY RUN] HZ-ANAウィンドウの確認をスキップ")
            print("✓ HZ-ANAウィンドウを確認しました")

            # 2. 処理対象ファイルの直接選択
            print("\n2. 処理対象ファイルの選択...")
            self.target_files = self.get_target_files_directly()
            if not self.target_files:
                return

            # === 以下は従来のフォルダ選択方式（コメントアウト） ===
            # # 2. 処理対象ディレクトリの選択
            # print("\n2. 処理対象フォルダの選択...")
            # directories = self.get_target_directories()
            # if not directories:
            #     return
            #
            # # 3. 処理対象ファイルの収集
            # print("\n3. 処理対象ファイルの収集中...")
            # self.target_files = self.collect_target_files(directories)
            # if not self.target_files:
            #     messagebox.showwarning("警告", "処理対象のファイルが見つかりませんでした。")
            #     return

            # 3. 処理開始の確認
            dry_run_note = (
                "\n\n[DRY RUN モード - 実際の処理は行いません]" if self.dry_run else ""
            )
            result = messagebox.askyesno(
                "処理開始確認",
                f"{len(self.target_files)}個のファイルを処理します。\n"
                "処理を開始しますか？\n\n"
                "注意: 処理中はマウスやキーボードを操作しないでください。\n"
                "処理はOK押下3秒後に開始します。開始までにHZ-ANAウィンドウを一番上に持ってきてください(アクティブ化してください)\n"
                f"{dry_run_note}",
            )

            if not result:
                print("処理がキャンセルされました。")
                return
            time.sleep(3)

            # 4. HZ-ANAウィンドウのアクティブ化
            print("\n4. HZ-ANAウィンドウのアクティブ化...")
            if not self.activate_hz_ana_window():
                if not self.dry_run:
                    messagebox.showerror(
                        "エラー", "HZ-ANAウィンドウのアクティブ化に失敗しました。"
                    )
                    return

            # 5. ファイル処理の実行
            print("\n5. ファイル処理の開始...")
            successful_count = 0
            failed_files = []

            for i, file_path in enumerate(self.target_files, 1):
                print(f"\n進捗: {i}/{len(self.target_files)}")

                if self.process_single_file(file_path):
                    successful_count += 1
                else:
                    failed_files.append(file_path)
                    print(
                        f"エラー: ファイル処理に失敗しました - {os.path.basename(file_path)}"
                    )

                # 処理間の待機時間
                time.sleep(FILE_SETTINGS["PROCESS_INTERVAL"])

            # 6. 処理結果の報告
            print(f"\n=== 処理完了 ===")
            print(f"成功: {successful_count}件")
            print(f"失敗: {len(failed_files)}件")

            result_message = f"処理が完了しました。\n\n成功: {successful_count}件\n失敗: {len(failed_files)}件"

            if failed_files:
                result_message += f"\n\n失敗したファイル:\n"
                for file_path in failed_files[:5]:  # 最初の5件のみ表示
                    result_message += f"- {os.path.basename(file_path)}\n"
                if len(failed_files) > 5:
                    result_message += f"... 他{len(failed_files) - 5}件"

            messagebox.showinfo("処理完了", result_message)

        except KeyboardInterrupt:
            print("\n処理が中断されました。")
            messagebox.showinfo("中断", "処理が中断されました。")
        except Exception as e:
            print(f"\n予期しないエラーが発生しました: {e}")
            messagebox.showerror("エラー", f"予期しないエラーが発生しました:\n{e}")


def main():
    """メイン関数"""
    processor = MeasurementDataProcessor()
    processor.run()


if __name__ == "__main__":
    main()

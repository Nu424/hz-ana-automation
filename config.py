"""
測定データ処理自動化ツール - 設定ファイル

このファイルでは、ツールの動作パラメータを調整できます。
実際の環境に合わせて値を変更してください。
"""

# PyAutoGUI設定
PYAUTOGUI_SETTINGS = {
    "PAUSE": 0.05,  # 各操作間の待機時間（秒）
    "FAILSAFE": True,  # マウスを左上角に移動すると緊急停止
}

# ウィンドウ設定
WINDOW_SETTINGS = {
    "TARGET_WINDOW_NAME": "HZ-ANA",  # 対象ウィンドウ名
    "ACTIVATION_WAIT_TIME": 0.3,  # ウィンドウアクティブ化後の待機時間
    "USE_PYWIN32_FIRST": True,  # pywin32を優先的に使用するか
    "ACTIVATION_RETRY_COUNT": 2,  # アクティブ化の再試行回数
    "ACTIVATION_RETRY_INTERVAL": 0.5,  # 再試行間の待機時間
}

# ファイル処理設定
FILE_SETTINGS = {
    "TARGET_EXTENSIONS": [".mdp"],  # 処理対象ファイル拡張子
    "DIALOG_WAIT_TIME": 1.0,  # ダイアログ表示後の待機時間
    "FILE_OPEN_WAIT_TIME": 0.3,  # ファイルオープン後の待機時間
    "EXPORT_WAIT_TIME": 0.3,  # エクスポート処理後の待機時間
    "PROCESS_INTERVAL": 0.3,  # ファイル処理間の待機時間
    "DATA_RESET_WAIT_TIME": 0.5,  # データリセット後の待機時間
}

# GUI操作設定（実際のソフトウェアに合わせて調整が必要）
GUI_OPERATIONS = {
    "EXPORT_TAB_COUNT": 6,  # エクスポート画面での出力形式選択までのTab数
    "EXPORT_SELECTION_COUNT": 3,  # 出力形式選択の回数
    "EXPORT_BUTTON_TAB_COUNT": 10,  # 書き出しボタンまでのTab数
    "KEY_PRESS_INTERVAL": 0.03,  # キー操作間の短い待機時間
}

# デバッグ設定
DEBUG_SETTINGS = {
    "VERBOSE_OUTPUT": True,  # 詳細な出力を表示
    "DRY_RUN": False,  # True の場合、実際のGUI操作を行わない（テスト用）
    "MAX_FILES_PER_SESSION": None,  # 一度に処理するファイル数の上限（None = 無制限）
}

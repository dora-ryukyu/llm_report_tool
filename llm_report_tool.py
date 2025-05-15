import tkinter as tk
from tkinter import scrolledtext, messagebox, StringVar, BooleanVar, END, filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
# ttkbootstrap.scrolled.ScrolledText を使用
from ttkbootstrap.scrolled import ScrolledText
from ttkbootstrap.tooltip import ToolTip # ToolTip をインポート
import requests
import json
import threading
import queue
import os
import base64
from pathlib import Path # Path をインポート
import mimetypes

# --- 設定ファイル関連 ---
CONFIG_FILE = "config.json"
PDF_ENGINE_OPTIONS = ["pdf-text", "mistral-ocr", "native"]

def load_config():
    """設定ファイルを読み込む。存在しない場合はデフォルト設定で作成する。"""
    default_config = {
        "openrouter_api_key": "",
        "models": [ "google/gemini-2.0-flash-exp:free", "meta-llama/llama-4-maverick:free", "qwen/qwen3-235b-a22b:free", "deepseek/deepseek-chat-v3-0324:free", "deepseek/deepseek-r1:free" ],
        "pdf_engine": PDF_ENGINE_OPTIONS[0]
    }
    if not os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2, ensure_ascii=False)
            print(f"設定ファイル '{CONFIG_FILE}' をデフォルト設定で作成しました。")
            return default_config
        except IOError as e:
            print(f"エラー: 設定ファイル '{CONFIG_FILE}' の作成に失敗しました: {e}")
            return default_config
    else:
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            # デフォルト値で不足しているキーを補完
            for key, value in default_config.items():
                config.setdefault(key, value)
            # モデルリストの検証と修正
            if not isinstance(config.get("models"), list) or not config.get("models"):
                config["models"] = default_config["models"]
                print(f"警告: 設定ファイル '{CONFIG_FILE}' の 'models' が不正または空です。デフォルトのモデルリストを使用します。")
            # PDFエンジンの検証と修正
            if config.get("pdf_engine") not in PDF_ENGINE_OPTIONS:
                print(f"警告: 設定ファイル '{CONFIG_FILE}' の 'pdf_engine' ('{config.get('pdf_engine')}') が無効です。デフォルトの '{default_config['pdf_engine']}' を使用します。")
                config["pdf_engine"] = default_config["pdf_engine"]
            return config
        except (json.JSONDecodeError, IOError) as e:
            print(f"エラー: 設定ファイル '{CONFIG_FILE}' の読み込みに失敗しました: {e}")
            messagebox.showerror("設定エラー", f"設定ファイル '{CONFIG_FILE}' の読み込みに失敗しました。\nデフォルト設定を使用します。\n\n詳細: {e}")
            return default_config

# --- 定数 ---
STRUCTURE_OPTIONS = {
    "連続文章": "セクション（見出し）は設けず、一つの連続した文章としてください。",
    "セクション分け": "適切なセクション（見出し）を設けて、構造化されたレポートを作成してください。",
    "箇条書き": "主要なポイントを箇条書き形式で記述してください。"
}
TONE_OPTIONS = ["である/だ調", "です/ます調"]
MATERIAL_TYPES = ["PDF", "画像", "テキスト資料", "資料なし"]

# --- ファイルエンコード関数 ---
def encode_file_to_base64(file_path):
    """ファイルを読み込み、Base64エンコードされた文字列を返す"""
    try:
        with open(file_path, "rb") as file:
            return base64.b64encode(file.read()).decode('utf-8')
    except FileNotFoundError:
        messagebox.showerror("エラー", f"ファイルが見つかりません:\n{file_path}")
        return None
    except Exception as e:
        messagebox.showerror("エラー", f"ファイルの読み込みまたはエンコード中にエラーが発生しました:\n{e}")
        return None

def get_mime_type(file_path):
    """ファイルパスからMIMEタイプを取得する"""
    mime_type, _ = mimetypes.guess_type(file_path)
    return mime_type or 'application/octet-stream' # 不明な場合はデフォルト

# --- GUIクラス ---
class PromptGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        self.api_key_var = StringVar(value=self.config.get("openrouter_api_key", ""))
        self.available_models = self.config.get("models", ["(モデルなし)"])
        self.pdf_engine_var = StringVar(value=self.config.get("pdf_engine", PDF_ENGINE_OPTIONS[0]))

        if not self.available_models or self.available_models[0] == "(モデルなし)":
             messagebox.showwarning("設定警告", f"設定ファイル '{CONFIG_FILE}' から有効なモデルを読み込めませんでした。")
             self.available_models = ["(モデルなし)"] # モデルがない場合でもリストを初期化

        root.title("LLM Report Tool"); root.geometry("750x1000")
        self.result_queue = queue.Queue() # API結果受け渡し用キュー
        self.material_type_var = StringVar(value=MATERIAL_TYPES[3]) # デフォルトは資料なし
        self.pdf_path_var = StringVar(); self.image_path_var = StringVar()

        # --- メインレイアウト (PanedWindow) ---
        main_pane = ttk.PanedWindow(root, orient=VERTICAL)
        main_pane.pack(fill=BOTH, expand=YES, padx=5, pady=5)

        # --- 上部ペイン (設定エリア) ---
        settings_frame = ttk.Frame(main_pane); main_pane.add(settings_frame, weight=3) # 設定エリアの初期比率を調整
        settings_notebook = ttk.Notebook(settings_frame); settings_notebook.pack(fill=BOTH, expand=YES)

        # --- 「基本設定」タブ ---
        self.basic_settings_tab = ttk.Frame(settings_notebook, padding=5)
        settings_notebook.add(self.basic_settings_tab, text=" 基本設定 ")
        self.basic_settings_tab.columnconfigure(1, weight=1); row_idx = 0 # 列1を伸縮可能に

        # 資料の種類
        ttk.Label(self.basic_settings_tab, text="資料の種類:").grid(row=row_idx, column=0, sticky=W, padx=5, pady=(0, 5))
        material_type_frame = ttk.Frame(self.basic_settings_tab)
        material_type_frame.grid(row=row_idx, column=1, sticky=W, padx=5, pady=(0, 5))
        for type_name in MATERIAL_TYPES:
            rb = ttk.Radiobutton(material_type_frame, text=type_name, variable=self.material_type_var,
                                 value=type_name, command=self.toggle_material_input_area, bootstyle="toolbutton")
            rb.pack(side=LEFT, padx=(0, 5))
        row_idx += 1
        # 資料入力エリア
        self.material_input_frame = ttk.Frame(self.basic_settings_tab)
        self.material_input_frame.grid(row=row_idx, column=0, columnspan=2, sticky=NSEW, pady=5)
        self.material_input_frame.columnconfigure(1, weight=1)
        self.basic_settings_tab.rowconfigure(row_idx, weight=1)
        row_idx += 1
        # PDF選択UI
        self.pdf_frame = ttk.Frame(self.material_input_frame)
        self.pdf_frame.columnconfigure(1, weight=1)
        pdf_btn = ttk.Button(self.pdf_frame, text="PDFを選択...", command=self.select_pdf_file, bootstyle=OUTLINE)
        pdf_btn.grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.pdf_label = ttk.Label(self.pdf_frame, text="(PDFファイルが選択されていません)", anchor=W, justify=LEFT, wraplength=400)
        self.pdf_label.grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        ToolTip(self.pdf_label, text="")
        # 画像選択UI
        self.image_frame = ttk.Frame(self.material_input_frame)
        self.image_frame.columnconfigure(1, weight=1)
        img_btn = ttk.Button(self.image_frame, text="画像を選択...", command=self.select_image_file, bootstyle=OUTLINE)
        img_btn.grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.image_label = ttk.Label(self.image_frame, text="(画像ファイルが選択されていません)", anchor=W, justify=LEFT, wraplength=400)
        self.image_label.grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        ToolTip(self.image_label, text="")
        # テキスト資料入力UI
        self.text_material_text = ScrolledText(self.material_input_frame, height=8, wrap=tk.WORD, autohide=True)
        # テーマ、文字数、構成、口調、指示者の意見
        ttk.Label(self.basic_settings_tab, text="テーマ:").grid(row=row_idx, column=0, sticky=W, padx=5, pady=1)
        self.theme_entry = ttk.Entry(self.basic_settings_tab); self.theme_entry.grid(row=row_idx, column=1, sticky=EW, padx=5, pady=1); row_idx += 1
        ttk.Label(self.basic_settings_tab, text="文字数:").grid(row=row_idx, column=0, sticky=W, padx=5, pady=1)
        self.word_count_entry = ttk.Entry(self.basic_settings_tab); self.word_count_entry.grid(row=row_idx, column=1, sticky=EW, padx=5, pady=1); row_idx += 1
        ttk.Label(self.basic_settings_tab, text="構成:").grid(row=row_idx, column=0, sticky=W, padx=5, pady=(3, 1))
        structure_radio_frame = ttk.Frame(self.basic_settings_tab); structure_radio_frame.grid(row=row_idx, column=1, sticky=W, padx=5, pady=(3, 0))
        self.structure_var = StringVar(value="連続文章")
        for i, text in enumerate(STRUCTURE_OPTIONS.keys()):
            ttk.Radiobutton(structure_radio_frame, text=text, variable=self.structure_var, value=text, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 5))
        row_idx += 1
        ttk.Label(self.basic_settings_tab, text="口調:").grid(row=row_idx, column=0, sticky=W, padx=5, pady=(3, 1))
        tone_radio_frame = ttk.Frame(self.basic_settings_tab); tone_radio_frame.grid(row=row_idx, column=1, sticky=W, padx=5, pady=(3, 0))
        self.tone_var = StringVar(value=TONE_OPTIONS[0])
        for i, text in enumerate(TONE_OPTIONS):
            ttk.Radiobutton(tone_radio_frame, text=text, variable=self.tone_var, value=text, bootstyle="toolbutton").pack(side=LEFT, padx=(0, 5))
        row_idx += 1
        ttk.Label(self.basic_settings_tab, text="指示者の意見・視点 (任意):").grid(row=row_idx, column=0, columnspan=2, sticky=W, padx=5, pady=(3, 1)); row_idx += 1
        self.instructor_opinion_text = ScrolledText(self.basic_settings_tab, height=8, wrap=tk.WORD, autohide=True)
        self.instructor_opinion_text.grid(row=row_idx, column=0, columnspan=2, sticky=EW+NS, padx=5, pady=1)
        self.basic_settings_tab.rowconfigure(row_idx, weight=1)

        # --- 「API設定」タブ ---
        api_settings_tab = ttk.Frame(settings_notebook, padding=5)
        settings_notebook.add(api_settings_tab, text=" API設定 ")
        api_settings_tab.columnconfigure(1, weight=1); row_idx_api = 0
        ttk.Label(api_settings_tab, text="APIキー:").grid(row=row_idx_api, column=0, sticky=W, padx=5, pady=2)
        self.api_key_entry = ttk.Entry(api_settings_tab, show='*', textvariable=self.api_key_var); self.api_key_entry.grid(row=row_idx_api, column=1, sticky=EW, padx=5, pady=2); row_idx_api += 1
        ttk.Label(api_settings_tab, text="モデル:").grid(row=row_idx_api, column=0, sticky=W, padx=5, pady=2)
        self.model_combobox = ttk.Combobox(api_settings_tab, values=self.available_models, state="readonly", width=30); self.model_combobox.grid(row=row_idx_api, column=1, sticky=EW, padx=5, pady=2)
        if self.available_models and self.available_models[0] != "(モデルなし)":
            self.model_combobox.current(0)
        row_idx_api += 1
        ttk.Label(api_settings_tab, text="PDFエンジン:").grid(row=row_idx_api, column=0, sticky=W, padx=5, pady=2)
        self.pdf_engine_combobox = ttk.Combobox(api_settings_tab, values=PDF_ENGINE_OPTIONS, state="readonly", textvariable=self.pdf_engine_var)
        self.pdf_engine_combobox.grid(row=row_idx_api, column=1, sticky=EW, padx=5, pady=2)
        current_engine = self.pdf_engine_var.get()
        if current_engine in PDF_ENGINE_OPTIONS:
            self.pdf_engine_combobox.current(PDF_ENGINE_OPTIONS.index(current_engine))
        else:
            self.pdf_engine_combobox.current(0)
        row_idx_api += 1


        # --- 下部ペイン (アクションボタンと出力エリア) ---
        bottom_frame = ttk.Frame(main_pane, padding=(0, 5, 0, 0)); main_pane.add(bottom_frame, weight=1) # 出力エリアの初期比率
        bottom_frame.rowconfigure(1, weight=1) # 出力ノートブックエリアを行方向に伸縮
        bottom_frame.columnconfigure(0, weight=1) # 出力ノートブックエリアを列方向に伸縮

        # --- アクションボタンフレーム ---
        action_frame = ttk.Frame(bottom_frame)
        action_frame.grid(row=0, column=0, sticky=EW, pady=(0, 5)) # 上部に配置

        generate_button = ttk.Button(action_frame, text="プロンプト生成", command=self.generate_prompt, bootstyle=SECONDARY)
        generate_button.pack(side=LEFT, padx=(0, 5))
        copy_button = ttk.Button(action_frame, text="コピー", command=self.copy_displayed_text, bootstyle="outline-secondary")
        copy_button.pack(side=LEFT, padx=(0, 10))
        self.api_execute_button = ttk.Button(action_frame, text="実行", command=self.start_api_request, bootstyle=PRIMARY)
        self.api_execute_button.pack(side=LEFT, padx=(0, 5))
        save_button = ttk.Button(action_frame, text="結果を保存", command=self.save_result_to_file, bootstyle="outline-info")
        save_button.pack(side=LEFT, padx=(0, 5))

        # --- ステータスラベル ---
        self.status_label = ttk.Label(action_frame, text="")
        self.status_label.pack(side=LEFT, padx=5, fill=X, expand=YES)

        # --- 文字数表示ラベル ---
        self.char_count_label = ttk.Label(action_frame, text="文字数: --", width=12, anchor=E)
        self.char_count_label.pack(side=RIGHT, padx=(5, 0))

        # --- 出力ノートブック ---
        self.output_notebook = ttk.Notebook(bottom_frame)
        self.output_notebook.grid(row=1, column=0, sticky=NSEW)

        prompt_output_tab = ttk.Frame(self.output_notebook)
        self.output_notebook.add(prompt_output_tab, text=" プロンプト ")
        self.output_text = ScrolledText(prompt_output_tab, wrap=tk.WORD, autohide=True, vbar=True, hbar=False, height=8)
        self.output_text.pack(fill=BOTH, expand=YES, padx=1, pady=1)

        result_output_tab = ttk.Frame(self.output_notebook)
        self.output_notebook.add(result_output_tab, text=" 実行結果 ")
        self.result_text = ScrolledText(result_output_tab, wrap=tk.WORD, autohide=True, vbar=True, hbar=False, height=8)
        self.result_text.pack(fill=BOTH, expand=YES, padx=1, pady=1)

        # --- 実行結果テキストエリアの変更イベントをバインド ---
        # ScrolledTextの実体は内部の .text ウィジェットなので、それにバインドする
        self.result_text.text.bind("<<Modified>>", self.update_char_count_realtime)


        # --- 初期化処理 ---
        self.toggle_material_input_area() # 初期表示を更新
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # 閉じる際の処理
        self.root.after(100, self.process_queue) # キュー監視を開始

    # --- メソッド ---

    # === on_closing メソッド ===
    def on_closing(self):
        """ウィンドウを閉じる際に設定を保存し、メッセージを表示"""
        current_config = {
            "openrouter_api_key": self.api_key_var.get(),
            "models": self.config.get("models", []), # 保存時は元のリストを維持
            "pdf_engine": self.pdf_engine_var.get()
        }
        save_success = False
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(current_config, f, indent=2, ensure_ascii=False)
            save_success = True
        except OSError as e:
            error_msg = f"設定保存エラー: {e}"
            print(error_msg)
            try:
                if self.root.winfo_exists():
                    self.status_label.config(text=error_msg, bootstyle=DANGER)
                    self.root.update_idletasks()
            except tk.TclError:
                 pass

        if save_success:
            try:
                if self.root.winfo_exists():
                    self.status_label.config(text=f"設定を'{CONFIG_FILE}'に保存しました", bootstyle="info")
                    self.root.update_idletasks()
                    self.root.after(1000, self.root.destroy)
                    return
            except tk.TclError:
                 pass

        if self.root.winfo_exists():
            self.root.destroy()

    def select_pdf_file(self):
        """PDFファイル選択ダイアログを開き、選択されたパスとファイル名を表示"""
        filepath = filedialog.askopenfilename(
            title="PDFファイルを選択",
            filetypes=[("PDFファイル", "*.pdf")]
        )
        if filepath:
            self.pdf_path_var.set(filepath)
            filename = Path(filepath).name
            self.pdf_label.config(text=filename)
            ToolTip(self.pdf_label, text=filepath)

    def select_image_file(self):
        """画像ファイル選択ダイアログを開き、選択されたパスとファイル名を表示"""
        filepath = filedialog.askopenfilename(
            title="画像ファイルを選択",
            filetypes=[("画像ファイル", "*.jpg *.jpeg *.png *.gif *.bmp *.webp")]
        )
        if filepath:
            self.image_path_var.set(filepath)
            filename = Path(filepath).name
            self.image_label.config(text=filename)
            ToolTip(self.image_label, text=filepath)

    def toggle_material_input_area(self):
        """選択された資料の種類に応じて、入力エリアの表示を切り替える"""
        self.pdf_frame.grid_forget()
        self.image_frame.grid_forget()
        self.text_material_text.pack_forget()
        selected_type = self.material_type_var.get()
        if selected_type == "PDF":
            self.pdf_frame.grid(row=0, column=0, columnspan=2, sticky=NSEW, padx=0, pady=0)
        elif selected_type == "画像":
            self.image_frame.grid(row=0, column=0, columnspan=2, sticky=NSEW, padx=0, pady=0)
        elif selected_type == "テキスト資料":
            self.text_material_text.pack(fill=BOTH, expand=YES, padx=5, pady=5)

    def _build_api_payload(self):
        """入力値に基づいてAPIリクエストのペイロード(messages部分)を構築する"""
        theme = self.theme_entry.get() or "(テーマ未設定)"
        word_count = self.word_count_entry.get() or "(文字数未設定)"
        selected_structure_key = self.structure_var.get()
        structure = STRUCTURE_OPTIONS.get(selected_structure_key, "（構成未選択）")
        selected_tone = self.tone_var.get()
        instructor_opinion_text = self.instructor_opinion_text.get("1.0", tk.END).strip()
        material_type = self.material_type_var.get()
        selected_pdf_engine = self.pdf_engine_var.get()

        if selected_pdf_engine not in PDF_ENGINE_OPTIONS:
            print(f"警告: 無効なPDFエンジン '{selected_pdf_engine}' が選択されました。デフォルトの '{PDF_ENGINE_OPTIONS[0]}' を使用します。")
            selected_pdf_engine = PDF_ENGINE_OPTIONS[0]
            self.pdf_engine_var.set(selected_pdf_engine)

        instruction_text_parts = [
            f"以下の条件に従いレポートを作成してください。\n",
            "## 条件",
            f"* **テーマ:** 「{theme}」",
            f"* **文字数:** 「{word_count}字」程度\n    * 指定文字数を大幅に超えたり、不足したりしないように注意してください。",
            f"* **構成:**\n    * {structure}",
            f"* **文体:**",
            f"    * 平易な言葉遣いを心がけ、専門用語や難解な語彙の多用は避けてください。",
            f"    * 口調は「{selected_tone}」を使用してください。",
        ]
        if instructor_opinion_text:
            instruction_text_parts.extend([
                "\n## 指示者の意見・視点",
                instructor_opinion_text
            ])
        instruction_text_parts.extend([
            "\n## 出力に関する厳守事項",
            "* **レポートの本文のみ**を生成してください。",
            "* **「承知しました」「レポートを作成します」といった応答や、レポート本文以外の説明、前置き、結びの挨拶、タイトル、氏名、日付、参考文献リスト（別途指示がない限り）などは一切含めないでください。**",
            "\n## レポート作成開始"
        ])
        final_instruction_text = "\n".join(instruction_text_parts)

        messages = []
        content_list = [{"type": "text", "text": final_instruction_text}]
        plugins = None

        if material_type == "PDF":
            pdf_path = self.pdf_path_var.get()
            if not pdf_path:
                messagebox.showwarning("警告", "PDFファイルが選択されていません。", parent=self.root)
                return None
            base64_pdf = encode_file_to_base64(pdf_path)
            if not base64_pdf: return None
            data_url = f"data:application/pdf;base64,{base64_pdf}"
            filename = Path(pdf_path).name
            content_list.insert(0, {"type": "text", "text": f"添付されたPDFファイル「{filename}」の内容を精読し、その情報を基にしてレポートを作成してください。"})
            content_list.append({
                 "type": "file",
                 "file": { "filename": filename, "file_data": data_url }
            })
            plugins = [{"id": "file-parser", "pdf": {"engine": selected_pdf_engine}}]
        elif material_type == "画像":
            image_path = self.image_path_var.get()
            if not image_path:
                messagebox.showwarning("警告", "画像ファイルが選択されていません。", parent=self.root)
                return None
            base64_image = encode_file_to_base64(image_path)
            if not base64_image: return None
            mime_type = get_mime_type(image_path)
            data_url = f"data:{mime_type};base64,{base64_image}"
            filename = Path(image_path).name
            content_list.insert(0, {"type": "text", "text": f"添付された画像ファイル「{filename}」の内容を理解し、その情報を基にしてレポートを作成してください。"})
            content_list.append({
                "type": "image_url",
                "image_url": { "url": data_url }
            })
        elif material_type == "テキスト資料":
            text_material = self.text_material_text.get("1.0", tk.END).strip()
            if text_material:
                content_list.insert(0, {"type": "text", "text": "以下のテキスト資料の内容を考慮してレポートを作成してください。"})
                content_list.append({"type": "text", "text": f"\n## 資料\n{text_material}"})
            else:
                content_list.insert(0, {"type": "text", "text": "あなた自身の知識ベースに基づいてレポートを作成してください。"})
                messagebox.showinfo("情報", "テキスト資料が空のため、モデル自身の知識に基づいてレポートを作成します。", parent=self.root)
        elif material_type == "資料なし":
            content_list.insert(0, {"type": "text", "text": "あなた自身の知識ベースに基づいてレポートを作成してください。"})

        messages.append({"role": "user", "content": content_list})
        payload = {"messages": messages}
        if plugins:
            payload["plugins"] = plugins
        return payload

    def generate_prompt(self):
        """プロンプト生成ボタンが押されたときの処理"""
        payload = self._build_api_payload()
        if payload:
            display_text = ""
            user_content = payload["messages"][0]["content"]
            instruction_texts = []
            file_info_texts = []
            if "plugins" in payload and payload["plugins"][0]["id"] == "file-parser":
                 pdf_path = self.pdf_path_var.get()
                 filename = Path(pdf_path).name if pdf_path else "(不明なPDF)"
                 file_info_texts.append(f"[添付ファイル: {filename} (エンジン: {payload['plugins'][0]['pdf']['engine']})]\n")
            elif any(item["type"] == "image_url" for item in user_content):
                 img_path = self.image_path_var.get()
                 filename = Path(img_path).name if img_path else "(不明な画像)"
                 file_info_texts.append(f"[添付画像: {filename}]\n")
            for item in user_content:
                if item["type"] == "text":
                    instruction_texts.append(item["text"])
            display_text += "\n\n".join(instruction_texts) + "\n\n" + "".join(file_info_texts)
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", display_text.strip())
            self.output_notebook.select(0)
            self.status_label.config(text="プロンプト生成完了", bootstyle="info")
            self.root.after(2000, lambda: self.status_label.config(text="", bootstyle="default"))
        else:
            self.output_text.delete("1.0", tk.END)
            self.output_text.insert("1.0", "[エラー: プロンプトを構築できませんでした。入力内容を確認してください。]")
            self.status_label.config(text="プロンプト生成エラー", bootstyle="danger")
            self.root.after(3000, lambda: self.status_label.config(text="", bootstyle="default"))

    def copy_displayed_text(self):
        """現在表示中のタブのテキストをクリップボードにコピーする"""
        try:
            current_tab_index = self.output_notebook.index("current")
            widget_to_copy = None
            if current_tab_index == 0:
                widget_to_copy = self.output_text
            elif current_tab_index == 1:
                widget_to_copy = self.result_text

            if widget_to_copy:
                text_to_copy = widget_to_copy.get("1.0", "end-1c")
                if text_to_copy.strip() and not text_to_copy.startswith("[エラー:") and not text_to_copy.startswith("エラーが発生しました:"):
                    self.root.clipboard_clear()
                    self.root.clipboard_append(text_to_copy)
                    self.status_label.config(text="表示中のテキストをコピーしました", bootstyle="info")
                    self.root.after(2000, lambda: self.status_label.config(text="", bootstyle="default"))
                elif text_to_copy.startswith("[エラー:") or text_to_copy.startswith("エラーが発生しました:"):
                     messagebox.showwarning("コピー不可", "エラーメッセージはコピーできません。", parent=self.root)
                else:
                     messagebox.showwarning("コピー不可", "コピーするテキストがありません。", parent=self.root)
            else:
                 messagebox.showerror("エラー", "コピー対象のウィジェットが見つかりません。", parent=self.root)
        except tk.TclError:
             messagebox.showerror("エラー", "タブ情報の取得に失敗しました。", parent=self.root)
        except Exception as e:
             messagebox.showerror("エラー", f"コピー中に予期せぬエラーが発生しました:\n{e}", parent=self.root)

    # --- 文字数リアルタイム更新用メソッド ---
    def update_char_count_realtime(self, event=None):
        """result_textの内容変更時に文字数を計算してラベルを更新する"""
        try:
            # ウィジェットが存在するか確認 (ウィンドウ закрытия 時のエラー防止)
            if not self.result_text.winfo_exists() or not self.char_count_label.winfo_exists():
                return

            current_content = self.result_text.get("1.0", tk.END).strip()
            char_count = len(current_content)
            self.char_count_label.config(text=f"文字数: {char_count}")

            # Text ウィジェットの変更フラグをリセット (重要)
            # これを行わないと、次の変更時に<<Modified>>イベントが発生しないことがある
            self.result_text.text.edit_modified(False)
        except tk.TclError:
             # ウィンドウ закрытия 中などに TclError が発生することがあるため無視
             pass
        except Exception as e:
            print(f"リアルタイム文字数更新中にエラー: {e}")


    def start_api_request(self):
        """APIリクエストを開始する前のチェックとスレッド起動"""
        api_key = self.api_key_var.get()
        if not api_key:
            messagebox.showerror("エラー", f"APIキーが設定されていません。\n'{CONFIG_FILE}'を確認するか、API設定タブで入力してください。", parent=self.root)
            return

        model = self.model_combobox.get()
        if not model or model == "(モデルなし)":
            messagebox.showerror("エラー", "モデルが選択されていません。", parent=self.root)
            return

        payload = self._build_api_payload()
        if not payload:
            return

        payload["model"] = model

        self.api_execute_button.config(state=DISABLED)
        self.status_label.config(text=f"APIリクエスト中 ({model})...", bootstyle=WARNING)
        # --- 文字数ラベルをリセット ---
        self.char_count_label.config(text="文字数: --")
        self.result_text.delete("1.0", END) # 以前の結果をクリア (ここでModified発生→リセットされる)
        self.output_notebook.select(1) # 実行結果タブを表示

        thread = threading.Thread(target=self._api_request_thread, args=(api_key, payload), daemon=True)
        thread.start()

    def _api_request_thread(self, api_key, payload):
        """APIリクエストを実行し、結果をキューに入れる (バックグラウンドスレッド)"""
        try:
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost",
                "X-Title": "LLM Report Tool"
            }
            response = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=300
            )
            if not response.ok:
                try:
                    error_data = response.json()
                    error_message = error_data.get("error", {}).get("message", response.text)
                except json.JSONDecodeError:
                    error_message = response.text
                detailed_error = f"APIエラー (HTTP {response.status_code}): {error_message}"
                self.result_queue.put(("error", detailed_error))
                return

            response_data = response.json()
            if response_data.get("choices") and len(response_data["choices"]) > 0:
                content = response_data["choices"][0].get("message", {}).get("content")
                if content:
                    self.result_queue.put(("success", content))
                else:
                    self.result_queue.put(("error", "APIレスポンスに有効な 'content' が見つかりませんでした。"))
            else:
                error_details = response_data.get("error", {}).get("message", f"予期しない応答形式:\n{response_data}")
                self.result_queue.put(("error", f"APIエラー: {error_details}"))
        except requests.exceptions.Timeout:
            self.result_queue.put(("error", "APIリクエストがタイムアウトしました。"))
        except requests.exceptions.RequestException as e:
            self.result_queue.put(("error", f"ネットワーク通信エラーが発生しました: {e}"))
        except json.JSONDecodeError:
            self.result_queue.put(("error", "APIからの応答をJSONとして解析できませんでした。"))
        except Exception as e:
            self.result_queue.put(("error", f"予期せぬエラーが発生しました: {e}"))


    def process_queue(self):
        """キューを定期的にチェックし、結果をGUIに反映する"""
        try:
            message_type, data = self.result_queue.get_nowait() # ノンブロッキングで取得

            if message_type == "success":
                # イベントハンドラが呼ばれるように、insert前にフラグをセットしておく
                # (deleteですでにフラグは立っているはずだが念のため)
                self.result_text.text.edit_modified(True)
                self.result_text.delete("1.0", END)
                self.result_text.insert(END, data)
                # insert 後に Modified イベントが発生し、update_char_count_realtime が呼ばれる

                # ステータスラベルのみ更新
                self.status_label.config(text="生成完了", bootstyle=SUCCESS)

            elif message_type == "error":
                # エラーメッセージを表示
                self.result_text.text.edit_modified(True)
                self.result_text.delete("1.0", END)
                self.result_text.insert(END, f"エラーが発生しました:\n\n{data}")
                # insert 後に Modified イベントが発生するが、エラー時は "--" にしたいので上書き
                self.char_count_label.config(text="文字数: --") # 文字数ラベルをリセット

                # ステータスラベル更新
                if "HTTP" in data: status_msg = data.split(":")[0]
                elif "タイムアウト" in data: status_msg = "タイムアウト"
                else: status_msg = "エラー発生"
                self.status_label.config(text=status_msg, bootstyle=DANGER)


            # API処理完了後、実行ボタンを再度有効化
            self.api_execute_button.config(state=NORMAL)

            # 一定時間後にステータスメッセージのみクリア (文字数ラベルはそのまま)
            self.root.after(3000, lambda: self.status_label.config(text="", bootstyle="default"))

        except queue.Empty:
            pass
        except Exception as e:
             print(f"キュー処理中にエラー発生: {e}")
             try:
                 self.status_label.config(text="内部エラー", bootstyle=DANGER)
                 self.char_count_label.config(text="文字数: --") # エラー時はリセット
                 self.api_execute_button.config(state=NORMAL)
             except tk.TclError:
                 pass
        finally:
            self.root.after(100, self.process_queue)

    def save_result_to_file(self):
        """実行結果テキストをファイルに保存する"""
        result_content = self.result_text.get("1.0", tk.END).strip()
        if not result_content or result_content.startswith("エラーが発生しました:"):
            messagebox.showwarning("保存不可", "保存する有効な実行結果がありません。", parent=self.root)
            return

        filepath = filedialog.asksaveasfilename(
            title="実行結果を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
            initialfile=f"レポート_{self.theme_entry.get()[:10]}.txt" if self.theme_entry.get() else "レポート.txt",
            parent=self.root
        )

        if not filepath:
            self.status_label.config(text="保存をキャンセルしました", bootstyle="warning")
            self.root.after(2000, lambda: self.status_label.config(text="", bootstyle="default"))
            return

        try:
            filename = Path(filepath).name
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(result_content)
            self.status_label.config(text=f"'{filename}' に保存しました", bootstyle="success")
            self.root.after(3000, lambda: self.status_label.config(text="", bootstyle="default"))
        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルの保存中にエラーが発生しました:\n{e}", parent=self.root)
            self.status_label.config(text="ファイル保存エラー", bootstyle="danger")
            self.root.after(3000, lambda: self.status_label.config(text="", bootstyle="default"))

# --- メイン処理 ---
if __name__ == "__main__":
    root = ttk.Window(themename="litera")
    app = PromptGeneratorGUI(root)
    root.mainloop()
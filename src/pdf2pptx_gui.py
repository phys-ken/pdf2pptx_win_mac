"""
PDFをPPTXに変換するGUIアプリケーション
ドラッグ＆ドロップで簡単に操作できるシンプルなインターフェース
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD
from pdf_converter import PDFConverter


class DragDropFrame(tk.Frame):
    """ドラッグ＆ドロップ対応のフレームクラス"""
    
    def __init__(self, master, on_drop_callback, **kwargs):
        super().__init__(master, **kwargs)
        self.on_drop_callback = on_drop_callback
        
        # ドラッグ＆ドロップエリアのラベル
        self.drop_label = tk.Label(
            self,
            text="PDFファイルをここにドラッグ＆ドロップしてください\n\nまたはクリックしてファイルを選択",
            font=("Arial", 14),
            padx=20,
            pady=20,
            bg="#f0f0f0",
            relief=tk.GROOVE,
            borderwidth=2
        )
        self.drop_label.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # クリックイベントのバインド（ファイル選択ダイアログを表示）
        self.drop_label.bind("<Button-1>", self._on_click)
        
        # ドラッグ＆ドロップの設定
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_drop)
        
        # 外観の設定
        self._set_default_appearance()
        
        # マウスイベントのバインド
        self.drop_label.bind("<Enter>", self._on_enter)
        self.drop_label.bind("<Leave>", self._on_leave)
    
    def _on_click(self, event):
        """クリックイベント処理"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.on_drop_callback(file_path)
    
    def _on_drop(self, event):
        """ドロップイベント処理"""
        # ドロップされたファイルパスを取得
        file_path = event.data
        
        # Windowsでは余分な{}や引用符が含まれることがある
        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]
        
        # 引用符を削除
        if file_path.startswith('"') and file_path.endswith('"'):
            file_path = file_path[1:-1]
        
        # 複数ファイルの場合は最初のファイルのみ使用
        if ' ' in file_path:
            file_paths = file_path.split(' ')
            for path in file_paths:
                # 引用符を削除して確認
                clean_path = path.strip('"')
                if clean_path.lower().endswith('.pdf'):
                    self.on_drop_callback(clean_path)
                    return
        
        # 単一ファイルの処理
        if file_path.lower().endswith('.pdf'):
            self.on_drop_callback(file_path)
        else:
            messagebox.showwarning("警告", "PDFファイルのみ変換できます。")
    
    def _set_default_appearance(self):
        """デフォルト外観の設定"""
        self.drop_label.config(
            bg="#f0f0f0",
            fg="#333333",
            text="PDFファイルをここにドラッグ＆ドロップしてください\n\nまたはクリックしてファイルを選択"
        )
    
    def _on_enter(self, event):
        """マウス入ってきた時のイベント"""
        self.drop_label.config(
            bg="#e0e0ff",
            fg="#0000aa",
            cursor="hand2"
        )
    
    def _on_leave(self, event):
        """マウス出て行った時のイベント"""
        self._set_default_appearance()


class PDF2PPTXApp(TkinterDnD.Tk):
    """PDFをPPTXに変換するメインアプリケーション"""
    
    def __init__(self):
        super().__init__()
        
        # ウィンドウの設定
        self.title("PDF to PPTX 変換")
        self.geometry("600x400")
        self.minsize(500, 350)
        
        # アプリアイコン設定
        icon_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                                "resources", "app_icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        
        # 変換エンジン
        self.converter = PDFConverter()
        
        # 変数初期化
        self.pdf_path = None
        self.output_folder = None
        self.conversion_in_progress = False
        self.conversion_thread = None
        
        # UI作成
        self._create_widgets()
        
        # アプリケーション全体のドロップターゲット登録
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_app_drop)
        
        # 終了時の処理を設定
        self.protocol("WM_DELETE_WINDOW", self._on_close)
    
    def _create_widgets(self):
        """ウィジェットの作成"""
        # メインフレーム
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 上部フレーム（ファイル選択用）
        top_frame = tk.Frame(main_frame)
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 0))
        
        # PDFファイル選択ボタン
        self.pdf_btn = tk.Button(
            top_frame,
            text="PDFファイルを選択",
            command=self._select_pdf
        )
        self.pdf_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 選択されたPDFファイル表示ラベル
        self.pdf_label = tk.Label(top_frame, text="選択されたファイル: なし", anchor='w')
        self.pdf_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 出力先選択ボタン
        self.output_btn = tk.Button(
            top_frame,
            text="出力先を選択 (オプション)",
            command=self._select_output_folder
        )
        self.output_btn.pack(side=tk.RIGHT)
        
        # 中央フレーム（ドラッグ＆ドロップエリア）
        self.drop_frame = DragDropFrame(main_frame, self._on_file_drop, bg="#f0f0f0")
        self.drop_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 下部フレーム（変換ボタンとプログレスバー）
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # プログレスバー
        self.progress = ttk.Progressbar(
            bottom_frame,
            orient=tk.HORIZONTAL,
            length=100,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # ステータスラベル
        self.status_label = tk.Label(bottom_frame, text="準備完了")
        self.status_label.pack(fill=tk.X, pady=(0, 10))
        
        # 変換ボタン
        self.convert_btn = tk.Button(
            bottom_frame,
            text="変換開始",
            command=self._start_conversion,
            state=tk.DISABLED,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold")
        )
        self.convert_btn.pack(fill=tk.X)
    
    def _on_app_drop(self, event):
        """アプリケーション全体へのドロップイベント"""
        self._on_file_drop(event.data)
    
    def _on_file_drop(self, file_path):
        """ファイルがドロップされたときの処理"""
        # 変換中は新しいファイルを受け付けない
        if self.conversion_in_progress:
            messagebox.showinfo("通知", "変換中です。しばらくお待ちください。")
            return
            
        # Windowsでは余分な{}や引用符が含まれることがある
        if isinstance(file_path, str):
            if file_path.startswith("{") and file_path.endswith("}"):
                file_path = file_path[1:-1]
            
            # 引用符を削除
            if file_path.startswith('"') and file_path.endswith('"'):
                file_path = file_path[1:-1]
        
        if file_path and file_path.lower().endswith('.pdf'):
            self.pdf_path = file_path
            self.pdf_label.config(text=f"選択されたファイル: {os.path.basename(file_path)}")
            self.convert_btn.config(state=tk.NORMAL)
            
            # デフォルトの出力先をPDFと同じフォルダに設定
            self.output_folder = os.path.dirname(file_path)
        else:
            messagebox.showwarning("警告", "PDFファイルのみ変換できます。")
    
    def _select_pdf(self):
        """PDFファイル選択ダイアログ"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self._on_file_drop(file_path)
    
    def _select_output_folder(self):
        """出力先フォルダ選択ダイアログ"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder = folder_path
            messagebox.showinfo("出力先", f"出力先フォルダを設定しました:\n{folder_path}")
    
    def _start_conversion(self):
        """変換処理の開始"""
        if not self.pdf_path:
            messagebox.showwarning("警告", "PDFファイルを選択してください。")
            return
        
        if not os.path.exists(self.pdf_path):
            messagebox.showerror("エラー", "選択されたPDFファイルが見つかりません。")
            return
        
        if self.conversion_in_progress:
            return
        
        # UI状態の更新
        self.conversion_in_progress = True
        self.convert_btn.config(state=tk.DISABLED)
        self.pdf_btn.config(state=tk.DISABLED)
        self.output_btn.config(state=tk.DISABLED)
        
        # プログレスバーリセット
        self.progress["value"] = 0
        
        # 変換処理を別スレッドで実行
        self.conversion_thread = threading.Thread(target=self._convert_pdf_thread)
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def _convert_pdf_thread(self):
        """別スレッドでPDF変換を実行"""
        try:
            # 変換実行
            self._update_status("開始", "変換を開始します...", 0)
            pptx_path, images_folder = self.converter.convert_pdf_to_pptx(
                self.pdf_path,
                self.output_folder,
                self._update_progress
            )
            
            if pptx_path and os.path.exists(pptx_path):
                # 成功
                self._update_status(
                    "完了",
                    f"変換が完了しました！\n"
                    f"PowerPointファイル: {os.path.basename(pptx_path)}\n"
                    f"画像フォルダ: {os.path.basename(images_folder)}"
                )
                # 成功メッセージ
                self.after(0, lambda: messagebox.showinfo(
                    "変換完了",
                    f"PDFの変換が完了しました！\n\n"
                    f"PowerPointファイル:\n{pptx_path}\n\n"
                    f"画像フォルダ:\n{images_folder}"
                ))
            else:
                # 失敗
                self._update_status("エラー", "変換に失敗しました。", None)
                self.after(0, lambda: messagebox.showerror(
                    "変換失敗", 
                    "PDFの変換に失敗しました。ファイルが破損しているか、サポートされていない形式の可能性があります。"
                ))
        except Exception as e:
            # エラーメッセージをユーザーフレンドリーにする
            error_msg = str(e)
            user_friendly_msg = error_msg
            
            # 特定のエラーメッセージに対するわかりやすい説明
            if "value must be an integral type" in error_msg:
                user_friendly_msg = (
                    "PDFのサイズ形式に問題があります。\n"
                    "可能であれば別のPDF編集ソフトでPDFを開き直して保存してから再試行してください。"
                )
            elif "password required" in error_msg.lower():
                user_friendly_msg = "パスワード付きPDFファイルは処理できません。パスワードを解除してから再試行してください。"
            elif "not a PDF file" in error_msg:
                user_friendly_msg = "選択されたファイルは有効なPDFファイルではありません。"
            elif "memory" in error_msg.lower():
                user_friendly_msg = "メモリ不足エラーが発生しました。PDFのサイズが大きすぎる可能性があります。"
            elif "permission" in error_msg.lower():
                user_friendly_msg = "ファイルの読み取りまたは書き込み権限がありません。"
            
            # エラー表示
            self.after(0, lambda: messagebox.showerror(
                "変換エラー", 
                f"PDFの変換中に問題が発生しました:\n\n{user_friendly_msg}\n\n"
                f"技術的詳細:\n{error_msg}"
            ))
            self._update_status("エラー", f"エラー: {user_friendly_msg}")
        
        finally:
            # UI状態の復元
            self.after(0, self._reset_ui)
    
    def _update_progress(self, status, message, progress=None):
        """進捗状況の更新（別スレッドから呼ばれる）"""
        self.after(0, lambda: self._update_status(status, message, progress))
    
    def _update_status(self, status, message, progress=None):
        """ステータス表示の更新（メインスレッドで実行）"""
        self.status_label.config(text=message)
        
        if progress is not None:
            self.progress["value"] = progress
    
    def _reset_ui(self):
        """UI状態のリセット"""
        self.conversion_in_progress = False
        self.conversion_thread = None
        self.convert_btn.config(state=tk.NORMAL)
        self.pdf_btn.config(state=tk.NORMAL)
        self.output_btn.config(state=tk.NORMAL)
    
    def _on_close(self):
        """アプリケーション終了時の処理"""
        if self.conversion_in_progress:
            if messagebox.askyesno("確認", "変換処理が実行中です。本当に終了しますか？"):
                # 終了を強制する
                self.quit()
        else:
            self.quit()


def main():
    """アプリケーション起動"""
    try:
        app = PDF2PPTXApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("起動エラー", f"アプリケーションの起動中にエラーが発生しました:\n{str(e)}")
        raise e


if __name__ == "__main__":
    main()
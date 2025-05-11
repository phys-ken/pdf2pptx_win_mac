"""
GUIアプリケーションのテストモジュール
"""
import os
import sys
import unittest
import tempfile
from unittest.mock import patch, MagicMock

# 親ディレクトリをsys.pathに追加して、src内のモジュールをインポートできるようにする
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# GUIモジュールのインポート
from pdf2pptx_gui import PDF2PPTXApp
from pdf_converter import PDFConverter


class TestPDF2PPTXGUIFunctions(unittest.TestCase):
    """GUIアプリケーションの機能テスト"""
    
    def setUp(self):
        """テスト前の準備"""
        # GUIを作成せずにテストするため、tkinterをモックに置き換える
        self.tk_mock = MagicMock()
        self.filedialog_mock = MagicMock()
        self.messagebox_mock = MagicMock()
        
        # テスト用の一時ディレクトリ
        self.temp_dir = tempfile.mkdtemp()
        self.test_pdf_path = os.path.join(self.temp_dir, "test.pdf")
        
        # テスト用のPDFファイルを作成（空のファイル）
        with open(self.test_pdf_path, "w") as f:
            f.write("Test PDF file (not a real PDF)")
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        if os.path.exists(self.temp_dir):
            import shutil
            shutil.rmtree(self.temp_dir)
    
    @patch('pdf2pptx_gui.filedialog')
    def test_select_pdf(self, filedialog_mock):
        """PDFファイル選択のテスト"""
        # モックの設定
        filedialog_mock.askopenfilename.return_value = self.test_pdf_path
        
        # テスト対象クラスの準備
        with patch('pdf2pptx_gui.tk.Tk'):
            app = PDF2PPTXApp()
            app.pdf_path = None
            app.pdf_label = MagicMock()
            app.convert_btn = MagicMock()
            
            # メソッド実行
            app._select_pdf()
            
            # 検証
            filedialog_mock.askopenfilename.assert_called_once()
            self.assertEqual(app.pdf_path, self.test_pdf_path)
            app.pdf_label.config.assert_called_once()
            app.convert_btn.config.assert_called_once_with(state="normal")
    
    @patch('pdf2pptx_gui.filedialog')
    def test_select_output_folder(self, filedialog_mock):
        """出力先フォルダ選択のテスト"""
        # モックの設定
        filedialog_mock.askdirectory.return_value = self.temp_dir
        
        # テスト対象クラスの準備
        with patch('pdf2pptx_gui.tk.Tk'):
            app = PDF2PPTXApp()
            
            with patch('pdf2pptx_gui.messagebox') as messagebox_mock:
                # メソッド実行
                app._select_output_folder()
                
                # 検証
                filedialog_mock.askdirectory.assert_called_once()
                self.assertEqual(app.output_folder, self.temp_dir)
                messagebox_mock.showinfo.assert_called_once()
    
    @patch('pdf2pptx_gui.PDFConverter')
    def test_start_conversion(self, converter_mock):
        """変換開始のテスト"""
        # モックの設定
        converter_instance = MagicMock()
        converter_mock.return_value = converter_instance
        converter_instance.convert_pdf_to_pptx.return_value = (
            os.path.join(self.temp_dir, "test.pptx"),
            os.path.join(self.temp_dir, "test_figs")
        )
        
        # テスト用データ準備
        with patch('pdf2pptx_gui.tk.Tk'):
            app = PDF2PPTXApp()
            app.pdf_path = self.test_pdf_path
            app.output_folder = self.temp_dir
            app.progress = MagicMock()
            app.status_label = MagicMock()
            app.convert_btn = MagicMock()
            app.pdf_btn = MagicMock()
            app.output_btn = MagicMock()
            
            # スレッドをモック化して同期的に実行
            with patch('pdf2pptx_gui.threading.Thread') as thread_mock:
                thread_mock.return_value = MagicMock()
                
                # メソッド実行
                app._start_conversion()
                
                # スレッド開始の検証
                thread_mock.assert_called_once()
                thread_mock.return_value.start.assert_called_once()
                
                # GUIの更新検証
                app.convert_btn.config.assert_called_with(state="disabled")
                app.pdf_btn.config.assert_called_with(state="disabled")
                app.output_btn.config.assert_called_with(state="disabled")


def manual_full_test():
    """手動でのフルテスト（実際のGUIを表示）"""
    from pdf2pptx_gui import main
    main()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        # GUIを実際に起動してテスト
        manual_full_test()
    else:
        # ユニットテストを実行
        unittest.main()
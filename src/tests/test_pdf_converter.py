"""
PDFコンバーターのテストモジュール
"""
import os
import sys
import unittest
import tempfile
import shutil

# 親ディレクトリをsys.pathに追加して、src内のモジュールをインポートできるようにする
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_converter import PDFConverter


class TestPDFConverter(unittest.TestCase):
    """PDFConverterクラスのテスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.converter = PDFConverter()
        self.temp_dir = tempfile.mkdtemp()
        
        # テスト用の空のPDFファイル（実際のテストでは存在するPDFファイルが必要）
        self.test_pdf_path = os.path.join(self.temp_dir, "test.pdf")
        with open(self.test_pdf_path, "w") as f:
            f.write("PDF test file (not a real PDF)")
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_initialization(self):
        """初期化のテスト"""
        self.assertIsNone(self.converter.temp_folder)
        self.assertIsNone(self.converter.output_folder)
        self.assertEqual(self.converter.image_format, "jpg")
        self.assertEqual(self.converter.dpi, 300)
    
    def test_setup_temp_folder(self):
        """一時フォルダ設定のテスト"""
        self.converter._setup_temp_folder(self.test_pdf_path)
        self.assertTrue(os.path.exists(self.converter.temp_folder))
        self.assertTrue(os.path.isdir(self.converter.temp_folder))
    
    def test_cleanup_temp_folder(self):
        """一時フォルダ削除のテスト"""
        self.converter._setup_temp_folder(self.test_pdf_path)
        temp_folder = self.converter.temp_folder
        self.converter._cleanup_temp_folder()
        self.assertFalse(os.path.exists(temp_folder))
    
    def test_invalid_file_extension(self):
        """無効なファイル拡張子のテスト"""
        # 拡張子なしのファイル
        invalid_file = os.path.join(self.temp_dir, "invalid_file")
        with open(invalid_file, "w") as f:
            f.write("Not a PDF")
        
        with self.assertRaises(ValueError):
            self.converter.convert_pdf_to_pptx(invalid_file)
    
    def test_file_not_found(self):
        """存在しないファイルのテスト"""
        non_existent_file = os.path.join(self.temp_dir, "nonexistent.pdf")
        with self.assertRaises(FileNotFoundError):
            self.converter.convert_pdf_to_pptx(non_existent_file)


def manual_test_with_file(pdf_path):
    """手動テスト用の関数 - 実際のPDFファイルを使ってテスト"""
    print(f"PDFファイル '{pdf_path}' を使って変換テスト開始...")
    
    converter = PDFConverter()
    
    def progress_callback(status, message, progress=None):
        if progress is not None:
            print(f"{status}: {message} - {progress:.1f}%")
        else:
            print(f"{status}: {message}")
    
    try:
        # 出力先をテンポラリディレクトリに設定
        output_dir = tempfile.mkdtemp()
        
        # 変換を実行
        pptx_path, images_folder = converter.convert_pdf_to_pptx(
            pdf_path, 
            output_folder=output_dir, 
            callback=progress_callback
        )
        
        print("\nテスト結果:")
        print(f"PPTXファイル: {pptx_path}")
        print(f"画像フォルダ: {images_folder}")
        
        # 成功確認
        assert os.path.exists(pptx_path), "PPTXファイルが作成されていません"
        assert os.path.exists(images_folder), "画像フォルダが作成されていません"
        
        # ファイルサイズ確認
        pptx_size = os.path.getsize(pptx_path)
        print(f"PPTXファイルサイズ: {pptx_size} バイト")
        assert pptx_size > 0, "PPTXファイルが空です"
        
        # 画像ファイル確認
        image_files = os.listdir(images_folder)
        print(f"作成された画像: {len(image_files)}枚")
        assert len(image_files) > 0, "画像が作成されていません"
        
        print("テスト成功！")
        return True
        
    except Exception as e:
        print(f"テスト失敗: {str(e)}")
        return False
    finally:
        # テンポラリディレクトリは削除しない (結果を確認できるように)
        pass


if __name__ == "__main__":
    # コマンドライン引数からPDFファイルのパスを取得
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
        if os.path.exists(pdf_file) and pdf_file.lower().endswith('.pdf'):
            manual_test_with_file(pdf_file)
        else:
            print("指定されたPDFファイルが見つからないか、PDFファイルではありません。")
            print("使用法: python test_pdf_converter.py <PDFファイルパス>")
    else:
        # 通常のunittestを実行
        unittest.main()
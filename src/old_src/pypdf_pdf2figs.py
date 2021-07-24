from pdf2image import convert_from_path
import os

pdf_file = "in.pdf"

for current, subfolders, subfiles in os.walk('.'):
    for fileName in subfiles:
        base, ext = os.path.splitext(fileName)
        if ext == '.pdf' or ext == '.PDF' :
                fullPath = current + "/" + fileName
                print(fullPath + "の変換を準備中...")
                load = convert_from_path(fullPath)
                if not os.path.exists("fig"):
                  # ディレクトリが存在しない場合、ディレクトリを作成する
                  os.makedirs("fig")
                img_file = "fig/outputfig_" + base
                for index, jpg in enumerate(load):
                  jpg_name = img_file + '_' + '{0:03d}.jpg'.format(index)
                  jpg.save(jpg_name)
                  print("変換中_{0:03d}.jpg".format(index))
                print("変換する")
        else:
                print(base + ext + "_さよなら")



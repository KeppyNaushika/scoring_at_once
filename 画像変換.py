import tkinter
import tkinter.filedialog
import tkinter.messagebox

import PIL
import PIL.Image

import glob
import os
import pdf2image
import time

def main():
  print(f"画像変換 for 一括採点 ver b.1.0\n\nCtrl+C で終了します\n")

  dict_length = {
    "A3": {
      "width": 800,
      "height": 1131
    },
    "A4": {
      "width": 567,
      "height": 800
    },
    "A5": {
      "width": 400,
      "height": 567
    },
    "A6": {
      "width": 283,
      "height": 400
    },
    "B4": {
      "width": 693,
      "height": 980
    },
    "B5": {
      "width": 490,
      "height": 693
    },
    "B6": {
      "width": 300,
      "height": 490
    },
    "B7": {
      "width": 245,
      "height": 300
    }
  }

  # select mode
  str_input = None
  while str_input not in ["1", "2"]:
    print(f"＝＝＝＝＝＝＝＝＝＝")
    print(f"変換モードを指定します")
    # print(f"｜1: 指定する1つの PDF ファイルを読み込みます")
    print(f"｜2: 指定するフォルダ内に含まれる全ての画像ファイルを読み込みます")
    # str_input = input(f"(1/2) >>> ")
    print(f"(1/2) >>> 2")
    str_input = 2

  # input files
  print(f"＝＝＝＝＝＝＝＝＝＝")
  if str_input == "1":
    print(f"変換元のファイルを指定します")
    path_pdf = tkinter.filedialog.askopenfilename(
      title="変換元の PDF ファイルを指定します",
      filetypes=[("PDF ドキュメント", ".pdf")],
      defaultextension="pdf"
    )
    if path_pdf in [None, ""]:
      print(f"ファイルが指定されませんでした")
      input(f"Enter キーを押して終了します >>> ")
      return
    print(f"ファイル: {path_pdf}")
    list_image = pdf2image.convert_from_path(path_pdf, poppler_path=f"{os.path.dirname(__file__)}/poppler-22.01.0/Library/bin")

  else:
    print(f"変換元のファイルが保存されているフォルダを指定します")
    path_input_dir = tkinter.filedialog.askdirectory(
      title="変換元の PDF ファイルを指定します"
    )
    if path_input_dir in [None, ""]:
      print(f"ファイルが指定されませんでした")
      input(f"Enter キーを押して終了します >>> ")
      return
    print(f"フォルダ: {path_input_dir}")
    list_image = [PIL.Image.open(path_file) for path_file in glob.glob(path_input_dir + "/*") if os.path.splitext(path_file)[1] in [".jpeg", ".jpg", ".png"]]
  print(f"{len(list_image)} 枚の画像を読み込みました\n")
  print(f"＝＝＝＝＝＝＝＝＝＝")

  # set pages
  str_pages = None
  while str_pages not in [str(i+1) for i in range(10)]:
    print(f"連続する答案の枚数を指定します")
    str_pages = input(F"(1 - 10) >>> ")
  list_size = []
  for index_page in range(int(str_pages)):
    str_size = None
    while str_size not in ["A3", "A4", "A5", "A6", "B4", "B5", "B6", "B7"]:
      print(f"{index_page + 1}枚目の答案用紙のサイズを指定します")
      print(f"｜次のいずれかから指定します")
      print(f"｜A3 / A4 / A5 / A6 / B4 / B5 / B6 / B7")
      str_size = input(f">>> ")
    print(f"{index_page + 1}枚目のサイズを{str_size}に指定しました\n")
    str_orient = None
    while str_orient not in ["tate", "yoko"]:
      print(f"{index_page + 1}枚目の答案用紙の向きを指定します")
      print(f"｜次のいずれかから指定します")
      print(f"｜tate / yoko")
      str_orient = input(f">>> ")
    if str_orient == "tate":
      list_size.append((dict_length[str_size]["width"], dict_length[str_size]["height"]))
    else:
      list_size.append((dict_length[str_size]["height"], dict_length[str_size]["width"]))
    print(f"{index_page + 1}枚目の向きを{str_orient}に指定しました\n")

  if len(list_size) == 1:
    list_output_image = list_image
  else:
    str_composite = None
    while str_composite not in ["1", "2", "3", "4"]:
      print(f"複数枚の答案を結合する方向を指定します")
      print(f"｜1: → 左から右 / 2: ↓ 上から下 / 3: ← 右から左 / 4: ↑ 下から上")
      str_composite = input(f"(1 - 4) >>> ")
  
  # 書き出し
  print(f"＝＝＝＝＝＝＝＝＝＝")
  print(f"出力先のフォルダを指定します")
  path_output_dir = tkinter.filedialog.askdirectory(
    title="出力先のフォルダを指定します"
  )
  if not isinstance(path_output_dir, str):
    print(f"フォルダが指定されませんでした")
    input(f"Enter キーを押して終了します >>> ")
    return
  print(f"フォルダ: {path_output_dir}")
  time.sleep(1)
  print(f"")
  index_save = 0
  for index_image, image in enumerate(list_image):
    if index_image % int(str_pages) == 0:
      if str_composite in ["1", "3"]:
        image_new = PIL.Image.new("RGB", (sum([size[0] for size in list_size]), max([size[1] for size in list_size])), (255, 255, 255, 0))
      elif str_composite in ["2", "4"]:
        image_new = PIL.Image.new("RGB", (max([size[0] for size in list_size]), sum([size[1] for size in list_size])), (255, 255, 255, 0))
    image_resized = image.resize((list_size[index_image % int(str_pages)][0], list_size[index_image % int(str_pages)][1]))
    if str_composite in ["1"]:
      image_new.paste(image_resized, (sum([size[0] for size in list_size[:index_image % int(str_pages)]]), 0))
    elif str_composite in ["2"]:
      image_new.paste(image_resized, (0, sum([size[1] for size in list_size[:index_image % int(str_pages)]])))
    elif str_composite in ["3"]:
      image_new.paste(image_resized, (sum([size[0] for size in list_size]) - sum([size[0] for size in list_size[:index_image % int(str_pages) + 1]]), 0))
    elif str_composite in ["4"]:
      image_new.paste(image_resized, (0, sum([size[1] for size in list_size]) - sum([size[1] for size in list_size[:index_image % int(str_pages) + 1]])))
    if index_image % int(str_pages) == int(str_pages) - 1:
      image_new.save(f"{path_output_dir}/{index_save}.png")
      index_save += 1
      print(f"\033[1A{index_save} 枚の画像を出力しました")
  print(f"正常に終了しました")
  time.sleep(1)
  input(f"Enter キーを押して終了します...")  

if __name__ == "__main__":
  main()
import tkinter
import tkinter.filedialog
import tkinter.messagebox

import PIL
import PIL.Image

import glob
import natsort
import os
import pdf2image
import sys
import time

def main():
  print(f"画像変換 for 一括採点\n\nCtrl+C で終了します")
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

  # project name
  print(f"\n＝＝＝＝＝＝＝＝＝＝")
  print("変換先のファイル名の連番の先頭に入れる文字列を入力して下さい")
  print("不要な場合は何も入力せず Enter キーを入力して下さい")
  str_project = input(">>> ")

  # select mode
  str_input = None
  while str_input not in ["1", "2"]:
    print(f"\n＝＝＝＝＝＝＝＝＝＝")
    print(f"変換モードを指定します")
    print(f"｜1: 指定する1つの PDF ファイルを読み込みます")
    print(f"｜2: 指定するフォルダ内に含まれる全ての画像ファイルを読み込みます")
    str_input = input(f"(1/2) >>> ")

  # input files
  print(f"\n＝＝＝＝＝＝＝＝＝＝")
  if str_input == "1":
    if not os.path.exists(f"{os.path.dirname(__file__)}/poppler-22.01.0/Library/bin"):
      print(f"poppler が存在しないため PDF を変換できません。変換モードを変更し画像ファイルを読み込んで下さい。")
      return False
    print(f"変換元のファイルを指定します")
    path_pdf = tkinter.filedialog.askopenfilename(
      title="変換元の PDF ファイルを指定します",
      filetypes=[("PDF ドキュメント", ".pdf")],
      defaultextension="pdf"
    )
    if path_pdf in [None, ""]:
      print(f"ファイルが指定されませんでした")
      return False
    print(f"ファイル: {path_pdf}")
    sys.stdout.write(f"ファイルを読み込んでいます。PC の性能と PDF ファイルの状態によっては、数分かかる場合があります...")
    sys.stdout.flush()
    list_image = pdf2image.convert_from_path(path_pdf, poppler_path=f"{os.path.dirname(__file__)}/poppler-22.01.0/Library/bin", thread_count=4)
    sys.stdout.write(f"\rファイルを読み込みが完了しました                                                                \r")
    sys.stdout.flush()


  else:
    print(f"変換元のファイルが保存されているフォルダを指定します")
    path_input_dir = tkinter.filedialog.askdirectory(
      title="変換元の PDF ファイルを指定します"
    )
    if path_input_dir in [None, ""]:
      print(f"ファイルが指定されませんでした")
      return False
    print(f"フォルダ: {path_input_dir}")
    sys.stdout.write(f"ファイルを読み込んでいます。PC の性能と PDF ファイルの状態によっては、数分かかる場合があります...")
    sys.stdout.flush()
    list_image = [PIL.Image.open(path_file) for path_file in natsort.natsorted(glob.glob(path_input_dir + "/*")) if os.path.splitext(path_file)[1] in [".jpeg", ".jpg", ".png"]]
    sys.stdout.write(f"\rファイルを読み込みが完了しました                                                                \r")
    sys.stdout.flush()
  print(f"\n{len(list_image)} 枚の画像を読み込みました")

  # set pages
  print(f"\n＝＝＝＝＝＝＝＝＝＝")
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
    print(f"{index_page + 1}枚目の向きを{str_orient}に指定しました")

  if len(list_size) == 1:
    str_composite = "1"
  else:
    str_composite = None
    while str_composite not in ["1", "2", "3", "4"]:
      print(f"複数枚の答案を結合する方向を指定します")
      print(f"｜1: → 左から右 / 2: ↓ 上から下 / 3: ← 右から左 / 4: ↑ 下から上")
      str_composite = input(f"(1 - 4) >>> ")
  
  # 書き出し
  print(f"\n＝＝＝＝＝＝＝＝＝＝")
  print(f"出力先のフォルダを指定します")
  path_output_dir = tkinter.filedialog.askdirectory(
    title="出力先のフォルダを指定します"
  )
  if path_output_dir in [None, ""]:
    print(f"フォルダが指定されませんでした")
    return False
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
    try:
      image_resized = image.resize((list_size[index_image % int(str_pages)][0], list_size[index_image % int(str_pages)][1]))
    except OSError:
      print(f"\nファイルの保存に関するエラーが発生しました")
      print(f"保存するファイルに書き込み権限が存在しない可能性があります")
      print(f"別のフォルダを指定して下さい")
      return False
    if str_composite in ["1"]:
      image_new.paste(image_resized, (sum([size[0] for size in list_size[:index_image % int(str_pages)]]), 0))
    elif str_composite in ["2"]:
      image_new.paste(image_resized, (0, sum([size[1] for size in list_size[:index_image % int(str_pages)]])))
    elif str_composite in ["3"]:
      image_new.paste(image_resized, (sum([size[0] for size in list_size]) - sum([size[0] for size in list_size[:index_image % int(str_pages) + 1]]), 0))
    elif str_composite in ["4"]:
      image_new.paste(image_resized, (0, sum([size[1] for size in list_size]) - sum([size[1] for size in list_size[:index_image % int(str_pages) + 1]])))
    if index_image % int(str_pages) == int(str_pages) - 1:
      image_new.save(f"{path_output_dir}/{str_project}{str(index_save).zfill(5)}.png")
      index_save += 1
      sys.stdout.write(f"\r{index_save}枚 / {len(list_image) // int(str_pages)}枚の画像を出力しました")
      sys.stdout.flush()
  return True

if __name__ == "__main__":
  bool_return = main()
  if bool_return:
    print(f"\n正常に終了しました")
  time.sleep(1)
  input(f"Enter キーを押して終了します...")
#####################################################
#                                                   #
#          Copyright(c) 2022 KeppyNaushika          #
#                                                   #
#        This software is released under the        #
#      GNU Affero General Public License v3.0,      #
#                    see LICENSE.                   #
#                                                   #
#      The github repository of this software:      #
# https://github.com/KeppyNaushika/scoring_at_once/ #
#                                                   #
#####################################################

import tkinter
import tkinter.filedialog
import tkinter.font
import tkinter.messagebox

import PIL
import PIL.Image
import PIL.ImageTk

import functools
import glob
import json
import os
import subprocess

def nothing_to_do():
  tkinter.messagebox.showinfo(
    "未実装", "ｱﾋｬﾋｬﾋｬﾋｬﾋｬ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ｱﾋｬﾋｬﾋｬﾋｬ\nｱﾋｬﾋｬﾋｬﾋｬﾋｬ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ｱﾋｬﾋｬﾋｬﾋｬ\nｱﾋｬﾋｬﾋｬﾋｬﾋｬ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ(ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ﾟ∀ﾟ)ｱﾋｬﾋｬﾋｬﾋｬ\n"
  )

# class: 子ウインドウ:
class SubWindow:
  def __init__(self, parent) -> None:
    self.parent = parent
    self.window = None  

  def this_window_close(self):
    self.window.withdraw()
    self.window = None
    self.parent.destroy()
    main()
    return "break"

  def sub_window_loop(func):
    def inner(self, *args, **kargs):
      self.parent.withdraw()
      if self.window:
        self.window.lift()
      else:
        self.window = tkinter.Toplevel(self.parent)
        self.window.title("一括採点")
        self.window.withdraw()
        if func(self, *args, **kargs) is None:
          self.window.deiconify()
          self.window.protocol("WM_DELETE_WINDOW", self.this_window_close)
          self.window.mainloop()
    return inner

  def check_dir_exist(self):
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
    name_project: str = dict_project["name"]
    path_dir    : str = dict_project["path_dir"]
    path_file   : str = dict_project["path_file"]
    if not os.path.exists(path_dir):
      tkinter.messagebox.showwarning(
        "フォルダが存在しません", 
        f"指定されたフォルダが存在しなかったため, フォルダを開くことができませんでした. \n"
        + f"答案スキャンデータが保存されているフォルダのパスが正しいことを確認して下さい. \n\n"
        + f"試験名: {name_project}"
      )
      return False
    if not os.path.exists(path_file):
      tkinter.messagebox.showwarning(
        "ファイルが存在しません", 
        f"指定されたファイルが存在しなかったため, ファイルを開くことができませんでした. \n"
        + f"模範解答スキャンデータが保存されているファイルのパスが正しいことを確認して下さい. \n\n"
        + f"試験名: {name_project}"
      )
      return False
    if not os.path.splitext(path_file)[1] in [".jpeg", ".jpg", ".png"]:
      print(os.path.splitext(path_file)[1])
      tkinter.messagebox.showwarning(
        "ファイルの拡張子が対応しません", 
        f"指定されたファイルの拡張子が jpeg, jpg, png 以外であったため, ファイルを開きませんでした. \n"
        + f"模範解答スキャンデータが保存されているファイル名が正しいことを確認して下さい. \n"
        + f"ファイルの形式が正しくない場合は, 外部のアプリケーションを利用してファイルを変換して下さい. \n"
        + f"ファイルの形式が正しい場合は, 拡張子を変更した上でもう一度実行して下さい. \n\n"
        + f"試験名: {name_project}"
      )
      return False
    if not os.path.exists(path_dir + "/.temp_saiten"):
      os.mkdir(path_dir + "/.temp_saiten")
      subprocess.check_call(["attrib", "+H", path_dir + "/.temp_saiten"])
    if not os.path.exists(path_dir + "/.temp_saiten"):
      os.mkdir(path_dir + "/.temp_saiten")
    if not os.path.exists(path_dir + "/.temp_saiten/answer_area.json"):
      dict_answer_area = {"id": None, "questions": []}
      with open(path_dir + "/.temp_saiten/answer_area.json", "w", encoding="utf-8") as f:
        json.dump(dict_answer_area, f, indent=2)
    if not os.path.exists(path_dir + "/.temp_saiten/model_answer"):
      os.mkdir(path_dir + "/.temp_saiten/model_answer")
    if not os.path.exists(path_dir + "/.temp_saiten/model_answer/model_answer.png"):
      if os.path.splitext(path_file)[1] in [".jpeg", ".jpg", ".png"]:
        img = PIL.Image.open(path_file)
        img.save(path_dir + "/.temp_saiten/model_answer/model_answer.png")
    if not os.path.exists(path_dir + "/.temp_saiten/answer"):
      os.mkdir(path_dir + "/.temp_saiten/answer")
    if os.path.exists(path_dir + "/.temp_saiten/load_picture.json"):
      with open(path_dir + "/.temp_saiten/load_picture.json", "r", encoding="utf-8") as f:
        dict_load_picture = json.load(f)
    else:
      dict_load_picture = {
        "answer": []
      }
    list_path_in_file_dir = [path.replace("\\", "/") for path in glob.glob(path_dir + "/*")]
    index_file = len(dict_load_picture["answer"])
    for path_in_file_dir in list_path_in_file_dir:
      if path_in_file_dir == path_file:
        continue
      elif path_in_file_dir in dict_load_picture["answer"]:
        continue
      elif os.path.splitext(path_in_file_dir)[1] in [".jpeg", ".jpg", ".png"]:
        img = PIL.Image.open(path_in_file_dir)
        img.save(path_dir + "/.temp_saiten/answer/" + str(index_file) + ".png")
        dict_load_picture["answer"].append(path_in_file_dir)        
      else:
        continue
      index_file += 1
    with open(path_dir + "/.temp_saiten/load_picture.json", "w", encoding="utf-8") as f:
      json.dump(dict_load_picture, f, indent=2)
    if index_file == 0:
      tkinter.messagebox.showwarning(
        "ファイルが存在しません", 
        f"指定されたフォルダ内に, 拡張子が *.jpeg, *.jpg, *.png であるファイルが存在しません. \n"
        + f"答案スキャンデータが保存されているフォルダ名が正しいことを確認して下さい. \n"
        + f"ファイルの形式が正しくない場合は, 外部のアプリケーションを利用してファイルを変換して下さい. \n"
        + f"ファイルの形式が正しい場合は, 拡張子を変更した上でもう一度実行して下さい. \n\n"
        + f"試験名: {name_project}"
      )
      return False
    tkinter.messagebox.showinfo(
      "答案スキャンデータが読み込みました", 
      f"{index_file} 件の答案スキャンデータが読みまれています. \n\n"
      + f"読み込まれるスキャンデータが少ない場合は以下の手順で確認して下さい. \n"
      + f"1. メインウインドウの［編集］ボタンをクリックして, 「試験を編集」ウインドウを開きます. \n"
      + f"2. 答案スキャンデータの保存されているフォルダのパスが正しいことを確認して下さい. \n"
      + f"3. 答案スキャンデータとして使用できるファイルは JPEG または PNG です. 拡張子が *.jpeg, *.jpg, *.png 以外のファイルは無視されます. \n"
      + f"4. ［適用］をクリックして, 答案データを再読み込みします. "
    )
    return True

  @sub_window_loop
  def add_project(self):
    def choose_dir():
      entry_path_dir.delete(0, "end")
      entry_path_dir.insert(0, tkinter.filedialog.askdirectory())
      self.window.lift()
    
    def choose_file():
      entry_path_file.delete(0, "end")
      entry_path_file.insert(0, tkinter.filedialog.askopenfilename())
      self.window.lift()

    def add_json():
      str_name = entry_name.get()
      str_path_dir = entry_path_dir.get()
      str_path_file = entry_path_file.get()
      if str_name == "":
        tkinter.messagebox.showwarning("試験名が入力されていません", "試験名が入力されていないため, 新しく試験を作成できません. \n試験名を入力して下さい. ")
        self.window.lift()
        return
      if str_path_dir == "":
        tkinter.messagebox.showwarning("フォルダパスが指定されていません", "答案ファイルが保存されているフォルダパスが指定されていないため, 新しく試験を作成できません. \nフォルダパスを指定して下さい. ")
        self.window.lift()
        return
      if str_path_file == "":
        tkinter.messagebox.showwarning("ファイルパスが指定されていません", "模範解答が保存されているファイルパスが指定されていないため, 新しく試験を追加できません. \nファイルパスを指定して下さい. ")
        self.window.lift()
        return
      with open("config.json", "r", encoding="utf-8") as f:
        dict_config = json.load(f)
      dict_config["projects"].append(
        {
          "name": str_name,
          "path_dir": str_path_dir,
          "path_file": str_path_file
        }
      )
      dict_config["index_projects_in_listbox"] = len(dict_config["projects"]) - 1
      with open("config.json", "w", encoding="utf-8") as f:
        json.dump(dict_config, f, indent=2)
      if not self.check_dir_exist():
        with open("config.json", "r", encoding="utf-8") as f:
          dict_config = json.load(f)
        dict_config["projects"].pop(len(dict_config["projects"]) - 1)
        if len(dict_config["projects"]) == 0:
          dict_config["index_projects_in_listbox"] = None
        else:
          dict_config["index_projects_in_listbox"] = 0
        with open("config.json", "w", encoding="utf-8") as f:
          json.dump(dict_config, f, indent=2)
        self.window.lift()
        return
      self.this_window_close()
      tkinter.messagebox.showinfo(
        "試験を追加しました",
        "採点データ等は, 指定したフォルダ内に作成された隠しフォルダ「.temp_saiten」内に保存されます. \n"
        + "予期せぬ動作を防ぐため, 本アプリ起動中は「.temp_saiten」や指定したフォルダを移動, 削除しないで下さい. "
      )

    self.window.title("試験を追加")
    frame_main = tkinter.Frame(self.window)
    frame_main.pack(expand=True, padx=20, pady=20)

    frame_form = tkinter.Frame(frame_main, width=80, height=10)
    label_vspace = tkinter.Label(frame_main, width=100, height=2)
    frame_btn = tkinter.Frame(frame_main, width=80, height=10)
    frame_form.grid(column=0, row=0)
    label_vspace.grid(column=0, row=1)
    frame_btn.grid(column=0, row=2)
    
    label_name = tkinter.Label(frame_form, width=80, text="試験名")
    label_name.grid(column=0, row=0)
    entry_name = tkinter.Entry(frame_form, width=80)
    entry_name.grid(column=0, row=1)
    label_vspace1 = tkinter.Label(frame_form, width=80, height=1)
    label_vspace1.grid(column=0, row=2)
    label_path_dir = tkinter.Label(frame_form, width=80, text="答案スキャンデータが保存されているフォルダのパス")
    label_path_dir.grid(column=0, row=3)
    frame_path_dir = tkinter.Frame(frame_form, width=80)
    frame_path_dir.grid(column=0, row=4)
    label_vspace2 = tkinter.Label(frame_form, width=80, height=1)
    label_vspace2.grid(column=0, row=5)
    label_path_file = tkinter.Label(frame_form, width=80, text="模範解答スキャンデータが保存されているファイルのパス")
    label_path_file.grid(column=0, row=6)
    frame_path_file = tkinter.Frame(frame_form, width=80)
    frame_path_file.grid(column=0, row=7)

    entry_path_dir = tkinter.Entry(frame_path_dir, width=60, textvariable="")
    entry_path_dir.grid(column=0, row=0)
    label_hspace_dir = tkinter.Label(frame_path_dir, width=3)
    label_hspace_dir.grid(column=1, row=0)
    btn_path_dir = tkinter.Button(frame_path_dir, width=15, text="フォルダを選択", command=choose_dir)
    btn_path_dir.grid(column=2, row=0)

    entry_path_file = tkinter.Entry(frame_path_file, width=60, textvariable="")
    entry_path_file.grid(column=0, row=0)
    label_hspace_file = tkinter.Label(frame_path_file, width=3)
    label_hspace_file.grid(column=1, row=0)
    btn_path_file = tkinter.Button(frame_path_file, width=15, text="模範解答を選択", command=choose_file)
    btn_path_file.grid(column=2, row=0)
    
    tkinter.Button(frame_btn, text="試験を追加", command=add_json, width=40, height=2).grid(column=0, row=0)
    tkinter.Button(frame_btn, text="キャンセル", command=self.this_window_close, width=40, height=2).grid(column=1, row=0)

  def edit_project(self):
    nothing_to_do()
    
  # 解答欄の位置を指定
  @sub_window_loop
  def select_area(self):
    if not self.check_dir_exist():
      tkinter.messagebox.showinfo(
        "設定を確認して下さい", 
        f"試験一覧の［編集］ボタンをクリックして, 試験の設定を確認して下さい. \n\n「解答欄の位置を指定」を終了します. "
      )
      return "break"

    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
    path_json_answer_area = dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
    path_file_model_answer = dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
    path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
    with open(path_json_answer_area, "r", encoding="utf-8") as f:
      dict_answer_area = json.load(f)

    def del_question(index_selected_question=None):
      self.index_selected_question = index_selected_question
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        dict_answer_area["questions"].pop(self.index_selected_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        load_question(self.index_selected_question)

    def up_question(index_selected_question=None):
      self.index_selected_question = index_selected_question
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        pop_question = dict_answer_area["questions"].pop(self.index_selected_question)
        self.index_selected_question = max(self.index_selected_question - 1, 0)
        dict_answer_area["questions"].insert(self.index_selected_question, pop_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        load_question(self.index_selected_question)

    def down_question(index_selected_question=None):
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        pop_question = dict_answer_area["questions"].pop(self.index_selected_question)
        self.index_selected_question = min(self.index_selected_question + 1, len(dict_answer_area["questions"]) - 1)
        dict_answer_area["questions"].insert(self.index_selected_question, pop_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        load_question(self.index_selected_question)

    def set_type(str_type):
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        dict_answer_area["questions"][self.index_selected_question]["type"] = str_type
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        load_question(self.index_selected_question)
    def set_question():
      set_type("設問")
    def set_name():
      set_type("氏名")
    def set_id():
      set_type("生徒番号")
    def set_stamp():
      set_type("採点者印")
    def set_subtotal():
      set_type("小計点")
    def set_total():
      set_type("合計点")

    
    self.canvas_model_answer_draw_rectangle = [0, 0, 0, 0]
    def canvas_model_answer_draw_rectangle_click(event):
      self.canvas_model_answer_draw_rectangle[0] = event.x
      self.canvas_model_answer_draw_rectangle[1] = event.y
      self.canvas_model_answer_draw_rectangle[2] = min(event.x + 1, canvas_model_answer.winfo_width())
      self.canvas_model_answer_draw_rectangle[3] = min(event.y + 1, canvas_model_answer.winfo_height())
      canvas_model_answer.coords("rectangle_new",
        self.canvas_model_answer_draw_rectangle[0],
        self.canvas_model_answer_draw_rectangle[1],
        self.canvas_model_answer_draw_rectangle[2],
        self.canvas_model_answer_draw_rectangle[3], 
      )
    def canvas_model_answer_draw_rectangle_drag(event):
      self.canvas_model_answer_draw_rectangle[2] = min(max(event.x, 0), canvas_model_answer.winfo_width())
      self.canvas_model_answer_draw_rectangle[3] = min(max(event.y, 0), canvas_model_answer.winfo_height())
      canvas_model_answer.coords("rectangle_new",
        self.canvas_model_answer_draw_rectangle[0],
        self.canvas_model_answer_draw_rectangle[1],
        self.canvas_model_answer_draw_rectangle[2],
        self.canvas_model_answer_draw_rectangle[3], 
      )
    def canvas_model_answer_draw_rectangle_release(event):
      with open(path_json_answer_area, "r", encoding="utf-8") as f:
        dict_answer_area = json.load(f)
      dict_answer_area["questions"].append(
        {
          "type": "設問", 
          "daimon": "",
          "shomon": "",
          "shimon": "",
          "haiten": None,
          "area": [
            min(self.canvas_model_answer_draw_rectangle[0], self.canvas_model_answer_draw_rectangle[2]),
            min(self.canvas_model_answer_draw_rectangle[1], self.canvas_model_answer_draw_rectangle[3]),
            max(self.canvas_model_answer_draw_rectangle[0], self.canvas_model_answer_draw_rectangle[2]),
            max(self.canvas_model_answer_draw_rectangle[1], self.canvas_model_answer_draw_rectangle[3])
          ]
        }
      )
      with open(path_json_answer_area, "w", encoding="utf-8") as f:
        json.dump(dict_answer_area, f, indent=2)
      self.index_selected_question = len(dict_answer_area["questions"]) - 1
      load_question()
      canvas_model_answer.coords("rectangle_new", 0, 0, 0, 0)
      
    def load_question(index_selected_question=None, event=None):
      def color_question_of_type(str_type):
        if str_type == "設問":
          return "green"
        elif str_type == "氏名":
          return "blue"
        elif str_type == "生徒番号":
          return "cyan"
        elif str_type == "小計点":
          return "magenta"
        elif str_type == "合計点":
          return "orange"
        elif str_type == "採点者印":
          return "red"

      canvas_question.delete("frame_question")
      canvas_model_answer.delete("field")
      canvas_model_answer.delete("number")
      with open(path_json_answer_area, "r", encoding="utf-8") as f:
        dict_answer_area = json.load(f)
      self.index_selected_question = index_selected_question

      # canvas_question への描画
      if len(dict_answer_area["questions"]) == 0:
        self.index_selected_question = None
        tkinter.Label(canvas_question, width=40, text="模範解答上で\nドラッグアンドドロップして\n枠を指定します", tags="question")
      else:
        list_frame_border_question = [tkinter.Frame(canvas_question, borderwidth=2, bg=color_question_of_type(question["type"])) for question in dict_answer_area["questions"]]
        for index_frame_border_question, frame_border_question in enumerate(list_frame_border_question):
          if index_frame_border_question == self.index_selected_question:
            list_frame_border_question[index_frame_border_question].configure(borderwidth=4)
        for index_question in range(len(dict_answer_area["questions"])):
          canvas_question.create_window(5, index_question * 40 + 4, window=list_frame_border_question[index_question], anchor=tkinter.NW, tags="frame_question")
        list_frame_question = [tkinter.Frame(list_frame_border_question[index_question], borderwidth=4, bg="white") for index_question in range(len(dict_answer_area["questions"]))]
        for index_frame_question, frame_question in enumerate(list_frame_question):
          frame_question.pack()
          if index_frame_question == self.index_selected_question:
            list_frame_question[index_frame_question].configure(borderwidth=2)
        list_label_question_index  = [tkinter.Label(list_frame_question[index_question], width=4, text=index_question) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_label_question_type   = [tkinter.Label(list_frame_question[index_question], width=10, text=question["type"]) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_entry_question_haiten = [tkinter.Entry(list_frame_question[index_question], width=5, text=str(question["haiten"])) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_entry_question_daimon = [tkinter.Entry(list_frame_question[index_question], width=5, text=question["daimon"]) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_entry_question_shomon = [tkinter.Entry(list_frame_question[index_question], width=5, text=question["shomon"]) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_entry_question_shimon = [tkinter.Entry(list_frame_question[index_question], width=5, text=question["shimon"]) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_btn_question_bold     = [tkinter.Button(list_frame_question[index_question], width=5, text="強調", command=functools.partial(load_question, index_question)) for index_question, question in enumerate(dict_answer_area["questions"])]
        list_btn_question_up       = [tkinter.Button(list_frame_question[index_question], width=5, text="上へ") for index_question, question in enumerate(dict_answer_area["questions"])]
        list_btn_question_down     = [tkinter.Button(list_frame_question[index_question], width=5, text="下へ") for index_question, question in enumerate(dict_answer_area["questions"])]
        list_btn_question_del      = [tkinter.Button(list_frame_question[index_question], width=5, text="削除") for index_question, question in enumerate(dict_answer_area["questions"])]
        for index_question in range(len(dict_answer_area["questions"])):
          list_label_question_index[index_question].grid(column=0, row=0)
          list_label_question_type[index_question].grid(column=1, row=0)
          list_entry_question_haiten[index_question].grid(column=2, row=0)
          list_entry_question_daimon[index_question].grid(column=3, row=0)
          list_entry_question_shomon[index_question].grid(column=4, row=0)
          list_entry_question_shimon[index_question].grid(column=5, row=0)
          list_btn_question_bold[index_question].grid(column=6, row=0)
          list_btn_question_up[index_question].grid(column=7, row=0)
          list_btn_question_down[index_question].grid(column=8, row=0)
          list_btn_question_del[index_question].grid(column=9, row=0)

      # canvas_model_answer への描画
      for index_question, question in enumerate(dict_answer_area["questions"]):
        color_reactangle = color_question_of_type(question["type"])
        if index_question == index_selected_question:
          color_width = 3
        else:
          color_width = 1
        canvas_model_answer.create_rectangle(
          question["area"][0], 
          question["area"][1], 
          question["area"][2], 
          question["area"][3], 
          outline=color_reactangle,
          width=color_width,
          fill=color_reactangle,
          stipple="gray12",
          tags="field"
        )
        canvas_model_answer.create_text(
          question["area"][0] - 10, 
          (question["area"][1] + question["area"][3]) // 2, 
          text=str(index_question),
          fill="green",
          tags="number"
        )


    self.scale_canvas_model_answer = 1.0
    def canvas_model_answer_scale_up(self):
      self.scale_canvas_model_answer += 0.1
    
    self.window.title("解答欄を指定")
    frame_main = tkinter.Frame(self.window)
    frame_main.grid(column=0, row=0)

    frame_question = tkinter.Frame(frame_main)
    frame_question.grid(column=0, row=0)
    frame_model_answer = tkinter.Frame(frame_main)
    frame_model_answer.grid(column=1, row=0)

    frame_table_question = tkinter.Frame(frame_question)
    frame_table_question.pack()
    frame_btn_list_question = tkinter.Frame(frame_question)
    frame_btn_list_question.pack()


    canvas_question = tkinter.Canvas(frame_table_question, bg="black", width=567, height=700)
    canvas_question.grid(row=0, column=0, sticky=(tkinter.N, tkinter.S))
    canvas_question.bind("<MouseWheel>", lambda eve:canvas_question.yview_scroll(int(-eve.delta/120), 'units'))
    yscrollbar_canvas_question = tkinter.Scrollbar(frame_table_question, orient=tkinter.VERTICAL, command=canvas_question.yview)
    yscrollbar_canvas_question.grid(row=0, column=1, sticky=(tkinter.N, tkinter.S))
    canvas_question.config(
      yscrollcommand=yscrollbar_canvas_question.set,
      scrollregion=(0, 0, 400, 1500)
    )

    btn_scale_up = tkinter.Button(frame_btn_list_question, width=6, text="拡大")
    btn_scale_up.grid(column=0, row=3)
    btn_scale_reset = tkinter.Button(frame_btn_list_question, width=6, text="100%")
    btn_scale_reset.grid(column=1, row=3)
    btn_scale_down = tkinter.Button(frame_btn_list_question, width=6, text="縮小")
    btn_scale_down.grid(column=2, row=3)
    btn_scale_mode = tkinter.Button(frame_btn_list_question, width=21, text="[ドラッグ] / 自動")
    btn_scale_mode.grid(column=0, row=4, columnspan=3)
    btn_scale_help = tkinter.Button(frame_btn_list_question, width=21, text="ヘルプ")
    btn_scale_help.grid(column=0, row=5, columnspan=3)
    btn_scale_back = tkinter.Button(frame_btn_list_question, width=21, text="戻る", command=self.this_window_close)
    btn_scale_back.grid(column=0, row=6, columnspan=3)      

    # canvas_model_answer: 模範解答画像
    frame_canvas_model_answer = tkinter.Frame(frame_model_answer)
    frame_canvas_model_answer.pack()

    canvas_model_answer = tkinter.Canvas(frame_canvas_model_answer, bg="black", width=567, height=700)
    canvas_model_answer.bind("<Control-MouseWheel>", lambda eve:canvas_model_answer.xview_scroll(int(-eve.delta/120), 'units'))
    canvas_model_answer.bind("<MouseWheel>", lambda eve:canvas_model_answer.yview_scroll(int(-eve.delta/120), 'units'))
    self.tk_image_model_answer = PIL.ImageTk.PhotoImage(file=path_file_model_answer)
    canvas_model_answer.create_image(0, 0, image=self.tk_image_model_answer, anchor="nw")
    yscrollbar_canvas_model_answer = tkinter.Scrollbar(frame_canvas_model_answer, orient=tkinter.VERTICAL, command=canvas_model_answer.yview)
    xscrollbar_canvas_model_answer = tkinter.Scrollbar(frame_canvas_model_answer, orient=tkinter.HORIZONTAL, command=canvas_model_answer.xview)
    yscrollbar_canvas_model_answer.pack(side="right", fill="y")
    xscrollbar_canvas_model_answer.pack(side="bottom", fill="x")
    canvas_model_answer.pack()
    canvas_model_answer.config(
      xscrollcommand=xscrollbar_canvas_model_answer.set,
      yscrollcommand=yscrollbar_canvas_model_answer.set,
      scrollregion=(0, 0, self.tk_image_model_answer.width(), self.tk_image_model_answer.height())
    )
    canvas_model_answer.coords("rectangle_new", 0, 0, 0, 0)
    canvas_model_answer.create_rectangle(0, 0, 0, 0, fill="red", tags="rectangle_new")
    canvas_model_answer.bind("<Button-1>", canvas_model_answer_draw_rectangle_click)
    canvas_model_answer.bind("<B1-Motion>", canvas_model_answer_draw_rectangle_drag)
    canvas_model_answer.bind("<ButtonRelease-1>", canvas_model_answer_draw_rectangle_release)
    
    if len(dict_answer_area["questions"]) == 0:
      self.index_selected_question = None
    else:
      self.index_selected_question = len(dict_answer_area["questions"]) - 1

    load_question()


#############################################################################################################################################

class MainFrame(tkinter.Frame):
  def __init__(self, root):
    super().__init__(root, width=800, height=500, borderwidth=2, relief="groove")
    self.root = root
    self.sub_window = SubWindow(self.root)
    self.pack()
    self.index_selected_exam = tkinter.IntVar(root)
    self.pack_propagate(0)
    self.create_listbox()
    self.btn_left()
    self.load_listbox_projects()

  # 試験一覧
  def create_listbox(self):
    frame_listbox = tkinter.Label(self)
    frame_listbox.grid(column=0, row=0, padx=5, pady=5)

    # 上部: ラベル
    label_listbox_header = tkinter.Label(frame_listbox, text="試験一覧", anchor="w")
    label_listbox_header.grid(column=0, row=0)

    # 中部: リストボックス
    self.listbox_projects = tkinter.Listbox(frame_listbox, width=60, height=20)
    self.listbox_projects.grid(column=0, row=1)
    self.listbox_projects.configure(
      activestyle=tkinter.DOTBOX,
      selectmode=tkinter.SINGLE,
      selectbackground="grey"
    )
    self.listbox_projects.bind("<<ListboxSelect>>", self.selected_element_in_listbox)

    # 下部: ボタン
    frame_listbox_footer = tkinter.Frame(frame_listbox)
    frame_listbox_footer.grid(column=0, row=2)
    tkinter.Button(frame_listbox_footer, text="追加", width=10, height=1, command=self.sub_window.add_project).grid(column=0, row=0)
    tkinter.Button(frame_listbox_footer, text="編集", width=10, height=1, command=self.sub_window.edit_project).grid(column=1, row=0)
    tkinter.Button(frame_listbox_footer, text="削除", width=10, height=1, command=self.del_project).grid(column=2, row=0)
    tkinter.Button(frame_listbox_footer, text="上へ", width=10, height=1, command=self.up_project).grid(column=3, row=0)
    tkinter.Button(frame_listbox_footer, text="下へ", width=10, height=1, command=self.down_project).grid(column=4, row=0)

  def write_index_to_config(self, index_projects_in_listbox):
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    dict_config["index_projects_in_listbox"] = index_projects_in_listbox
    with open("config.json", "w", encoding="utf-8") as f:
      json.dump(dict_config, f, indent=2)

  def selected_element_in_listbox(self, event):
    if self.listbox_projects.curselection() != ():
      index_projects_in_listbox = self.listbox_projects.curselection()[0]
      self.write_index_to_config(index_projects_in_listbox)

  def load_listbox_projects(self, *, parent=None):
    if parent is not None:
      self = parent
    self.listbox_projects.delete(0, tkinter.END)
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    if len(dict_config["projects"]) == 0:
      self.listbox_projects.insert(0, "［追加］をクリックして新しく試験を追加して下さい")
      self.listbox_projects.configure(state="disable")
      self.write_index_to_config(None)
    else:
      for project in dict_config["projects"]:
        self.listbox_projects.insert(tkinter.END, project["name"])
      index_projects_in_listbox = dict_config["index_projects_in_listbox"]
      self.listbox_projects.select_set(index_projects_in_listbox)
  
  def del_project(self):
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    index_projects_in_listbox = dict_config["index_projects_in_listbox"]
    if index_projects_in_listbox is None:
      tkinter.messagebox.showinfo(
        "試験が選択されていません. ", 
        "「試験一覧」より削除したい試験を選択して下さい. "
      )
    else:
      bool_del_project = tkinter.messagebox.askyesno(
        "試験を削除しますか？",
        f"この操作で, 答案スキャンデータ / 採点データが失われることはありませんが, 試験一覧からは表示されなくなり, 本アプリ上からはアクセスできなくなります. \n"
        + f"［追加］より同じフォルダ / ファイルを指定することで, 採点データ等を再び利用できます. \n"
        + f"採点データ等を完全に削除したい場合は, 本アプリ終了後, 答案スキャンデータが保存されているフォルダの隠しフォルダ「.temp_saiten」を手動で削除して下さい. \n\n"
        + f"試験名: {dict_config['projects'][index_projects_in_listbox]['name']}\n\n"
        + f"本当に試験を削除しますか？"
      )
      if bool_del_project:
        with open("config.json", "r", encoding="utf-8") as f:
          dict_config = json.load(f)
        dict_config["projects"].pop(index_projects_in_listbox)
        if len(dict_config["projects"]) == 0:
          dict_config["index_projects_in_listbox"] = None
        else:
          dict_config["index_projects_in_listbox"] = 0
        with open("config.json", "w", encoding="utf-8") as f:
          json.dump(dict_config, f, indent=2)
        self.load_listbox_projects()
  
  def up_project(self):
    nothing_to_do()
    self.load_listbox_projects()

  def down_project(self):
    nothing_to_do()
    self.load_listbox_projects()




  # btn_left: 操作ボタン
  def btn_left(self):
    frame_operate = tkinter.Frame(self)
    frame_operate.grid(column=1, row=0, padx=10, pady=10)
    tkinter.Frame(frame_operate, width=20, height=25).pack(expand=True)
    tkinter.Button(frame_operate, text="解答欄の位置を指定", command=self.sub_window.select_area, width=20, height=2).pack(expand=True)
    tkinter.Button(frame_operate, text="一括採点する", command=nothing_to_do, width=20, height=2).pack(expand=True)
    tkinter.Frame(frame_operate, width=20, height=25).pack(expand=True)
    tkinter.Button(frame_operate, text="書き出す", command=nothing_to_do, width=20, height=2).pack(expand=True)
    tkinter.Button(frame_operate, text="終了", command=self.root.destroy, width=20, height=2).pack(expand=True)  

def menu(root):
  menu_root = tkinter.Menu(root)

  menu_file = tkinter.Menu(menu_root, tearoff=0)
  menu_file.add_command(label="新しく試験を追加")
  menu_file.add_command(label="選択中の試験を編集")
  menu_file.add_separator()
  menu_file.add_command(label="構成設定をリセット")
  menu_file.add_separator()
  menu_file.add_command(label="終了")

  menu_edit = tkinter.Menu(menu_root, tearoff=0)
  menu_edit.add_command(label="選択中の試験の解答欄の位置を指定")
  menu_edit.add_command(label="選択中の試験を一括採点する")
  menu_edit.add_separator()
  menu_edit.add_command(label="終了")

  menu_help = tkinter.Menu(menu_root, tearoff=0)
  menu_help.add_command(label="ヘルプ")
  
  menu_root.add_cascade(label="ファイル", menu=menu_file)
  menu_root.add_cascade(label="編集", menu=menu_edit)
  menu_root.add_cascade(label="ヘルプ", menu=menu_help)
  root.config(menu=menu_root)

def make_config():
  dict_config = {
    "index_projects_in_listbox": None,
    "projects": []
  }
  with open("config.json", "w", encoding="utf-8") as f:
    json.dump(dict_config, f, indent=2)

def check_on_run():
  try:
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
  except FileNotFoundError:
    make_config()

def main():
  root = tkinter.Tk()
  root.title("一括採点")
  root.geometry("800x500")
  menu(root)
  MainFrame(root=root)
  root.mainloop()

if __name__ == "__main__":
  check_on_run()
  main()

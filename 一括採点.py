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

import openpyxl
import openpyxl.styles
import openpyxl.worksheet.datavalidation

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
      with open("config.json", "r", encoding="utf-8") as f:
        dict_config = json.load(f)
      if dict_config["index_projects_in_listbox"] is not None:
        dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
        path_dir    : str = dict_project["path_dir"]
        if os.path.exists(path_dir + "/.temp_saiten/配点.xlsx"):
          bool_del_xlsx = tkinter.messagebox.askokcancel(
            "配点ファイルが存在しています", 
            "配点ファイルに入力した情報を保存するには, ［配点を読み込む］をクリックする必要があります. \n"
            + "既に Excel で配点を入力されている場合で［配点を読み込む］をクリックしていない場合は, 入力した情報が破棄されます. \n\n"
            + "入力した配点を保存した上で操作を続行したい場合は, ［キャンセル］をクリックした後, ［配点を読み込む］をクリックして配点を読み込んでから, もう一度実行して下さい. \n\n"
            + "配点ファイルを削除してもよろしいですか？"
          )
          if not bool_del_xlsx:
            return None
          else:
            try:
              os.remove(path_dir + "/.temp_saiten/配点.xlsx")
            except PermissionError:
              tkinter.messagebox.showerror(
                "ファイルを削除できません",
                "ファイルを削除できませんでした. \n"
                + "ファイルを開いていませんか？\n"
                + "Excel を終了して, もう一度お試し下さい. "
              )
              return None
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
      dict_answer_area = {"questions": []}
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
    path_dir = dict_project["path_dir"]
    path_json_answer_area = dict_project["path_dir"] + "/.temp_saiten/answer_area.json"
    path_file_model_answer = dict_project["path_dir"] + "/.temp_saiten/model_answer/model_answer.png"
    path_dir_of_answers = dict_project["path_dir"] + "/.temp_saiten/answer"
    with open(path_json_answer_area, "r", encoding="utf-8") as f:
      dict_answer_area = json.load(f)

    def del_question():
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        dict_answer_area["questions"].pop(self.index_selected_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        reload_listbox_question()

    def up_question():
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        pop_question = dict_answer_area["questions"].pop(self.index_selected_question)
        self.index_selected_question = max(self.index_selected_question - 1, 0)
        dict_answer_area["questions"].insert(self.index_selected_question, pop_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        reload_listbox_question()

    def down_question():
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        pop_question = dict_answer_area["questions"].pop(self.index_selected_question)
        self.index_selected_question = min(self.index_selected_question + 1, len(dict_answer_area["questions"]) - 1)
        dict_answer_area["questions"].insert(self.index_selected_question, pop_question)
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        reload_listbox_question()

    def set_type(str_type):
      if self.index_selected_question is not None:
        with open(path_json_answer_area, "r", encoding="utf-8") as f:
          dict_answer_area = json.load(f)
        dict_answer_area["questions"][self.index_selected_question]["type"] = str_type
        with open(path_json_answer_area, "w", encoding="utf-8") as f:
          json.dump(dict_answer_area, f, indent=2)
        reload_listbox_question()
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

    def canvas_draw_rectangle_click(event):
      self.canvas_draw_rectangle[0] = event.x
      self.canvas_draw_rectangle[1] = event.y
      self.canvas_draw_rectangle[2] = min(event.x + 1, canvas.winfo_width())
      self.canvas_draw_rectangle[3] = min(event.y + 1, canvas.winfo_height())
      canvas.coords("rectangle_new",
        self.canvas_draw_rectangle[0],
        self.canvas_draw_rectangle[1],
        self.canvas_draw_rectangle[2],
        self.canvas_draw_rectangle[3], 
      )
    def canvas_draw_rectangle_drag(event):
      self.canvas_draw_rectangle[2] = min(max(event.x, 0), canvas.winfo_width())
      self.canvas_draw_rectangle[3] = min(max(event.y, 0), canvas.winfo_height())
      canvas.coords("rectangle_new",
        self.canvas_draw_rectangle[0],
        self.canvas_draw_rectangle[1],
        self.canvas_draw_rectangle[2],
        self.canvas_draw_rectangle[3], 
      )
    def canvas_draw_rectangle_release(event):
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
            min(self.canvas_draw_rectangle[0], self.canvas_draw_rectangle[2]),
            min(self.canvas_draw_rectangle[1], self.canvas_draw_rectangle[3]),
            max(self.canvas_draw_rectangle[0], self.canvas_draw_rectangle[2]),
            max(self.canvas_draw_rectangle[1], self.canvas_draw_rectangle[3])
          ]
        }
      )
      with open(path_json_answer_area, "w", encoding="utf-8") as f:
        json.dump(dict_answer_area, f, indent=2)
      self.index_selected_question = len(dict_answer_area["questions"]) - 1
      reload_listbox_question()
      canvas.coords("rectangle_new", 0, 0, 0, 0)

    def selected_listbox_question(*args, **kwargs):
      with open(path_json_answer_area, "r", encoding="utf-8") as f:
        dict_answer_area = json.load(f)      
      for index_question, question in enumerate(dict_answer_area["questions"]):
        if question["type"] == "設問":
          color_reactangle = "green"
        elif question["type"] == "氏名":
          color_reactangle = "blue"
        elif question["type"] == "生徒番号":
          color_reactangle = "cyan"
        elif question["type"] == "小計点":
          color_reactangle = "magenta"
        elif question["type"] == "合計点":
          color_reactangle = "orange"
        elif question["type"] == "採点者印":
          color_reactangle = "yellow"
        self.index_selected_question = listbox_question.curselection()[0]
        if index_question == listbox_question.curselection()[0]:
          color_reactangle = "red"
        canvas.create_rectangle(
          question["area"][0], 
          question["area"][1], 
          question["area"][2], 
          question["area"][3], 
          outline=color_reactangle,
          width=2,
          fill=color_reactangle,
          stipple="gray12",
          tags="field"
        )
        canvas.create_text(
          question["area"][0] - 10, 
          (question["area"][1] + question["area"][3]) // 2, 
          text=str(index_question),
          fill="green",
          tags="number"
        )

    def reload_listbox_question():
      listbox_question.configure(state=tkinter.NORMAL)
      listbox_question.delete(0, tkinter.END)
      with open(path_json_answer_area, "r", encoding="utf-8") as f:
        dict_answer_area = json.load(f)
      canvas.delete("field")
      canvas.delete("number")
      if len(dict_answer_area["questions"]) == 0:
        self.index_selected_question = None
        listbox_question.insert(tkinter.END, "模範解答の画像の上で")
        listbox_question.insert(tkinter.END, "ドラッグして")
        listbox_question.insert(tkinter.END, "解答欄を指定して下さい")
        listbox_question.configure(state=tkinter.DISABLED)
      else:
        self.index_selected_question = min(self.index_selected_question, len(dict_answer_area["questions"]) - 1)
        for index_question, question in enumerate(dict_answer_area["questions"]):
          listbox_question.insert(tkinter.END, f"枠{index_question} - {question['type']}")
        listbox_question.select_set(self.index_selected_question)
        selected_listbox_question()

    self.window.title("解答欄を指定")
    self.canvas_draw_rectangle = [0, 0, 0, 0]

    frame_main = tkinter.Frame(self.window)
    frame_main.grid(column=0, row=0)

    frame_question = tkinter.Frame(frame_main)
    frame_question.grid(column=0, row=0)
    frame_picture = tkinter.Frame(frame_main)
    frame_picture.grid(column=1, row=0)

    frame_listbox_question = tkinter.Frame(frame_question)
    frame_listbox_question.grid(column=0, row=0)
    frame_btn_list_question = tkinter.Frame(frame_question)
    frame_btn_list_question.grid(column=0, row=1)

    listbox_question = tkinter.Listbox(frame_listbox_question, width=20, height=30)
    listbox_question.pack(side="left")
    listbox_question.configure(
      activestyle=tkinter.DOTBOX,
      selectmode=tkinter.SINGLE,
      selectbackground="grey"
    )
    for index_question in range(len(dict_answer_area["questions"])):
      listbox_question.insert(tkinter.END, f"設問{index_question}")   
    listbox_question.bind("<MouseWheel>", lambda eve:listbox_question.yview_scroll(int(-eve.delta/120), 'units'))
    yscrollbar_table_question = tkinter.Scrollbar(frame_listbox_question, orient=tkinter.VERTICAL, command=listbox_question.yview)
    yscrollbar_table_question.pack(side="right", fill="y")
    listbox_question.config(
      yscrollcommand=yscrollbar_table_question.set
    )
    
    btn_list_question_del = tkinter.Button(frame_btn_list_question, width=6, text="削除", command=del_question)
    btn_list_question_del.grid(column=0, row=0)
    btn_list_question_up = tkinter.Button(frame_btn_list_question, width=6, text="上へ", command=up_question)
    btn_list_question_up.grid(column=1, row=0)
    btn_list_question_down = tkinter.Button(frame_btn_list_question, width=6, text="下へ", command=down_question)
    btn_list_question_down.grid(column=2, row=0)
    btn_list_question_que = tkinter.Button(frame_btn_list_question, width=6, text="設問", command=set_question)
    btn_list_question_que.grid(column=0, row=1)
    btn_list_question_name = tkinter.Button(frame_btn_list_question, width=6, text="氏名", command=set_name)
    btn_list_question_name.grid(column=1, row=1)
    btn_list_question_id = tkinter.Button(frame_btn_list_question, width=6, text="生徒番号", command=set_id)
    btn_list_question_id.grid(column=2, row=1)
    btn_list_question_id = tkinter.Button(frame_btn_list_question, width=6, text="採点者印", command=set_stamp)
    btn_list_question_id.grid(column=0, row=2)
    btn_list_question_subtotal = tkinter.Button(frame_btn_list_question, width=6, text="小計点", command=set_subtotal)
    btn_list_question_subtotal.grid(column=1, row=2)
    btn_list_question_total = tkinter.Button(frame_btn_list_question, width=6, text="合計点", command=set_total)
    btn_list_question_total.grid(column=2, row=2)
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

    frame_canvas = tkinter.Frame(frame_picture)
    frame_canvas.pack()

    canvas = tkinter.Canvas(frame_canvas, bg="black", width=567, height=800)
    canvas.bind("<Control-MouseWheel>", lambda eve:canvas.xview_scroll(int(-eve.delta/120), 'units'))
    canvas.bind("<MouseWheel>", lambda eve:canvas.yview_scroll(int(-eve.delta/120), 'units'))
    self.tk_image_model_answer = PIL.ImageTk.PhotoImage(file=path_file_model_answer)
    canvas.create_image(0, 0, image=self.tk_image_model_answer, anchor="nw")
    yscrollbar_canvas = tkinter.Scrollbar(frame_canvas, orient=tkinter.VERTICAL, command=canvas.yview)
    xscrollbar_canvas = tkinter.Scrollbar(frame_canvas, orient=tkinter.HORIZONTAL, command=canvas.xview)
    yscrollbar_canvas.pack(side="right", fill="y")
    xscrollbar_canvas.pack(side="bottom", fill="x")
    canvas.pack()
    canvas.config(
      xscrollcommand=xscrollbar_canvas.set,
      yscrollcommand=yscrollbar_canvas.set,
      scrollregion=(0, 0, self.tk_image_model_answer.width(), self.tk_image_model_answer.height())
    )
    
    listbox_question.bind("<<ListboxSelect>>", selected_listbox_question)
    canvas.coords("rectangle_new", 0, 0, 0, 0)
    canvas.create_rectangle(0, 0, 0, 0, fill="red", tags="rectangle_new")
    canvas.bind("<Button-1>", canvas_draw_rectangle_click)
    canvas.bind("<B1-Motion>", canvas_draw_rectangle_drag)
    canvas.bind("<ButtonRelease-1>", canvas_draw_rectangle_release)
    
    if len(dict_answer_area["questions"]) == 0:
      self.index_selected_question = None
    else:
      self.index_selected_question = len(dict_answer_area["questions"]) - 1

    reload_listbox_question()

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

  def make_xlsx(self):
    tkinter.messagebox.showinfo(
      "配点を入力します",
      "配点の入力は, 本ソフトウェア上ではなく Excel 等の表計算ソフトウェアを使用して行います. \n\n"
      + "配点を登録するために 配点.xlsx ファイルを作成して開きます. "
    )
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
    path_dir = dict_project["path_dir"]
    with open(path_dir + "/.temp_saiten/answer_area.json") as f:
      dict_answer_area = json.load(f)
      if os.path.exists(path_dir + "/.temp_saiten/配点.xlsx"):
        bool_del_xlsx = tkinter.messagebox.askokcancel(
          "配点ファイルが存在しています", 
          "配点ファイルに入力した情報を保存するには, ［配点を読み込む］をクリックする必要があります. \n"
          + "既に Excel で配点を入力されている場合で［配点を読み込む］をクリックしていない場合は, 入力した情報が破棄されます. \n\n"
          + "入力した配点を保存した上で操作を続行したい場合は, ［キャンセル］をクリックした後, ［配点を読み込む］をクリックして配点を読み込んでから, もう一度実行して下さい. \n\n"
          + "配点ファイルを削除してもよろしいですか？"
        )
        if not bool_del_xlsx:
          return None

    if len(dict_answer_area["questions"]) == 0:
      tkinter.messagebox.showwarning(
        "解答欄の位置が指定されていません",
        "解答欄の位置が指定されていません. \n"
        + "［解答欄の位置を指定］をクリックして解答欄の位置を指定してから, もう一度お試し下さい. "
      )
      return      
    workbook_haiten = openpyxl.Workbook()
    workbook_haiten.remove(workbook_haiten["Sheet"])
    workbook_haiten.create_sheet(title="配点登録")
    sheet_haiten = workbook_haiten["配点登録"]
    sheet_haiten.cell(1, 2).value = "種類"
    sheet_haiten.cell(1, 3).value = "大問"
    sheet_haiten.cell(1, 4).value = "小問"
    sheet_haiten.cell(1, 5).value = "枝問"
    sheet_haiten.cell(1, 6).value = "配点"
    side = openpyxl.styles.Side(style="thin", color="000000")
    border_up_down = openpyxl.styles.Border(top=side, bottom=side)
    datavalidation_whole = openpyxl.worksheet.datavalidation.DataValidation(type="whole")
    datavalidation_textlength10 = openpyxl.worksheet.datavalidation.DataValidation(type="textLength", operator="lessThanOrEqual", formula1=10)
    for index_question, question in enumerate(dict_answer_area["questions"]):
      sheet_haiten.cell(index_question + 2, 1).value = index_question
      sheet_haiten.cell(index_question + 2, 1).border = border_up_down
      sheet_haiten.cell(index_question + 2, 2).value = question["type"]
      sheet_haiten.cell(index_question + 2, 2).border = border_up_down
      sheet_haiten.cell(index_question + 2, 3).value = question["daimon"]
      sheet_haiten.cell(index_question + 2, 3).border = border_up_down
      if question["type"] in ["設問", "小計点"]:
        sheet_haiten.cell(index_question + 2, 3).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="bfffff")
        sheet_haiten.cell(index_question + 2, 3).protection = openpyxl.styles.Protection(locked=False)
        datavalidation_textlength10.add(sheet_haiten.cell(index_question + 2, 3))
      else:
        sheet_haiten.cell(index_question + 2, 3).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
      sheet_haiten.cell(index_question + 2, 4).value = question["shomon"]
      sheet_haiten.cell(index_question + 2, 4).border = border_up_down
      if question["type"] in ["設問"]:
        sheet_haiten.cell(index_question + 2, 4).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="cfefef")
        sheet_haiten.cell(index_question + 2, 4).protection = openpyxl.styles.Protection(locked=False)
        datavalidation_textlength10.add(sheet_haiten.cell(index_question + 2, 4))
      else:
        sheet_haiten.cell(index_question + 2, 4).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
      sheet_haiten.cell(index_question + 2, 5).value = question["shimon"]
      sheet_haiten.cell(index_question + 2, 5).border = border_up_down
      if question["type"] in ["設問"]:
        sheet_haiten.cell(index_question + 2, 5).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="bfffff")
        sheet_haiten.cell(index_question + 2, 5).protection = openpyxl.styles.Protection(locked=False)
        datavalidation_textlength10.add(sheet_haiten.cell(index_question + 2, 5))
      else:
        sheet_haiten.cell(index_question + 2, 5).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
      sheet_haiten.cell(index_question + 2, 6).value = question["haiten"]
      sheet_haiten.cell(index_question + 2, 6).border = border_up_down
      if question["type"] in ["設問"]:
        sheet_haiten.cell(index_question + 2, 6).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="ffbfbf")
        sheet_haiten.cell(index_question + 2, 6).protection = openpyxl.styles.Protection(locked=False)
        datavalidation_whole.add(sheet_haiten.cell(index_question + 2, 6))
      else:
        sheet_haiten.cell(index_question + 2, 6).fill = openpyxl.styles.PatternFill(patternType="solid", fgColor="cccccc")
    sheet_haiten.cell(index_question + 3, 5).value = "配点合計"
    sheet_haiten.cell(index_question + 3, 6).value = f"=SUMIF(B2:B{index_question + 2}, \"設問\", F2:F{index_question + 2})"
    sheet_haiten.protection.selectLockedCells   = True  # ロックされたセルの選択
    sheet_haiten.protection.selectUnlockedCells = False # ロックされていないセルの選択
    sheet_haiten.protection.formatCells         = True  # セルの書式設定
    sheet_haiten.protection.formatColumns       = True  # 列の書式設定
    sheet_haiten.protection.formatRows          = True  # 行の書式設定
    sheet_haiten.protection.insertColumns       = True  # 列の挿入
    sheet_haiten.protection.insertRows          = True  # 行の挿入
    sheet_haiten.protection.insertHyperlinks    = True  # ハイパーリンクの挿入
    sheet_haiten.protection.deleteColumns       = True  # 列の削除
    sheet_haiten.protection.deleteRows          = True  # 行の削除
    sheet_haiten.protection.sort                = True  # 並べ替え
    sheet_haiten.protection.autoFilter          = True  # フィルター
    sheet_haiten.protection.pivotTables         = True  # ピボットテーブルレポート
    sheet_haiten.protection.objects             = False # オブジェクトの編集
    sheet_haiten.protection.scenarios           = True  # シナリオの編集
    sheet_haiten.protection.enable()
    workbook_haiten.security.lockStructure = True

    try:
      workbook_haiten.save(path_dir + "/.temp_saiten/配点.xlsx")
    except PermissionError:
      tkinter.messagebox.showerror(
        "ファイルを保存できません",
        "ファイルを保存できませんでした. \n"
        + "既にファイルを開いていませんか？\n"
        + "Excel を終了して, もう一度お試し下さい. "
      )
    else:
      os.startfile(path_dir + "/.temp_saiten/配点.xlsx")

  def read_xlsx(self):
    with open("config.json", "r", encoding="utf-8") as f:
      dict_config = json.load(f)
    dict_project = dict_config["projects"][dict_config["index_projects_in_listbox"]]
    path_dir = dict_project["path_dir"]
    with open(path_dir + "/.temp_saiten/answer_area.json") as f:
      dict_answer_area = json.load(f)
    if not os.path.exists(path_dir + "/.temp_saiten/配点.xlsx"):
      tkinter.messagebox.showerror(
        "ファイルが見つかりません",
        "配点.xlsx が見つかりません. \n"
        + "［配点を入力する］をクリックして, ファイルを生成し, 配点を入力して保存して下さい. "
      )
    else:
      try:
        workbook_haiten = openpyxl.load_workbook(path_dir + "/.temp_saiten/配点.xlsx")
        os.remove(path_dir + "/.temp_saiten/配点.xlsx")
      except PermissionError:
        tkinter.messagebox.showerror(
          "ファイルを操作できません",
          "ファイルを操作できませんでした. \n"
          + "ファイルを開いていませんか？\n"
          + "Excel を終了して, もう一度お試し下さい. "
        )
        return
      sheet_haiten = workbook_haiten["配点登録"]
      for index_question in range(len(dict_answer_area["questions"])):
        dict_answer_area["questions"][index_question]["daimon"] = sheet_haiten.cell(index_question + 2, 3).value
        dict_answer_area["questions"][index_question]["shomon"] = sheet_haiten.cell(index_question + 2, 4).value
        dict_answer_area["questions"][index_question]["shimon"] = sheet_haiten.cell(index_question + 2, 5).value
        if sheet_haiten.cell(index_question + 2, 5).value == "":
          dict_answer_area["questions"][index_question]["haiten"] = None
        else:
          dict_answer_area["questions"][index_question]["haiten"] = sheet_haiten.cell(index_question + 2, 6).value
      with open(path_dir + "/.temp_saiten/answer_area.json", "w", encoding="utf-8") as f:
        json.dump(dict_answer_area, f, indent=2)
      message_sample = ""
      for i in range(1000):
        message_sample += "ああああああああああ"
      tkinter.messagebox.showinfo(
        "配点を読み込みました",
        "読み込んだ内容は保存し, 配点.xlsx は削除しました. \n"
        + "再び配点を編集するには, ［配点を入力する］をクリックして下さい. \n"
      )

  # btn_left: 操作ボタン
  def btn_left(self):
    frame_operate = tkinter.Frame(self)
    frame_operate.grid(column=1, row=0, padx=10, pady=10)
    tkinter.Button(frame_operate, text="解答欄の位置を指定", command=self.sub_window.select_area, width=20, height=2).grid(column=0, row=0, columnspan=2)
    tkinter.Button(frame_operate, text="配点を\n入力する", command=self.make_xlsx, width=9, height=2).grid(column=0, row=1, sticky="WE")
    tkinter.Button(frame_operate, text="配点を\n読み込む", command=self.read_xlsx, width=9, height=2).grid(column=1, row=1, sticky="WE")
    tkinter.Frame(frame_operate, width=20, height=25).grid(column=0, row=2, columnspan=2)
    tkinter.Button(frame_operate, text="一括採点する", command=nothing_to_do, width=20, height=2).grid(column=0, row=3, columnspan=2)
    tkinter.Frame(frame_operate, width=20, height=25).grid(column=0, row=4, columnspan=2)
    tkinter.Button(frame_operate, text="書き出す", command=nothing_to_do, width=20, height=2).grid(column=0, row=5, columnspan=2)
    tkinter.Button(frame_operate, text="終了", command=self.root.destroy, width=20, height=2).grid(column=0, row=6, columnspan=2)  

def menu(root):
  menu_root = tkinter.Menu(root)

  menu_file = tkinter.Menu(menu_root, tearoff=0)
  menu_file.add_command(label="新しく試験を追加")
  menu_file.add_command(label="選択中の試験を編集")
  menu_file.add_command(label="選択中の試験を削除")
  menu_file.add_separator()
  menu_file.add_command(label="構成設定をリセット")
  menu_file.add_separator()
  menu_file.add_command(label="終了")

  menu_edit = tkinter.Menu(menu_root, tearoff=0)
  menu_edit.add_command(label="選択中の試験の解答欄の位置を指定")
  menu_edit.add_command(label="選択中の試験の配点を入力する")
  menu_edit.add_command(label="選択中の試験の配点を読み込む")
  menu_edit.add_command(label="選択中の試験を一括採点する")

  menu_help = tkinter.Menu(menu_root, tearoff=0)
  menu_help.add_command(label="ヘルプ")
  menu_help.add_command(label="バージョン情報")
  
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
    return True
  except FileNotFoundError:
    bool_accept_terms = tkinter.messagebox.askyesno(
      "Accept the terms? - 規約に同意しますか？",
      "Copyright(c) 2022 KeppyNaushika\n\n"
      + "This software is released under the GNU Affero General Public License v3.0, see LICENSE. \n\n"
      + "このソフトウェアは, GNU Affro General Public License version3 の下, 提供されています. \n\n"
      + "ライセンスを遵守する限り, 商用利用, 変更, 頒布, 特許利用, 私的利用が認められますが, "
      + "利用にあたって開発者は責任を負いませんしいかなる保証も提供しません. \n\n"
      + "同梱されている LICENSE をお読みいただき, 同意される場合は［はい］をクリックして下さい. \n\n"
      + "尚, 本ソフトウェアにおける Microsoft 製品についての記述は, マイクロソフトの商標およびブランドガイドラインに準拠しています. \n\n"
      + "The github repository of this software:\n"
      + "https://github.com/KeppyNaushika/scoring_at_once/"
    )
    if bool_accept_terms:
      tkinter.messagebox.showinfo(
        "表計算ソフトをご用意下さい. ",
        "本ソフトウェアでは, 一部で Microsoft Excel 等の表計算ソフトウェアを利用します. \n\n"
        + "あらかじめインストールの上, ご利用下さい. "
      )
      tkinter.messagebox.showinfo(
        "ごめんなさい",
        "本ソフトウェアは, alpha版です. \n\n"
        + "一部動作しない機能がございます. "
      )
      make_config()
      return True
    else:
      return False

def main():
  root = tkinter.Tk()
  root.title("一括採点 - alpha版")
  root.geometry("800x500")
  menu(root)
  MainFrame(root=root)
  root.mainloop()

if __name__ == "__main__":
  if check_on_run():
    main()
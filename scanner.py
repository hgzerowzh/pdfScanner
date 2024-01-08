import fitz  # PyMuPDF
import os, sys
import time
import random
import glob
import hashlib
import requests
import openpyxl
import pandas as pd
import threading
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, wait, ALL_COMPLETED
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import xml.etree.ElementTree as ET


class Scanner:
    def __init__(self) -> None:
        self.pdf_path = None
        self.writer = None
        self.highlights_by_color = {}
        self.pages_by_color = {}

    def rgb2color(self, rgb):
        """translate RGB to color name

        Args:
            rgb (str): rgb
        Returns:
            str: color name
        """
        color_names = {
            (0.9804, 0.9804, 0.0): "Yellow",
            (0.5647, 1.0, 0.5647): "Green",
            (0.0, 0.502, 1.0): "Blue",
            (1.0, 0.5647, 1.0): "Pink",
            (0.7294, 0.3333, 1.0): "Purple",
            (1.0, 0.5647, 0.5647): "Red",
            (1.0, 0.8471, 0.5647): "Orange",
            (0.8471, 0.8471, 0.8627): "Gray",
            (0.5647, 0.8471, 1.0): "Light Blue",
            (0.8471, 1.0, 0.8471): "Light Green",
            (0.8471, 0.5647, 0.0): "Brown"
        }
        closest_color = min(color_names.keys(), key=lambda c: sum((a-b)**2 for a, b in zip(c, rgb)))
        return color_names[closest_color]


    def scan_pdf(self, pdf_path, writer):
        """scan highlight in pdf file

        Args:
            pdf_path (_type_): _description_
            writer (_type_): _description_
        """
        self.pdf_path = pdf_path
        self.writer = writer
        self.highlights_by_color.clear()
        self.pages_by_color.clear()

        doc = fitz.open(self.pdf_path)

        try:
            for page_num, page in enumerate(doc):
                annotations = page.annots()
                if annotations:
                    for annot in annotations:
                        if annot.type[0] == 8:  
                            color = annot.colors['stroke']  
                            highlight = page.get_text("text", clip=annot.rect).strip()
                            if color not in self.highlights_by_color:
                                self.highlights_by_color[color] = []
                                self.pages_by_color[color] = []
                            if highlight in self.highlights_by_color[color]: # 去重
                                continue
                            self.highlights_by_color[color].append(highlight)
                            self.pages_by_color[color].append(page_num + 1)
        except Exception as e:
            print(e)
            print("some err occured.")
        

class Translator:
    """Translator Class
    """
    def __init__(self, view_workers, row_workers, trans_id_pool, base_url,words_book_url, sleep_time, books, tag_name, cookies) -> None:
        self.view_workers = view_workers
        self.row_workers = row_workers
        self.btn_states = "normal"
        self.trans_id_pool = trans_id_pool
        self.BASE_URL = base_url
        self.words_book_url = words_book_url
        self.sleep_time = sleep_time
        self.books = books
        self.dictionary = {}    # 加载的词典
        self.youdao_book = ET.Element('youdao_wordbook.xml')
        self.tag_name = tag_name
        self.cookies = cookies
        self.add_youdao_workbook_headers = {
            'Cookie': self.cookies,
            'Host': 'dict.youdao.com',
            'Upgrade-Insecure-Requests': '1',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Accept': 'application/json, text/plain, */*',
            'Referer': 'https://dict.youdao.com',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 11_1_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36'
        }
        
      
    def load_translate_books(self):
        """load local dictionary

        Returns:
            None
        """
        n = 0
        for book_item in self.books: 
            my_dictionary = pd.read_csv(book_item, sep='⬄', header=0, names=['word', 'interpretation'])
            def check(series):
                return series['word'].strip()
            my_dictionary['word'] = my_dictionary.apply(check, axis=1)
            my_dictionary.set_index('word', inplace=True)
            self.dictionary[n] = my_dictionary
            n += 1


    def start(self):
        """start translater
        """
        # output_display.delete(1.0, tk.END)  # clear output
        trans_label_var.set("加载词典中......")
        self.load_translate_books()
        threading.Thread(target=self.deal_excel).start()


    def get_trans_id(self):
        key = random.choice(list(self.trans_id_pool))
        value = self.trans_id_pool[key]
        return (str(key), value)


    def btn_change(self, state):
        """translate button click change

        Args:
            state (str): disable or norma
        """
        if state == "disable":
            btn_start_translate.config(state=tk.DISABLED)
            btn_open_file.config(state=tk.DISABLED) 
        elif state == "normal":
            btn_start_translate.config(state=tk.NORMAL)
            btn_open_file.config(state=tk.NORMAL) 


    def check_path(self, path):
        """check selectd path

        Args:
            path (str): selected path

        Returns:
            bool: True or False
        """
        if not os.path.exists(path):
            messagebox.showwarning("警告", "路径不存在!")
            file_path = filedialog.askopenfile()
            if file_path:
                output_file_path_var.set(file_path.name)
                self.btn_change("normal")
                return True
            else:
                return False
        if not path.endswith(('.xls', '.xlsx')):
            messagebox.showwarning("警告", "请选择一个有效的Excel文件!")
            self.btn_change("normal")
            return False
        return True
        

    def deal_excel(self):
        """main logic to deal excel
        """
        path = output_file_path_var.get().strip()
        if not self.check_path(path): 
            return

        self.btn_change("disable")
        trans_label_var.set("正在拼命翻译中, 请耐心等待......")
        output_display.insert(tk.END, f"[Info] 翻译开始, 请等待一段时间...\n")

        wb = openpyxl.load_workbook(path)

        # 构造线程池对所有工作表进行操作
        t_poll = ThreadPoolExecutor(max_workers=self.view_workers)
        thread_list = []
        for worksheet_name in wb.sheetnames:
            output_display.insert(tk.END, f"[Info] 开始翻译: {worksheet_name}\n")
            ws = wb[worksheet_name]
            # 将未翻译的sheet翻译
            if ws.max_column % 3 != 0:
                f = t_poll.submit(self.insert_translation_columns, ws)
                thread_list.append(f)   
        wait(thread_list, return_when=ALL_COMPLETED)
            
        # 保存为Excel文件
        wb.save(path)
        output_display.insert(tk.END, f"[Success] 翻译完成!\n")
        current_xview = output_display.xview()
        output_display.see(tk.END)
        output_display.xview_moveto(current_xview[0])
        
        # 生成有道词典单词本xml文件
        xml_str = ET.tostring(self.youdao_book, encoding='unicode', method='xml')
        xml_str = xml_str.replace('&lt;', '<').replace('&gt;', '>').replace("youdao_wordbook.xml", "wordbook")
        bookfile = "%s/youdao_wordbook.xml" % os.path.dirname(path)
        with open(bookfile, 'wb') as f:
            f.write(xml_str.encode('utf-8'))
        output_display.insert(tk.END, f"[Success] 单词本: {bookfile}\n")
        trans_label_var.set("翻译完成!")
        self.btn_change("normal")


    def insert_translation_columns(self, ws):
        """insert translated info to specific cell in column

        Args:
            ws (str): worksheet name
        """
        max_col = ws.max_column
        for col in range(1, 2*max_col+1, 3):
            ws.insert_cols(col+1)
            # 开启线程池进行多线程翻译
            executor = ThreadPoolExecutor(max_workers=self.row_workers)
            t_list = []
            # 从第二行开始，翻译并填充翻译列的内容
            for row in range(2, ws.max_row + 1):
                text = ws.cell(row=row, column=col).value
                if text:
                    f = executor.submit(self.trans_row, text, ws, row, col)
                    t_list.append(f)
            wait(t_list, return_when=ALL_COMPLETED)

    def trans_row(self, text, ws, row, col):
        """translate one cell in Excel

        Args:
            text (str): content in excel cell
            ws (w): worksheet
            row (int): row
            col (int): column
        """
        # 进行翻译
        ret = self.translate_local(str(text))
        if ret:
            translation = str(ret)
        else:
            translation = self.translate_baidu_api(str(text))
            
        # 添加到单词本
        if youdao_wordbook_check_var.get():
            self.add_word_youdao(str(text).strip())
            
        # 处理翻译内容
        if len(translation) > 150:
            translation = translation[:150]
            
        # if "sup>" in translation:
        #     translation = translation.split("sup>")[2]       
        # if " ➞ " in translation:
        #     translation = translation.split("➞")[0]
        # # if "：" in translation:
        # #     translation = translation.split("：")[0]
        # # if ":" in translation:
        # #     translation = translation.split(":")[0]
        # if "/" in translation:
        #     translation = translation.split("/")[0]
            
        # 创建有道词典单词本xml文件
        self.create_words_book(str(text), translation)
            
        current_xview = output_display.xview()
        output_display.see(tk.END)
        output_display.xview_moveto(current_xview[0])
        # print(translation)
        if translation:
            ws.cell(row=row, column=col+1).value = translation
        time.sleep(float(self.sleep_time))

    def create_words_book(self, word, trans):
        """generate youdao wordbook xml object

        Args:
            word (str): word
            trans (str): translate result
        """
        item = ET.SubElement(self.youdao_book, 'item')
        word_elem = ET.SubElement(item, 'word')
        word_elem.text = f'<![CDATA[{word}]]>'
        trans_elem = ET.SubElement(item, 'trans')
        trans_elem.text = f'<![CDATA[{trans}]]>'
        phonetic_elem = ET.SubElement(item, 'phonetic')
        phonetic_elem.text = '<![CDATA[]]>'
        tags_elem = ET.SubElement(item, 'tags')
        tags_elem.text = '<![CDATA[%s]]>' % self.tag_name
        progress_elem = ET.SubElement(item, 'progress')
        progress_elem.text = '0'

    def translate_baidu_api(self, text):
        """use baidu translate api 

        Args:
            text (str): word
        Returns:
            str: translate result
        """
        APP_ID, SECRET_KEY = self.get_trans_id()
        salt = 'random_salt'
        sign = APP_ID + text + salt + SECRET_KEY
        m = hashlib.md5()
        m.update(sign.encode('utf-8'))
        sign = m.hexdigest()
        params = {
            'q': text,
            'from': 'en',
            'to': 'zh',
            'appid': APP_ID,
            'salt': salt,
            'sign': sign
        }
        response = requests.get(self.BASE_URL, params=params)
        result = response.json()
        try:
            dst = str(result['trans_result'][0]['dst'])
            output_display.insert(tk.END, f"[Info] {text}: {dst}\n")
            return dst
        except KeyError:
            print('翻译失败:', result)
            return None

    def translate_local(self, word):
        """use local dictionary to translate

        Args:
            word (str): word
        Returns:
            str: translated word 
        """
        for i in self.dictionary.keys():
            if word in self.dictionary[i].index.values:
                result = self.dictionary[i].loc[word].values.tolist()[0]
                output_display.insert(tk.END, f"[Info] {word}: {result}\n")
                return result
            else:
                continue
        return None
    
    def add_word_youdao(self, word):
        """add word to youdao wordbook

        Args:
            word (str): wrod
        Returns:
            None
        """
        url = f"{self.words_book_url}{word}"
        try:
            response = requests.get(url, headers=self.add_youdao_workbook_headers)
            data = response.json()
            if data['code'] == 0:
                pass
            else:
                print("添加失败: %s" % word)
        except Exception as e:
            print("异常: %s" % word)
            return {'error': f'发生异常: {str(e)}'}


class ScannerGui(Scanner):
    """PDF Scanner Gui Class

    Args:
        Scanner (Scanner): Scanner class
    """
    def __init__(self, name, size, output_file, translator) -> None:
        super().__init__()
        self.window = tk.Tk()
        self.name = name
        self.size = size
        self.output_file = output_file
        self.translator = translator
        self.dir_path = None
        self.output_path = None

    def run(self):
        self.window.title(self.name)
        self.window.geometry(self.size)

        global selected_directory, btn_start_scan, output_display, output_file_path_var, btn_start_translate, btn_open_file, trans_label, trans_label_var, youdao_wordbook_check_var

        # 目录选择和显示
        selected_directory = tk.StringVar(self.window)
        tk.Entry(self.window, textvariable=selected_directory, width=25).grid(row=0, column=0, padx=(10,2), pady=5, sticky='ew')
        tk.Button(self.window, text="选择文件", command=self.select_file).grid(row=0, column=1, padx=(0, 0), pady=5, sticky='ew')
        tk.Button(self.window, text="选择目录", command=self.select_directory).grid(row=0, column=2, padx=(0, 0), pady=5, sticky='ew')

        # 开始扫描按钮
        btn_start_scan = tk.Button(self.window, text="开始扫描", command=self.start_scan, state=tk.DISABLED)
        btn_start_scan.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='ew')

        # 带有垂直和水平滚动条的输出显示区域
        frame_output = tk.Frame(self.window)
        frame_output.grid(row=2, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')

        v_scrollbar = tk.Scrollbar(frame_output, orient=tk.VERTICAL)
        h_scrollbar = tk.Scrollbar(frame_output, orient=tk.HORIZONTAL)
        
        output_display = tk.Text(frame_output, wrap=tk.NONE, yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set, height=14)
        v_scrollbar.config(command=output_display.yview)
        h_scrollbar.config(command=output_display.xview)
        
        output_display.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        frame_output.grid_columnconfigure(0, weight=1)
        frame_output.grid_rowconfigure(0, weight=1)

        # 输出文件路径和打开按钮
        output_file_path_var = tk.StringVar(self.window)
        tk.Entry(self.window, textvariable=output_file_path_var, width=25).grid(row=3, column=0, padx=(10,2), pady=2, sticky='ew')
        btn_open_file = tk.Button(self.window, text="打开", command=self.open_output_file)
        btn_open_file.grid(row=3, column=1, padx=(0,0), pady=5, sticky='ew')

        # 翻译
        trans_label_var = tk.StringVar(self.window)
        trans_label = tk.Label(self.window, textvariable=trans_label_var, width=25, anchor=tk.W)
        trans_label.grid(row=4, column=0, padx=(10,0), pady=5, sticky='ew')
        btn_start_translate = tk.Button(self.window, text="翻译", command=self.translator.start, state=tk.NORMAL)
        btn_start_translate.grid(row=3, column=2, padx=(0,0), pady=5, sticky='ew')
        
        # 有道单词本
        youdao_wordbook_check_var = tk.BooleanVar(self.window)
        tk.Checkbutton(self.window, text="单词本", variable=youdao_wordbook_check_var).grid(row=4, column=2, columnspan=1, padx=1, pady=1, sticky='ew')

        # 配置网格以适当地展开
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_columnconfigure(1, weight=1)
        self.window.grid_columnconfigure(2, weight=1)
        self.window.mainloop()

    def select_file(self):
        """select file button
        """
        file_name = filedialog.askopenfile()
        if file_name:
            selected_directory.set(file_name.name)
            btn_start_scan.config(state=tk.NORMAL)

    def select_directory(self):
        """select directory button
        """
        dir_path = filedialog.askdirectory()
        if dir_path:
            selected_directory.set(dir_path)
            btn_start_scan.config(state=tk.NORMAL)

    def start_scan(self):
        """start scan button
        """
        if not selected_directory.get():
            messagebox.showerror("Error", "请先选择对应的文件或目录")
            return
        btn_start_scan.config(state=tk.DISABLED)
        output_display.delete(1.0, tk.END)  # clear output
        if os.path.isfile(selected_directory.get()):
            target = self.scan_file
        elif os.path.isdir(selected_directory.get()):
            target = self.scan_directory
        threading.Thread(target=target).start()

    def scan_file(self):
        """scan pdf file hightlight info and extract it into excel
        """
        file_path = selected_directory.get()
        self.dir_path = os.path.dirname(file_path)
        self.output_path= self.dir_path + "/" + self.output_file
        if not file_path.endswith('.pdf'):
            output_display.insert(tk.END, "[Error] 请选择PDF文件.\n")
        sheet_name = os.path.splitext(os.path.basename(file_path))[0][:31]
        if os.path.exists(self.output_path):
            try:
                book = load_workbook(filename=self.output_path)
                if sheet_name in book.sheetnames:
                    if len(book.sheetnames)>1:
                        del book[sheet_name]
                        book.save(self.output_path)
                        mode='a'
                    else:
                        os.remove(self.output_path)
                        mode='w'
                else:
                    mode='a'
            except Exception as e:
                print("[ERROR] Error occurred: ", e)
        else:
            mode='w'
        with pd.ExcelWriter(self.output_path, engine='openpyxl', mode=mode) as writer:
            super().scan_pdf(file_path, writer)
            self.deal_excel(sheet_name)
        self.show_output()

    def scan_directory(self):
        """scan pdf files in directory and extract it into excel
        """
        dir_path = selected_directory.get()
        os.chdir(dir_path) 
        pdf_files = glob.glob('*.pdf')
        if not pdf_files:
            output_display.insert(tk.END, "[Error] 选定的目录中未发现PDF文件.\n")
        self.output_path = os.path.join(dir_path, self.output_file)
        with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
            for pdf_file in pdf_files:
                super().scan_pdf(pdf_file, writer)
                sheet_name = os.path.splitext(os.path.basename(self.pdf_path))[0][:31]
                self.deal_excel(sheet_name)
        self.show_output()


    def deal_excel(self, sheet_name):
        """Convert data to DataFrame for Excel

        Args:
            sheet_name (str): Excel sheet name
        """      
        if self.highlights_by_color: 
            output_display.insert(tk.END, f"[Info] 处理中: {self.pdf_path}\n")
            max_rows = max(len(texts) for texts in self.highlights_by_color.values())
            df = pd.DataFrame()
            for color in self.highlights_by_color.keys():
                col_name = self.rgb2color(color)
                df[col_name] = self.highlights_by_color[color] + [''] * (max_rows - len(self.highlights_by_color[color]))
                df[col_name + " Page"] = self.pages_by_color[color] + [''] * (max_rows - len(self.pages_by_color[color]))
            df.to_excel(self.writer, index=False, sheet_name=sheet_name)
        else:
            print(f"No highlighted text found in {self.pdf_path}.")
            output_display.insert(tk.END, f"[Warn] 无高亮: {self.pdf_path}\n")

    def show_output(self):
        """deal excel sheet and show output info
        """
        self.sort_sheets()
        output_display.insert(tk.END, f"[Success] 归类完成! 文件: {self.output_path}.\n\n\n")
        output_file_path_var.set(self.output_path)
        btn_start_scan.config(state=tk.NORMAL) 
        current_xview = output_display.xview()
        output_display.see(tk.END)
        output_display.xview_moveto(current_xview[0])

    def sort_sheets(self):
        """sort sheet by sheetname
        """
        sheet = load_workbook(self.output_path)
        sorted_sheet_names = sorted(sheet.sheetnames)
        print(sorted_sheet_names)
        for idx, sheet_name in enumerate(sorted_sheet_names):
            sheet[sheet_name].index = idx
        sheet.save(self.output_path)

    def open_output_file(self):
        """open output excel file 
        """
        path = output_file_path_var.get().strip()
        if not os.path.exists(path):
            messagebox.showwarning("警告", "路径不存在!")
            file_path = filedialog.askopenfile()
            if file_path:
                output_file_path_var.set(file_path.name)
            return
        if not path.endswith(('.xls', '.xlsx')):
            messagebox.showwarning("警告", "请选择一个有效的Excel文件!")
            return
        file_path = output_file_path_var.get().strip()
        if sys.platform.startswith("darwin"):
            os.system(f'open "{file_path}"')
        elif sys.platform.startswith("win32"):
            os.system(f'start "" "{file_path}"')


def main():
    # 网页有道翻译的cookie, 填入自己的cookie
    youdao_cookie = "填入自己的cookie" 
    
    # 百度翻译的api, 可以填入多个
    id_pool = {
        '填入自己的key1': '填入自己的secret2',
        '填入自己的key2': '填入自己的secret2',
    }
    
    # 本地字典的路径, 目前仅支持以下两本字典, 将仓库中的字典下载之后, 将路径换成自己的路径
    translate_books = [
        "/Users/user/Desktop/trans/英汉大词典_del_ipa_edited.txt",
        "/Users/user/Desktop/trans/英汉大词典_edited.txt",
    ]
    
    translator = Translator(
        view_workers=1,
        row_workers=6,
        trans_id_pool=id_pool,
        base_url='https://fanyi-api.baidu.com/api/trans/vip/translate',                     # 百度翻译api
        words_book_url='https://dict.youdao.com/wordbook/webapi/v2/ajax/add?lan=en&word=',  # 有道单词本api
        sleep_time='0.1',
        books=translate_books,
        tag_name="mytest",            # 生成的有道单词本xml中的tag, 可以自己指定
        cookies=youdao_cookie
    )
    scanner_gui = ScannerGui(
        name="pdfScanner v1.1(添加网易单词本)",
        size="380x375",
        output_file="output.xlsx",    # 生成excel文件的名字
        translator=translator,
    )
    scanner_gui.run()

if __name__ == "__main__":
    main()

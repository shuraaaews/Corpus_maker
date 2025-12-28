import sys
import os
import logging
import pathlib
from subprocess import Popen, PIPE
from docx import Document as Ddoc
from odfdo import Document as Dodf
from pathlib import  Path
import tkinter as tk
from tkinter import messagebox
import tkinter.ttk as ttk
import configparser
import json
from striprtf.striprtf import rtf_to_text



logger = logging.getLogger('corpus_maker_log')
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                                datefmt='%Y-%m-%d %H:%M:%S')
logfile = logging.FileHandler('logfile.log')
logfile.setFormatter(formatter)
logger.addHandler(logfile)



class TreeFrame(tk.Frame):
    def __init__(self, master, path, smode):
        super().__init__(master)
        abspath = os.path.abspath(path)
        self.nodes = {}
        self.tree = ttk.Treeview(self)
        self.tree.heading("#0", text=abspath, anchor=tk.W)
        ysb = ttk.Scrollbar(self, orient=tk.VERTICAL,
                            command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL,
                            command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set, selectmode = smode)
        self.tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
        ysb.grid(row=0, column=1, sticky=tk.N + tk.S)
        xsb.grid(row=1, column=0, sticky=tk.E + tk.W)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.tree.bind("<<TreeviewOpen>>", self.open_node)
        self.populate_node("", abspath)

    def populate_node(self, parent, abspath):
        for entry in os.listdir(abspath):
            entry_path = os.path.join(abspath, entry)
            node = self.tree.insert(parent, tk.END, text=entry, open=False)
            if os.path.isdir(entry_path):
                self.nodes[node] = entry_path
                self.tree.insert(node, tk.END)

    def open_node(self, event):
        item = self.tree.focus()
        abspath = self.nodes.pop(item, False)
        if abspath:
            children = self.tree.get_children(item)
            self.tree.delete(children)
            self.populate_node(item, abspath)

    def walk_node_up(self, select_node, lst_dir_des):
        node_dir = self.tree.parent(select_node)
        if node_dir !='':
            lst_dir_des.append(node_dir)
            self.walk_node_up(node_dir, lst_dir_des)
        else:
            return lst_dir_des

       


    @property
    def selected_id(self):
        selection = self.tree.selection()
        lst_id = [lf for lf in selection]
        if (len(lst_id)==1) and (self.tree.item(selection, 'text') in DIR_DCT):
            lst_dir_des=[lst_id[0]]
            self.walk_node_up(selection, lst_dir_des)
            lst_items = [self.tree.item(li, 'text') for li in lst_dir_des]
        elif (len(lst_id)==1) and (self.tree.item(selection, 'text') not in DIR_DCT):
            lst_dir_des=[lst_id[0]]
            self.walk_node_up(selection, lst_dir_des)
            lst_items = [self.tree.item(li, 'text') for li in lst_dir_des]
        elif (len(lst_id)>1):
            lst_items = []
            for row in lst_id:
                lst_dir_des=[row]
                self.walk_node_up(row, lst_dir_des)
                lst_items_one = [self.tree.item(li, 'text') for li in lst_dir_des]
                lst_items.append(lst_items_one)

        return lst_items if selection else None





class App(tk.Tk):
    def __init__(self, source, destination):
        super().__init__()
        self.title("CorpusMaker")
        self.geometry('1400x700')


        menu = tk.Menu(self)
        menu.add_command(label="About", command = self.show_about)
        menu.add_command(label="Quit", command=self.destroy)
        self.config(menu=menu)


        self.frame = ttk.Frame(self, relief  = tk.RIDGE)
        self.frame.grid(row=0, column=0, sticky='nsew')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(2, weight=1)

        self.frame_a = ttk.Frame(self.frame, relief  = tk.RIDGE, borderwidth = 15)
        self.frame_b = ttk.Frame(self.frame, relief  = tk.RIDGE, borderwidth = 15)
        self.frame_button = ttk.Frame(self.frame, relief  = tk.RIDGE, borderwidth = 15)
        self.frame_a.grid(row=0, column=0, sticky=tk.N + tk.S + tk.W + tk.E)
        self.frame_b.grid(row=0, column=2, sticky=tk.N + tk.S + tk.W + tk.E)
        self.frame_button.grid(row=0, column=1, sticky=tk.N + tk.S + tk.E + tk.W)
        self.frame_a.rowconfigure(0, weight=1)
        self.frame_a.columnconfigure(0, weight=1)
        self.frame_b.rowconfigure(0, weight=1)
        self.frame_b.columnconfigure(0, weight=1)

        self.progressbar = ttk.Progressbar(self.frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progressbar.grid(row=1, columnspan=3, sticky=tk.S + tk.W + tk.E)

        self.frame_a_tree = TreeFrame(self.frame_a, source, smode = "extended")

        self.btn_right = tk.Button(self.frame_button, width = 10, text=">",
                                   command=self.convert_txt)
        self.frame_b_tree = TreeFrame(self.frame_b, destination, smode = "browse")

        self.frame_a_tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
        self.frame_b_tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
        self.btn_right.pack(expand=True, ipadx=5)

    def reset_progressbar(self):
        self.progressbar.config(value=0)

    def show_about(self):
        about_message = "Corpus Maker"
        about_detail = ("""by Alexandr Evsunin
        For convert .docx(.odt) into .txt
        See detail in Readme.... """)
        messagebox.showinfo(
            title='About', message=about_message, detail=about_detail)

    def convert_txt(self):

        destination_dir = self.frame_b_tree.selected_id
        source_files = self.frame_a_tree.selected_id
        if destination_dir is None or source_files is None:
            return
        destination_dir.reverse()
        dir_dst = Path(destination).joinpath(*destination_dir)

        source_lst = []
        for el in source_files:
            if isinstance(el, list):
                el.reverse()
                file_s = Path(source).joinpath(*el)
                source_lst.append(file_s)
                continue
            else:
                source_files.reverse()
                file_s = Path(source).joinpath(*source_files)
                source_lst.append(file_s)
                break    

        if any(dr.is_dir() for dr in source_lst) or dir_dst.is_file():
            messagebox.showwarning("WARNING !!!!!", """Для конвертации выбирать только файлы !!!
                А для записи - только папку !!!""")
            return
        file_new_lst = [dir_dst.name]
        for count in range(10):
            file_new_lst.append(dir_dst.parents[count].name)
            if dir_dst.parents[count].name == Path(destination).name:
                break
        file_new = ""
        file_new_lst = file_new_lst[:-1]
        for _ in range(len(file_new_lst)):
            try:
                file_new += DIR_DCT[file_new_lst.pop()]
            except:
                file_new = ""
                messagebox.showwarning("WARNING !!!!!", "Несоответствие в именах папок и config.ini")
                return
                
        lst_dest_files = [_ for _ in dir_dst.iterdir()]
        if any(_.is_dir() for _ in lst_dest_files):
            messagebox.showwarning("WARNING !!!!!", "Есть вложенные каталоги !!!")
            return
        lst_dest_stem_files = [x.stem for x in dir_dst.iterdir()]
        
        if len(lst_dest_stem_files)>0:
            max_dest_files_number = max([int(x.lstrip(file_new))
                                         for x in lst_dest_stem_files])
            number = max_dest_files_number +1
        else:
            number = 1

        self.convert_doc_txt(source_lst, dir_dst, file_new, number)

    def convert_doc_txt(self, source_lst, dir_dst, file_new, number):
        lst_err = []    
        self.progressbar['maximum'] = len(source_lst)
        for count, f in enumerate(source_lst, 1):
            file_txt = self.document_to_text(Path(f).name, Path(f))
            if file_txt is None:
                lst_err.append(Path(f).name)
                if verbose:
                    logger.warning(f' {Path(f).name} convert FALIED !!!')
                continue
            else:
                if verbose:
                    logger.info(f' {Path(f).name} convert succsesful')
                file_destination = os.path.abspath(
                        os.path.join(Path(dir_dst),file_new + str(number) + '.txt'))
                with open(file_destination, 'wt', encoding = 'utf-8') as ff:
                    ff.write(file_txt)
            self.progressbar['value'] = count
            number += 1
            self.update_idletasks()
        self.after(500, self.reset_progressbar)
        self.frame_b_tree.tree.delete(*self.frame_b_tree.tree.get_children())
        self.frame_b_tree.populate_node("",destination)
        if messagebox.askyesno("INFO!", f"""Успешно конвертировано: {count - len(lst_err)}\n
        Ошибка конвертации: {len(lst_err)}\n
        Удалить конвертированные файлы из источника ?""", default = messagebox.NO):
            lst_del = [f for f in source_lst if Path(f).name not in lst_err]
            if messagebox.askyesno("WARNING !!!!!", "ВЫ УВЕРЕНЫ ? ФАЙЛЫ БУДУТ УДАЛЕНЫ БЕЗВОЗВРАТНО !"):
                for fd in lst_del:
                    fd.unlink(missing_ok=True)
                self.frame_a_tree.tree.delete(*self.frame_a_tree.tree.get_children())
                self.frame_a_tree.populate_node("",source)
            else:
                pass
        else:
            pass
                
            
   
        
    def document_to_text(self, filename, file_path):
        try:
            if filename[-4:] == ".doc" and os.name == "posix":
                cmd = ['antiword', file_path]
                p = Popen(cmd, stdout=PIPE)
                stdout, stderr = p.communicate()
                return stdout.decode('utf-8', 'ignore')
            elif filename[-5:] == ".docx":
                document = Ddoc(file_path)
                newparatextlist = []
                for paratext in document.paragraphs:
                    newparatextlist.append(paratext.text)
                return '\n\n'.join(newparatextlist)
            elif filename[-4:] == ".odt":
                document = Dodf(file_path)
                text_content = document.get_formatted_text()
                return text_content
            elif filename[-4:] == ".rtf":
                with open(file_path, 'r') as ff:
                    rtf_str = ff.read()
                text_content = rtf_to_text(rtf_str, encoding = 'utf-8', errors='ignore')
                return text_content
        except:
            messagebox.showwarning("WARNING !!!!!", f"Не удается правильно декодировать файл {filename}")
            return
        


def main(source, destination, verbose):
    app = App(source, destination)
    app.mainloop()


if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Access values
    try:
        source = config['instalation']['source']
        destination = config['instalation']['destination']
        if config['instalation']['verbose']=="True":
            verbose = True
        else:
            verbose = False
        DIR_DCT = json.loads(config['tags']['DIR_DCT'])
        assert Path(source).is_dir()
        assert Path(destination).is_dir()

    except:
        messagebox.showwarning("WARNING !!!!!", "Не удается правильно прочитать файл config.ini")
        sys.exit()
    main(source, destination, verbose)

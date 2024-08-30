from kivy.lang import Builder
from kivymd.app import MDApp
from kivy.uix.boxlayout import BoxLayout
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDRectangleFlatButton
from kivymd.uix.card import MDCard
from tkinter import Tk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
import os


# import keyboard

class MainApp(MDApp):

    def build(self):
        self.theme_cls.theme_style = "Dark"
        # self.theme_cls.theme_style = "Light"
        self.theme_cls.primary_palette = "Yellow"

        # return Builder.load_file("s_kivy.kv")
        return Builder.load_string(
            '''
Screen:
    MDCard:
    
        size_hint : None, None
        size:300 , 500
        pos_hint:{'center_x':.5 , 'center_y':.5}
        
        BoxLayout:
            orientation:'vertical'
            spacing:20
            padding:20
            MDRaisedButton:
                id:spdb
                text:"Select Pictures Directory"
                pos_hint:{"center_x":0.5 , 'center_y':.5}
                on_press:app.selectPhotoDirectory()
            MDLabel:
                id:spd
                hint_text:"Pictures Directory"
                size_hint_x:.5
                height:100
                font_size:18
                pos_hint:{'x' : .25}
            MDRaisedButton:
                id:sefb
                text:"Select excel file"
                pos_hint:{"center_x":0.5 , 'center_y':.5}
                on_press:app.selectExcelFile()
            MDLabel:
                id:sef
                hint_text:"Percent"
                size_hint_x:.5
                height:100
                font_size:18
                pos_hint:{'x' : .25}
                
            BoxLayout:
                size_hint: None, None
                pos_hint: {'top': 1}
                
                BoxLayout:
                    orientation:"vertical"
                    size_hint: None, None
                    MDTextField:
                        id:oldext
                        hint_text:"old ext"
                        pos_hint:{"center_x":0.75 , 'center_y':.5}
    
                
                BoxLayout:
                    orientation:"vertical"
                    size_hint: None, None
                    MDTextField:
                        id:newext
                        hint_text:"new ext"
                        pos_hint:{"center_x":0.75 , 'center_y':.5}

            
            MDRectangleFlatButton:
                text:"Submit"
                size_hint_x:.7
                pos_hint:{"center_x":0.5 , 'center_y':.5}
                on_press:app.calculate()
            
            '''
        )

    def calculate(self):
        #TODO:>>>>
        print("")
        label = self.root.ids.spd
        label.text = tempdir
        def prefix_remover(path, prefix):
            for file in os.listdir(path):
                if file.startswith(prefix):
                    new_name = file.replace(prefix, "")
                    os.rename(os.path.join(path, file), os.path.join(path, new_name))

        def converter(pics_path, excel_full_path, new_ext, prefix):
            print("please Wait")
            workbook = load_workbook(excel_full_path)
            sheet = workbook.active
            for row in sheet.iter_rows(values_only=True):
                for file_name in os.listdir(pics_path):
                    prefix_remover(pics_path, prefix)
                    file_raw_name = file_name.split(".")[0]
                    if str(row[0]) == file_raw_name:
                        new_name = (str(row[2]) + "." + new_ext)
                        os.rename(os.path.join(pics_path, file_name),
                                  os.path.join(pics_path, new_name))
            print("Done")

        with open('settings') as f:
            json_data = load(f)

        pictures_path = json_data["pictures path"]
        excel_path = json_data["excel path"]
        new_extension = json_data["new extension"]

        converter(pics_path=pictures_path,
                  excel_full_path=excel_path,
                  new_ext=new_extension,
                  prefix="LF")

    def selectPhotoDirectory(self):
        root = Tk()
        root.withdraw()

        def search_for_file_path():
            currdir = os.getcwd()
            tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
            if len(tempdir) > 0:
                print("You chose: %s" % tempdir)
            label = self.root.ids.spd
            label.text = tempdir
            # ============
            return tempdir

        search_for_file_path()

    def selectExcelFile(self):
        root = Tk()
        root.withdraw()

        def search_for_file_path():
            currdir = os.getcwd()
            tempdir = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls *.xlsb *.xlsm")])
            if len(tempdir) > 0:
                print("You chose: %s" % tempdir)

            label = self.root.ids.sef
            label.text = tempdir
            # ============
            return tempdir

        search_for_file_path()


MainApp().run()

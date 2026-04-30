import os
import pandas as pd
import gdown
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.graphics import Color, Line, Rectangle

class CellLabel(Label):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        with self.canvas.before:
            Color(0, 0, 0, 0)  # nền trong suốt
            self.rect = Rectangle(pos=self.pos, size=self.size)
            Color(1, 1, 1, 1)  # viền trắng
            self.line = Line(rectangle=(self.x, self.y, self.width, self.height), width=1)

        self.bind(pos=self.update_rect, size=self.update_rect)

    def update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size
        self.line.rectangle = (self.x, self.y, self.width, self.height)

class HomeScreen(Screen):
    def __init__(self, app, **kwargs):
        super().__init__(**kwargs)
        self.app = app
        layout = BoxLayout(orientation='vertical')

        self.code_input = TextInput(hint_text="Nhập mã vạch", multiline=False)
        self.name_input = TextInput(hint_text="Nhập tên hàng", multiline=False)
        self.search_btn = Button(text="Tra cứu")
        self.search_btn.bind(on_press=self.search_item)

        layout.add_widget(self.code_input)
        layout.add_widget(self.name_input)
        layout.add_widget(self.search_btn)

        self.result = Label(text="Kết quả sẽ hiển thị ở đây")
        layout.add_widget(self.result)

        self.add_widget(layout)

    def search_item(self, instance):
        df = self.app.df
        if df is None:
            self.result.text = "Không thể đọc dữ liệu Excel."
            return

        code = self.code_input.text.strip()
        name = self.name_input.text.strip()

        if code:
            code = code.replace(".0", "")
            item = df[df['Barcode'] == code]
        elif name:
            item = df[df['Tên sản phẩm'].str.contains(name, case=False, na=False)]
        else:
            self.result.text = "Vui lòng nhập mã vạch hoặc tên hàng."
            return

        self.app.result_screen.show_result(item)
        self.app.sm.current = "result"

class ResultScreen(Screen):
    def __init__(self, app, **kwargs):
        super().__init__(**kwargs)
        self.app = app
        layout = BoxLayout(orientation='vertical')

        self.scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.grid = GridLayout(cols=8, size_hint_y=None, size_hint_x=None,
                               spacing=0, padding=0)
        self.grid.bind(minimum_height=self.grid.setter('height'),
                       minimum_width=self.grid.setter('width'))
        self.scroll.add_widget(self.grid)

        layout.add_widget(self.scroll)

        self.back_btn = Button(text="Quay lại")
        self.back_btn.bind(on_press=self.go_back)
        layout.add_widget(self.back_btn)

        self.add_widget(layout)

    def show_result(self, item):
        self.grid.clear_widgets()
        if not item.empty:
            cols = [
                ("Barcode", 150),
                ("Tên sản phẩm", 300),
                ("Giá lẻ", 100),
                ("Hàng bình thuờng", 120),
                ("Tên NCC", 200),
                ("Ngành", 100),
                ("Nhóm", 100),
                ("Thể loại", 120)
            ]

            # Tiêu đề
            for c, w in cols:
                self.grid.add_widget(CellLabel(
                    text=c, bold=True,
                    size_hint_x=None, width=w,
                    size_hint_y=None, height=50,
                    halign="left", valign="middle",
                    text_size=(w-10, 50),
                    padding_x=5
                ))

            # Dữ liệu
            for _, row in item.iterrows():
                for c, w in cols:
                    val = row[c]

                    if c in ["Barcode", "Giá lẻ"]:
                        if pd.notna(val):
                            val = str(val).replace(".0", "")
                        else:
                            val = ""

                    if c == "Hàng bình thuờng":
                        self.grid.add_widget(CellLabel(
                            text=str(val),
                            size_hint_x=None, width=w,
                            size_hint_y=None, height=50,
                            halign="left", valign="middle",
                            text_size=(w-10, 50),
                            padding_x=5,
                            bold=True, color=(1,0,0,1)
                        ))
                    elif c == "Tên NCC":
                        self.grid.add_widget(CellLabel(
                            text=str(val),
                            size_hint_x=None, width=w,
                            size_hint_y=None, height=80,
                            halign="left", valign="top",
                            text_size=(w-10, 80),
                            padding_x=5,
                            shorten=False
                        ))
                    elif c == "Tên sản phẩm":
                        self.grid.add_widget(CellLabel(
                            text=str(val),
                            size_hint_x=None, width=w,
                            size_hint_y=None, height=100,
                            halign="left", valign="top",
                            text_size=(w-10, 100),
                            padding_x=5,
                            shorten=False
                        ))
                    else:
                        self.grid.add_widget(CellLabel(
                            text=str(val),
                            size_hint_x=None, width=w,
                            size_hint_y=None, height=50,
                            halign="left", valign="middle",
                            text_size=(w-10, 50),
                            padding_x=5,
                            shorten=False
                        ))
        else:
            self.grid.add_widget(CellLabel(text="Không tìm thấy mặt hàng.",
                                           size_hint_y=None, height=50,
                                           text_size=(200, 50),
                                           valign="middle"))

    def go_back(self, instance):
        self.app.sm.current = "home"

class StockApp(App):
    def build(self):
        # ID file Excel trên Google Drive
        FILE_ID = "128wfi3wHWjd690Cxk-3p5EiIuav-4Zpx"
        url = f"https://drive.google.com/uc?id={FILE_ID}"
        EXCEL_FILE = "TONKHO.xlsx"

        try:
            # tải file từ Google Drive
            gdown.download(url, EXCEL_FILE, quiet=False)

            # đọc dữ liệu Excel
            self.df = pd.read_excel(EXCEL_FILE, header=5)
            self.df['Barcode'] = self.df['Barcode'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        except Exception as e:
            self.df = None
            print("Không thể đọc file Excel:", e)

        self.sm = ScreenManager()
        self.home_screen = HomeScreen(self, name="home")
        self.result_screen = ResultScreen(self, name="result")

        self.sm.add_widget(self.home_screen)
        self.sm.add_widget(self.result_screen)

        return self.sm

if __name__ == "__main__":
    StockApp().run()

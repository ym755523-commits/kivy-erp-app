# -*- coding: utf-8 -*-
from kivy.app import App
from kivy.core.text import LabelBase
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
import sqlite3
import os

# تسجيل الخط العربي
LabelBase.register(
    name='Amiri',
    fn_regular='Amiri-Regular.ttf',
    fn_bold='Amiri-Bold.ttf'
)

def ar(text):
    """دالة بسيطة للعربية — Kivy يدعم RTL تلقائياً"""
    return text


class TestScreen(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation='vertical', padding=20, spacing=10, **kwargs)
        
        self.add_widget(Label(
            text='مرحباً بكم في نظام الجمارك',
            font_name='Amiri',
            font_size='30sp',
            size_hint_y=None, height=80
        ))
        
        self.add_widget(Label(
            text='Customs ERP - Test Build',
            font_size='20sp',
            size_hint_y=None, height=60
        ))
        
        self.input = TextInput(
            hint_text='اكتب اسماً...',
            font_name='Amiri',
            font_size='20sp',
            size_hint_y=None, height=60
        )
        self.add_widget(self.input)
        
        btn = Button(
            text='اختبار قاعدة البيانات',
            font_name='Amiri',
            font_size='20sp',
            size_hint_y=None, height=60
        )
        btn.bind(on_press=self.test_db)
        self.add_widget(btn)
        
        self.status = Label(
            text='اضغط الزر لاختبار SQLite',
            font_name='Amiri',
            font_size='18sp',
            size_hint_y=None, height=60
        )
        self.add_widget(self.status)
        
        self.add_widget(Label())  # مسافة
    
    def test_db(self, *args):
        try:
            conn = sqlite3.connect('test.db')
            c = conn.cursor()
            c.execute('CREATE TABLE IF NOT EXISTS test (id INTEGER, name TEXT)')
            c.execute('INSERT INTO test VALUES (1, ?)', (self.input.text or 'بدون اسم',))
            conn.commit()
            c.execute('SELECT * FROM test')
            rows = c.fetchall()
            conn.close()
            self.status.text = f'✅ نجح! السجلات: {len(rows)}'
        except Exception as e:
            self.status.text = f'❌ خطأ: {e}'


class TestApp(App):
    def build(self):
        self.title = 'Test Customs ERP'
        return TestScreen()


if __name__ == '__main__':
    TestApp().run()

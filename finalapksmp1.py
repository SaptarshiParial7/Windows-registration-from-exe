from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.metrics import dp
from datetime import datetime
import openpyxl
from datetime import timedelta

from openpyxl import Workbook
import os

# New directory to save the Excel data (current working directory)
FILE_SAVE_PATH = os.getcwd()

if not os.path.exists(FILE_SAVE_PATH):
    os.makedirs(FILE_SAVE_PATH)

class CalendarWidget(GridLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cols = 7
        self.spacing = 5
        self.selected_date = None

        # Header row
        for day in ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]:
            self.add_widget(Label(text=day, size_hint=(None, None), size=(dp(40), dp(40))))

        self.populate_calendar()

    def populate_calendar(self):
        self.clear_widgets()
        for day in ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]:
            self.add_widget(Label(text=day, size_hint=(None, None), size=(dp(40), dp(40))))

        today = datetime.now()
        first_day = today.replace(day=1)
        start_day = first_day.weekday()
        days_in_month = (first_day.replace(month=(today.month % 12) + 1, day=1) - timedelta(days=1)).day

        # Add empty cells for the start of the week
        for _ in range(start_day):
            self.add_widget(Label(size_hint=(None, None), size=(dp(40), dp(40))))

        for day in range(1, days_in_month + 1):
            btn = Button(text=str(day), size_hint=(None, None), size=(dp(40), dp(40)))
            btn.bind(on_press=self.select_date)
            self.add_widget(btn)

    def select_date(self, instance):
        today = datetime.now()
        self.selected_date = today.replace(day=int(instance.text)).strftime("%Y-%m-%d")
        print(f"Selected Date: {self.selected_date}")


class OTPApp(App):
    def build(self):
        self.layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Title
        self.layout.add_widget(Label(text="User Registration", font_size=24, size_hint=(1, None), height=dp(50)))

        # Name input
        self.name_input = TextInput(hint_text="Enter your name", multiline=False, size_hint=(1, None), height=dp(40))
        self.layout.add_widget(self.name_input)

        # Email input
        self.email_input = TextInput(hint_text="Enter your email", multiline=False, size_hint=(1, None), height=dp(40))
        self.layout.add_widget(self.email_input)

        # Phone input
        self.phone_input = TextInput(hint_text="Enter your phone", multiline=False, size_hint=(1, None), height=dp(40))
        self.layout.add_widget(self.phone_input)

        # Calendar widget
        scroll_view = ScrollView(size_hint=(1, 0.6))
        self.calendar = CalendarWidget(size_hint=(None, None), size=(dp(280), dp(400)))
        scroll_view.add_widget(self.calendar)
        self.layout.add_widget(scroll_view)

        # Generate OTP button
        self.generate_button = Button(text="Generate OTP", size_hint=(1, None), height=dp(50))
        self.generate_button.bind(on_press=self.generate_otp)
        self.layout.add_widget(self.generate_button)

        # OTP input
        self.otp_input = TextInput(hint_text="Enter OTP", multiline=False, size_hint=(1, None), height=dp(40))
        self.layout.add_widget(self.otp_input)

        # Verify button
        self.verify_button = Button(text="Verify OTP", size_hint=(1, None), height=dp(50))
        self.verify_button.bind(on_press=self.verify_otp)
        self.layout.add_widget(self.verify_button)

        return self.layout

    def generate_otp(self, instance):
        if not self.name_input.text.strip() or not self.email_input.text.strip() or not self.phone_input.text.strip() or not self.calendar.selected_date:
            popup = Popup(title="Error",
                          content=Label(text="Please fill in all fields and select a date."),
                          size_hint=(0.8, 0.4))
            popup.open()
            return

        self.otp = str(datetime.now().microsecond % 1000000).zfill(6)
        print(f"Generated OTP: {self.otp}")

        popup = Popup(title="OTP Generated",
                      content=Label(text=f"Your OTP is: {self.otp}"),
                      size_hint=(0.8, 0.4))
        popup.open()

    def verify_otp(self, instance):
        entered_otp = self.otp_input.text.strip()
        if entered_otp == self.otp:
            self.save_data()
        else:
            popup = Popup(title="Error",
                          content=Label(text="Invalid OTP! Please try again."),
                          size_hint=(0.8, 0.4))
            popup.open()

    def save_data(self):
        name = self.name_input.text.strip()
        email = self.email_input.text.strip()
        phone = self.phone_input.text.strip()
        date = self.calendar.selected_date
        save_time = datetime.now().strftime("%H:%M:%S")  # Capture the current time

        if not name or not email or not phone or not date:
            popup = Popup(title="Error",
                          content=Label(text="Please fill in all fields and select a date."),
                          size_hint=(0.8, 0.4))
            popup.open()
            return

        # Generate file path in the current working directory
        file_name = os.path.join(FILE_SAVE_PATH, "UserData.xlsx")

        try:
            if not os.path.exists(file_name):
                wb = Workbook()
                sheet = wb.active
                sheet.title = "User Data"
                sheet.append(["Name", "Email", "Phone", "Date", "Time"])  # Add header row
                wb.save(file_name)

            # Append data to the Excel file
            wb = openpyxl.load_workbook(file_name)
            sheet = wb.active
            sheet.append([name, email, phone, date, save_time])
            wb.save(file_name)

            popup = Popup(title="Success",
                          content=Label(text="Your data has been saved successfully!"),
                          size_hint=(0.8, 0.4))
            popup.open()
            self.reset_form()

        except Exception as e:
            popup = Popup(title="Error",
                          content=Label(text=f"Error saving data: {e}"),
                          size_hint=(0.8, 0.4))
            popup.open()

    def reset_form(self):
        self.name_input.text = ""
        self.email_input.text = ""
        self.phone_input.text = ""
        self.otp_input.text = ""
        self.calendar.populate_calendar()


if __name__ == "__main__":
    OTPApp().run()

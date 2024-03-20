import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinter import ttk
import openai
import os
from datetime import datetime, timedelta
import sv_ttk
import win32com.client as com
import win32com.client as win32
import win32com.client

#client = openai()

def generate_appointment_details(prompt, self):
  OPENAI_API_KEY = self.api_to_entry.get()
  try:
    completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-0125",
        messages=[{
            "role":
            "system",
            "content":
            "You are an appointment assistant. You provide appointment information for the given prompt."
        }, {
            "role": "user",
            "content": prompt
        }],
        max_tokens=200,
        temperature=0.7)
    # Assuming the structure of the completion object includes 'choices' and you want the 'content' of the first choice.
    # Note: For ChatCompletion, the response structure might differ, so adjust the following line accordingly.
    return completion.choices[0].message.content.strip()
  except Exception as e:
    print(f"An error occurred: {e}")
    return ""


def generate_email_body(prompt, self):
  OPENAI_API_KEY = self.api_to_entry.get()
  try:
    completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-0125",
        messages=[{
            "role":
            "system",
            "content":
            "You are an email assistant. You provide email body for the given prompt without subject."
        }, {
            "role": "user",
            "content": prompt
        }],
        max_tokens=200,
        temperature=0.7)
    # Assuming the structure of the completion object includes 'choices' and you want the 'content' of the first choice.
    # Note: For ChatCompletion, the response structure might differ, so adjust the following line accordingly.
    return completion.choices[0].message.content.strip()
  except Exception as e:
    print(f"An error occurred: {e}")
    return ""


def schedule_outlook_appointment(subject, start_datetime, duration_minutes,
                                 location, body):
  try:
    import win32com.client as com
    import win32com.client as win32
    outlook = win32com.client.Dispatch("Outlook.Application")
    appointment = outlook.CreateItem(1)  # 1 = olAppointmentItem
    appointment.Subject = subject
    appointment.Start = start_datetime
    appointment.Duration = duration_minutes
    appointment.Location = location
    appointment.Body = body
    appointment.Save()
    print("Appointment scheduled in Outlook.")
  except Exception as e:
    print(f"An error occurred while scheduling the appointment: {e}")


def send_email(to, subject, body):
  print("Email sent. This is a simulation.")


class ApplicationGUI:

  def __init__(self, root):
    self.root = root
    self.root.title("GPT-3 Assistant")
    self.create_appointment_widgets()
    self.create_email_widgets()

  def create_virtual_keyboard(self):
    # Keyboard creation logic goes here
    pass

  def create_appointment_widgets(self):
    ttk.Label(self.root, text="Schedule Appointment").pack(pady=(10, 0))

    # Brief Description
    ttk.Label(self.root, text="Brief Description:").pack()
    self.appointment_description_entry = ttk.Entry(self.root, width=50)
    self.appointment_description_entry.pack()

    # Start Date and Time
    ttk.Label(self.root, text="Start Date and Time (YYYY-MM-DD HH:MM):").pack()
    self.appointment_start_entry = ttk.Entry(self.root, width=50)
    self.appointment_start_entry.pack()

    # Duration in Minutes
    ttk.Label(self.root, text="Duration (Minutes):").pack()
    self.appointment_duration_entry = ttk.Entry(self.root, width=50)
    self.appointment_duration_entry.pack()

    # Location (Optional)
    ttk.Label(self.root, text="Location (Optional):").pack()
    self.appointment_location_entry = ttk.Entry(self.root, width=50)
    self.appointment_location_entry.pack()

    # Details/Body
    ttk.Label(self.root, text="Details:").pack()
    self.appointment_detail_text = scrolledtext.ScrolledText(self.root, height=5, width=50)
    self.appointment_detail_text.pack()

    # Buttons
    ttk.Button(self.root, text="Generate Appointment Details", command=self.generate_appointment_details).pack(pady=(5, 0))
    ttk.Button(self.root, text="Schedule Appointment", command=self.schedule_appointment).pack(pady=(5, 10))


  def create_email_widgets(self):
    ttk.Label(self.root, text="Compose Email").pack()
    ttk.Label(self.root, text="Outlook Account Name:").pack()
    self.email_account_entry = tk.Entry(self.root)
    self.email_account_entry.pack()
    ttk.Label(self.root, text="To:").pack()
    self.email_to_entry = tk.Entry(self.root)
    self.email_to_entry.pack()
    ttk.Label(self.root, text="Subject:").pack()
    self.email_subject_entry = tk.Entry(self.root)
    self.email_subject_entry.pack()
    ttk.Label(self.root, text="Brief Description:").pack()
    self.email_description_entry = tk.Entry(self.root, width=50)
    self.email_description_entry.pack()
    self.email_body_text = scrolledtext.ScrolledText(self.root,
                                                     height=5,
                                                     width=50)
    self.email_body_text.pack()
    ttk.Button(self.root,
               text="Generate Email Content",
               command=self.generate_email_content).pack()
    ttk.Button(self.root, text="Send Email", command=self.write_email).pack()


#end of class

  def generate_appointment_details(self):
    prompt = self.appointment_description_entry.get()
    details = generate_appointment_details(prompt, self)
    self.appointment_detail_text.delete('1.0', tk.END)
    self.appointment_detail_text.insert('1.0', details)

  def generate_email_content(self):
    description = self.email_description_entry.get()
    prompt = f"Write a detailed email about: {description}"
    generated_body = generate_email_body(prompt, self)
    self.email_body_text.delete("1.0", tk.END)
    self.email_body_text.insert("1.0", generated_body)

  def schedule_appointment(self):
     try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        appointment = outlook.CreateItem(1)  # 1 = olAppointmentItem
        appointment.Subject = subject
        appointment.Start = start_datetime
        appointment.Duration = duration_minutes
        appointment.Location = location
        appointment.Body = body
        appointment.Save()
        print("Appointment scheduled in Outlook.")
     except Exception as e:
        print(f"An error occurred while scheduling the appointment: {e}")

  def write_email(self):
    

    account = self.email_account_entry.get()
    to = self.email_to_entry.get()
    subject = self.email_subject_entry.get()
    body = self.email_body_text.get("1.0", tk.END).strip()
    try:
      import win32com.client as com
      import win32com.client as win32
      # Start an instance of Outlook
      outlook = win32com.client.Dispatch("Outlook.Application")

      # Create a new email item
      olmailitem=0 #size of the new email
      mail = outlook.CreateItem(olmailitem)  # 0 is the code for a mail item (see Outlook's Item Types)

      # Set the recipient, subject, and body of the email
      mail.To = to
      mail.Subject = subject
      mail.Body = body
      #mail.GetInspector 

      # Optionally, you can add an attachment to the email like this:
      # attachment_path = "path_to_your_attachment"
      # mail.Attachments.Add(attachment_path)

      # Send the email
      mail.Display()
      #mail.Send()

      print("Email sent successfully!")

    except Exception as e:
      print(f"An error occurred: {e}")

  def open_outlook():
    try:
        subprocess.call([r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'])
        os.system(r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE');
    except:
        print("Outlook didn't open successfully")
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            flag = 1
            break
        else:
            flag = 0
    if (flag == 1):
        send_notification()
    else:
        open_outlook()
        send_notification()

  def create_virtual_keyboard(self):
    # Define the keyboard layout
    keyboard_layout = [['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'],
                       ['q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p'],
                       ['a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l'],
                       ['z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.', '?'],
                       ['space', 'backspace']]

    # Create a frame to contain the keyboard
    keyboard_frame = tk.Frame(self.root)
    keyboard_frame.pack(pady=20)

    # Create buttons for each key in the layout
    for row in keyboard_layout:
      row_frame = tk.Frame(keyboard_frame)
      row_frame.pack()

      for key in row:
        if key == 'space':
          btn = tk.Button(row_frame,
                          text='Space',
                          width=20,
                          command=lambda: self.key_pressed(' '))
        elif key == 'backspace':
          btn = tk.Button(row_frame,
                          text='←',
                          width=20,
                          command=self.backspace_pressed)
        else:
          btn = tk.Button(row_frame,
                          text=key.upper(),
                          width=5,
                          command=lambda k=key: self.key_pressed(k))
        btn.pack(side='left', padx=3, pady=3)

  def key_pressed(self, key):
    focused_widget = self.root.focus_get()
    if isinstance(focused_widget, tk.Entry) or isinstance(
        focused_widget, scrolledtext.ScrolledText):
      focused_widget.insert(tk.INSERT, key)

  def backspace_pressed(self):
    focused_widget = self.root.focus_get()
    if isinstance(focused_widget, tk.Entry) or isinstance(
        focused_widget, scrolledtext.ScrolledText):
      if isinstance(focused_widget, tk.Entry):
        current_text = focused_widget.get()
        focused_widget.delete(0, tk.END)
        focused_widget.insert(0, current_text[:-1])
      else:  # For ScrolledText widget
        focused_widget.delete("insert -1 chars", tk.INSERT)

if __name__ == "__main__":
  root = tk.Tk()
  app = ApplicationGUI(root)
  sv_ttk.use_dark_theme()
  root.mainloop()

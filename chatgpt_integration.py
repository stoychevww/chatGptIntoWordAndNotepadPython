import os
import openai
import win32gui
import win32com.client
import tkinter as tk
from keyboard import add_hotkey
from tkinter import scrolledtext

# Set up OpenAI API Key
openai.api_key = 'API_KEY'





# Define file paths to store questions and answers
WORD_SAVE_PATH = "word_chat_history.txt"
NOTEPAD_SAVE_PATH = "notepad_chat_history.txt"


# Check if the active window is Word or Notepad
def get_active_application():
    try:
        word = win32com.client.Dispatch("Word.Application")
        if word.Visible:
            return "Word"
    except:
        pass

    hwnd = win32gui.GetForegroundWindow()
    window_title = win32gui.GetWindowText(hwnd)

    if "Notepad" in window_title:
        return "Notepad"

    return None


# Save chat to a file based on the active application
def save_chat(app_name, question, answer):
    save_path = WORD_SAVE_PATH if app_name == "Word" else NOTEPAD_SAVE_PATH
    with open(save_path, 'a') as f:
        f.write(f"Q: {question}\nA: {answer}\n\n")


# ChatGPT Query Function

def ask_chatgpt(question):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": question}],
            max_tokens=60000  # Reduce token usage per response
        )

        # Extract the response content
        answer = response['choices'][0]['message']['content'].strip()
        return answer
    except Exception as e:
        return str(e)




# Create the Bubble UI
def show_bubble_ui():
    app_name = get_active_application()
    if not app_name:
        return

    window = tk.Tk()
    window.title("ChatGPT Bubble")
    window.geometry("300x300+1000+100")
    window.configure(bg='black')
    window.attributes("-topmost", True)

    question_label = tk.Label(window, text="Ask ChatGPT:", fg='white', bg='black')
    question_label.pack(pady=10)

    question_box = tk.Entry(window, width=40)
    question_box.pack(pady=10)

    def on_submit():
        question = question_box.get()
        answer = ask_chatgpt(question)
        save_chat(app_name, question, answer)

        answer_box.delete(1.0, tk.END)
        answer_box.insert(tk.END, answer)

    submit_button = tk.Button(window, text="Submit", command=on_submit)
    submit_button.pack(pady=10)

    answer_box = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=40, height=10, bg='black', fg='white')
    answer_box.pack(pady=10)

    window.mainloop()


# Set up Keyboard Shortcut
add_hotkey("ctrl+alt+c", show_bubble_ui)

# Keep the script running
while True:
    pass

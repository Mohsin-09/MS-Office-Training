import tkinter as tk
from tkinter import messagebox
from plyer import notification
import schedule
import time
import threading
import os
from datetime import datetime

month_data = {
    "objective": "Build foundational skills in Microsoft Word, PowerPoint, and Excel for creating documents, presentations, and managing data.",
    "days": [
        {
            "day": 1,
            "objective": "Familiarize with Word's interface and basic document creation.",
            "tasks": [
                "Create a new Word document titled 'About Me'. Write a short paragraph introducing yourself. Save the document as 'Day1_AboutMe.docx'."
            ],
            "quizzes": [
                "What is Word's default font?"
            ]
        },
        {
            "day": 2,
            "objective": "Learn to format text using basic tools.",
            "tasks": [
                "Open 'Day1_AboutMe.docx' and apply different font styles, sizes, and colors to your text. Use bold, italic, and underline to emphasize important words."
            ],
            "quizzes": [
                "How do you format text as bold?"
            ]
        },
        {
            "day": 3,
            "objective": "Practice writing and formatting essays.",
            "tasks": [
                "Write a half-page essay about your favorite hobby. Adjust the line spacing to 1.5, align the text to 'Justify', and indent the first line of each paragraph."
            ],
            "quizzes": [
                "How do you create a new document?"
            ]
        },
        {
            "day": 4,
            "objective": "Learn to insert and format images.",
            "tasks": [
                "Insert an image related to your hobby into the essay. Resize the image and apply a style (e.g., shadow, reflection). Wrap text around the image using the 'Tight' option."
            ],
            "quizzes": [
                "What is the shortcut to save a document?"
            ]
        },
        {
            "day": 5,
            "objective": "Understand document formatting options.",
            "tasks": [
                "Change your document's orientation to 'Landscape', set custom margins (1-inch all around), and add page numbers at the bottom center."
            ],
            "quizzes": [
                "How can you adjust the margins in Word?"
            ]
        },
        {
            "day": 6,
            "objective": "Learn to create and format tables.",
            "tasks": [
                "Create a table listing your weekly schedule. Include columns for 'Day', 'Activity', and 'Time'. Apply a table style to enhance its appearance."
            ],
            "quizzes": [
                "How do you insert a table in Word?"
            ]
        },
        {
            "day": 7,
            "objective": "Compile and finalize a document.",
            "tasks": [
                "Compile all your work into a single document named 'Month_WordProject.docx'. Ensure consistent formatting and include a cover page."
            ],
            "quizzes": [
                "How can you wrap text around an image?"
            ]
        },
        {
            "day": 8,
            "objective": "Get started with PowerPoint presentations.",
            "tasks": [
                "Create a new presentation titled 'My Interests'. Add a title slide and two additional slides with titles 'Hobbies' and 'Goals'."
            ],
            "quizzes": [
                "How do you create a new slide?"
            ]
        },
        {
            "day": 9,
            "objective": "Enhance slides with bullet points and formatting.",
            "tasks": [
                "On the 'Hobbies' slide, list your hobbies using bullet points. Change the bullet style and text formatting for emphasis."
            ],
            "quizzes": [
                "How do you format bullet points?"
            ]
        },
        {
            "day": 10,
            "objective": "Learn to incorporate visuals and shapes.",
            "tasks": [
                "Add relevant images to each slide and draw a shape (e.g., arrow or star) to highlight key points. Apply effects to the shapes."
            ],
            "quizzes": [
                "What is a design theme?"
            ]
        },
        {
            "day": 11,
            "objective": "Use SmartArt and charts for visual representation.",
            "tasks": [
                "Create a SmartArt graphic illustrating your weekly routine. Insert a simple bar chart showing the time spent on each hobby."
            ],
            "quizzes": [
                "How do you insert an image?"
            ]
        },
        {
            "day": 12,
            "objective": "Add animations and transitions for effect.",
            "tasks": [
                "Apply transitions between slides and animations to text and images. Experiment with timing and effects."
            ],
            "quizzes": [
                "How do you add a transition between slides?"
            ]
        },
        {
            "day": 13,
            "objective": "Review and practice presentation skills.",
            "tasks": [
                "Review your presentation for consistency and practice delivering it. Save it as 'Month_PowerPointProject.pptx'."
            ],
            "quizzes": [
                "What is SmartArt used for?"
            ]
        },
        {
            "day": 14,
            "objective": "Incorporate multimedia into presentations.",
            "tasks": [
                "Add a relevant video clip to your presentation and set it to play on click. Insert background music to play across all slides."
            ],
            "quizzes": [
                "What is the shortcut for starting a slideshow?"
            ]
        },
        {
            "day": 15,
            "objective": "Create and format a personal budget in Excel.",
            "tasks": [
                "Create a new workbook for a 'Personal Budget'. Input headers for 'Item', 'Cost', and 'Category'."
            ],
            "quizzes": [
                "What is Excel’s default number format?"
            ]
        },
        {
            "day": 16,
            "objective": "Enter and format data in Excel.",
            "tasks": [
                "Enter at least 10 expenses into your budget and format currency cells to display as dollars. Apply cell borders and shading to headers."
            ],
            "quizzes": [
                "How do you adjust column width?"
            ]
        },
        {
            "day": 17,
            "objective": "Use formulas for calculations.",
            "tasks": [
                "Calculate the total cost of your expenses using a formula. Use subtraction to find the remaining balance if you have a $500 budget."
            ],
            "quizzes": [
                "How do you perform addition in Excel?"
            ]
        },
        {
            "day": 18,
            "objective": "Utilize Excel functions for data analysis.",
            "tasks": [
                "Use the SUM function to total expenses, AVERAGE to find average cost, MIN to find the smallest expense, and MAX for the largest."
            ],
            "quizzes": [
                "What function would you use to find the average of a list of numbers?"
            ]
        },
        {
            "day": 19,
            "objective": "Sort and filter data effectively.",
            "tasks": [
                "Sort your expenses alphabetically and then by cost. Apply a filter to display only expenses above $50."
            ],
            "quizzes": [
                "How do you sort data in Excel?"
            ]
        },
        {
            "day": 20,
            "objective": "Create visual representations of data.",
            "tasks": [
                "Create a column chart representing your expenses by category. Customize the chart's title and legend."
            ],
            "quizzes": [
                "How do you create a chart in Excel?"
            ]
        },
        {
            "day": 21,
            "objective": "Track study hours and visualize progress.",
            "tasks": [
                "Create a 'Weekly Study Tracker' spreadsheet logging hours studied each day. Include totals and create a line chart to visualize your progress."
            ],
            "quizzes": [
                "What is conditional formatting?"
            ]
        },
        {
            "day": 22,
            "objective": "Write a structured report in Word.",
            "tasks": [
                "Write a report on a topic of your choice with at least three sections. Apply 'Heading 1' and 'Heading 2' styles appropriately and insert an automatic table of contents."
            ],
            "quizzes": [
                "How do you apply styles in Word?"
            ]
        },
        {
            "day": 23,
            "objective": "Use conditional formatting and data validation in Excel.",
            "tasks": [
                "In your budget spreadsheet, apply conditional formatting to highlight expenses over $100 in red. Use data validation to restrict the 'Cost' column to accept only numerical values."
            ],
            "quizzes": [
                "What is a pivot table?"
            ]
        },
        {
            "day": 24,
            "objective": "Learn about Mail Merge in Word.",
            "tasks": [
                "Create a newsletter template and use Mail Merge to insert different recipient names and addresses from an Excel file."
            ],
            "quizzes": [
                "What is mail merge used for?"
            ]
        },
        {
            "day": 25,
            "objective": "Customize presentations using Slide Master.",
            "tasks": [
                "Modify the Slide Master to include a footer with your name and date on all slides. Change the default font for headings and body text."
            ],
            "quizzes": [
                "How do you customize a slide master?"
            ]
        },
        {
            "day": 26,
            "objective": "Analyze data using Pivot Tables in Excel.",
            "tasks": [
                "Use a provided data set to create a Pivot Table that summarizes total sales by region."
            ],
            "quizzes": [
                "How do you create a Pivot Table?"
            ]
        },
        {
            "day": 27,
            "objective": "Practice giving presentations using different delivery techniques.",
            "tasks": [
                "Deliver your PowerPoint presentation to a friend or family member, focusing on eye contact, clear speaking, and engaging your audience."
            ],
            "quizzes": [
                "What is the importance of rehearsing a presentation?"
            ]
        },
        {
            "day": 28,
            "objective": "Prepare a portfolio of your work in Word.",
            "tasks": [
                "Compile selected works (essays, presentations, budgets) into a portfolio document. Include a table of contents and introduction."
            ],
            "quizzes": [
                "How can you create a table of contents in Word?"
            ]
        },
        {
            "day": 29,
            "objective": "Final review of skills learned throughout the month.",
            "tasks": [
                "Review all your documents, presentations, and spreadsheets. Make improvements where necessary and prepare for a final showcase."
            ],
            "quizzes": [
                "What have you learned this month?"
            ]
        },
        {
            "day": 30,
            "objective": "Present your portfolio and reflect on your learning.",
            "tasks": [
                "Present your portfolio to a peer or instructor, explaining each component and what you learned. Collect feedback."
            ],
            "quizzes": [
                "How can feedback improve your skills?"
            ]
        }
    ]
}

class LearningApp:
    def __init__(self, root):
        self.root = root
        self.root.title("30-Day Learning Program")
        self.root.geometry("900x350")
        self.root.configure(bg="black")

        # Variables
        self.theme_var = tk.StringVar(value="Light")
        self.notification_time = tk.StringVar(value="09:00 AM")
        self.current_day = self.load_current_day()  # Track the current day

        # Load settings from file
        self.load_settings()

        # Menu bar
        self.menu_bar = tk.Menu(root)
        root.config(menu=self.menu_bar)

        # Add settings and analytics to menu
        settings_menu = tk.Menu(self.menu_bar, tearoff=0)
        settings_menu.add_command(label="Settings", command=self.open_settings)
        self.menu_bar.add_cascade(label="Menu", menu=settings_menu)

        analytics_menu = tk.Menu(self.menu_bar, tearoff=0)
        analytics_menu.add_command(label="Show Analytics", command=self.display_analytics)
        self.menu_bar.add_cascade(label="Analytics", menu=analytics_menu)

        # Add start button
        start_button = tk.Button(root, text="Start Learning Program", command=self.start_learning_program, bg="blue", fg="white")
        start_button.pack(pady=20)

        # Add frame for tasks
        self.task_frame = tk.Frame(root, bg="black")
        self.task_frame.pack(pady=10)

        # Start the notification in a separate thread
        threading.Thread(target=self.notify, daemon=True).start()

    def load_settings(self):
        if os.path.exists("settings.txt"):
            with open("settings.txt", "r") as f:
                lines = f.readlines()
                if lines:
                    self.theme_var.set(lines[0].strip())  # Theme
                    self.notification_time.set(lines[1].strip())  # Notification time

    def save_settings(self):
        with open("settings.txt", "w") as f:
            f.write(f"{self.theme_var.get()}\n{self.notification_time.get()}")

    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("300x300")
        settings_window.configure(bg="black")

        # Theme selection
        theme_label = tk.Label(settings_window, text="Select Theme:", bg="black", fg="lightblue")
        theme_label.pack(pady=5)

        for theme in ["Light", "Dark"]:
            rb = tk.Radiobutton(settings_window, text=theme, variable=self.theme_var, value=theme, command=self.change_theme, bg="black", fg="lightblue")
            rb.pack(anchor=tk.W)

        # Notification time input
        notification_label = tk.Label(settings_window, text="Set Notification Time (HH:MM AM/PM):", bg="black", fg="lightblue")
        notification_label.pack(pady=5)

        notification_entry = tk.Entry(settings_window, textvariable=self.notification_time, bg="lightblue", fg="black")
        notification_entry.pack(pady=5)

        reset_button = tk.Button(settings_window, text="Reset Program", command=self.reset_program, bg="red", fg="white")
        reset_button.pack(pady=20)
        
        save_button = tk.Button(settings_window, text="Save", command=lambda: self.save_and_close(settings_window), bg="blue", fg="white")
        save_button.pack(pady=10)

    def change_theme(self):
        theme = self.theme_var.get()
        if theme == "Light":
            self.root.configure(bg="lightgray")
            self.task_frame.configure(bg="lightgray")
        else:
            self.root.configure(bg="black")
            self.task_frame.configure(bg="black")

    def save_and_close(self, window):
        self.save_settings()
        window.destroy()
        self.change_theme()  # Change theme when settings are saved


    def reset_program(self):
        # Reset the contents of the required files
        with open("current_day.txt", "w") as f:
            f.write("1")  # Reset day to 0

        with open("day_completed.txt", "w") as f:
            f.write("")  # Clear day completed file

        with open("on.txt", "w") as f:
            f.write("no")  # Set on.txt to "no"

        with open("settings.txt", "w") as f:
            f.write("")  # Clear settings.txt

        with open("analytics.txt", "w") as f:
            f.write("")  # Clear analytics.txt

        # Set the title back to the initial title
        self.root.title("30 Days Learning MS Office")

        # Optional message to inform the user
        messagebox.showinfo("Reset", "Program has been reset!")

    def start_learning_program(self):
            # Load current day
        with open("current_day.txt", "r") as f:
            current_day = int(f.read().strip())

        if current_day == 0:
            # Program hasn't started yet
            self.root.title("30 Days Learning MS Office")
        else:
            # Program started, update the title with the current day
            self.root.title(f"Learning MS Office - Day {current_day}")
        self.current_day = self.load_current_day()  # Reload the current day
        self.show_day_data(self.current_day)

    def show_day_data(self, day):
         # Update title to reflect the current day
        self.root.title(f"Learning MS Office - Day {day}")

        if day > len(month_data["days"]):
            messagebox.showinfo("Program Completed", "Congratulations! You've completed the learning program.")
            return

        day_data = month_data["days"][day - 1]
        self.clear_tasks()

        # Check if objective_label exists, otherwise create it
        if not hasattr(self, 'objective_label'):
            self.objective_label = tk.Label(self.root, bg="black", fg="lightblue")
            self.objective_label.pack(pady=10)
            self.objective_label.config(text=f"Objective: {day_data['objective']}")

        self.task_vars = []
        for task in day_data['tasks']:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(self.task_frame, text=task, variable=var, bg="black", fg="lightblue")
            cb.pack(anchor=tk.W)
            self.task_vars.append(var)

        # Check for quiz question and ensure label exists
        if day_data['quizzes']:
            quiz_question = day_data['quizzes'][0]
            if quiz_question:
                # Check if quiz_label exists, otherwise create it
                if not hasattr(self, 'quiz_label') or not self.quiz_label.winfo_exists():
                    self.quiz_label = tk.Label(self.task_frame, bg="black", fg="lightblue")
                    self.quiz_label.pack(pady=10)

                # Update quiz label with new text
                self.quiz_label.config(text=f"Quiz: {quiz_question}")

                self.quiz_vars = []
                for answer in ["Correct", "Incorrect"]:
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.task_frame, text=answer, variable=var, bg="black", fg="lightblue")
                    cb.pack(anchor=tk.CENTER)
                    self.quiz_vars.append(var)
            else:
                print("Quiz question is empty!")

        if not hasattr(self, 'submit_button'):
            self.submit_button = tk.Button(self.root, text="Submit Tasks", command=lambda: self.submit_tasks(day_data, day), bg="blue", fg="white")
            self.submit_button.pack(pady=10)
        else:
            self.submit_button.config(command=lambda: self.submit_tasks(day_data, day))

    def clear_tasks(self):
        for widget in self.task_frame.winfo_children():
            widget.destroy()

    def submit_tasks(self, day_data, day):
        completed_today = []
        for i, cb in enumerate(self.task_vars):  # Change this line to loop through task_vars
            if cb.get():  # Get the associated variable from the BooleanVar
                completed_today.append(day_data['tasks'][i])

        if completed_today:
            messagebox.showinfo("Tasks Completed", f"Completed Tasks:\n" + "\n".join(completed_today))
            self.update_completion_file(day)
            self.show_next_day(day_data)
        else:
            messagebox.showwarning("No Tasks Completed", "You must complete at least one task.")

    def update_completion_file(self, day):
        with open("day_completed.txt", "a") as f:
            f.write(f"{day},")  # Append completed day

    def show_next_day(self, day_data):
        day_number = month_data["days"].index(day_data) + 1
        if day_number < len(month_data["days"]):
            self.clear_tasks()
            self.current_day = day_number + 1  # Move to the next day
            self.show_day_data(self.current_day)
            self.save_current_day(self.current_day)  # Save the current day
        else:
            messagebox.showinfo("Program Completed", "Congratulations! You've completed the learning program.")

    def save_current_day(self, day):
        with open("current_day.txt", "w") as f:
            f.write(str(day))  # Save the current day to file

    def load_current_day(self):
        if os.path.exists("current_day.txt"):
            with open("current_day.txt", "r") as f:
                return int(f.read().strip())
        return 1  # Default to day 1 if no file exists

    
    def notify(self):
        while True:
            now = time.strftime("%I:%M %p")  # Get the current time in HH:MM AM/PM format
            if now == self.notification_time.get():
                notification.notify(
                    title="Learning Reminder",
                    message="It's time to complete your tasks!",
                    app_name="LearningApp"
                )
            time.sleep(60)  # Check every minute

    def display_analytics(self):
        if os.path.exists("analytics.txt"):
            with open("analytics.txt", "r") as f:
                analytics_data = f.read()
            messagebox.showinfo("Learning Progress Analytics", analytics_data)
        else:
            messagebox.showinfo("Analytics", "No analytics data available.")

if __name__ == "__main__":
    root = tk.Tk()
    app = LearningApp(root)
    root.mainloop()

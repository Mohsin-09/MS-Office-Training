# 30-Day MS Office Learning Program

This project is a **desktop learning assistant** built with **Python (Tkinter)** that helps beginners learn Microsoft Office applications (Word, PowerPoint, and Excel) through a 30‑day interactive program.  
It guides users with **daily objectives, tasks, and quizzes**, and keeps track of progress using simple text files.

## Features
- **30-Day Curriculum**: Each day contains objectives, practical tasks, and quizzes to reinforce learning.
- **Progress Tracking**: Uses `.txt` files to store the current day, completed days, analytics, and settings.
- **Daily Notifications**: Sends desktop reminders using the `plyer` library to stay on schedule.
- **Theme Support**: Switch between Light and Dark themes from the Settings menu.
- **Analytics View**: Review learning progress and completed tasks.

## Project Structure
```
MSOfficeLearningProgram/
├─ main.py             # Main Tkinter application
├─ current_day.txt     # Stores the current learning day
├─ day_completed.txt   # Tracks all completed days
├─ on.txt              # Indicates whether the program is active
├─ settings.txt        # Stores theme and notification settings
├─ analytics.txt       # Stores progress analytics
└─ Syllabus of msoffice (semester).docx   # (Optional) Detailed syllabus
```

## Installation
1. Clone the repository or download the files:
   ```bash
   git clone <repo-url>
   cd MSOfficeLearningProgram
   ```

2. Install dependencies:
   ```bash
   pip install tkinter plyer schedule
   ```
   *(Tkinter comes pre-installed with most Python distributions.)*

3. Run the application:
   ```bash
   python main.py
   ```

## Usage
- Click **Start Learning Program** to begin the 30-day journey.
- Each day, tasks and quizzes will appear in the app.
- Use checkboxes to mark completed tasks and quizzes.
- Daily notifications remind you to complete your work (time configurable in **Settings**).
- Reset the program anytime via **Menu → Settings → Reset Program**.

## Data Storage
No database is required. The program uses plain text files:
- **current_day.txt** – Current learning day (e.g., `5`)
- **day_completed.txt** – Comma-separated list of completed days
- **on.txt** – Tracks whether the program is active
- **settings.txt** – Theme preference and notification time
- **analytics.txt** – Summary of progress and performance

## Dependencies
- **Python 3**
- **Tkinter** – GUI library
- **Plyer** – For system notifications
- **Schedule** – (Optional) for scheduling tasks

## License
This project is released under the MIT License.

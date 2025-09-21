# MS Office Teaching Assistant

MS Office Teaching Assistant is a **desktop application** built with **Python (Tkinter)** to help instructors create and manage quizzes, tasks, and to-dos for teaching MS Office (Word, Excel, PowerPoint, etc.) to juniors.  
It provides a structured way to assign daily tasks, track progress, and analyze performance.

## Features
- **Quizzes & Tasks**: Create and manage quizzes or tasks to help students learn MS Office.
- **Daily Progress Tracking**: Track current day tasks and completed days.
- **Analytics**: View performance statistics of students and completion rates.
- **Settings Management**: Configure the application according to your teaching needs.
- **Text File Storage**: All data is saved as simple `.txt` files instead of using a database.

## Project Structure
```
MSOfficeTeachingAssistant/
├─ main.py             # Main application entry point using Tkinter
├─ database.py         # Handles reading/writing to text files
├─ analytics.py        # Generates and displays progress analytics
├─ settings.py         # Application configuration and preferences
├─ current_day.py      # Logic for managing daily quizzes and tasks
├─ day_completed.py    # Tracks completed days and progress
└─ data/               # Folder containing text files (data store)
   ├─ quizzes.txt
   ├─ tasks.txt
   ├─ progress.txt
   └─ settings.txt
```

## Installation
1. Clone the repository:
   ```bash
   git clone <repo-url>
   cd MSOfficeTeachingAssistant
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   *Typical dependencies: `tkinter`, `pandas` (for analytics)*

3. Run the application:
   ```bash
   python main.py
   ```

## Usage
- Launch the app and use the **Settings** menu to configure course details.
- Add new **quizzes**, **tasks**, or **to-dos** for each day.
- Monitor student progress and review **analytics** for performance insights.
- Track completed days to ensure consistent learning progress.

## Data Storage
Instead of a database, the application stores all information in **plain text files** under the `data/` folder:
- `quizzes.txt` – Quiz questions and answers
- `tasks.txt` – Daily tasks and to-dos
- `progress.txt` – Completed days and user progress
- `settings.txt` – Application configuration

These files can be easily edited or backed up manually.

## Technologies Used
- **Python 3**
- **Tkinter** for the GUI
- **Text files** for lightweight storage
- **Pandas/Matplotlib** (optional) for analytics

## License
This project is licensed under the MIT License.

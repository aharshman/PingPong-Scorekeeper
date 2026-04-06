# 🏓 Ping Pong Scorekeeper

A Python app for tracking ping pong games with a clean scoreboard interface and automatic match history stored in Excel.

---

## 🚀 Features

- Live score tracking
- Fault system (2 faults = opponent point)
- Custom game length
- Win-by-2 rule enforcement
- Automatic match logging
- Simple, easy-to-use interface

---

## 📊 Match Data (Excel)

All matches are automatically saved to an Excel file:

**`ping_pong_data.xlsx`**

Each game records:
- Date
- Player names
- Final score
- Winner
- Game length
- Fault counts for each player

The file is created automatically the first time you run the app, and updates after every completed game.

---

## 🖥️ Tech Used

- Python
- Tkinter (GUI)
- OpenPyXL (Excel handling)
- Pandas & Matplotlib (stats viewer)

---

## ▶️ How to Run (Windows)

1. Download this repository as a ZIP file and extract it.

2. Open the folder in Command Prompt.

3. Install required packages:


pip install -r requirements.txt


4. Run the app:


python PingPong.py


---

## 📌 Notes

- Default game is played to 15
- Games must be won by 2 points
- Match data is saved automatically after each game

---

## 💡 Future Improvements

- Player statistics dashboard
- Match filtering
- Undo last action
- UI improvements

---

## 📬 Author

Built as a personal project to practice building real-world Python applications and trackin

Built as a personal project to improve Python and real-world application development.

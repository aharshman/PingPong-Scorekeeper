# 🏓 Ping Pong Scorekeeper

A Python-based ping pong scoreboard and match tracker with a clean UI, fault tracking, and automatic Excel logging.

## 🚀 Features

- Live score tracking
- Fault system (2 faults = opponent point)
- Custom game length
- Win-by-2 rule enforcement
- Match history saved to Excel
- Simple and clean GUI (Tkinter)

## 📊 Stats Tracking

Each match records:
- Date
- Players
- Final score
- Winner
- Game length
- Fault counts

## 🖥️ Tech Used

- Python
- Tkinter (GUI)
- OpenPyXL (Excel logging)
- Pandas & Matplotlib (for stats viewer)

## ▶️ How to Run

1. Clone the repo:

git clone https://github.com/YOUR_USERNAME/PingPong-Scorekeeper.git

cd PingPong-Scorekeeper


2. Create virtual environment:

python3 -m venv venv
source venv/bin/activate


3. Install dependencies:

pip install -r requirements.txt


4. Run the app:

python3 PingPong.py


## 📌 Notes

- Default game is played to 15
- Must win by 2 points
- Data is stored in `ping_pong_data.xlsx`

## 💡 Future Improvements

- Player statistics dashboard
- Match filtering
- Undo button
- Enhanced UI/UX

---

## 📬 Author

Built as a personal project to improve Python and real-world application development.

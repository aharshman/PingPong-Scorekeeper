import pandas as pd
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import ttk
from tkinter import messagebox

# --- Load Excel Data ---
data_file = "PingPongData.xlsx"

try:
    df = pd.read_excel(data_file, sheet_name="Matches", parse_dates=['Date'])
except FileNotFoundError:
    messagebox.showerror("Error", f"Could not find {data_file}")
    exit()

# --- Functions ---
def show_player_stats():
    player = player_var.get()
    if not player:
        messagebox.showwarning("Warning", "Select a player")
        return

    # Filter matches
    player_matches = df[(df['Player 1'] == player) | (df['Player 2'] == player)]

    # Clear previous table results
    for row in tree.get_children():
        tree.delete(row)

    total_wins = 0
    total_faults = 0
    total_points = 0  # Score column format: "X-Y"

    # Record vs other players
    record_vs = {}

    for _, match in player_matches.iterrows():
        # Determine points for this player
        score_text = match['Score']
        try:
            p1_score, p2_score = map(int, score_text.split('-'))
        except:
            p1_score = p2_score = 0

        points = p1_score if match['Player 1'] == player else p2_score
        total_points += points

        # Faults
        faults = match['Faults P1'] if match['Player 1']==player else match['Faults P2']
        total_faults += faults

        # Wins/Losses
        winner = match['Winner']
        if winner == player:
            total_wins += 1
            tag = "win"
        else:
            tag = "loss"

        # Record vs opponent
        opponent = match['Player 2'] if match['Player 1']==player else match['Player 1']
        if opponent not in record_vs:
            record_vs[opponent] = {'wins':0,'losses':0}
        if winner == player:
            record_vs[opponent]['wins'] += 1
        else:
            record_vs[opponent]['losses'] += 1

        # Insert row into table
        tree.insert("", END, values=(
            match['Date'].strftime("%Y-%m-%d"),
            match['Player 1'],
            match['Player 2'],
            match['Score'],
            match['Winner'],
            faults
        ), tags=(tag,))

    games_played = len(player_matches)
    points_per_game = total_points / games_played if games_played else 0
    faults_per_game = total_faults / games_played if games_played else 0

    # --- Update summary ---
    summary_label.config(
        text=f"{player} → Wins: {total_wins}, Total Faults: {total_faults}, Total Points: {total_points}"
    )

    per_game_label.config(
        text=f"Points/Game: {points_per_game:.2f}, Faults/Game: {faults_per_game:.2f}"
    )

    # Update record vs other players
    record_text = "\n".join([f"{op}: {v['wins']}W-{v['losses']}L" for op,v in record_vs.items()])
    record_label.config(text=f"Record vs Others:\n{record_text}")

def plot_top_faults():
    fault_totals = {}
    for player in set(df['Player 1']).union(set(df['Player 2'])):
        p1_faults = df[df['Player 1']==player]['Faults P1'].sum()
        p2_faults = df[df['Player 2']==player]['Faults P2'].sum()
        fault_totals[player] = p1_faults + p2_faults

    names = list(fault_totals.keys())
    faults = list(fault_totals.values())

    plt.figure(figsize=(8,5))
    plt.bar(names, faults, color='#4e0000')
    plt.title("Total Faults by Player")
    plt.ylabel("Faults")
    plt.xlabel("Player")
    plt.show()

# --- GUI ---
root = Tk()
root.title("Ping Pong Stats Viewer")
root.geometry("950x650")

Label(root, text="🏓 Ping Pong Stats Viewer 🏓", font=("Courier", 24, "bold")).pack(pady=10)

# Player selection
player_var = StringVar()
Label(root, text="Select Player:", font=("Courier", 14)).pack(pady=5)
player_menu = ttk.Combobox(root, textvariable=player_var, font=("Courier", 12))
player_menu['values'] = list(set(df['Player 1']).union(set(df['Player 2'])))
player_menu.pack(pady=5)

Button(root, text="Show Stats", font=("Courier", 12, "bold"), command=show_player_stats).pack(pady=5)
Button(root, text="Plot Top Faults", font=("Courier", 12, "bold"), command=plot_top_faults).pack(pady=5)

# Summary labels
summary_label = Label(root, text="", font=("Courier", 14, "bold"), fg="blue")
summary_label.pack(pady=5)

per_game_label = Label(root, text="", font=("Courier", 14, "bold"), fg="purple")
per_game_label.pack(pady=2)

record_label = Label(root, text="", font=("Courier", 12), justify=LEFT)
record_label.pack(pady=5)

# Table for matches
columns = ("Date", "Player 1", "Player 2", "Score", "Winner", "Faults")
tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, anchor=CENTER, width=130)
tree.tag_configure('win', background='#d0ffd0')   # light green for wins
tree.tag_configure('loss', background='#ffd0d0')  # light red for losses
tree.pack(expand=True, fill=BOTH, pady=10, padx=10)

root.mainloop()
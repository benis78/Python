import os
import subprocess

# Stier til MT4-terminal for hver bruger
mt4_paths = [
    r"C:\Program Files (x86)\TradeMax Global MT4 Terminal\terminal.exe" for i in range(1, 11)
]

def start_mt4_terminals():
    """
    Start MT4-terminalen for hver bruger.
    """
    for path in mt4_paths:
        if os.path.exists(path):
            try:
                subprocess.Popen(path, shell=True)
                print(f"MT4 terminal startet for sti: {path}")
            except Exception as e:
                print(f"Kunne ikke starte MT4 terminal for sti '{path}': {e}")
        else:
            print(f"Stien '{path}' eksisterer ikke.")

if __name__ == "__main__":
    start_mt4_terminals()

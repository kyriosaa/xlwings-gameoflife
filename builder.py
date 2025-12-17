import xlwings
import os 

PATTERNS_DIR = "patterns"

# scans folder and returns .txt filenames in the patterns folder
def list_patterns():
    if not os.path.exists(PATTERNS_DIR):
        print(f"[ERROR]   No {PATTERNS_DIR} found. Creating...")
        os.makedirs(PATTERNS_DIR)
        print(f"[STATUS]  Created {PATTERNS_DIR} folder. Add your patterns there.")
        return []
    
    files = [f for f in os.listdir(PATTERNS_DIR) if f.endswith(".txt")]
    return files

# reads txt file and converts to coords
def load_pattern(filename):
    path = os.path.join(PATTERNS_DIR, filename)
    coords = []
    
    with open(path, 'r') as f:
        lines = f.readlines()
        
    for r, line in enumerate(lines):
        for c, char in enumerate(line):
            if char == 'O':
                coords.append((r, c))
    return coords

def main():
    GRID_SIZE = 50
    
    try:
        wb = xlwings.Book('game.xlsx')
        sheet = wb.sheets[0]
    except FileNotFoundError:
        print("[ERROR]   Pls open the game.xlsx file first.")
        return
    
    patterns = list_patterns()
    
    print("--- Available Patterns ---")
    for i, p in enumerate(patterns):
        print(f"{i+1}: {p}")
        
    try:
        selection = int(input("Choose a pattern: ")) - 1
        if selection < 0 or selection >= len(patterns):
            raise ValueError
    except ValueError:
        print("[ERROR]   Invalid number.")
        return
    
    selected_file = patterns[selection]
    print(f"[STATUS]  Loading {selected_file}...")
    
    start_r = 5
    start_c = 5
    
    sheet.range((1,1), (GRID_SIZE,GRID_SIZE)).clear_contents()
    
    coords = load_pattern(selected_file)
    for r, c in coords:
        sheet.cells(start_r + r, start_c + c).value = 1
    
    print(f"[SUCCESS] {selected_file} successfully loaded.")
    
if __name__ == "__main__":
    main()
    
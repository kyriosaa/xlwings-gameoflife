# the engine takes a look at the board state before it runs the game
# so draw your design or import one to the excel file before running the engine

import xlwings
import numpy
import time
from ui_config import UI_MAP, FORMULA_CONFIG

# calc number of "alive" neighbors for every cell
# every cell has 8 neighbors, instead of visiting a specific cell and looking at its neighbors (x+1, y), (x-1, y)
# we can just shift the entire grid so the neighbor moves into the position of the current cell
def get_neighbors(grid):
    neighbors = (
        numpy.roll(grid, 1, axis=0) 
        + numpy.roll(grid, -1, axis=0) 
        + numpy.roll(grid, 1, axis=1) 
        + numpy.roll(grid, -1, axis=1) 
        + numpy.roll(numpy.roll(grid, 1, axis=0), 1, axis=1)
        + numpy.roll(numpy.roll(grid, 1, axis=0), -1, axis=1)
        + numpy.roll(numpy.roll(grid, -1, axis=0), 1, axis=1)
        + numpy.roll(numpy.roll(grid, -1, axis=0), -1, axis=1)
    )
    return neighbors

# conway game of life ruleset
def update(grid):
    neighbors = get_neighbors(grid)
    # DEATH - if a cell is alive (1) and has either fewer than 2 neighbors or more than 3 neighbors, set the cell to 0
    new_grid = numpy.where((grid == 1) & ((neighbors < 2) | (neighbors > 3)), 0, grid)
    # BIRTH - if a cell is dead (0) and has exactly 3 neighbors, set the cell to 1
    new_grid = numpy.where((grid == 0) & (neighbors == 3), 1, new_grid)
    return new_grid

def clear(grid):
    new_grid = numpy.where(grid, 0, grid)
    return new_grid

def draw_ui(sheet):
    if not UI_MAP:
        return
    
    for color_rgb, cell_list in UI_MAP.items():
        for cell_address in cell_list:
            try:
                current_color = sheet.range(cell_address).color
                
                if current_color != color_rgb:
                    sheet.range(cell_address).color = color_rgb
            except Exception as error:
                print(f"[ERROR]   Could not apply UI to {cell_address}: {error}")
                return False
    return True
            
def draw_formula(sheet):
    if not FORMULA_CONFIG:
        return
    for element in FORMULA_CONFIG:
        try:
            cell_range = sheet.range(element['range'])
            
            if 'h_align' in element:
                cell_range.api.HorizontalAlignment = element['h_align']
            if 'v_align' in element:
                cell_range.api.VerticalAlignment = element['v_align']
            if 'font_name' in element:
                cell_range.font.name = element['font_name']
            if 'font_size' in element:
                cell_range.font.size = element['font_size']
            if 'formula' in element:
                cell_range.formula = element['formula']
            if element.get('merge', False):
                cell_range.merge()
        except Exception as error:
            print(f"[ERROR]   Could not apply formula to {element.get('range', 'unknown')}: {error}")
            return False
    return True
            
def main():
    try:
        wb = xlwings.Book('game.xlsx')
        sheet = wb.sheets[0]
    except FileNotFoundError:
        print("[ERROR]   Pls open the game.xlsx file first.")
        return
    
    GRID_SIZE = 50 # 50x50
    DELAY = 0.05
    
    # read the board state from excel
    # yea make sure u draw your figure or import a design first
    print("[STATUS]  Reading current board state...")
    raw_data = sheet.range((2,2), (GRID_SIZE + 1, GRID_SIZE + 1)).value

    # cleaning
    grid = numpy.array(raw_data)
    grid[grid == None] = 0
    grid = grid.astype(int)
    grid[grid != 1] = 0
    
    print("[STATUS]  Drawing UI...")
    if not draw_ui(sheet):
        print("[ERROR]   Failed to draw UI. Exiting...")
        return
    if not draw_formula(sheet):
        print("[ERROR]   Failed to draw UI. Exiting...")
        return
    print("[SUCCESS] UI successfully drawn.")
    print("[STATUS]  Engine running. Press Ctrl+C to stop.")
    
    try:
        while True:
            sheet.range('B2').value = grid
            grid = update(grid)
            time.sleep(DELAY)
    except KeyboardInterrupt:
        grid = clear(grid)
        sheet.range('B2').value = grid
        print("[STATUS]  Engine stopped.")

if __name__ == "__main__":
    main()
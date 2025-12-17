# the engine takes a look at the board state before it runs the game
# so draw your design or import one to the excel file before running the engine

import xlwings
import numpy
import time

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

def main():
    try:
        wb = xlwings.Book('game.xlsx')
        sheet = wb.sheets[0]
    except FileNotFoundError:
        print("[ERROR] pls open the game.xlsx file first")
        return
    
    GRID_SIZE = 50 # 50x50
    DELAY = 0.1
    
    # read the board state from excel
    # yea make sure u draw your figure or import a design first
    print("Reading current board state")
    raw_data = sheet.range((1,1), (GRID_SIZE, GRID_SIZE)).value

    # cleaning
    grid = numpy.array(raw_data)
    grid[grid == None] = 0
    grid = grid.astype(int)
    grid[grid != 1] = 0
    
    print("Engine running")
    print("--- Press Ctrl+C to stop ---")
    
    try:
        while True:
            sheet.range('A1').value = grid
            grid = update(grid)
            time.sleep(DELAY)
    except KeyboardInterrupt:
        grid = clear(grid)
        sheet.range('A1').value = grid
        print("Engine stopped")

if __name__ == "__main__":
    main()
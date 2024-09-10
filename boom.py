# import random

# def create_board(rows, cols, num_mines):
#     # 게임판 생성
#     board = [[' ' for _ in range(cols)] for _ in range(rows)]

#     # 무작위로 지뢰 배치
#     for _ in range(num_mines):
#         row = random.randint(0, rows - 1)
#         col = random.randint(0, cols - 1)
#         while board[row][col] == '*':
#             row = random.randint(0, rows - 1)
#             col = random.randint(0, cols - 1)
#         board[row][col] = '*'

#     return board

# def print_board(board, revealed):
#     for i in range(len(board)):
#         for j in range(len(board[0])):
#             if revealed[i][j]:
#                 print(board[i][j], end=' ')
#             else:
#                 print(' ', end=' ')
#         print()

# def count_adjacent_mines(board, row, col):
#     count = 0
#     for r in range(max(0, row - 1), min(row + 2, len(board))):
#         for c in range(max(0, col - 1), min(col + 2, len(board[0]))):
#             if (r != row or c != col) and board[r][c] == '*':
#                 count += 1
#     return count

# def reveal_empty_cells(board, row, col, revealed):
#     if row < 0 or row >= len(board) or col < 0 or col >= len(board[0]) or revealed[row][col]:
#         return
#     revealed[row][col] = True
#     if board[row][col] == ' ':
#         for r in range(row - 1, row + 2):
#             for c in range(col - 1, col + 2):
#                 reveal_empty_cells(board, r, c, revealed)

# def main():
#     rows = 5
#     cols = 5
#     num_mines = 5
#     revealed = [[False for _ in range(cols)] for _ in range(rows)]

#     board = create_board(rows, cols, num_mines)
#     print_board(board, revealed)

#     while True:
#         try:
#             row = int(input("행을 선택하세요 (0부터 {}까지): ".format(rows - 1)))
#             col = int(input("열을 선택하세요 (0부터 {}까지): ".format(cols - 1)))
#             if board[row][col] == '*':
#                 print("지뢰를 밟았습니다! 게임 종료!")
#                 break
#             else:
#                 reveal_empty_cells(board, row, col, revealed)
#                 print_board(board,revealed)
#         except (ValueError, IndexError):
#             print("잘못된 입력입니다.")

# if __name__ == "__main__":
#     main()

def count_adjacent_mines(grid, row, col):
    # 주변 지뢰의 수를 세는 함수
    count = 0
    directions = [(1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (1, -1), (-1, 1), (-1, -1)]
    for dr, dc in directions:
        nr, nc = row + dr, col + dc
        if 0 <= nr < len(grid) and 0 <= nc < len(grid[0]) and grid[nr][nc] == '*':
            count += 1
    return count

def calculate_adjacent_mines(grid):
    # 각 칸의 주변 지뢰의 수를 계산하는 함수
    rows = len(grid)
    cols = len(grid[0])
    result = [[0 for _ in range(cols)] for _ in range(rows)]

    for row in range(rows):
        for col in range(cols):
            if grid[row][col] != '*':
                result[row][col] = count_adjacent_mines(grid, row, col)

    return result

def print_grid_with_mines(grid):
    for row in grid:
        print(row)

def print_result(result):
    for row in result:
        print(row)

def main():
    # 입력 grid
    grid = [
        ['-','*','-','-','-'],
        ['-','-','-','*','-'],
        ['-','*','-','-','-'],
        ['-','-','-','-','-']
    ]

    # 각 칸의 주변 지뢰의 수 계산
    result = calculate_adjacent_mines(grid)

    # 결과 출력
    print("입력 격자판:")
    print_grid_with_mines(grid)
    print("\n각 칸의 주변 지뢰의 수:")
    print_result(result)

if __name__ == "__main__":
    main()
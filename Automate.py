# import openpyxl as xl
# import math


# # makes array
# def array(wb,sheet,cell):
#     a = []
#     for column in range(1, sheet.max_column + 1):
#         if cell.value == 'zzzbreak':
#             break
#         for row in range(1, sheet.max_row + 1):
#             cell = sheet.cell(row, column)
#             a.append(cell.value)
#             if cell.value == 'zzzbreak':
#                 break
#
#     return a
#
#
# def add(a):
#     flag = True
#     while flag:
#         p = input("Which anime would you like to add: ").strip()
#         if p == "close" or p == "" or p == "no":
#             flag = False
#         else:
#             a.append(p)
#             print("Anime has been added")
#
#
# def delete(a):
#     flag = True
#     while flag:
#         p = input("Which anime would you like to remove: ").strip()
#         if p == "close" or p == "" or p == "no":
#             flag = False
#         else:
#             try:
#                 a.remove(p)
#                 print("Anime has been removed")
#             except ValueError:
#                 print("Anime is not in list")
#
# '''
# def arrange(wb,sheet,cell,a):
#
#     a.sort()
#     i = 0
#     for column in range(1, 7):
#         if i == len(a):
#             break
#         for row in range(1, sheet.max_row + 1):
#             cell = sheet.cell(row, column)
#             cell.value = a[i]
#             print(cell.value)
#             i += 1
#             if i == len(a):
#                 break
# '''
#
#
# def arrange(ARRAY,sheet):
#     i = 0
#     end = len(ARRAY)
#
#     for column in range(1,15):
#         if i == end:
#             break;
#         for row in range(1,8):
#             if i == end:
#                 break;
#             cell = sheet.cell(row, column)
#             cell.value = ARRAY[i]
#             i+= 1
#
#
# ARRAY = []
# for i in range(92):
#     ARRAY.append(1)
#
#
# wb = xl.load_workbook('List of Animes.xlsx')
# sheet = wb['Ratings']
# cell = sheet['a1']
#
#
# a = (array(wb,sheet,cell))
#
#
# f = True
# while f:
#     m = input("Enter 1 to add or 2 to delete: ")
#     if m == "close" or m == "" or m == "no":
#         f = False
#     elif m == "1":
#         add(a)
#     elif m == "2":
#         delete(a)
#
#
# arrange(ARRAY,sheet)
# wb.save('List of Animes.xlsx')
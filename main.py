# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

#
# def print_hi(name):
#     # Use a breakpoint in the code line below to debug your script.
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
import os
from window import Window

path = os.path.dirname(os.path.abspath(__file__)) + '\dbase'

if not os.path.exists(path):
    os.makedirs(path)
window = Window(
    width=800,
    height=600,
    title="Обученность"
)

window.draw_menu()
window.run()
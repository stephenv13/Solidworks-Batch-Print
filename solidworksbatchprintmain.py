'''
Solidworks Batch Print by Stephen Vogelmeier

This script does the following:

1.      Ask user for text file name

2.      Read text file to pull PN

3.      Search windows explorer for PN

4.      Open the file, check to see if file is open/closed, then move to next part number

5.      Print or call macro to print

6.      Close file

7.      Repeat for every PN in text file

'''

import os
import re
import subprocess
import win32gui

#This sets the given users file location to the default desktop location.
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']),'Desktop')

sldwrks_location = 'C:\\"Program Files"\\"SOLIDWORKS Corp"\\SOLIDWORKS\\SLDWORKS.exe'
print_macro_location = 'C:\\Users\\Stephen\\Desktop\\Solidworks\\B3PrintAllSheets11x17.swp'
close_macro_location = 'C:\\Users\\Stephen\\Desktop\\Solidworks\\CloseOpenDocument.swp'

open_windows = []
open_parts = []

'''
This function prints welcome info for the list and asks the user to input their text file name with part numbers.
'''
def welcome() -> str:
    print('-----------------------------------')
    print('Welcome to Solidworks Batch Print!')
    print('-----------------------------------\n')

    print('To begin, please place a text file with all part numbers on your desktop!\n')
    #file_name = str(input('Please enter the text file name ("XXXXX.txt"): \n'))
    file_name = 'partnumbers2.txt'
    return file_name

'''
This function opens the given text files, utilizes regex to strip part numbers from the text, and store them in a list.
'''
def open_file(file_name: str) -> list:
    file = open(f'{desktop}\{file_name}','r')

    part_numbers = re.findall(r'A\d\d\d\d\d\d\d',file.read())
    
    return part_numbers

'''
This function takes in the name of a file and searches the given file path.
'''
def find(name: str, path: str) -> str:
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)

'''
This function checks the Windows UI for all open windows and stores them in a list.
'''
def winEnumHandler(hwnd, ctx) -> None:
    if win32gui.IsWindowVisible(hwnd):
        open_windows.append(win32gui.GetWindowText(hwnd))


'''
This function checks the Windows UI to see which parts have an 'open window' already. It will add these to a list using a regex call.
It will then take the list of open windows and strip them down to a more usable "part number" name.
'''
def check_windows(open_parts: list) -> list:
    stripped_parts = []
    win32gui.EnumWindows(winEnumHandler,None)
    for window in open_windows:
        if re.search(r'A\d\d\d\d\d\d\d - Sheet\d', window) != None:
            open_parts.append(window)

    for part in open_parts:
        stripped_parts.append(part[0:8])
    return stripped_parts

'''
This function opens the called part number.
'''
def open_part(part_number: str) -> None:
    file_path = find(f'{part_number}.SLDDRW',desktop)
    print(f'Opening: {file_path}')

    doc = subprocess.Popen(["start", file_path],shell=True)
    

'''
This function will take in a called part number and the list of already opened parts, It will check to see if that part number
is already opened. If it is, it will exit. If it is not, it will wait until the part number is open.
'''
def check_open(part_number: str, open_parts: list) -> None:
    open = False
    
    while not open:
        open_parts = check_windows(open_parts)

        if part_number in open_parts:
            print(f'{part_number} Opened!')
            open = True
            break

# Main Body
if __name__ == '__main__':

    file_name = welcome()
    part_numbers = open_file(file_name)

    '''
    Loop through all part numbers in the file, 
    '''
    for part_number in part_numbers:
        
        open_part(part_number)
        check_open(part_number, open_parts)

    
    """
    Call the solidworks print macro utilizing the SOLIDWORKS API and macro file location to print each of the part drawings
    """
    subprocess.call(print_macro_location)
    os.popen(f' \m {print_macro_location}')
    subprocess.Popen(['C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe', print_macro_location],shell=True)
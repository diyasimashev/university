import pandas as pd
import openpyxl


def student_greeting():
    print('\nWelcome, dear student!')


def student_menu():
    x = int(input("""\nIn order to get the information you are interested in, enter the number of the corresponding menu: \n1 - Attendance\n2 - Exams\n3 - Grades\n4 - Schedule\n5 - Subjects\n6 - Tasks\nIf you want to terminate the program, enter 7: """))
    if x == 1:
        x1 = int(input("""\nTo get information about your attendance, enter the subject number: \n1 - Algorithms and Data Structure\n2 - Human-Computer Interaction\n3 - Mathematics\n4 - Programming Languages\n5 - The English Language\n6 - The German Language: """))
        if 1 <= x1 <= 2 or 4 <= x1 <= 6:
            print("\nThe teacher hasn't marked anyone yet.")
            comeback()
        elif x1 == 3:
            z = pd.read_excel('teacher.xlsx', sheet_name='sheet3')
            x1 = z[z["Student"] == 'Жапарбекова Айэлина']
            print('\nAT - attended, AB - was absent')
            print('\n', x1)
            comeback()
        else:
            print("\nThe data you entered was not found. Try again.")
            comeback()
    elif x == 2:
        z = pd.read_excel('director.xlsx', sheet_name='sheet2')
        x1 = int(input("""\nTo get information about the exams, enter the subject number: \n1 - Algorithms and Data Structure\n2 - Human-Computer Interaction\n3 - Mathematics\n4 - Programming Languages\n5 - The English Language\n6 - The German Language: """))
        if x1 == 1:
            x1 = z[z["Subject"] == 'Algorithms and Data Structures']
            print('\n', x1)
            comeback()
        elif x1 == 2:
            x1 = z[z["Subject"] == 'Human-Computer Interaction']
            print('\n', x1)
            comeback()
        elif x1 == 3:
            x1 = z[z["Subject"] == 'Mathematics']
            print('\n', x1)
            comeback()
        elif x1 == 4:
            x1 = z[z["Subject"] == 'Programming Languages']
            print('\n', x1)
            comeback()
        elif x1 == 5:
            x1 = z[z["Subject"] == 'The English Language']
            print('\n', x1)
            comeback()
        elif x1 == 6:
            x1 = z[z["Subject"] == 'The German Language']
            print('\n', x1)
            comeback()
        else:
            print("\nThe data you entered was not found. Try again.")
            comeback()
    elif x == 3:
        z = pd.read_excel('teacher.xlsx', sheet_name='sheet1')
        x1 = int(input("""\nTo get information about your grades, enter the subject number: \n1 - Algorithms and Data Structure\n2 - Human-Computer Interaction\n3 - Mathematics\n4 - Programming Languages\n5 - The English Language\n6 - The German Language: """))
        if 1 <= x1 <= 2 or 4 <= x1 <= 6:
            print("\nThe teacher hasn't graded anyone yet")
            comeback()
        elif x1 == 3:
            x1 = z[z['Student'] == 'Жапарбекова Айэлина']
            print('\n', x1)
            comeback()
        else:
            print("\nThe data you entered was not found. Try again.")
            comeback()
    elif x == 4:
        x1 = int(input("""\nTo get information about the schedule, enter the day number: \n1 - Monday\n2 - Tuesday\n3 - Wednesday\n4 - Thursday\n5 - Friday\n6 - Saturday: """))
        if x1 == 1:
            z = pd.read_excel('director.xlsx', sheet_name='sheet3', usecols=['Start time of lessons', 'Monday'])
            print('\n', z[z.Monday.notna()])
            comeback()
        elif x1 == 2:
            z = pd.read_excel('director.xlsx', sheet_name='sheet3', usecols=['Start time of lessons', 'Tuesday'])
            print('\n', z[z.Tuesday.notna()])
            comeback()
        elif x1 == 3:
            z = pd.read_excel('director.xlsx', sheet_name='sheet3', usecols=['Start time of lessons', 'Wednesday'])
            print('\n', z[z.Wednesday.notna()])
            comeback()
        elif x1 == 4:
            z = pd.read_excel('director.xlsx', sheet_name='sheet3', usecols=['Start time of lessons', 'Thursday'])
            print('\n', z[z.Thursday.notna()])
            comeback()
        elif x1 == 5:
            print("\nYou don't have any lessons on this day, lucky!")
            comeback()
        elif x1 == 6:
            z = pd.read_excel('director.xlsx', sheet_name='sheet3', usecols=['Start time of lessons', 'Saturday'])
            print('\n', z[z.Saturday.notna()])
            comeback()
        else:
            print("\nThe data you entered was not found. Try again.")
            comeback()
    elif x == 5:
        z = openpyxl.open('director.xlsx', read_only=True)
        z.active = 0
        x1 = z.active
        for row in range(1, x1.max_row + 1):
            x2 = x1[row][0].value
            print(x2)
        comeback()
    elif x == 6:
        z = pd.read_excel('teacher.xlsx', sheet_name='sheet2')
        x1 = int(input("""\nTo get information about the homework, enter the subject number: \n1 - Algorithms and Data Structure\n2 - Human-Computer Interaction\n3 - Mathematics\n4 - Programming Languages\n5 - The English Language\n6 - The German Language: """))
        if 1 <= x1 <= 2 or 4 <= x1 <= 6:
            print("\nThere are no assignments yet.")
            comeback()
        elif x1 == 3:
            x2 = z[z["Group"] == 'WIN-1-21']
            print('\n', x2)
            comeback()
        else:
            print("\nThe data you entered was not found. Try again.")
            comeback()
    elif x == 7:
        for i in range(1):
            print("\nThe work of the program is completed. We are waiting for your next appearance.")
            break
    else:
        student_menu()


def comeback():
    x = input('To return to the menu, enter "c": ')
    if x == 'c':
        for i in range(1):
            student_menu()
            break
    else:
        for i in range(1):
            comeback()

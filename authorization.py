import directior_page
import student_page
import teacher_page


def authorization():
    x = input('\nEnter your status (s - student/ t - teacher/d - director)? ')
    if x == 's':
        x1 = input("Key word: ")
        if x1 == '111':
            student_page.student_greeting()
            student_page.student_menu()
        else:
            print('The data you entered was not found. Try again.')
            authorization()
    elif x == 't':
        x1 = input("Key word: ")
        if x1 == '222':
            teacher_page.teacher_greeting()
            teacher_page.teacher_menu()
        else:
            print('The data you entered was not found. Try again.')
            authorization()
    elif x == 'd':
        x1 = input("Key word: ")
        if x1 == '333':
            directior_page.director_greeting()
            directior_page.director_menu()
        else:
            print('The data you entered was not found. Try again.')
            authorization()
    else:
        print("Sorry, but we didn't find this type of account. Try again.")
        authorization()

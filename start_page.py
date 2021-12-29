import authorization

print("Welcome to INAI!")


def instruction():
    a = open('instruction.txt', 'r')
    data = a.read()
    print(data)
    a.close()
    instruction1()


def instruction1():
    b = input('To start the program, enter "s": ')
    if b == 's':
        authorization.authorization()
    else:
        print('The data you entered is invalid. Try again.')
        instruction1()


instruction()

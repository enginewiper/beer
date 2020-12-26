#build menuing system

helpMsg = '''
    h - print this message
    q - quit
'''
while True:
    menuChoice = input('Enter option: ')
    if menuChoice == 'h':
        print(helpMsg)
    elif menuChoice == 'q':
        break
    else:
        print("I don't know that! (h for help)")
import random

def TalPlus():
    return random.randint(10,75)

def TalMinus():
    x=random.randint(10,50)
    y=random.randint(0,10)
    return x,y

def TalGange():
    return random.randint(0,15)

def TalDivider():
    x=random.randint(1,15)
    y=random.randint(1,15)
    return x,y

def TalGennemsnit():
    while True:
        x = random.randint(1,15)
        y = random.randint(1,15)
        z = random.randint(1,15)
        tjekLigeTal=(x+y+z)%3
        if tjekLigeTal == 0:
            return x, y, z


def int_input(prompt):
    while True:
        try:
            svar = int(input(prompt))
            if svar == 1234:
                quit()
            return svar
        except ValueError:
            print("Kun tal må bruges, Prøv igen eller skriv 1234 for afslut!")


def Korrekt():
    print('Du har svaret rigtigt {} gange, næste'.format(antalOpgaver))

def KorrektEnd():
    print('Du har svaret rigtigt på alle opgaverne. Godt klaret vi ses igen i morgen :)')


def Plus():
    for svar in range(antalOpgaver):
        x=TalPlus()
        y=TalPlus()
        resultat = x+y
        svar = int_input('Hvad er {}+{} = '.format(x,y))
        if svar == resultat:
            print('Rigtigt :)')
        while not svar == resultat:
            svar = int_input('Prøv igen :( {}+{} = '.format(x,y))
            if svar == resultat:
                print('Rigtigt :)')
    Korrekt()


def Minus():
    for svar in range(antalOpgaver):
        x,y = TalMinus()
        resultat = x-y
        svar = int_input('Hvad er {}-{} = '.format(x,y))
        if svar == resultat:
            print('Rigtigt :)')
        while not svar == resultat:
            svar = int_input('Prøv igen :( hvad er {}-{} = '.format(x,y))
            if svar == resultat:
                print('Rigtigt :)')
    Korrekt()


def Gange():
    for svar in range(antalOpgaver):
        x=TalGange()
        y=TalGange()
        resultat = x*y
        svar = int_input('Hvad er {}x{} = '.format(x,y))
        if svar == resultat:
            print('Rigtigt :)')
        while not svar == resultat:
            svar = int_input('Prøv igen :( hvad er {}x{} = '.format(x,y))
            if svar == resultat:
                print('Rigtigt :)')
    Korrekt()

def Divider():
    for svar in range(antalOpgaver):
        x,y = TalDivider()
        resultat = x*y
        svar = int_input('Hvad er {}/{} = '.format(resultat,y))
        if svar == x:
            print('Rigtigt :)')
        while not svar == x:
            svar = int_input('Prøv igen :( hvad er {}/{} = '.format(resultat,y))
            if svar == resultat:
                print('Rigtigt :)')
    Korrekt()

def Gennemsnit():
    for svar in range(antalOpgaver):
        x,y,z=TalGennemsnit()
        resultat = (x+y+z)/3
        svar = int_input('Hvad er gennemsnittet af {},{},{} = '.format(x,y,z))
        if svar == resultat:
            print('Rigtigt :)')
        while not svar == resultat:
            svar = int_input('Prøv igen :( hvad er gannemsnittet af {},{},{} = '.format(x,y,z))
            if svar == resultat:
                print('Rigtigt :)')
    KorrektEnd()
        
 

antalOpgaver=5

# #antalOpgaver = int(input('Antal plus opgaver'))
Plus()
# #antalOpgaver = int(input('Antal minus opgaver'))
Minus()
# #antalOpgaver = int(input('Antal gange opgaver'))
Gange()

Divider()

Gennemsnit()

exit()
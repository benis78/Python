import numpy as np
from scipy.optimize import fsolve
import openpyxl

# Funktion til at beregne radius
def equation(r, m, n, o, A, B, C):
    term_m = m * np.arccos(1 - (A**2) / (2 * r**2)) if m > 0 else 0
    term_n = n * np.arccos(1 - (B**2) / (2 * r**2)) if n > 0 else 0
    term_o = o * np.arccos(1 - (C**2) / (2 * r**2)) if o > 0 else 0
    return term_m + term_n + term_o - 2 * np.pi

# Initialisering af Excel Workbook og Sheet
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Radius Results"

# Tilføjer overskrifter til Excel-arket
ws.append(["m", "n", "o", "Radius"])

# Sidelængder
A = 1050 # Spændeelement
B = 2110 # Elementlængde
C = 6  # Fuge

# Iterer over m og n værdier
for m in range(0, 7):  # m spændeelement fra 0 til 6
    for n in range(5, 51):  # n element fra 5 til 50
        o = m + n   # o fuge beregnet som m + n 
        # Antag en startværdi for r
        initial_guess = 2000
        # Løs ligningen for at finde r
        r_solution = fsolve(equation, initial_guess, args=(m, n, o, A, B, C))
        # Tilføj resultater til Excel-arket
        ws.append([m, n, o, float(r_solution[0])])

# Gem workbook til en Excel-fil
wb.save("Radius_Results.xlsx")

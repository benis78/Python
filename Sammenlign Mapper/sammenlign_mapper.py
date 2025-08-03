import os
from tkinter import Tk, filedialog

def vælg_mappe(prompt):
    root = Tk()
    root.withdraw()
    print(prompt)
    valgt_mappe = filedialog.askdirectory(title=prompt)
    return valgt_mappe

def sammenlign_og_find_unikke_pdf(mappe1, mappe2):
    filer1 = {
        f: os.path.getsize(os.path.join(mappe1, f))
        for f in os.listdir(mappe1)
        if os.path.isfile(os.path.join(mappe1, f)) and f.lower().endswith(".pdf")
    }
    filer2 = {
        f: os.path.getsize(os.path.join(mappe2, f))
        for f in os.listdir(mappe2)
        if os.path.isfile(os.path.join(mappe2, f)) and f.lower().endswith(".pdf")
    }

    unikke_pdf_i_mappe1 = [
        f for f in filer1
        if f not in filer2 or filer1[f] != filer2[f]
    ]
    return unikke_pdf_i_mappe1

def slet_pdf_filer(filer, fra_mappe):
    for fil in filer:
        sti = os.path.join(fra_mappe, fil)
        try:
            os.remove(sti)
            print(f"Slettet PDF: {fil}")
        except Exception as e:
            print(f"Kunne ikke slette {fil}: {e}")

def main():
    mappe1 = vælg_mappe("Vælg den første mappe (PDF'er herfra kan blive slettet)")
    mappe2 = vælg_mappe("Vælg den anden mappe (bruges til sammenligning)")

    unikke_pdf = sammenlign_og_find_unikke_pdf(mappe1, mappe2)
    slet_pdf_filer(unikke_pdf, mappe1)

    print("\nFærdig!")

if __name__ == "__main__":
    main()

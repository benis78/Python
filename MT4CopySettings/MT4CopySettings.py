import os
import shutil
import tkinter as tk
from tkinter import filedialog

# Stier til bruger mapper
paths = {
    "Template": [
        fr"C:\Users\U{i}\AppData\Roaming\MetaQuotes\Terminal\AB75DD8A03E8CC693E1336EB0D50BA2D\templates" for i in range(1, 11)
    ],
    "Indicator": [
        fr"C:\Users\U{i}\AppData\Roaming\MetaQuotes\Terminal\AB75DD8A03E8CC693E1336EB0D50BA2D\MQL4\Indicators" for i in range(1, 11)
    ],
    "Expert": [
        fr"C:\Users\U{i}\AppData\Roaming\MetaQuotes\Terminal\AB75DD8A03E8CC693E1336EB0D50BA2D\MQL4\Experts" for i in range(1, 11)
    ],
    "Preset": [
        fr"C:\Users\U{i}\AppData\Roaming\MetaQuotes\Terminal\AB75DD8A03E8CC693E1336EB0D50BA2D\MQL4\Presets" for i in range(1, 11)
    ]
}

def copy_to_user_folders(source_paths, user_paths):
    """
    Kopier en eller flere filer eller mapper til de forskellige brugermapper.
    :param source_paths: Liste over stier til de filer eller mapper, der skal kopieres.
    :param user_paths: Liste over brugermapper, hvor filerne eller mapperne skal kopieres til.
    """
    for source_path in source_paths:
        if not os.path.exists(source_path):
            print(f"Kilden '{source_path}' eksisterer ikke.")
            continue

        for user_path in user_paths:
            if os.path.exists(user_path):
                destination = os.path.join(user_path, os.path.basename(source_path))
                try:
                    if os.path.isfile(source_path):
                        shutil.copy2(source_path, destination)
                        print(f"Fil '{source_path}' kopieret til '{destination}'")
                    elif os.path.isdir(source_path):
                        if os.path.exists(destination):
                            shutil.rmtree(destination)
                        shutil.copytree(source_path, destination)
                        print(f"Mappe '{source_path}' kopieret til '{destination}'")
                    else:
                        print(f"Kilden '{source_path}' er hverken en fil eller en mappe.")
                except Exception as e:
                    print(f"Kunne ikke kopiere til '{user_path}': {e}")
            else:
                print(f"Brugerens sti '{user_path}' eksisterer ikke.")

def select_source():
    """
    Åbn en filudvælger for at vælge en eller flere filer eller mapper til kopiering.
    """
    root = tk.Tk()
    root.title("Vælg type af filer")

    def set_type_and_copy(file_type):
        root.withdraw()
        source_paths = filedialog.askopenfilenames() or [filedialog.askdirectory()]
        source_paths = [path for path in source_paths if path]  # Fjern tomme stier
        if source_paths:
            copy_to_user_folders(source_paths, paths[file_type])
        else:
            print("Ingen filer eller mapper valgt.")

    # Tilføj knapper til valg af type
    template_button = tk.Button(root, text="Template", command=lambda: set_type_and_copy("Template"))
    template_button.pack(pady=5)

    indicator_button = tk.Button(root, text="Indicator", command=lambda: set_type_and_copy("Indicator"))
    indicator_button.pack(pady=5)

    expert_button = tk.Button(root, text="Expert", command=lambda: set_type_and_copy("Expert"))
    expert_button.pack(pady=5)

    preset_button = tk.Button(root, text="Preset", command=lambda: set_type_and_copy("Preset"))
    preset_button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    select_source()

import subprocess
import re
import time
from pyautogui import click, moveTo, size

def check_internet_connection(url='http://www.google.com', timeout=5):
    """
    Tjekker, om der er internetforbindelse ved at forsøge at åbne en given URL.
    """
    import urllib.request
    while True:
        try:
            urllib.request.urlopen(url, timeout=timeout)
            return True
        except urllib.request.URLError:
            print("Ingen internetforbindelse. Prøver igen...")
            time.sleep(5)

def check_adapter_connection(adapter_identifier="Ethernet adapter Ethernet 3:"):
    """
    Tjekker, om den angivne netværksadapter har en gyldig IPv4-adresse.
    Hvis der findes en adresse, antages forbindelsen at være aktiv.
    """
    try:
        output = subprocess.check_output("ipconfig", shell=True, stderr=subprocess.STDOUT).decode('utf-8', errors="ignore")
    except subprocess.CalledProcessError as e:
        print("Fejl ved kørsel af ipconfig:", e)
        return False

    # Søg efter adaptersektionen baseret på det angivne identifier
    pattern = re.compile(r"({0}.*?)(?=\n\S|\Z)".format(re.escape(adapter_identifier)), re.DOTALL)
    match = pattern.search(output)
    if not match:
        print(f"Adapter '{adapter_identifier}' blev ikke fundet i ipconfig-outputtet.")
        return False

    adapter_info = match.group(1)
    # Find en linje med IPv4-adressen (fx "IPv4 Address. . . . . . . . . . . : 10.240.240.1(Preferred)")
    ip_regex = re.compile(r"IPv4[ .]*[A-Za-z]*[ .]*:[ ]*([\d\.]+)")
    ip_match = ip_regex.search(adapter_info)
    
    if ip_match:
        ip_address = ip_match.group(1)
        print(f"Adapter '{adapter_identifier}' har IP: {ip_address}")
        return True
    else:
        print(f"Adapter '{adapter_identifier}' har ingen gyldig IPv4-adresse.")
        return False

def connect_vpn():
    """
    Udfører pyautogui-sekvensen for at forsøge at oprette VPN-forbindelsen.
    Denne version udregner koordinaterne relativt til skærmstørrelsen.
    
    Her antages det, at de oprindelige koordinater (4915, 1415)
    gælder for en skærm med basisopløsning base_width x base_height.
    """
    # Definer basisopløsningen (tilpas efter din opsætning)
    base_width = 5120
    base_height = 1440

    # Hent den aktuelle skærmstørrelse
    current_width, current_height = size()

    # Udregn forholdet mellem de oprindelige koordinater og basisopløsningen
    x_ratio = 4915 / base_width  # For eksempel: 4915/5120 ≈ 0.959
    y_ratio = 1415 / base_height  # For eksempel: 1415/1440 ≈ 0.984

    # Udregn de nye koordinater baseret på den aktuelle skærmstørrelse
    icon_x = current_width * x_ratio
    icon_y = current_height * y_ratio

    print(f"Flytter musen til ({icon_x:.0f}, {icon_y:.0f}) ud fra skærmstørrelsen ({current_width}, {current_height})")

    # Udfør musehandlinger med de udregnede koordinater
    click()  # Sørg for, at vinduet har fokus
    time.sleep(1)
    moveTo(icon_x, icon_y, duration=1)
    time.sleep(1)
    click()  # Klik på VPN-ikonet
    # Klik på "Opret forbindelse" – her kan du også udregne koordinaterne relativt,
    # hvis du har et forholdstal for denne placering. For nu bruges en relativ flytning:
    # Eksempelvis antages at "opret forbindelse"-knappen ligger lidt til venstre for ikonet.
    # Du kan justere fordelingen, fx ved at trække et fast antal pixels fra:
    connection_button_x = icon_x - (current_width * 0.005)  # Juster forholdstallet efter behov
    connection_button_y = icon_y  # Eller tilpas hvis knappen er placeret højere/lavere
    click(connection_button_x, connection_button_y, duration=1.0)
    # Flyt musen væk, fx til midten af skærmen
    moveTo(current_width / 2, current_height / 2)

# --------------------- HOVEDPROGRAM ---------------------

# Tjek først for internetforbindelse
if check_internet_connection():
    print("Der er internetforbindelse.")
else:
    print("Ingen internetforbindelse. Afbryder...")
    exit(1)

# Forsøg at etablere forbindelse via den angivne netværksadapter op til 3 gange.
max_attempts = 3
attempt = 1
adapter_identifier = "Ethernet adapter Ethernet 3:"  # Tilpas hvis nødvendigt

while attempt <= max_attempts:
    if check_adapter_connection(adapter_identifier):
        print("Netværksadapter 3 har forbindelse.")
        break
    else:
        print(f"Forsøg {attempt}: Netværksadapter 3 har ikke forbindelse. Forsøger at oprette VPN-forbindelsen...")
        connect_vpn()
        # Giv Windows tid til at oprette forbindelsen (tilpas ventetid om nødvendigt)
        time.sleep(10)
        attempt += 1

# Hvis forbindelsen stadig ikke er oprettet efter 3 forsøg:
if not check_adapter_connection(adapter_identifier):
    print("Kunne ikke oprette forbindelse via netværksadapter 3 efter 3 forsøg.")
    user_choice = input("Tryk 'c' for at prøve igen manuelt, eller en hvilken som helst anden tast for at afslutte: ")
    if user_choice.lower() == 'c':
        print("Vent venligst, prøv at oprette forbindelsen manuelt eller gennem yderligere kommandoer.")
    else:
        print("Afslutter programmet.")
        exit(1)

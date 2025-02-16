import os
import socket
import struct
import time
import platform

def connect_vpn(vpn_name: str):
    """
    Opretter VPN-forbindelse på Windows eller Linux.
    :param vpn_name: VPN-navn i Windows' OpenVPN GUI.
    """
    print("Forbinder til VPN...")
    if platform.system() == "Windows":
        os.system("taskkill /IM openvpn-gui.exe /F")  # Luk eksisterende VPN GUI
        os.system(f'start /B "C:\\Program Files\\OpenVPN\\bin\\openvpn-gui.exe"')
        time.sleep(2)  # Vent på GUI-opstart
        os.system(f'"C:\\Program Files\\OpenVPN\\bin\\openvpn-gui.exe" --command connect {vpn_name}')
    else:
        os.system(f"sudo openvpn --config {vpn_name} --daemon")  # Linux OpenVPN
    time.sleep(10)  # Vent på VPN-forbindelsen
    print("VPN-forbindelse etableret!")

def disconnect_vpn(vpn_name: str):
    """
    Afbryder VPN-forbindelsen på Windows eller Linux.
    """
    print("Afbryder VPN...")
    if platform.system() == "Windows":
        os.system(f'"C:\\Program Files\\OpenVPN\\bin\\openvpn-gui.exe" --command disconnect {vpn_name}')
        time.sleep(2)
        os.system("taskkill /IM openvpn-gui.exe /F")  # Luk OpenVPN GUI
    else:
        os.system("sudo pkill openvpn")  # Linux OpenVPN
    print("VPN afbrudt!")

def send_wol(mac_address: str, target_ip: str, port: int = 9):
    """
    Sender en Wake-on-LAN magic packet til en given MAC-adresse via unicast.
    :param mac_address: MAC-adressen på enheden, der skal vækkes (format: "AA:BB:CC:DD:EE:FF").
    :param target_ip: IP-adressen på den enhed, der skal vækkes.
    :param port: Portnummer til WOL (default: 9, kan også være 7).
    """
    try:
        # Fjern eventuelle separatorer og konverter til bytes
        mac_bytes = bytes.fromhex(mac_address.replace("-", "").replace(":", ""))
        if len(mac_bytes) != 6:
            raise ValueError("Ugyldig MAC-adresse format. Brug f.eks. 'AA:BB:CC:DD:EE:FF'")
    except ValueError as e:
        print(f"Fejl: {e}")
        return
    
    # Skab en magic packet: 6x FF efterfulgt af MAC-adressen 16 gange
    magic_packet = b'\xff' * 6 + mac_bytes * 16
    
    try:
        # Opret en UDP socket
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.sendto(magic_packet, (target_ip, port))
        print(f"Magic Packet sendt til {mac_address} via unicast til {target_ip}:{port}")
    except Exception as e:
        print(f"Kunne ikke sende Magic Packet: {e}")

# Eksempel på brug
if __name__ == "__main__":
    vpn_name = "er"  # Erstat med navnet på din OpenVPN-profil (uden .ovpn)
    mac = "98:EE:CB:9D:0A:49"  # Erstat med den rigtige MAC-adresse
    target_ip = "192.168.2.107"  # Erstat med IP-adressen til den enhed, der skal vækkes
    
    connect_vpn(vpn_name)  # Forbind til VPN
    send_wol(mac, target_ip)
    disconnect_vpn(vpn_name)  # Afbryd VPN efter WOL

# # Eksempel på brug
# if __name__ == "__main__":
#     vpn_config = "C:\\Program Files\\OpenVPN\\config\\Home\\er.ovpn"  # Erstat med den rigtige sti
#     vpn_name = "er"  # Erstat med navnet på din OpenVPN-forbindelse
#     mac = "98:EE:CB:9D:0A:49"  # Erstat med den rigtige MAC-adresse
#     target_ip = "192.168.2.107"  # Erstat med IP-adressen til den enhed, der skal vækkes
    
#     connect_vpn(vpn_config, vpn_name)  # Forbind til VPN
#     send_wol(mac, target_ip)
#     disconnect_vpn(vpn_name)  # Afbryd VPN efter WOL
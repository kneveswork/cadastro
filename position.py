import pyautogui
import time

print("Posicione o cursor no local desejado e espere 5 segundos...")
time.sleep(5)  # Tempo para voce mover o mouse
print(pyautogui.position())  # Mostra as coordenadas x, y

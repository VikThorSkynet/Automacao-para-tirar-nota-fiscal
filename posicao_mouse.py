import pyautogui
import time

# Mensagem de início
print("Posicione o mouse no local desejado. Capturando a posição em 5 segundos...")

# Aguarda 5 segundos
time.sleep(5)

#pyautogui.scroll(-800)

#time.sleep(1)

#pyautogui.scroll(-500)

#time.sleep(1)

#pyautogui.scroll(-600)

time.sleep(5)

# Captura a posição atual do mouse
x, y = pyautogui.position()
print(f"Posição capturada do mouse: x={x}, y={y}")



# cpf = x=499, y=215
# campo_b = x=529, y=323
# campo_c = x=1281, y=300
# click_emitir = x=1137, y=495
# numero_nf = x=513, y=429 ate x=549, y=430

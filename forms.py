import pyautogui as pygui
from time import sleep
import webbrowser

webbrowser.open('https://docs.google.com/forms/u/0/')

pygui.click(1825, 35, duration=2)
sleep(2)
participantes = pygui.locateCenterOnScreen('participantesfesta.png', confidence=0.8)
pygui.moveTo(participantes)
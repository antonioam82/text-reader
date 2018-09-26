#LECTOR DE ARCHIVOS DE TEXTO
from VALID import ns
import win32com.client as wc

speak=wc.Dispatch("Sapi.SpVoice")

def conti():
    pre=ns(input("¿Desea continuar?: "))
    return pre

while True:
    texto=(input("Introduce el nombre del fichero a leer: ")+".txt")
    try:
        fichero=open(texto,"r")
    except:
        print("No se encontró el archivo",texto)
        contin=conti()
        if contin=="n":
            break
        else:
            continue

    print("REPRODUCIENDO TEXTO")
    for linea in fichero:
        if linea[-1]==('\n'):
            linea=linea[:-1]
        speak.Speak(linea)

    contin=conti()
    if contin=="n":
        break

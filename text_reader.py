#LECTOR DE ARCHIVOS DE TEXTO
import win32com.client as wc


def conti():
    pre=input("¿Desea continuar?: ")
    while pre!="n" and pre!="s":
        pre=input("Introduzca solo \'n\' o \'s\' según su opción: ")
    return pre
    

speak=wc.Dispatch("Sapi.SpVoice")

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

    print("REPRODUCIENDO TEXTO",texto)
    for linea in fichero:
        if linea[-1]==('\n'):
            linea=linea[:-1]
        speak.Speak(linea)

    print("LECTURA FINALIZADA")
    contin=conti()
    if contin=="n":
        break

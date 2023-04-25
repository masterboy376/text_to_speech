from win32com.client import Dispatch

if __name__ == '__main__':
    speak = Dispatch("SAPI.SpVoice")
    while True:
        st = input(">>>> ")
        speak.Speak(st)
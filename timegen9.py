#Imports necessarios para ter acesso aos eventos do windows, para manipular os aqruivos .csv, para administrar o tempo e ler sinais do sistema.
import sqlite3
import win32con
import sys
import ctypes
import ctypes.wintypes
import time
import signal


conn = sqlite3.connect('timegen.db')
try:
    conn.execute('''CREATE TABLE events
           (TIMESTAMP INTEGER NOT NULL,
            EVENT_TIME REAL NOT NULL,
            EVENT_TYPE TEXT NOT NULL,
            SHORT_NAME TEXT NOT NULL,
           WINDOW_TITLE TEXT)''')
    print ("Table created successfully");
except:
    pass


#Ctypes é uma biblioteca de funções externas para python. Permite chamar funções em arquivos .dll ou em bibliotecas compartilhadas.
user32 = ctypes.windll.user32
ole32 = ctypes.windll.ole32
kernel32 = ctypes.windll.kernel32




#A função WINFUNCTYPE cria definições de funções de callback usando a convenção stdcall
WinEventProcType = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LONG,
    ctypes.wintypes.LONG,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
)


#Define os tipos de eventos a serem capturados
eventTypes = {
    win32con.EVENT_SYSTEM_FOREGROUND: "Foreground",
    win32con.EVENT_SYSTEM_CAPTURESTART: "Click"
}


#Retorna informações a respeito da thread e do processo correntes
processFlag = getattr(win32con, 'PROCESS_QUERY_LIMITED_INFORMATION',
                      win32con.PROCESS_QUERY_INFORMATION)

threadFlag = getattr(win32con, 'THREAD_QUERY_LIMITED_INFORMATION',
                     win32con.THREAD_QUERY_INFORMATION)


#Armazena o timestamp do último evento para mostrar o tempo entre eventos
lastTime = 0


#[Temporário] Recebe o sinal de Ctrl + C para parar o programa e salvar o .csv
def signal_handler(signal, frame):
    print ("terminating")
    cur = conn.execute('select SHORT_NAME, count(EVENT_TYPE) from events group by 1')
    res = [dict(SHORT_NAME=row[0],
            EVENT_COUNT=row[1]) for row in cur.fetchall()]
    print(res)
    conn.close()
    sys.exit()


#Salva no arquivo de logs o timestamp do evento, o tempo gasto em cada evento,
#o tipo de evento e o nome de cada janela aberta
def log(tstamp,eventTime,eventType,windowShortName,windowTitle):
    conn.execute("INSERT INTO events (TIMESTAMP, EVENT_TIME, EVENT_TYPE, SHORT_NAME, WINDOW_TITLE) \
      VALUES (?,?,?,?,?)", [tstamp, eventTime, eventType, windowShortName, windowTitle])
    conn.commit()


#Imprime uma mensagem de erro    
def logError(msg):
    sys.stdout.write(msg + '\n')


#Retorna o ID do processo e define as mensagens de erro caso ocorram
def getProcessID(dwEventThread, hwnd):
    
    hThread = kernel32.OpenThread(threadFlag, 0, dwEventThread)

    if hThread:
        try:
            processID = kernel32.GetProcessIdOfThread(hThread)
            if not processID:
                logError("Couldn't get process for thread %s: %s" %
                         (hThread, ctypes.WinError()))
        finally:
            kernel32.CloseHandle(hThread)
    else:
        errors = ["No thread handle for %s: %s" %
                  (dwEventThread, ctypes.WinError(),)]

        if hwnd:
            processID = ctypes.wintypes.DWORD()
            threadID = user32.GetWindowThreadProcessId(
                hwnd, ctypes.byref(processID))
            if threadID != dwEventThread:
                logError("Window thread != event thread? %s != %s" %
                         (threadID, dwEventThread))
            if processID:
                processID = processID.value
            else:
                errors.append(
                    "GetWindowThreadProcessID(%s) didn't work either: %s" % (
                    hwnd, ctypes.WinError()))
                processID = None
        else:
            processID = None

        if not processID:
            for err in errors:
                logError(err)

    return processID


#Retorna o nome do arquivo do processo especificado
def getProcessFilename(processID):
    hProcess = kernel32.OpenProcess(processFlag, 0, processID)
    if not hProcess:
        logError("OpenProcess(%s) failed: %s" % (processID, ctypes.WinError()))
        return None

    try:
        filenameBufferSize = ctypes.wintypes.DWORD(4096)
        filename = ctypes.create_unicode_buffer(filenameBufferSize.value)
        kernel32.QueryFullProcessImageNameW(hProcess, 0, ctypes.byref(filename),
                                            ctypes.byref(filenameBufferSize))

        return filename.value
    finally:
        kernel32.CloseHandle(hProcess)


#Define o que ocorre quando um evento de troca de janela ou clique é capturado
def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread,
             dwmsEventTime):
    global lastTime
    length = user32.GetWindowTextLengthW(hwnd)
    title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title, length + 1)

    processID = getProcessID(dwEventThread, hwnd)

    shortName = '?'
    if processID:
        filename = getProcessFilename(processID)
        if filename:
            shortName = '\\'.join(filename.rsplit('\\', 2)[-2:])

    if hwnd:
        hwnd = hex(hwnd)
    elif idObject == win32con.OBJID_CURSOR:
        hwnd = '<Cursor>'

    log(dwmsEventTime, float(dwmsEventTime - lastTime)/1000, eventTypes.get(event, hex(event)),
        shortName, title.value)

    lastTime = dwmsEventTime


#Define os eventos a serem capturados
def setHook(WinEventProc, eventType):
    return user32.SetWinEventHook(
        eventType,
        eventType,
        0,
        WinEventProc,
        0,
        0,
        win32con.WINEVENT_OUTOFCONTEXT
    )


#Inicia o loop que captura os eventos
def main():
    ole32.CoInitialize(0)

    WinEventProc = WinEventProcType(callback)
    user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE

    hookIDs = [setHook(WinEventProc, et) for et in eventTypes.keys()]
    if not any(hookIDs):
        print ('SetWinEventHook failed')
        sys.exit(1)

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        user32.TranslateMessageW(msg)
        user32.DispatchMessageW(msg)

    for hookID in hookIDs:
        user32.UnhookWinEvent(hookID)
    ole32.CoUninitialize()


#Recebe o sinal de Ctrl + C ao final da execução 
signal.signal(signal.SIGINT, signal_handler)


#Chama a função main e escreve os headers do arquivo .csv
if __name__ == '__main__':
    #writer.writerow(["timestamp","event_time","event_type", "short_name","title"])
    main()

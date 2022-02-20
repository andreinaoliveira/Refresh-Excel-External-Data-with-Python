import os
import sys
import time
import win32com.client
from ftplib import FTP
from datetime import datetime

localPath = 'C:/caminho/nome-arquivo.xlsx'

# Update Excel
try:
    xlapp = win32com.client.DispatchEx('Excel.Application')
    wb = xlapp.Workbooks.Open(localPath)
    wb.RefreshAll()
    xlapp.CalculateUntilAsyncQueriesDone()
    time.sleep(5)
    xlapp.DisplayAlerts = False
    wb.Save()
    wb.Close()
    xlapp.quit()
except Exception:
    print('Erro ao atualizar o arquivo. Cheque se o arquivo local foi corrompido em ' + localPath)
    os.system("pause")
    sys.exit()


# FTP Connection
try:
    ftp = FTP('informar-servidor.com.br')
    ftp.login(user='informar-usuario', passwd='informar-senha')
    ftp.cwd('/FTPE8/FFA')
except Exception:
    print('Erro ao conectar ao FTP. Cheque o caminho.')
    os.system("pause")
    sys.exit()


# FTP Upload File
try:
    data_e_hora_atuais = datetime.now()
    data_e_hora_em_texto = data_e_hora_atuais.strftime('%Y-%m-%d-%H-%M')
    filename = "informar-pasta" + data_e_hora_em_texto + ".xlsx"
    ftp.storbinary('STOR '+filename, open(localPath, 'rb'))
    ftp.quit()
except Exception:
    print('Erro ao transferir arquivo da base local para o FTP.')
    os.system("pause")
    sys.exit()

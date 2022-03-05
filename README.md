# Refresh Excel External Data with Python
Esse script automaticamente atualiza os dados de um arquivo Excel existente na máquina local e salva uma cópia em um servidor FTP. Uso ideal para atualizar listas do SharePoint exportadas no formato .iqy <br>

![image](https://user-images.githubusercontent.com/51168329/154853113-5e82be4f-cda7-48cf-9f11-a27614f93f5b.png)
![image](https://user-images.githubusercontent.com/51168329/154853303-da56958e-ac10-4524-9098-878b8e75d180.png)

## Pré-requisito
Substituir os seguintes campos no código
- <b>localPath = ''</b>: informar o caminho do arquivo e o arquivo juntamente com a sua extensão.
- <b>ftp = FTP('')</b>: servidor FTP
- <b>ftp.login(user='', passwd='')</b>: usuário e senha do FTP
- <b>filename = ""</b>: pasta que será salvo o arquivo no FTP


## Como o script funciona
### Atualizando Excel
- localpath representa o caminho e o arquivo excel local que será atualziando. Entrando em Update Excel, será atualizado e salvo os dados do arquivo aquivo local. Se você busca apenas atualizar o arquivo sem salvá-lo em outra pasta, o cósigo abaixo já supre suas necessidades.
~~~ 
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
~~~

### Estabelecendo Conexão FTP
- O código abaixo estabelece uma conexão com o servicor FTP. Para isso é necessário informar o servidor, usuário e senha.
~~~ 
# FTP Connection
try:
    ftp = FTP('informar-servidor.com.br')
    ftp.login(user='informar-usuario', passwd='informar-senha')
    ftp.cwd('/FTPE8/FFA')
except Exception:
    print('Erro ao conectar ao FTP. Cheque o caminho.')
    os.system("pause")
    sys.exit()
~~~ 

### Salvando no FTP
- Com a conexão estabelecida o código abaixo irá criar uma cópia do arquivo atualizado e irá salvá-lo no caminho informado em "filename". Para não haver conflito entre nomes do arquivo, o mesmo será nomeado com a data e hora do momento em que o script está sendo rodado.

~~~~
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
~~~~

## Arquivo python (.py) em executável (.exe)

Para não haver a necessidade de executar o arquivo sempre que necessário é possível executar o código abaixo no terminal para gerar um executável no arquivo python.

~~~
pyinstaller --onefile .\script.py
~~~

Com o executável gerado você pode adicioná-lo na rotina do windows. Essa matéria pode ajudar <a href="Como agendar uma tarefa no Windows">Como agendar uma tarefa no Windows</a>.

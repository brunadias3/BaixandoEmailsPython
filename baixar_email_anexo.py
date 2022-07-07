import win32com.client as win32
from pathlib import Path

# Cria a pasta Destino
destino = Path.cwd() / "emails"
destino.mkdir(parents = True, exist_ok = True)

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

#ACESSANDO NOMES DE PASTAS
root_folders = outlook.Folders.Item(1)

for folder in root_folders.Folders:
    print(folder.Name)

# Caixa de entrada
inbox =  outlook.GetDefaultFolder(6)

# Pasta Específica
inbox = root_folders.Folders["teste"]

messages = inbox.items

for m in messages:
    subject = m.Subject
    body = m.body
    attachments = m.Attachments

    pasta_destino = destino / str(subject).replace(':','').replace('/','')
    pasta_destino.mkdir(parents=True,exist_ok=True)

    Path(pasta_destino / 'Corpo_email.txt').write_text(str(body))

    for att in attachments:
        att.SaveAsFile(pasta_destino / str(att))
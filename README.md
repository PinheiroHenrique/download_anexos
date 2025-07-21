# download_anexos
Download de anexos automatizados
# Outlook PDF Attachment Saver

Script Python para baixar automaticamente anexos PDF específicos de mensagens não lidas em uma subpasta do Outlook, evitando sobrescrever arquivos já existentes.

## Funcionalidades

- Conecta ao Outlook via COM Automation (win32com).
- Filtra mensagens não lidas em uma subpasta específica.
- Filtra anexos PDF que contenham palavras-chave definidas no nome do arquivo.
- Evita sobrescrever arquivos no destino, gerando nomes únicos automaticamente.
- Imprime status detalhado do processo no terminal.

## Pré-requisitos

- Windows (o script depende do Outlook instalado e do COM).
- Python 3.x instalado.
- Biblioteca `pywin32` instalada:
  
  ```bash
  pip install pywin32

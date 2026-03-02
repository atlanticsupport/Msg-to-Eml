import sys
import os
import subprocess
import extract_msg
from email.message import EmailMessage

def convert_msg_to_eml(msg_path, eml_path):
    """
    Usa o `extract_msg` para ler o ficheiro MSG e compõe de forma segura
    uma representação `.eml` na pasta destino.
    """
    msg = extract_msg.Message(msg_path)
    eml = EmailMessage()
    
    # Cabeçalhos base
    if msg.sender:
        eml['From'] = msg.sender
    if msg.to:
        eml['To'] = msg.to
    if msg.cc:
        eml['Cc'] = msg.cc
    if msg.subject:
        eml['Subject'] = msg.subject
    if msg.date:
        eml['Date'] = msg.date

    # Corpo do texto (text/html e fallback plain/text)
    body_text = msg.body
    body_html = msg.htmlBody

    if body_text:
        eml.set_content(body_text)

    if body_html:
        # Tenta decodificar o HTML caso seja recebido em bytes
        html_str = body_html.decode('utf-8', errors='ignore') if isinstance(body_html, bytes) else body_html
        if body_text:
            eml.add_alternative(html_str, subtype='html')
        else:
            eml.set_content(html_str, subtype='html')

    # Adicionar os anexos do msg para que fiquem no eml
    for attachment in msg.attachments:
        try:
            content = attachment.data
            filename = getattr(attachment, 'longFilename', None) or getattr(attachment, 'shortFilename', None) or "attachment"
            eml.add_attachment(content, maintype='application', subtype='octet-stream', filename=filename)
        except Exception:
            pass # Ignora caso de erro em arquivo anexo comrrompido

    with open(eml_path, 'wb') as f:
        f.write(bytes(eml))
        
    msg.close()

def main():
    # py2app com argv_emulation=True guarda os caminhos arrastados no sys.argv[1:]
    if len(sys.argv) < 2:
        print("Uso: MSGtoOutlook <arquivo.msg>")
        sys.exit(1)

    # Iterar porque o py2app por vezes injeta argumentos inesperados
    msg_path = None
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.msg'):
            msg_path = arg
            break

    if not msg_path:
        print("Erro: Ficheiro .msg não foi detetado nos argumentos do sistema.")
        sys.exit(1)

    if not os.path.exists(msg_path):
        print(f"Erro: Ficheiro não existe ({msg_path})")
        sys.exit(1)

    try:
        # Formatar nome do ficheiro e gravar em /tmp/
        filename = os.path.basename(msg_path)
        name, _ = os.path.splitext(filename)
        eml_filename = f"{name}.eml"
        eml_path = os.path.join("/tmp", eml_filename)

        # Converter
        convert_msg_to_eml(msg_path, eml_path)

        # Usar subprocess para abrir EML no Microsoft Outlook para macOS
        subprocess.run(['open', '-a', 'Microsoft Outlook', eml_path], check=True)

    except Exception as e:
        print(f"Erro inesperado: {e}")
        sys.exit(1)

if __name__ == '__main__':
    main()

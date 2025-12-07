import firebase_admin
from firebase_admin import credentials, firestore
import time
import os
import json
import requests
import win32print
import win32api
import socket
import uuid
import mimetypes # <--- IMPORTANTE: Adicionado para detectar extensões

# --- CONFIGURAÇÕES GERAIS ---
ARQUIVO_CONFIG = "config.json"
PASTA_DOWNLOAD = "arquivos_temp"

if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

# --- 1. GERENCIAMENTO DE IDENTIDADE DO PC ---
def obter_identidade_pc():
    if os.path.exists(ARQUIVO_CONFIG):
        with open(ARQUIVO_CONFIG, 'r') as f:
            try:
                config = json.load(f)
                if "pc_id" in config:
                    return config
            except:
                pass 
    
    novo_id = str(uuid.uuid4())
    nome_pc = socket.gethostname()
    
    dados_iniciais = {
        "pc_id": novo_id,
        "nome_amigavel": nome_pc,
        "medico_dono_email": "nao_vinculado"
    }
    
    with open(ARQUIVO_CONFIG, 'w') as f:
        json.dump(dados_iniciais, f)
    
    return dados_iniciais

# --- 2. FUNÇÕES DE IMPRESSORA (COM FILTRO DE ATIVAS) ---
def listar_impressoras_ativas():
    impressoras_ativas = []
    flags_enum = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    
    try:
        lista_raw = win32print.EnumPrinters(flags_enum)
        for p in lista_raw:
            nome_impressora = p[2]
            try:
                handle = win32print.OpenPrinter(nome_impressora)
                info = win32print.GetPrinter(handle, 2)
                win32print.ClosePrinter(handle)

                status = info.get('Status', 0)
                atributos = info.get('Attributes', 0)

                # Verifica se está offline
                is_offline = (status & 0x80) or (atributos & 0x400) # PRINTER_STATUS_OFFLINE
                is_error = status & 0x02

                if not is_offline and not is_error:
                    impressoras_ativas.append(nome_impressora)
            except:
                # Se der erro ao ler status, ignora ou assume que não está ativa
                pass

    except Exception as e:
        print(f"Erro ao listar impressoras: {e}")
    
    return impressoras_ativas

def imprimir_arquivo(caminho_arquivo, nome_impressora):
    print(f"Tentando imprimir: {caminho_arquivo} na impressora: {nome_impressora}")
    
    try:
        # O comando "printto" usa o programa padrão do Windows para aquela extensão
        # Ex: Se for .pdf usa o Adobe, se for .txt usa o Bloco de Notas, etc.
        win32api.ShellExecute(
            0,
            "printto",
            caminho_arquivo,
            f'"{nome_impressora}"',
            ".",
            0
        )
        return True
    except Exception as e:
        print(f"Erro na impressão (ShellExecute): {e}")
        return False

# --- 3. CONEXÃO FIREBASE ---
try:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)
    
    # Mantido conforme sua instrução: valut
    db = firestore.client(database_id='cloud-valut-storage')
    
except Exception as e:
    print(f"Erro fatal ao conectar no Firebase: {e}")
    exit()

# --- 4. LOOP PRINCIPAL ---
def iniciar_robo():
    print("🤖 INICIANDO ROBÔ (Suporte Multiformato)")
    
    config = obter_identidade_pc()
    meu_id = config['pc_id']
    print(f"ID do PC: {meu_id}")

    # Atualiza status inicial
    doc_pc_ref = db.collection('dispositivos_online').document(meu_id)
    doc_pc_ref.set({
        'nome': config['nome_amigavel'],
        'impressoras': listar_impressoras_ativas(),
        'status': 'online',
        'ultimo_visto': firestore.SERVER_TIMESTAMP
    }, merge=True)

    def on_snapshot(col_snapshot, changes, read_time):
        for change in changes:
            if change.type.name == 'ADDED':
                doc = change.document
                dados = doc.to_dict()

                if (dados.get('pc_alvo_id') == meu_id and 
                    dados.get('status') == 'pendente'):
                    
                    print(f"🔔 Novo arquivo recebido! ID: {doc.id}")
                    
                    url = dados.get('url_arquivo')
                    printer_escolhida = dados.get('impressora_alvo')
                    
                    # Verifica impressora
                    ativas = listar_impressoras_ativas()
                    if printer_escolhida not in ativas:
                        print(f"❌ Impressora '{printer_escolhida}' não encontrada ou offline.")
                        # Tenta mandar mesmo assim se não achar nas ativas?
                        # Se quiser forçar, remova o 'continue' abaixo.
                        doc.reference.update({'status': 'erro_impressora_offline'})
                        continue

                    try:
                        # 1. Baixa o arquivo PRIMEIRO para ver o cabeçalho
                        r = requests.get(url)
                        r.raise_for_status()

                        # 2. Descobre a extensão real (PDF, PNG, JPG, TXT...)
                        content_type = r.headers.get('Content-Type')
                        extensao = mimetypes.guess_extension(content_type)
                        
                        # Fallback se não descobrir
                        if not extensao:
                            if '.pdf' in url.lower(): extensao = '.pdf'
                            elif '.png' in url.lower(): extensao = '.png'
                            elif '.jpg' in url.lower(): extensao = '.jpg'
                            else: extensao = '.pdf' # Último caso
                            
                        # Correção para imagens JPEG que as vezes vem como .jpe
                        if extensao == '.jpe': extensao = '.jpg'

                        # 3. Salva com a extensão correta
                        nome_local = f"{doc.id}{extensao}"
                        caminho_completo = os.path.join(PASTA_DOWNLOAD, nome_local)
                        
                        # Limpa arquivo anterior se existir
                        if os.path.exists(caminho_completo):
                            os.remove(caminho_completo)

                        with open(caminho_completo, 'wb') as f:
                            f.write(r.content)
                        
                        print(f"📂 Arquivo salvo como: {nome_local} (Tipo detectado: {content_type})")

                        # 4. Manda imprimir
                        sucesso = imprimir_arquivo(caminho_completo, printer_escolhida)
                        
                        if sucesso:
                            doc.reference.update({'status': 'impresso'})
                            print("✅ Comando de impressão enviado.")
                        else:
                            doc.reference.update({'status': 'erro_driver'})
                            
                    except Exception as e:
                        print(f"Erro no processamento: {e}")
                        doc.reference.update({'status': 'erro_download'})

    query = db.collection('fila_impressao').where('pc_alvo_id', '==', meu_id).where('status', '==', 'pendente')
    watch = query.on_snapshot(on_snapshot)

    while True:
        time.sleep(60)
        # Heartbeat
        doc_pc_ref.update({
            'ultimo_visto': firestore.SERVER_TIMESTAMP,
            'impressoras': listar_impressoras_ativas()
        })

if __name__ == "__main__":
    iniciar_robo()
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

# --- CONFIGURAÇÕES GERAIS ---
ARQUIVO_CONFIG = "config.json"
PASTA_DOWNLOAD = "arquivos_temp"

# Cria pasta temporária se não existir
if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

# --- 1. GERENCIAMENTO DE IDENTIDADE DO PC ---
def obter_identidade_pc():
    """
    Verifica se este PC já tem um ID registrado no config.json.
    Se não tiver, cria um novo ID único e salva.
    Isso garante que mesmo se o médico mudar o nome do PC, o ID é o mesmo.
    """
    if os.path.exists(ARQUIVO_CONFIG):
        with open(ARQUIVO_CONFIG, 'r') as f:
            try:
                config = json.load(f)
                if "pc_id" in config:
                    return config
            except:
                pass # Arquivo vazio ou corrompido
    
    # Se chegou aqui, é a primeira vez que roda ou config sumiu
    novo_id = str(uuid.uuid4()) # Gera um ID único universal
    nome_pc = socket.gethostname()
    
    dados_iniciais = {
        "pc_id": novo_id,
        "nome_amigavel": nome_pc, # O médico poderá mudar isso no site depois
        "medico_dono_email": "nao_vinculado" # Vamos usar isso no futuro para segurança
    }
    
    with open(ARQUIVO_CONFIG, 'w') as f:
        json.dump(dados_iniciais, f)
    
    return dados_iniciais

# --- 2. FUNÇÕES DE IMPRESSORA ---
def listar_impressoras():
    """Lista todas as impressoras instaladas no Windows"""
    impressoras = []
    # O parametro 2 (PRINTER_ENUM_LOCAL) lista as locais, 4 (CONNECTIONS) as de rede
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    
    try:
        lista_raw = win32print.EnumPrinters(flags)
        for p in lista_raw:
            impressoras.append(p[2]) # p[2] é o nome da impressora
    except Exception as e:
        print(f"Erro ao listar impressoras: {e}")
    
    return impressoras

def imprimir_arquivo(caminho_arquivo, nome_impressora):
    """Manda imprimir numa impressora ESPECÍFICA"""
    print(f"Tentando imprimir em: {nome_impressora}")
    
    try:
        # Método seguro usando ShellExecute com flag de impressora específica
        # A sintaxe 'printto' permite escolher o device
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
        print(f"Erro na impressão: {e}")
        return False

# --- 3. CONEXÃO FIREBASE ---
# Certifique-se que o arquivo serviceAccountKey.json está na pasta
try:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)
    
    # AQUI ESTÁ A CORREÇÃO:
    # Usamos o nome exato que aparece no seu print (incluindo o 'valut')
    db = firestore.client(database_id='cloud-valut-storage')
    
except Exception as e:
    print(f"Erro fatal ao conectar no Firebase. Verifique a chave. {e}")
    exit()

# --- 4. LOOP PRINCIPAL ---
def iniciar_robo():
    print("🤖 INICIANDO ROBÔ DE IMPRESSÃO 2.0")
    
    # 1. Quem sou eu?
    config = obter_identidade_pc()
    meu_id = config['pc_id']
    meu_nome = config['nome_amigavel']
    print(f"ID do PC: {meu_id}")
    print(f"Nome: {meu_nome}")

    # 2. Quais minhas impressoras?
    minhas_impressoras = listar_impressoras()
    print(f"Impressoras encontradas: {minhas_impressoras}")

    # 3. Atualiza o Firebase (HEARTBEAT)
    # Isso diz pro seu site: "Estou online e essas são minhas impressoras"
    doc_pc_ref = db.collection('dispositivos_online').document(meu_id)
    doc_pc_ref.set({
        'nome': meu_nome,
        'impressoras': minhas_impressoras,
        'status': 'online',
        'ultimo_visto': firestore.SERVER_TIMESTAMP
    }, merge=True)

    print("✅ Status atualizado na nuvem. Aguardando impressões...")

    # 4. O Listener (Ouvido)
    # Agora escutamos apenas documentos destinados a ESTE PC (pc_alvo == meu_id)
    def on_snapshot(col_snapshot, changes, read_time):
        for change in changes:
            if change.type.name == 'ADDED':
                doc = change.document
                dados = doc.to_dict()

                # Segurança: Só processa se for pra mim e estiver pendente
                if (dados.get('pc_alvo_id') == meu_id and 
                    dados.get('status') == 'pendente'):
                    
                    print(f"🔔 Novo pedido de impressão recebido! ID: {doc.id}")
                    
                    url = dados.get('url_arquivo')
                    printer_escolhida = dados.get('impressora_alvo') # O site vai mandar esse nome
                    
                    # Verifica se a impressora pedida ainda existe
                    if printer_escolhida not in listar_impressoras():
                        print(f"❌ Erro: Impressora '{printer_escolhida}' não encontrada.")
                        doc.reference.update({'status': 'erro_impressora_inexistente'})
                        continue

                    # Baixa
                    nome_local = f"{doc.id}.pdf"
                    caminho_completo = os.path.join(PASTA_DOWNLOAD, nome_local)
                    
                    try:
                        r = requests.get(url)
                        with open(caminho_completo, 'wb') as f:
                            f.write(r.content)
                        
                        # Imprime
                        sucesso = imprimir_arquivo(caminho_completo, printer_escolhida)
                        
                        if sucesso:
                            doc.reference.update({'status': 'impresso'})
                            print("✅ Impressão enviada com sucesso.")
                        else:
                            doc.reference.update({'status': 'erro_driver'})
                            
                    except Exception as e:
                        print(f"Erro no download/processamento: {e}")
                        doc.reference.update({'status': 'erro_download'})

    # Conecta o listener na coleção de 'fila_impressao'
    # Filtra apenas onde pc_alvo_id == meu_id para economizar banda
    query = db.collection('fila_impressao').where('pc_alvo_id', '==', meu_id).where('status', '==', 'pendente')
    watch = query.on_snapshot(on_snapshot)

    while True:
        # A cada 1 minuto, atualiza o status "online" para o site saber que não caiu
        time.sleep(60)
        doc_pc_ref.update({'ultimo_visto': firestore.SERVER_TIMESTAMP})

if __name__ == "__main__":
    iniciar_robo()
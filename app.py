from flask import Flask, request, send_file, jsonify, render_template, after_this_request
import fitz
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.graphics.barcode import code128
from reportlab.lib.utils import ImageReader
import io
from PIL import Image
import re
import time
import os
import traceback
import atexit
import tempfile
from flask_cors import CORS

HTML_TEMPLATE = """

"""

# ==============================
# PRE-CLEAN: remove páginas-sobra
# ==============================
def preclean_pdf_remove_overflow_by_blocks(input_pdf_path: str) -> str:
    """
    Remove páginas de 'sobra' (continuações/fragmentos da DANFE) antes do processamento principal.
    Heurística baseada em blocos de texto (sem rasterizar):

      - poucos blocos (<= 3)
      - blocos não passam de ~40% da altura, nem ocupam mais de 1-2 faixas verticais
      - texto muito semelhante/contido no da página anterior

    Retorna: caminho do PDF limpo (arquivo *_precleaned.pdf) ou o original se nada removido.
    """
    import fitz  # PyMuPDF

    def norm_text(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "")).strip().lower()

    KEEP_HEADERS = [
        "danfe", "fim do danfe", "chave de acesso",
        "destinatário", "remetente", "pedido criado", "pieces", "peso",
        "item", "conteúdo", "atributos", "quant"
    ]

    doc = fitz.open(input_pdf_path)
    if len(doc) == 0:
        doc.close()
        return input_pdf_path

    to_drop = set()
    prev_text_norm = ""

    for i, page in enumerate(doc):
        W, H = page.rect.width, page.rect.height
        text_norm = norm_text(page.get_text("text"))

        # Se parecer claramente uma etiqueta/DANFE com cabeçalho, mantém
        if any(h in text_norm for h in KEEP_HEADERS):
            prev_text_norm = text_norm
            continue

        blocks_raw = page.get_text("blocks") or []
        blocks = []
        total_area = W * H
        area_text = 0.0

        # Filtra blocos com texto real e descarta ruídos muito pequenos
        for b in blocks_raw:
            if len(b) < 5 or not isinstance(b[4], str) or not b[4].strip():
                continue
            x0, y0, x1, y1 = map(float, b[:4])
            bw, bh = max(0.0, x1 - x0), max(0.0, y1 - y0)
            if bw * bh < 0.0003 * total_area:
                continue
            blocks.append((x0, y0, x1, y1))
            area_text += bw * bh

        if not blocks:
            prev_text_norm = text_norm
            continue

        # Métricas geométricas
        max_y = max(y1 for _, _, _, y1 in blocks) / H
        max_x = max(x1 for _, _, x1, _ in blocks) / W

        # Dispersão vertical/horizontal em "faixas"
        V_BANDS, H_COLS = 8, 4
        v_occ = set()
        h_occ = set()
        for x0, y0, x1, y1 in blocks:
            v0 = int((y0 / H) * V_BANDS); v1 = int((y1 / H) * V_BANDS)
            h0 = int((x0 / W) * H_COLS);  h1 = int((x1 / W) * H_COLS)
            for v in range(max(0, v0), min(V_BANDS - 1, max(0, v1)) + 1):
                v_occ.add(v)
            for h in range(max(0, h0), min(H_COLS - 1, max(0, h1)) + 1):
                h_occ.add(h)

        num_blocks = len(blocks)

        # Similaridade com a página anterior (para detectar "continuação"/repetição)
        similar_prev = False
        if prev_text_norm:
            if text_norm and (text_norm in prev_text_norm or prev_text_norm in text_norm):
                similar_prev = True
            else:
                a, b = set(prev_text_norm.split()), set(text_norm.split())
                jacc = (len(a & b) / (len(a | b) or 1))
                if jacc >= 0.60:
                    similar_prev = True

        # Regras para detectar páginas fragmentadas/problemáticas
        
        # 1. Fragmentos pequenos concentrados no topo
        is_small_fragment = (num_blocks <= 3 and max_y <= 0.40 and len(v_occ) <= 2)
        
        # 2. Continuação/repetição da página anterior
        is_continuation = similar_prev
        
        # 3. Páginas com conteúdo espalhado/desorganizado (NOVO)
        # - Muitos blocos pequenos espalhados
        # - Sem estrutura principal (sem "DANFE" completo)
        # - Densidade baixa (muito espaço vazio entre blocos)
        has_main_structure = "danfe" in text_norm and ("destinatário" in text_norm or "documento auxiliar" in text_norm)
        
        if not has_main_structure and num_blocks >= 3:
            # Calcular densidade de conteúdo
            total_block_area = sum((x1-x0)*(y1-y0) for x0,y0,x1,y1 in blocks)
            page_area = W * H
            density = total_block_area / page_area
            
            # Verificar se blocos estão muito espalhados
            y_positions = [y0 for x0,y0,x1,y1 in blocks] + [y1 for x0,y0,x1,y1 in blocks]
            y_spread = (max(y_positions) - min(y_positions)) / H if y_positions else 0
            
            # Página fragmentada se:
            # - Densidade baixa/média (< 35% da página ocupada) OU
            # - Blocos espalhados (> 70% da altura) E sem estrutura
            is_scattered_fragment = ((density < 0.35 and y_spread > 0.70) or
                                   (density < 0.20 and y_spread > 0.60))
        else:
            is_scattered_fragment = False
        
        # 4. Páginas com apenas códigos/fragmentos de produtos (NOVO)
        # - Contém códigos de produto mas sem DANFE
        # - Texto curto e fragmentado
        has_product_codes = bool(re.search(r'i\d{2}[a-z0-9]{6,10}', text_norm))  # Mais flexível
        is_product_fragment = (has_product_codes and not has_main_structure and len(text_norm) < 600)
        
        # Decisão final
        should_remove = (is_small_fragment or is_continuation or
                        is_scattered_fragment or is_product_fragment)
        
        if should_remove:
            to_drop.add(i)
            reasons = []
            if is_small_fragment: reasons.append("SmallFragment")
            if is_continuation: reasons.append("Continuation")
            if is_scattered_fragment: reasons.append("ScatteredFragment")
            if is_product_fragment: reasons.append("ProductFragment")
            
            print(f"[preclean] Página {i+1} marcada para remoção - {', '.join(reasons)}")
            if is_scattered_fragment:
                print(f"  └─ Densidade: {density:.3f}, Espalhamento Y: {y_spread:.3f}")

        prev_text_norm = text_norm

    if not to_drop:
        doc.close()
        return input_pdf_path

    cleaned = fitz.open()
    for i in range(len(doc)):
        if i in to_drop:
            print(f"[preclean] removendo página {i+1} (fragmento/continuação)")
            continue
        cleaned.insert_pdf(doc, from_page=i, to_page=i)

    base, ext = os.path.splitext(input_pdf_path)
    out_path = base + "_precleaned" + ext
    cleaned.save(out_path)
    cleaned.close()
    doc.close()
    return out_path

app = Flask(__name__)
CORS(app)  # Adiciona suporte CORS para permitir requisições de diferentes origens

# Configurar limite de tamanho de upload para 50MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB em bytes

# Lista para rastrear arquivos temporários
temp_files = []  

# Função para limpar arquivos temporários
def cleanup_temp_files():
    for file_path in temp_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Arquivo temporário removido: {file_path}")
        except Exception as e:
            print(f"Erro ao remover arquivo temporário {file_path}: {str(e)}")
    temp_files.clear()

# Registrar função de limpeza para ser executada quando o aplicativo for encerrado
atexit.register(cleanup_temp_files)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/processar-pdf', methods=['POST'])
@app.route('/api/processar-pdf', methods=['POST'])
def processar_pdf():
    # Criar arquivos temporários com o módulo tempfile para garantir limpeza adequada
    input_pdf_fd, input_pdf = tempfile.mkstemp(suffix='.pdf', prefix='temp_input_')
    output_pdf_fd, output_pdf = tempfile.mkstemp(suffix='.pdf', prefix='temp_output_')
    
    # Fechar os descritores de arquivo criados pelo tempfile
    os.close(input_pdf_fd)
    os.close(output_pdf_fd)
    
    # Adicionar à lista de arquivos temporários para garantir limpeza
    temp_files.append(input_pdf)
    temp_files.append(output_pdf)
    
    try:
        # Verifica se foi enviado um arquivo
        if 'arquivo' not in request.files:
            return jsonify({'erro': 'Nenhum arquivo enviado'}), 400
            
        arquivo = request.files['arquivo']
        
        # Verifica se o nome do arquivo está vazio
        if arquivo.filename == '':
            return jsonify({'erro': 'Nome do arquivo vazio'}), 400
        
        # Salva o arquivo temporariamente
        arquivo.save(input_pdf)
        
        # PRE-CLEAN: Remove páginas de sobra/fragmentos antes do processamento principal
        print("[INFO] Iniciando pré-limpeza do PDF...")
        cleaned_pdf = preclean_pdf_remove_overflow_by_blocks(input_pdf)
        
        # Se foi criado um arquivo limpo, adiciona à lista de arquivos temporários
        if cleaned_pdf != input_pdf:
            temp_files.append(cleaned_pdf)
            print(f"[INFO] PDF pré-limpo salvo em: {cleaned_pdf}")
        else:
            print("[INFO] Nenhuma página removida na pré-limpeza")
        
        # Processa o PDF (usando o arquivo limpo se disponível)
        extracted_data = extract_text_from_pdf(cleaned_pdf)
        if extracted_data:
            print(f"[PROCESSAMENTO] Dados extraídos com sucesso: {len(extracted_data)} DANFEs encontradas")
            pdf_gerado = create_individual_page_pdf(output_pdf, extracted_data, cleaned_pdf)
            
            if not pdf_gerado:
                print("[PROCESSAMENTO] ERRO: Falha na geração do PDF final")
                # Limpar arquivos temporários em caso de erro
                try:
                    if os.path.exists(input_pdf):
                        os.remove(input_pdf)
                        if input_pdf in temp_files:
                            temp_files.remove(input_pdf)
                    if os.path.exists(output_pdf):
                        os.remove(output_pdf)
                        if output_pdf in temp_files:
                            temp_files.remove(output_pdf)
                    if cleaned_pdf != input_pdf and os.path.exists(cleaned_pdf):
                        os.remove(cleaned_pdf)
                        if cleaned_pdf in temp_files:
                            temp_files.remove(cleaned_pdf)
                except Exception as cleanup_error:
                    print(f"Erro ao limpar arquivos após falha: {str(cleanup_error)}")
                
                return jsonify({
                    'erro': 'Falha na geração do PDF', 
                    'mensagem': 'Os dados foram extraídos mas houve erro na geração do PDF final. Verifique se o PDF contém dados válidos.'
                }), 500
            
            # Registra função para limpar os arquivos após o request
            @after_this_request
            def cleanup_after_request(response):
                try:
                    # Remover os arquivos temporários após o envio da resposta
                    if os.path.exists(input_pdf):
                        os.remove(input_pdf)
                        if input_pdf in temp_files:
                            temp_files.remove(input_pdf)
                        print(f"Arquivo temporário removido após request: {input_pdf}")
                    if os.path.exists(output_pdf):
                        os.remove(output_pdf)
                        if output_pdf in temp_files:
                            temp_files.remove(output_pdf)
                        print(f"Arquivo temporário removido após request: {output_pdf}")
                    # Remover arquivo pré-limpo se existir
                    if cleaned_pdf != input_pdf and os.path.exists(cleaned_pdf):
                        os.remove(cleaned_pdf)
                        if cleaned_pdf in temp_files:
                            temp_files.remove(cleaned_pdf)
                        print(f"Arquivo pré-limpo removido após request: {cleaned_pdf}")
                except Exception as e:
                    print(f"Erro ao remover arquivos temporários após request: {str(e)}")
                return response
            
            # Envia o arquivo processado como resposta
            try:
                response = send_file(
                    output_pdf,
                    mimetype='application/pdf',
                    as_attachment=True,
                    download_name='processado.pdf'
                )
                
                return response
            except Exception as e:
                # Se houver erro ao enviar o arquivo, tenta remover os temporários
                print(f"Erro ao enviar o arquivo: {str(e)}")
                try:
                    if os.path.exists(input_pdf):
                        os.remove(input_pdf)
                        temp_files.remove(input_pdf)
                    if os.path.exists(output_pdf):
                        os.remove(output_pdf)
                        temp_files.remove(output_pdf)
                except Exception as cleanup_error:
                    print(f"Erro ao limpar arquivos após falha de envio: {str(cleanup_error)}")
                raise
        else:
            # Se não extraiu dados, remove o arquivo de entrada
            try:
                if os.path.exists(input_pdf):
                    os.remove(input_pdf)
                    if input_pdf in temp_files:
                        temp_files.remove(input_pdf)
                if os.path.exists(output_pdf):
                    os.remove(output_pdf)
                    if output_pdf in temp_files:
                        temp_files.remove(output_pdf)
                # Remover arquivo pré-limpo se existir
                if cleaned_pdf != input_pdf and os.path.exists(cleaned_pdf):
                    os.remove(cleaned_pdf)
                    if cleaned_pdf in temp_files:
                        temp_files.remove(cleaned_pdf)
            except Exception as e:
                print(f"Erro ao remover arquivos temporários: {str(e)}")
                
            return jsonify({
                'erro': 'Nenhum dado extraído do PDF', 
                'mensagem': 'O PDF enviado não parece conter o formato esperado. Certifique-se de que o PDF contém uma DANFE com a chave de acesso e itens.'
            }), 400
            
    except Exception as e:
        # Log do erro completo para debug
        error_trace = traceback.format_exc()
        print(error_trace)
        
        # Limpar arquivos temporários em caso de erro
        try:
            if os.path.exists(input_pdf):
                os.remove(input_pdf)
                if input_pdf in temp_files:
                    temp_files.remove(input_pdf)
            if os.path.exists(output_pdf):
                os.remove(output_pdf)
                if output_pdf in temp_files:
                    temp_files.remove(output_pdf)
            # Remover arquivo pré-limpo se existir
            if 'cleaned_pdf' in locals() and cleaned_pdf != input_pdf and os.path.exists(cleaned_pdf):
                os.remove(cleaned_pdf)
                if cleaned_pdf in temp_files:
                    temp_files.remove(cleaned_pdf)
        except Exception as cleanup_error:
            print(f"Erro ao limpar arquivos temporários: {str(cleanup_error)}")
        
        # Mensagem de erro mais amigável para o usuário
        error_message = str(e)
        if "already being used by another process" in error_message:
            user_message = "O arquivo está sendo usado por outro processo. Por favor, tente novamente em alguns instantes."
        elif "Permission denied" in error_message:
            user_message = "Erro de permissão ao acessar os arquivos. Por favor, tente novamente."
        else:
            user_message = "Ocorreu um erro ao processar o PDF. Por favor, verifique se o formato está correto e tente novamente."
            
        return jsonify({
            'erro': str(e),
            'mensagem': user_message
        }), 500

# Resto do código permanece igual
def extract_text_from_pdf(input_pdf):
    inicio = time.time()
    doc = fitz.open(input_pdf)
    extracted_data = []
    page_num = 0
    
    print(f"[EXTRAÇÃO] Iniciando extração de {doc.page_count} páginas")
    
    while page_num < doc.page_count:
        page = doc.load_page(page_num)
        text = page.get_text("text")
        
        print(f"[EXTRAÇÃO] Analisando página {page_num + 1}")

        # Verificação mais flexível para DANFE
        if not ("DANFE" in text.upper() or text.startswith("DANFE")):
            print(f"[EXTRAÇÃO] Página {page_num + 1} não contém DANFE, pulando...")
            page_num += 1
            continue

        try:
            # Busca mais robusta pela chave de acesso
            chave_acesso = None
            chave_patterns = ["CHAVE DE ACESSO", "CHAVE DE ACESSO:", "CHAVE ACESSO"]
            
            for pattern in chave_patterns:
                if pattern in text:
                    chave_acesso_index = text.index(pattern)
                    chave_linha = text[chave_acesso_index + len(pattern):].strip().split('\n')[0]
                    # Limpar a chave de acesso (remover espaços e caracteres especiais)
                    chave_acesso = ''.join(c for c in chave_linha if c.isdigit())
                    if len(chave_acesso) >= 40:  # Chave de acesso válida tem 44 dígitos
                        break
                    else:
                        chave_acesso = None
            
            if not chave_acesso:
                print(f"[EXTRAÇÃO] ERRO: Chave de acesso não encontrada na página {page_num + 1}")
                page_num += 1
                continue
                
            print(f"[EXTRAÇÃO] Chave de acesso encontrada: {chave_acesso[:10]}...")

            # Busca mais robusta pelos itens
            item_patterns = ["ITEM", "CÓDIGO", "DESCRIÇÃO"]
            item_index = -1
            
            for pattern in item_patterns:
                if pattern in text:
                    item_index = text.index(pattern)
                    break
            
            if item_index == -1:
                print(f"[EXTRAÇÃO] ERRO: Seção de itens não encontrada na página {page_num + 1}")
                page_num += 1
                continue

            texto_completo = text[item_index:]

            # Verificar próxima página para continuação
            proxima_pagina = page_num + 1
            if proxima_pagina < doc.page_count:
                next_page = doc.load_page(proxima_pagina)
                if not next_page.get_images():
                    next_text = next_page.get_text("text")
                    if next_text and not "DANFE" in next_text.upper():
                        texto_completo += "\n" + next_text
                        print(f"[EXTRAÇÃO] Incluindo continuação da página {proxima_pagina + 1}")

            # Processamento de Itens - Abordagem híbrida melhorada
            linhas = texto_completo.strip().split('\n')
            itens = []
            item_atual = []
            
            print(f"[EXTRAÇÃO] Processando {len(linhas)} linhas para extrair itens...")
            print(f"[EXTRAÇÃO] DEBUG - Primeiras 10 linhas do texto:")
            for i, linha in enumerate(linhas[:10]):
                print(f"[EXTRAÇÃO] Linha {i}: '{linha.strip()}'")
            
            quantidade_atual = "1"  # Quantidade padrão
            
            for i, linha in enumerate(linhas[1:]):  # Pular primeira linha (cabeçalho)
                linha_limpa = linha.strip()
                
                # Pular linhas vazias ou cabeçalhos conhecidos
                if not linha_limpa or linha_limpa in ["CONTEÚDO", "ATRIBUTOS", "QUANT.", "DESCRIÇÃO", "CÓDIGO"]:
                    continue
                
                # Detectar quantidade após QUANT.
                if i > 0 and "QUANT." in linhas[i].upper():
                    # Procurar quantidade na próxima linha ou na mesma linha
                    linha_quant = linha_limpa
                    if linha_quant.isdigit():
                        quantidade_atual = linha_quant
                        print(f"[EXTRAÇÃO] Quantidade encontrada: '{quantidade_atual}'")
                        continue
                    # Se não encontrou na linha atual, procurar na próxima
                    elif i + 1 < len(linhas):
                        proxima_linha = linhas[i + 1].strip()
                        if proxima_linha.isdigit():
                            quantidade_atual = proxima_linha
                            print(f"[EXTRAÇÃO] Quantidade encontrada na próxima linha: '{quantidade_atual}'")
                
                # Detectar início de novo item - deve ser um código que começa com letra/número longo
                # ou um número sequencial pequeno APENAS se não temos item atual
                is_new_item = False
                
                # Se não temos item atual e é um número pequeno, pode ser início de item
                if not item_atual and linha_limpa.isdigit() and len(linha_limpa) <= 2 and int(linha_limpa) <= 50:
                    is_new_item = True
                    print(f"[EXTRAÇÃO] Detectado início de novo item (número sequencial): '{linha_limpa}'")
                
                # Se parece com um código (letras + números, mais de 5 caracteres)
                elif re.match(r'^[A-Za-z0-9]{5,}', linha_limpa) and not item_atual:
                    is_new_item = True
                    print(f"[EXTRAÇÃO] Detectado início de novo item (código): '{linha_limpa}'")
                
                if is_new_item:
                    # Salvar item anterior se existir
                    if item_atual and len(item_atual) >= 2:
                        codigo = item_atual[0]
                        conteudo = " ".join(item_atual[1:])
                        if codigo and conteudo and len(codigo) > 3:  # Validar código mínimo
                            itens.append([codigo, conteudo, quantidade_atual])
                            print(f"[EXTRAÇÃO] Item adicionado: Código='{codigo}', Desc='{conteudo[:30]}...', Qtd='{quantidade_atual}'")
                    item_atual = []
                    quantidade_atual = "1"  # Reset quantidade para próximo item
                
                if linha_limpa:
                    # Adicionar linha ao item atual
                    item_atual.append(linha_limpa)
                    print(f"[EXTRAÇÃO] Linha adicionada ao item: '{linha_limpa[:50]}...'")
            
            # Processar último item se existir
            if item_atual and len(item_atual) >= 2:
                codigo = item_atual[0]
                conteudo = " ".join(item_atual[1:])
                if codigo and conteudo and len(codigo) > 3:
                    itens.append([codigo, conteudo, quantidade_atual])
                    print(f"[EXTRAÇÃO] Último item adicionado: Código='{codigo}', Desc='{conteudo[:30]}...', Qtd='{quantidade_atual}'")
            
            print(f"[EXTRAÇÃO] Total de itens extraídos: {len(itens)}")

            # Validar se extraiu dados válidos
            if itens:
                extracted_data.append([chave_acesso, itens])
                print(f"[EXTRAÇÃO] Sucesso: {len(itens)} itens extraídos da página {page_num + 1}")
            else:
                print(f"[EXTRAÇÃO] AVISO: Nenhum item válido extraído da página {page_num + 1}")

        except ValueError as e:
            print(f"[EXTRAÇÃO] ERRO: Falha ao extrair dados na página {page_num + 1}: {str(e)}")
        except Exception as e:
            print(f"[EXTRAÇÃO] ERRO INESPERADO na página {page_num + 1}: {str(e)}")

        page_num += 2

    doc.close()
    fim = time.time()
    print(f"[EXTRAÇÃO] Concluída em {fim - inicio:.2f}s - {len(extracted_data)} DANFEs processadas")
    
    # Validação final
    if not extracted_data:
        print("[EXTRAÇÃO] ERRO: Nenhum dado válido foi extraído do PDF")
    
    return extracted_data

def create_individual_page_pdf(output_pdf, data, input_pdf):
    inicio = time.time()
    
    # Validação inicial dos dados
    if not data:
        print("[GERAÇÃO] ERRO: Nenhum dado fornecido para gerar PDF")
        return False
        
    print(f"[GERAÇÃO] Iniciando geração de PDF com {len(data)} DANFEs")
    
    doc = fitz.open(input_pdf)
    c = canvas.Canvas(output_pdf, pagesize=(799, 1197))
    width, height = c._pagesize
    
    paginas_geradas = 0

    for i, row in enumerate(data):
        try:
            chave_acesso, itens = row
            
            # Validações robustas
            if not chave_acesso or not itens:
                print(f"[GERAÇÃO] AVISO: DANFE {i+1} tem dados inválidos (chave: {bool(chave_acesso)}, itens: {len(itens) if itens else 0})")
                continue
                
            if len(chave_acesso) < 40:
                print(f"[GERAÇÃO] AVISO: Chave de acesso inválida na DANFE {i+1}: {chave_acesso}")
                continue
                
            print(f"[GERAÇÃO] Processando DANFE {i+1}: {len(itens)} itens")

            # Gerar código de barras
            try:
                barcode = code128.Code128(chave_acesso, barHeight=1.8 * cm, barWidth=0.05 * cm)
                c.saveState()
                c.rotate(90)
                barcode.drawOn(c, height - 14.00 * cm - 0.80 * cm, -width + 0.50 * cm)
                c.restoreState()
                print(f"[GERAÇÃO] Código de barras gerado para DANFE {i+1}")
            except Exception as e:
                print(f"[GERAÇÃO] ERRO ao gerar código de barras para DANFE {i+1}: {str(e)}")
                continue

            # Texto da chave de acesso
            text_x = width - 0.10 * cm
            text_y = height - 12.0 * cm
            c.saveState()
            c.translate(text_x, text_y)
            c.rotate(90)
            c.drawString(0, 0, chave_acesso)
            c.restoreState()

            # Preparar dados da tabela com validação
            table_data = []
            itens_validos = 0
            
            for item in itens:
                if len(item) >= 3:
                    codigo, conteudo, quantidade = item[0], item[1], item[2]
                    
                    # Validar que os dados não estão vazios
                    if codigo and conteudo:
                        # Função para corrigir palavras cortadas
                        def corrigir_palavras_cortadas(texto):
                            # Dicionário de correções comuns
                            correcoes = {
                                r'\bU\s+nissex\b': 'Unissex',
                                r'\bSkat\s+ista\b': 'Skatista', 
                                r'\bMa\s+sculino\b': 'Masculino',
                                r'\bFe\s+minino\b': 'Feminino',
                                r'\bPre\s+mium\b': 'Premium',
                                r'\bCas\s+ual\b': 'Casual',
                                r'\bCam\s+pus\b': 'Campus',
                                r'\bTê\s+nis\b': 'Tênis',
                                r'\bSka\s+te\b': 'Skate',
                                r'\bCon\s+fortável\b': 'Confortável',
                                r'\bLan\s+çamento\b': 'Lançamento',
                                r'\bDia\s+a\s+Dia\b': 'Dia a Dia',
                                r'\bSu\s+per\b': 'Super',
                                r'\bLi\s+nha\b': 'Linha',
                                r'\bMo\s+retto\b': 'Moretto'
                            }
                            
                            texto_corrigido = texto
                            for padrao, correcao in correcoes.items():
                                texto_corrigido = re.sub(padrao, correcao, texto_corrigido, flags=re.IGNORECASE)
                            
                            return texto_corrigido
                        
                        # Limpar e normalizar a descrição
                        conteudo_limpo = re.sub(r'\s+', ' ', conteudo.strip())  # Remove espaços extras
                        conteudo_limpo = re.sub(r'[■□▪▫]', '', conteudo_limpo)  # Remove caracteres especiais
                        conteudo_limpo = corrigir_palavras_cortadas(conteudo_limpo)  # Corrige palavras cortadas
                        
                        # Incluir código junto com a descrição
                        produto_completo = f"Código: {codigo}\n{conteudo_limpo}"
                        
                        # Função para quebrar texto sem cortar palavras
                        def quebrar_texto_inteligente(texto, largura_maxima=112):  # Alterado para 112 caracteres
                            linhas = []
                            palavras = texto.split()
                            linha_atual = ""
                            
                            for palavra in palavras:
                                # Se adicionar a palavra não ultrapassar o limite
                                if len(linha_atual + " " + palavra) <= largura_maxima:
                                    if linha_atual:
                                        linha_atual += " " + palavra
                                    else:
                                        linha_atual = palavra
                                else:
                                    # Se a linha atual não está vazia, adiciona às linhas
                                    if linha_atual:
                                        linhas.append(linha_atual)
                                        linha_atual = palavra
                                    else:
                                        # Palavra muito longa, força quebra
                                        linhas.append(palavra[:largura_maxima])
                                        linha_atual = palavra[largura_maxima:]
                            
                            # Adiciona a última linha se não estiver vazia (FORA do loop)
                            if linha_atual:
                                linhas.append(linha_atual)
                            
                            return "\n".join(linhas)
                        
                        # Quebrar conteúdo em linhas para melhor formatação
                        produto_quebrado = quebrar_texto_inteligente(produto_completo, 112)  # Usar largura de 112
                        table_data.append([produto_quebrado, quantidade])
                        itens_validos += 1
                        print(f"[GERAÇÃO] Item formatado: {len(produto_quebrado.split())} linhas")
                    else:
                        print(f"[GERAÇÃO] Item inválido ignorado na DANFE {i+1}: código='{codigo}', conteúdo='{conteudo}'")
                else:
                    print(f"[GERAÇÃO] Item com formato inválido ignorado na DANFE {i+1}: {item}")

            if not table_data:
                print(f"[GERAÇÃO] ERRO: Nenhum item válido encontrado na DANFE {i+1}")
                continue
                
            print(f"[GERAÇÃO] {itens_validos} itens válidos processados para DANFE {i+1}")

            # Criar tabela
            table_width = width * 0.98
            col_widths = [table_width * 0.85, table_width * 0.15]  # Melhor distribuição: 85% descrição, 15% quantidade
            table = Table(table_data, colWidths=col_widths)

            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),  # Centralizar quantidade
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 12),  # Fonte 12pt para máximo aproveitamento
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),  # Padding mínimo inferior (2pt)
                ('TOPPADDING', (0, 0), (-1, -1), 2),    # Padding mínimo superior (2pt)
                ('LEFTPADDING', (0, 0), (-1, -1), 3),   # Padding mínimo esquerdo (3pt)
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),  # Padding mínimo direito (3pt)
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('NOSPLIT', (0, 0), (-1, -1)),
                ('WORDWRAP', (0, 0), (-1, -1)),
                ('ROWHEIGHT', (0, 0), (-1, -1), 100),
                ('LEADING', (0, 0), (-1, -1), 14)  # Espaçamento mínimo entre linhas (14pt)
            ])
            table.setStyle(style)

            # Buscar imagem com lógica melhorada
            img_height = 0
            pagina_com_imagem = None
            
            # Buscar em múltiplas páginas relacionadas à DANFE atual
            paginas_para_buscar = [i * 2, i * 2 + 1]
            if i * 2 + 2 < doc.page_count:
                paginas_para_buscar.append(i * 2 + 2)
                
            for pagina_num in paginas_para_buscar:
                if pagina_num < doc.page_count:
                    page = doc.load_page(pagina_num)
                    if page.get_images():
                        text = page.get_text("text")
                        # Aceitar páginas com imagem que não sejam DANFE principal
                        if not ("DANFE" in text.upper() and "CHAVE DE ACESSO" in text.upper()):
                            pagina_com_imagem = page
                            print(f"[GERAÇÃO] Imagem encontrada na página {pagina_num + 1} para DANFE {i+1}")
                            break

            # Processar imagem se encontrada
            if pagina_com_imagem:
                try:
                    pix = pagina_com_imagem.get_pixmap(alpha=False, dpi=200)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    img_bytes = io.BytesIO()
                    img.save(img_bytes, format='JPEG')
                    img_bytes.seek(0)
                    img_reader = ImageReader(img_bytes)

                    margem_direita = 1.5 * cm
                    margem_inferior = 0.1 * cm
                    img_width = width - margem_direita
                    img_height = height - margem_inferior - table.wrap(0, width)[1] - 2 * cm

                    c.drawImage(img_reader, 0, height - img_height, width=img_width, height=img_height, preserveAspectRatio=True, anchor='nw')
                    print(f"[GERAÇÃO] Imagem adicionada com sucesso para DANFE {i+1}")
                except Exception as e:
                    print(f"[GERAÇÃO] ERRO ao processar imagem para DANFE {i+1}: {str(e)}")
                    img_height = 0
            else:
                print(f"[GERAÇÃO] Nenhuma imagem encontrada para DANFE {i+1}")

            # Posicionar tabela
            try:
                if len(table_data) > 4:
                    c.showPage()
                    table.wrapOn(c, width, height)
                    table_y = height - table.wrap(0, width)[1] - 1 * cm
                    table.drawOn(c, 0.1 * cm, table_y)
                else:
                    table.wrapOn(c, width, height)
                    table_y = height - img_height - table.wrap(0, width)[1] - 1 * cm
                    table.drawOn(c, 0.1 * cm, table_y)
                    
                print(f"[GERAÇÃO] Tabela posicionada com sucesso para DANFE {i+1}")
            except Exception as e:
                print(f"[GERAÇÃO] ERRO ao posicionar tabela para DANFE {i+1}: {str(e)}")
                continue

            # Adicionar contador de páginas (P1, P2, P3, etc.)
            try:
                contador_texto = f"P{paginas_geradas + 1}"
                c.setFont("Helvetica-Bold", 14)
                c.setFillColor(colors.black)
                # Posicionar no final da página, canto inferior direito
                contador_x = width - 2 * cm
                contador_y = 0.3 * cm  # Bem próximo à borda inferior
                c.drawString(contador_x, contador_y, contador_texto)
                print(f"[GERAÇÃO] Contador de página adicionado: {contador_texto}")
            except Exception as e:
                print(f"[GERAÇÃO] ERRO ao adicionar contador de página para DANFE {i+1}: {str(e)}")

            c.showPage()
            paginas_geradas += 1
            print(f"[GERAÇÃO] DANFE {i+1} concluída com sucesso")
            
        except Exception as e:
            print(f"[GERAÇÃO] ERRO INESPERADO ao processar DANFE {i+1}: {str(e)}")
            continue

    # Finalizar PDF
    if paginas_geradas > 0:
        c.save()
        print(f"[GERAÇÃO] PDF salvo com sucesso: {paginas_geradas} páginas geradas")
    else:
        print("[GERAÇÃO] ERRO: Nenhuma página foi gerada com sucesso")
        return False
        
    doc.close()
    fim = time.time()
    print(f"[GERAÇÃO] Concluída em {fim - inicio:.2f}s - {paginas_geradas}/{len(data)} DANFEs processadas")
    return True

if __name__ == '__main__':
    # Registrar limpeza de arquivos temporários quando o servidor for encerrado
    atexit.register(cleanup_temp_files)
    app.run(debug=True, port=5000)
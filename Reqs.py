import win32com.client
import pandas as pd
import time
import pythoncom
import os
from datetime import datetime

# Caminho padrão do arquivo Excel (parametrizável)
ARQUIVO_PADRAO = r"\\br03file\pcoudir\Operacoes\10. Planning Raw Material\Gerenciamento de materiais\Atividades diarias\Robo Atualizacao de Datas Fornecedores\Alterar_pedidos.xlsx"
LOG_PASTA = r"\\br03file\pcoudir\Operacoes\10. Planning Raw Material\Gerenciamento de materiais\Atividades diarias\Robo Atualizacao de Datas Fornecedores\Log"

# Variável global para armazenar o caminho do arquivo Excel selecionado
arquivo_excel = None

# Callbacks para integração com a interface gráfica (opcional)
progress_callback = None
status_callback = None
log_callback = None

def set_callbacks(progress_cb=None, status_cb=None, log_cb=None):
    """Configura callbacks para comunicação com a interface gráfica."""
    global progress_callback, status_callback, log_callback
    progress_callback = progress_cb
    status_callback = status_cb
    log_callback = log_cb

def set_arquivo_excel(caminho):
    """Define o caminho do arquivo Excel a ser processado."""
    global arquivo_excel
    arquivo_excel = caminho

def emit_progress(value):
    """Emite progresso para a interface."""
    if progress_callback:
        progress_callback(value)

def emit_status(message, status_type="info"):
    """Emite status para a interface."""
    if status_callback:
        status_callback(message, status_type)

def emit_log(msg):
    """Emite mensagens de log para a interface."""
    if log_callback:
        log_callback(msg)
    else:
        print(msg)  # Fallback para console se não houver callback

# ------------------- Integração SAP -------------------
def conectar_sap():
    """Estabelece conexão com o SAP GUI via Scripting."""
    emit_log("🔄 Conectando ao SAP...")
    emit_status("Conectando ao SAP", "running")
    
    pythoncom.CoInitialize()
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        
        emit_log("✅ Conexão SAP estabelecida")
        emit_status("Conectado ao SAP", "success")
        return session
        
    except Exception as e:
        emit_log(f"❌ Erro ao conectar SAP: {e}")
        emit_status("Erro na conexão SAP", "error")
        return None

def esperar_objeto(session, objeto_id, tentativas=5, intervalo=0.5):
    """Aguarda a disponibilidade de um objeto no SAP GUI."""
    for t in range(tentativas):
        try:
            return session.findById(objeto_id)
        except:
            emit_log(f"🔄 Aguardando objeto {objeto_id}... ({t+1}/{tentativas})")
            time.sleep(intervalo)
    raise Exception(f"Objeto {objeto_id} não encontrado após {tentativas} tentativas")

def limpar_tela_sap(session):
    """Fecha pop-ups residuais e reinicia o contexto da transação."""
    try:
        # Fechar possíveis popups
        for i in range(5):
            try:
                session.findById("wnd[1]/tbar[0]/btn[12]").press()
            except:
                break
        
        # Resetar transação
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        
    except Exception as e:
        emit_log(f"⚠️ Erro ao limpar tela SAP: {e}")

def verificar_erro_sap(session):
    """Verifica mensagens de erro na barra de status do SAP."""
    try:
        sbar = session.findById("wnd[0]/sbar")
        if sbar and sbar.MessageType in ['E','A']:
            return f"Erro SAP: {sbar.Text.strip()}"
        return None
    except:
        return None

# ------------------- Utilitários -------------------
def formatar_data(valor):
    """Formata datas no padrão SAP (DD.MM.YYYY)."""
    if pd.isna(valor):
        return ""
    
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%d.%m.%Y")
    
    try:
        data = pd.to_datetime(str(valor), dayfirst=True, errors="coerce")
        if pd.notna(data):
            return data.strftime("%d.%m.%Y")
    except:
        pass
    
    return str(valor)

def validar_colunas_excel(df):
    """Valida a presença das colunas obrigatórias na planilha Excel."""
    colunas_necessarias = ["Requisicao", "NovaQtd", "NovaData"]
    colunas_existentes = df.columns.tolist()
    
    # Verificar se todas as colunas necessárias existem
    colunas_faltantes = [col for col in colunas_necessarias if col not in colunas_existentes]
    
    if colunas_faltantes:
        raise ValueError(f"Colunas obrigatórias não encontradas: {', '.join(colunas_faltantes)}\n"
                        f"Colunas disponíveis: {', '.join(colunas_existentes)}")
    
    return True

# ------------------- Processamento -------------------
def atualizar_requisicao(session, req, qtd, data, tentativas=2):
    """Atualiza uma requisição no SAP (ME52N) com quantidade e datas informadas."""
    data_atual = datetime.now().strftime("%d.%m.%Y")
    
    for t in range(1, tentativas+1):
        try:
            emit_log(f"📋 Abrindo requisição {req} (tentativa {t}/{tentativas})")
            limpar_tela_sap(session)

            # Abrir ME52N (Alterar Requisição de Compra)
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nme52n"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)

            # Buscar requisição específica
            session.findById("wnd[0]/tbar[1]/btn[17]").press()
            session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").text = str(req)
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1)

            # Garantir navegação para a aba "Datas de Entrega"
            aba_datas = "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" \
            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" \
            "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT4"
            emit_log("📑 Acessando aba Datas de Entrega...")
            session.findById(aba_datas).select()
            time.sleep(0.5)
            
            # Verificar erros SAP
            erro = verificar_erro_sap(session)
            if erro:
                raise Exception(erro)

            # Alterar quantidade
            campo_qtd = "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT4/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/txtMEREQ3321-MENGE"
            emit_log(f"✏️ Alterando quantidade para {qtd}")
            cell_qtd = esperar_objeto(session, campo_qtd)
            cell_qtd.text = str(int(qtd))
            cell_qtd.setFocus()
            cell_qtd.caretPosition = len(str(int(qtd)))
            session.findById("wnd[0]").sendVKey(0)

            # Alterar data de remessa
            campo_data = "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT4/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/ctxtMEREQ3321-EEIND"
            emit_log(f"📅 Alterando data para {data}")
            cell_data = esperar_objeto(session, campo_data)
            cell_data.text = data
            cell_data.caretPosition = 2
            session.findById("wnd[0]").sendVKey(0)

            
        # Definir data de liberação para a data corrente
            campo_data = "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/" \
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/" \
                         "subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT4/" \
                         "ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/ctxtMEREQ3321-FRGDT"
            cell_data = esperar_objeto(session, campo_data)
            cell_data.text = data_atual
            cell_data.setFocus()
            cell_data.caretPosition = 2
            session.findById("wnd[0]").sendVKey(0)

            # Marcar como Fixado
            chk_fixado = "wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT4/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:3321/chkMEREQ3321-FIXKZ"
            emit_log("📌 Marcando como fixado")
            session.findById(chk_fixado).selected = True
            session.findById(chk_fixado).setFocus()

            # Salvar alterações
            emit_log("💾 Salvando alterações")
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            time.sleep(1)
            
            # Verificar se salvou com sucesso
            erro_save = verificar_erro_sap(session)
            if erro_save:
                raise Exception(erro_save)
            
            emit_log(f"✅ Requisição {req} atualizada com sucesso!")
            return "SUCESSO", f"Requisição {req} atualizada"

        except Exception as e:
            emit_log(f"❌ Erro na tentativa {t}: {e}")
            if t < tentativas:
                emit_log("🔄 Tentando novamente em 2 segundos...")
                time.sleep(2)
            else:
                return "ERRO", f"Falha ao atualizar requisição {req}: {e}"

def salvar_logs(resultados):
    """Salva o resultado da execução em arquivo CSV."""
    try:
        if not os.path.exists(LOG_PASTA):
            os.makedirs(LOG_PASTA)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo = os.path.join(LOG_PASTA, f"log_requisicoes_{timestamp}.csv")
        
        df_log = pd.DataFrame(resultados)
        df_log.to_csv(arquivo, sep=";", index=False, encoding="utf-8-sig")
        
        emit_log(f"📝 Log salvo em {arquivo}")
        return arquivo
        
    except Exception as e:
        emit_log(f"⚠️ Erro ao salvar logs: {e}")
        return None

# ------------------- Execução principal -------------------
def main():
    """Ponto de entrada do robô SAP para atualização de requisições."""
    global arquivo_excel
    
    emit_log("🚀 Iniciando processo de atualização de requisições...")
    emit_status("Carregando dados", "running")
    emit_progress(0)

    # Determinar qual arquivo usar
    arquivo_usar = arquivo_excel if arquivo_excel else ARQUIVO_PADRAO
    
    # Verificar se arquivo Excel existe
    if not os.path.exists(arquivo_usar):
        error_msg = f"❌ Arquivo {arquivo_usar} não encontrado"
        emit_log(error_msg)
        emit_status("Arquivo não encontrado", "error")
        raise FileNotFoundError(error_msg)

    try:
        # Ler dados do Excel
        emit_log(f"📊 Carregando dados do Excel: {os.path.basename(arquivo_usar)}")
        df = pd.read_excel(arquivo_usar, sheet_name="Req")
        
        # Validar colunas
        validar_colunas_excel(df)
        
        df = df[["Requisicao", "NovaQtd", "NovaData"]].dropna()
        
        total_registros = len(df)
        emit_log(f"✅ {total_registros} registros carregados para processamento")
        emit_progress(10)
        
    except ValueError as ve:
        error_msg = f"❌ Erro de validação: {ve}"
        emit_log(error_msg)
        emit_status("Erro de validação", "error")
        raise Exception(error_msg)
    except Exception as e:
        error_msg = f"❌ Erro ao ler Excel: {e}"
        emit_log(error_msg)
        emit_status("Erro no Excel", "error")
        raise Exception(error_msg)

    if df.empty:
        emit_log("⚠️ Nenhuma requisição válida encontrada para processar")
        emit_status("Nenhum dado para processar", "warning")
        return

    # Conectar ao SAP
    session = conectar_sap()
    if not session:
        error_msg = "❌ Não foi possível conectar ao SAP"
        emit_log(error_msg)
        emit_status("Falha na conexão SAP", "error")
        raise ConnectionError(error_msg)

    emit_progress(20)
    
    # Processar cada requisição
    resultados = []
    sucessos = 0
    erros = 0
    
    for idx, row in df.iterrows():
        req = row["Requisicao"]
        qtd = row["NovaQtd"]
        nova_data = formatar_data(row["NovaData"])
        
        # Calcular progresso (20% a 90%)
        progresso = 20 + int((idx / total_registros) * 70)
        emit_progress(progresso)
        emit_status(f"Processando requisição {req}", "running")
        
        try:
            # Processar requisição
            status, msg = atualizar_requisicao(session, req, qtd, nova_data)
            
            if status == "SUCESSO":
                sucessos += 1
            else:
                erros += 1
            
            resultados.append({
                "Requisicao": req,
                "NovaQtd": qtd,
                "NovaData": nova_data,
                "Status": status,
                "Mensagem": msg,
                "Data_Execucao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            
        except Exception as e:
            erros += 1
            error_msg = f"Erro crítico na requisição {req}: {e}"
            emit_log(f"❌ {error_msg}")
            
            resultados.append({
                "Requisicao": req,
                "NovaQtd": qtd,
                "NovaData": nova_data,
                "Status": "ERRO",
                "Mensagem": error_msg,
                "Data_Execucao": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

        emit_log("-" * 60)
        
    # Finalizar processamento
    emit_progress(95)
    emit_log("💾 Salvando logs de execução...")
    
    # Salvar resultados
    salvar_logs(resultados)
    emit_progress(100)
    
    # Status final
    if erros == 0:
        emit_log(f"🎉 Processo concluído com sucesso! {sucessos} requisições processadas")
        emit_status("Concluído com sucesso", "success")
    elif sucessos > 0:
        emit_log(f"⚠️ Processo finalizado com ressalvas: {sucessos} sucessos, {erros} erros")
        emit_status("Finalizado com ressalvas", "warning")
    else:
        emit_log(f"❌ Processo finalizado com erros: {erros} falhas")
        emit_status("Finalizado com erros", "error")
    
    emit_log(f"📊 Resumo: {sucessos} sucessos, {erros} erros de {total_registros} registros")

if __name__ == "__main__":
    # Execução
    try:
        main()
    except Exception as e:
        print(f"Erro na execução: {e}")
        input("Pressione Enter para sair...")

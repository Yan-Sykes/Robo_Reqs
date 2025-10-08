# SAP Robot – Atualização de Requisições (GUI)

Aplicação em Python com interface gráfica (PyQt6) para automatizar a atualização de requisições no SAP (ME52N) a partir de um arquivo Excel. O robô conecta-se ao SAP GUI via COM (pywin32), lê os dados do Excel (pandas) e registra logs de execução.


## Principais recursos
- Interface moderna para iniciar/parar a execução e acompanhar logs e progresso
- Seleção de arquivo Excel customizado ou uso de caminho padrão pré-configurado
- Validação de colunas obrigatórias no Excel
- Atualização automatizada de quantidade, datas (remessa e liberação) e marcação “Fixado”
- Logs detalhados em CSV e logs locais por execução


## Requisitos
- Sistema operacional: Windows (obrigatório, devido à automação via SAP GUI Scripting/COM)
- SAP GUI for Windows instalado e com SAP GUI Scripting habilitado
  - Servidor/Perfil: permitir scripting
  - Cliente (SAP Logon): Opções > Accessibility & Scripting > Enable scripting
- Sessão do SAP aberta e autenticada antes de iniciar o robô
- Python 3.8+ (recomendado 3.10+)

Dependências de Python (instaladas via requirements.txt):
- pywin32 (integração COM com SAP GUI)
- pandas (leitura e tratamento do Excel)
- openpyxl (engine para .xlsx)
- xlrd (leitura de .xls, quando aplicável)
- PyQt6 (interface gráfica)
- requests (carregamento de ícones/logos via URL)


## Instalação
1) Instale o Python 3 para Windows e garanta o pip disponível no PATH.
2) No Prompt de Comando, navegue até a pasta do projeto e instale as dependências:

   ```bat
   pip install -r requirements.txt
   ```

Se ocorrer erro relacionado ao pywin32/COM, rode o pós-installer (algumas versões):

```bat
python -m pywin32_postinstall -install
```


## Execução (GUI)
1) Abra e mantenha uma sessão do SAP GUI conectada (logada) no ambiente desejado.
2) No Prompt de Comando, execute:

```bat
python RoboSAP_GUI.py
```

3) Na interface:
- Opcional: clique em “Procurar” para escolher um Excel específico. Sem seleção, será usado o caminho padrão definido no código.
- Clique em “Executar”. Acompanhe o progresso e os logs.
- Ao finalizar, será exibido um resumo e os logs serão salvos automaticamente.


## Estrutura do Excel
A planilha a ser processada deve conter:
- Aba: `Req`
- Colunas obrigatórias (nomes exatos):
  - `Requisicao`
  - `NovaQtd`
  - `NovaData`

Notas:
- `NovaData` pode estar em formato de data reconhecível; o robô converte para `DD.MM.YYYY` (formato SAP).
- Linhas vazias ou com campos críticos ausentes são ignoradas.


## Caminhos padrão e Logs
- Caminho padrão do Excel (editar em Reqs.py se necessário):
  - Constante `ARQUIVO_PADRAO`
- Pasta de logs CSV (editar em Reqs.py):
  - Constante `LOG_PASTA`
- Logs locais da interface (GUI):
  - `%USERPROFILE%\SAP_Robo_Logs` (um .log por execução)
- Configuração persistente da interface (último arquivo selecionado):
  - `%USERPROFILE%\SAP_Robo_Logs\config.json`


## Personalização
Edite o arquivo `Reqs.py` para ajustar:
- `ARQUIVO_PADRAO`: caminho do Excel padrão na rede ou local
- `LOG_PASTA`: pasta onde o CSV de resultados será salvo
- Parâmetros de tentativas/intervalos de automação, se necessário


## Build do executável (opcional)
Já existe um arquivo `RoboSAP_GUI.spec` no projeto. Para gerar um executável com o PyInstaller:

```bat
pip install pyinstaller
pyinstaller --clean --noconfirm RoboSAP_GUI.spec
```

O executável será gerado em `dist/` (por exemplo, `dist/RoboSAP_GUI/RoboSAP_GUI.exe`).


## Solução de problemas
- SAP GUI Scripting desabilitado:
  - Habilite no servidor/perfil e no cliente (SAP Logon) em Accessibility & Scripting
- “Arquivo não encontrado”:
  - Verifique o caminho definido em `ARQUIVO_PADRAO` ou selecione manualmente o Excel na GUI
  - Confirme se o caminho de rede está acessível e você possui permissões
- Erros COM/pywin32:
  - Reinstale `pywin32` e rode `python -m pywin32_postinstall -install`
  - Execute o prompt “Como administrador” se necessário
- Excel inválido / colunas ausentes:
  - Garanta a aba `Req` e as colunas `Requisicao`, `NovaQtd`, `NovaData`
- Sessão do SAP não encontrada:
  - Abra o SAP GUI e autentique-se no ambiente antes de clicar em “Executar”


## Licença / Uso
Projeto direcionado a uso interno. Ajuste conforme as políticas da sua organização.

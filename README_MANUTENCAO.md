# README de manutenção – GAC Printer

Guia rápido para quem vai dar manutenção no projeto. O código está concentrado em `main.py`; abaixo estão os pontos importantes para entender o fluxo atual e sugestões de como quebrar em módulos.

## Visão geral do fluxo
- **Login/assinatura (Supabase REST):** `show_login_dialog()` abre o diálogo, valida e-mail/senha numérica em `supabase_check_subscription_rest()` e grava `_auth_session`. Rechecagens periódicas usam `schedule_subscription_recheck()`/`periodic_recheck()`.
- **Seleção de impressora:** `choose_printer_dialog()` chama `set_selected_printer()` e atualiza a default do Windows para que PDFs sigam a mesma impressora.
- **Monitoramento de pasta:** `select_folder()` escolhe o diretório e inicia uma thread (`monitor_etiquetas_shopee()`) que varre a pasta a cada 3s em busca de `.zip`, `.txt` (ZPL) ou `.pdf`, processando o primeiro encontrado.
- **Processamento de arquivos:** `process_zip()` extrai e prioriza TXT→PDF; `process_txt()` envia ZPL em RAW via `send_to_printer()`; `process_pdf()` chama `print_pdf()` e apaga o arquivo.
- **RAR Amazon (opcional):** quando a opção "Imprimir Amazon (.rar)" está marcada na UI, `process_amazon_rar()` extrai o `.rar`, ignora o PDF e imprime somente o `.zpl`.
- **Impressão de PDF:** `print_pdf()` tenta, na ordem configurada, os métodos: 1) `ShellExecute`, 2) `PowerShell`, 3) automação com `pyautogui` (Ctrl+P, Alt+F4 opcional), 4) `SumatraPDF` silencioso (`print_pdf_method_4_silent()`).
- **UI Tkinter:** montada em `main()`, com controles para fechar janelas pós-impressão, clique extra, método de PDF, status de assinatura (exibe data + dias restantes e fica vermelha a <7 dias), seleção de impressora e botões de monitoramento/teste.

## Dependências e ambiente
- **SO:** Windows (usa `win32print`, `win32api`, `winreg`, automação de janelas).
- **Python:** 3.x com `pywin32`, `pyautogui`, `tkinter` (builtin), `urllib`, `zipfile`, `subprocess`, `shutil`, `tempfile`.
- **SumatraPDF (opcional):** buscado em `Program Files`, `%LOCALAPPDATA%`, diretório atual ou ao lado do binário gerado (PyInstaller). `print_pdf_method_4_silent` usa `-print-to` ou `-print-to-default`.
- **RAR (Amazon):** precisa de `rarfile`+`unrar` ou `UnRAR.exe`/`WinRAR` disponível no PATH para extrair. Sem isso, a opção Amazon logará erro e seguirá adiante.
- **Variáveis de ambiente:**
  - `SUPABASE_URL` (default já definido no código)
  - `SUPABASE_ANON_KEY` (default já definido)
  - `SUPABASE_SUBS_TABLE` (default: `Clientes_Printer`)
  - `SUBS_RECHECK_MINUTES` (intervalo em minutos; default 60 na UI, 720 na rechecagem agendada)

## Como rodar em desenvolvimento
1) Instale dependências: `pip install pywin32 pyautogui pyinstaller` (Tkinter vem com Python no Windows).  
2) Garanta que as variáveis do Supabase estejam exportadas se for validar assinatura.  
3) Execute `python main.py`. A janela de login aparece antes da UI principal.  
4) Escolha a impressora, selecione a pasta e deixe a thread de monitoramento rodando.

## Build/distribuição
- Usa PyInstaller com o spec `main.spec` (ícone e inclusão opcional do Sumatra portátil).
- Scripts de apoio: `build.bat` ou `build.ps1` (ambos aceitam `--clean`/`-Clean` para limpar `build/` e `dist/`).
- Saída esperada: `dist/main.exe` ou `dist/main/main.exe` dependendo da versão do PyInstaller.

## Pontos de atenção atuais
- Há chamadas para `supabase_check_subscription()` (linhas ~289 e ~367) que não existem; só há a variante REST (`supabase_check_subscription_rest`). Corrigir antes de dividir em módulos.
- Estado global abundante: `selected_printer_name`, `_auth_session`, `monitorando`, `fechar_telas`, `clicar_apos_fechar`, `Método_impressão_pdf`. Facilita regressões se manipulado fora de ordem.
- Automação com `pyautogui` depende de foco de janela; evitar rodar headless. Flags na UI controlam fechamento automático e clique residual.
- O loop de monitoramento é um `while` com `time.sleep(3)` sem cancelamento mais sofisticado; travamentos silenciosos podem ocorrer em exceções não tratadas.
- Leva PDFs/TXTs à impressora e apaga os arquivos depois; cuidado em ambiente de teste, use cópias.

## Sugestão de modularização (passos mínimos)
1) **config.py:** carregar/env validar `SUPABASE_*`, `SUBS_RECHECK_MINUTES`, métodos padrão de impressão.
2) **auth.py:** `_http_request`, `supabase_signup_rest`, `supabase_check_subscription_rest`, `supabase_check_status_only`, helpers de data (`_format_expire_date_brt`), sessão.
3) **printing.py:** `send_to_printer`, métodos 1–4, descoberta de Sumatra, flags `fechar_telas`/`clicar_apos_fechar`, seleção de impressora.
4) **processing.py:** `process_zip/txt/pdf`, tratamento de diretórios temporários e remoção segura.
5) **monitor.py:** thread de varredura, controle start/stop.
6) **ui/** (pacote): `login_dialog`, `printer_dialog`, `main_window`, ligação dos callbacks ao estado compartilhado.
7) **tests/manual:** scripts mínimos para testar impressão fake (ex.: gerar PDF/TXT temporário e simular monitoramento sem deletar originais).

## Checklist rápido ao mexer
- Atualize os defaults/envs do Supabase em um único lugar (após modularizar, em `config.py`).
- Reforce logging (stdout hoje é a única fonte). Planeje um logger que escreva em arquivo no diretório do executável.
- Teste todos os métodos de PDF após mudanças de foco/threads; `pyautogui` é sensível a timeouts.
- Valide se a impressora padrão está sendo setada corretamente quando o usuário troca na UI.
- Reempacote com PyInstaller e confirme se o Sumatra portátil entrou em `dist/` quando presente.

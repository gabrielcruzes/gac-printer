import os
import zipfile
import win32print
import win32api
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import time
import pyautogui
import threading
import subprocess
import shutil
import tempfile
import json
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime, timezone
from datetime import timedelta
import platform
import sys

# -------------------- Config / Estado global --------------------
# Impressora Selecionada
selected_printer_name = None

# Autenticação / Supabase (config via variáveis de ambiente)
SUPABASE_URL = os.getenv("SUPABASE_URL", "https://hteijwzysrzupmgvgeao.supabase.co").rstrip("/")
SUPABASE_ANON_KEY = os.getenv("SUPABASE_ANON_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh0ZWlqd3p5c3J6dXBtZ3ZnZWFvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDE3MzI1ODEsImV4cCI6MjA1NzMwODU4MX0.emRVNWm9viTLpBYX8RBGiWI6bsOpiO1FNqRXeiVW-ZY")
SUPABASE_SUBS_TABLE = os.getenv("SUPABASE_SUBS_TABLE", "Clientes_Printer")
_auth_session = {
    "email": None,
    "expires_at": None,
}

# ---- Utilidades HTTP (sem dependências externas) ----
def _http_request(method, url, headers=None, data=None, timeout=15):
    req = urllib.request.Request(url, method=method)
    if headers:
        for k, v in headers.items():
            req.add_header(k, v)
    if data is not None:
        if isinstance(data, (dict, list)):
            data = json.dumps(data).encode("utf-8")
            req.add_header("Content-Type", "application/json")
        elif isinstance(data, str):
            data = data.encode("utf-8")
    try:
        with urllib.request.urlopen(req, data=data, timeout=timeout) as resp:
            body = resp.read()
            try:
                return resp.getcode(), json.loads(body.decode("utf-8"))
            except Exception:
                return resp.getcode(), body.decode("utf-8")
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except Exception:
            return e.code, e.read().decode("utf-8")
    except Exception as e:
        return None, str(e)

# ---- Supabase: login e checagem de assinatura ----
def supabase_login(email, password):
    # REST apenas: login via Auth desativado
    return False, "Login via Auth desativado (REST apenas)", None

def supabase_signup_rest(email):
    """Cria um registro na tabela Clientes_Printer via REST (status=false).
    Retorna (ok, erro, data)."""
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return False, "SUPABASE_URL/ANON_KEY não configurados", None
    url = f"{SUPABASE_URL}/rest/v1/{SUPABASE_SUBS_TABLE}"
    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }
    payload = {"email": email, "status": False, "expires_at": None}
    status, data = _http_request("POST", url, headers=headers, data=payload)
    try:
        print(f"Supabase signup (REST) HTTP status: {status}")
        print(f"Supabase signup (REST) response: {data}")
    except Exception:
        pass
    if status in (200, 201):
        return True, None, data
    else:
        txt = json.dumps(data) if isinstance(data, dict) else str(data)
        if txt and ('duplicate' in txt.lower() or 'unique' in txt.lower()):
            return True, None, data
        err = None
        if isinstance(data, dict):
            err = data.get('message') or data.get('hint') or data.get('details')
        return False, err or f"Falha ao criar registro (HTTP {status})", data

def supabase_check_subscription_rest(email, password):
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return False, "SUPABASE_URL/ANON_KEY nao configurados", None
    # Clientes_Printer: id(int8), email(text), status(bool), expires_at(date), senha(int8)
    rest = f"{SUPABASE_URL}/rest/v1/{SUPABASE_SUBS_TABLE}?email=eq.{urllib.parse.quote(email)}&select=status,expires_at,senha&limit=1"
    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Accept": "application/json",
        "Prefer": "count=exact",
    }
    status, data = _http_request("GET", rest, headers=headers)
    # Caso e-mail não exista na base
    if status == 200 and isinstance(data, list) and len(data) == 0:
        return False, "login inválido", None
    if status == 200 and isinstance(data, list) and len(data) > 0:
        row = data[0]
        active = bool(row.get("status", False))
        expires_at = row.get("expires_at")
        # valida senha numérica (int8)
        try:
            provided = int(str(password).strip())
        except Exception:
            # Mensagem solicitada: senha invalida
            return False, "senha invalida", expires_at
        stored = row.get("senha")
        try:
            stored_int = int(stored) if stored is not None else None
        except Exception:
            stored_int = None
        # Senha incorreta ou não cadastrada
        if stored_int is None or provided != stored_int:
            return False, "senha invalida", expires_at
        if not active:
            return False, "Assinatura inativa", expires_at
        if expires_at:
            try:
                # Timezone do Brasil (São Paulo) sem horário de verão ativo (UTC-03)
                tz_brt = timezone(timedelta(hours=-3))
                now_brt = datetime.now(tz_brt)

                exp_raw = str(expires_at)
                if 'T' not in exp_raw and len(exp_raw) == 10:
                    # Campo date (YYYY-MM-DD). Tratar como inclusivo até 23:59:59.999999 no BRT
                    y, m, d = map(int, exp_raw.split('-'))
                    exp = datetime(y, m, d, 23, 59, 59, 999999, tzinfo=tz_brt)
                else:
                    # datetime ISO
                    exp = datetime.fromisoformat(exp_raw.replace('Z', '+00:00'))
                    if exp.tzinfo is None:
                        # Considera como BRT se vier sem tz
                        exp = exp.replace(tzinfo=tz_brt)
                    else:
                        # Converte para BRT para comparação
                        exp = exp.astimezone(tz_brt)

                if exp < now_brt:
                    # Mensagem solicitada: assinatura expirada
                    return False, "assinatura expirada", expires_at
            except Exception:
                pass
        return True, None, expires_at
    else:
        err = data if isinstance(data, str) else json.dumps(data)
        return False, f"Falha ao checar assinatura (HTTP {status}): {err}", None

def supabase_check_status_only(email):
    """Checa apenas o campo status no Supabase para o e-mail informado.
    Retorna (ok, err, status_bool). ok=True apenas quando a consulta deu certo; o
    fechamento do app deve ocorrer somente quando status_bool for False.
    """
    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        return False, "SUPABASE_URL/ANON_KEY nao configurados", None
    url = f"{SUPABASE_URL}/rest/v1/{SUPABASE_SUBS_TABLE}?email=eq.{urllib.parse.quote(email)}&select=status&limit=1"
    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Accept": "application/json",
        "Prefer": "count=exact",
    }
    status_code, data = _http_request("GET", url, headers=headers)
    if status_code == 200 and isinstance(data, list) and len(data) > 0:
        try:
            st = bool(data[0].get("status", False))
        except Exception:
            st = False
        return True, None, st
    else:
        err = data if isinstance(data, str) else json.dumps(data)
        return False, f"Falha ao checar status (HTTP {status_code}): {err}", None

## n8n removido a pedido; nenhum webhook necessário

# ---- Impressoras: listagem e seleção ----
def list_installed_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = win32print.EnumPrinters(flags)
    names = [p[2] for p in printers if len(p) >= 3]
    seen = set()
    ordered = []
    for n in names:
        if n not in seen:
            seen.add(n)
            ordered.append(n)
    return ordered

def set_selected_printer(name):
    global selected_printer_name
    selected_printer_name = name
    try:
        # Define como padrão do Windows para que impressão de PDF também use
        win32print.SetDefaultPrinter(selected_printer_name)
    except Exception:
        pass
    print(f"Impressora Selecionada: {selected_printer_name}")

def choose_printer_dialog(root):
    top = tk.Toplevel(root)
    top.title("Selecione a impressora")
    top.grab_set()
    top.geometry("520x180")
    tk.Label(top, text="Escolha uma impressora para usar:", font=("Arial", 11)).pack(pady=10)
    printers = list_installed_printers()
    try:
        current_default = win32print.GetDefaultPrinter()
    except Exception:
        current_default = ""
    initial = (selected_printer_name or current_default or (printers[0] if printers else ""))
    var = tk.StringVar(value=initial)
    combo = ttk.Combobox(top, textvariable=var, values=printers, width=60, state="readonly")
    combo.pack(pady=10)
    info = tk.Label(top, text=f"Padrão atual: {current_default or 'desconhecida'}", fg="gray")
    info.pack(pady=2)

    def confirm():
        sel = var.get().strip()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma impressora.")
            return
        set_selected_printer(sel)
        top.destroy()

    tk.Button(top, text="Confirmar", bg="#4CAF50", fg="white", command=confirm).pack(pady=10)
    root.wait_window(top)

# ---- Login (UI) ----
def show_login_dialog(root):
    supabase_configured = bool(SUPABASE_URL and SUPABASE_ANON_KEY)
    dlg = tk.Toplevel(root)
    dlg.title("Login - GAC Printer")
    dlg.geometry("420x300")
    dlg.grab_set()

    tk.Label(dlg, text="E-mail:").pack(pady=(12, 2))
    email_var = tk.StringVar()
    email_entry = tk.Entry(dlg, textvariable=email_var, width=40)
    email_entry.pack()
    email_entry.focus_set()

    tk.Label(dlg, text="Senha:").pack(pady=(8, 2))
    pass_var = tk.StringVar()
    pass_entry = tk.Entry(dlg, textvariable=pass_var, width=40, show="*")
    pass_entry.pack()

    status_lbl = tk.Label(dlg, text="", fg="gray")
    status_lbl.pack(pady=8)

    def do_login():
        email = email_var.get().strip()
        password = pass_var.get().strip()
        if not email or not password:
            messagebox.showwarning("Aviso", "Informe e-mail e senha.")
            return
        status_lbl.config(text="Autenticando...")
        dlg.update_idletasks()
        ok = False
        err = None
        session = None
        if supabase_configured:
            ok, err, session = supabase_login(email, password)
        else:
            if email == "admin" and password == "soeusei12345":
                ok, session = True, {"access_token": "dev", "user": {"id": "dev"}}
            else:
                ok, err = False, "Ambiente sem Supabase configurado. Use admin/admin para testes."

        if ok and session:
            _auth_session["access_token"] = session.get("access_token")
            user = session.get("user", {})
            _auth_session["user_id"] = user.get("id") or session.get("user_id")
            _auth_session["email"] = email

            if supabase_configured:
                ok_sub, sub_err, exp_at = supabase_check_subscription(_auth_session["access_token"], email)
                if not ok_sub:
                    messagebox.showerror("Assinatura", sub_err or "Assinatura inválida")
                    status_lbl.config(text=sub_err or "Assinatura inválida", fg="red")
                    return
                _auth_session["expires_at"] = str(exp_at) if exp_at else None

            dlg.destroy()
        else:
            messagebox.showerror("Login", err or "Falha no login")
            status_lbl.config(text=err or "Falha no login", fg="red")

    # Linha de botões
    btn_row = tk.Frame(dlg)
    btn_row.pack(pady=10)
    def do_login_rest():
        email = email_var.get().strip()
        password = pass_var.get().strip()
        if not email or not password:
            messagebox.showwarning("Aviso", "Informe e-mail e senha.")
            return
        status_lbl.config(text="Validando credenciais...")
        dlg.update_idletasks()
        if supabase_configured:
            ok_sub, sub_err, exp_at = supabase_check_subscription_rest(email, password)
            try:
                print(f"REST check status: ok={ok_sub}, err={sub_err}, exp={exp_at}")
            except Exception:
                pass
            if not ok_sub:
                messagebox.showerror("Assinatura", sub_err or "Assinatura inválida")
                status_lbl.config(text=sub_err or "Assinatura inválida", fg="red")
                return
            _auth_session["email"] = email
            _auth_session["senha"] = password
            _auth_session["expires_at"] = str(exp_at) if exp_at else None
            dlg.destroy()

    tk.Button(btn_row, text="Entrar", bg="#4CAF50", fg="white", width=14, command=do_login_rest).pack(side=tk.LEFT, padx=6)
    # Aguarda o usuario finalizar o login antes de prosseguir
    try:
        dlg.wait_window()
    except Exception:
        pass

    def do_signup():
        messagebox.showinfo("Cadastro", "Criação de conta desativada. Use e-mail e senha já cadastrados.")
        return

def _format_expire_date_brt(exp_raw):
    try:
        tz_brt = timezone(timedelta(hours=-3))
        s = str(exp_raw)
        if 'T' not in s and len(s) == 10:
            y, m, d = map(int, s.split('-'))
            dt = datetime(y, m, d, 23, 59, 59, 999999, tzinfo=tz_brt)
        else:
            dt = datetime.fromisoformat(s.replace('Z', '+00:00'))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=tz_brt)
            else:
                dt = dt.astimezone(tz_brt)
        return dt.strftime('%d/%m/%Y')
    except Exception:
        return str(exp_raw) if exp_raw else ''

def schedule_subscription_recheck(root, label_var):
    try:
        # Intervalo configurável (minutos) via env; padrão 12 horas
        minutes = int(os.getenv('SUBS_RECHECK_MINUTES', str(12*60)))
        interval_ms = max(1, minutes) * 60 * 1000
    except Exception:
        interval_ms = 12 * 60 * 60 * 1000

    def _check():
        try:
            if not _auth_session.get('access_token') or not _auth_session.get('email'):
                return
            ok, err, exp_at = supabase_check_subscription(_auth_session['access_token'], _auth_session['email'])
            if not ok:
                messagebox.showerror('Assinatura', err or 'Assinatura inválida/expirada')
                root.destroy()
                return
            # Atualiza label com data
            if exp_at:
                label_var.set(f"Assinatura: válida até {_format_expire_date_brt(exp_at)}")
        finally:
            try:
                root.after(interval_ms, _check)
            except Exception:
                pass

    # agenda primeira checagem
    root.after(interval_ms, _check)

def periodic_recheck(root, label_var, minutes=60):
    """Revalida periodicamente apenas o status do cliente.
    Fecha a aplicação somente se o status estiver False. Ignora expiração/data.
    """
    try:
        interval_ms = max(1, int(minutes)) * 60 * 1000
    except Exception:
        interval_ms = 60 * 60 * 1000

    def _do():
        try:
            email = _auth_session.get('email')
            if not email:
                return
            ok, err, status_bool = supabase_check_status_only(email)
            if ok:
                if status_bool is False:
                    messagebox.showerror('Assinatura', 'Assinatura inativa. Encerrando o aplicativo.')
                    root.destroy()
                    return
                # status true => mantém aberto; tenta mostrar validade se conhecida
                exp_at = _auth_session.get('expires_at')
                if exp_at:
                    try:
                        label_var.set(f"Assinatura: válida até {_format_expire_date_brt(exp_at)}")
                    except Exception:
                        pass
                else:
                    label_var.set('Assinatura: ativa')
            else:
                # Em caso de falha de rede ou erro HTTP, não fecha o app.
                # Mantém o status anterior para evitar falsos positivos.
                pass
        finally:
            try:
                root.after(interval_ms, _do)
            except Exception:
                pass

    root.after(interval_ms, _do)

import json
import urllib.request
import urllib.error
from datetime import datetime, timezone
import platform

# Variáveis globais
monitorando = False
fechar_telas = True
# Controla se deve clicar na tela após fechar (para evitar "último clique")
clicar_apos_fechar = False
Método_impressão_pdf = 1  # 1=ShellExecute, 2=PowerShell, 3=Automação, 4=SumatraPDF

# Função para enviar ZPL para a impressora padrão
def send_to_printer(zpl_data):
    try:
        printer_name = selected_printer_name or win32print.GetDefaultPrinter()
        printer_handle = win32print.OpenPrinter(printer_name)
        win32print.StartDocPrinter(printer_handle, 1, ("ZPL Print", None, "RAW"))
        win32print.StartPagePrinter(printer_handle)
        win32print.WritePrinter(printer_handle, zpl_data.encode('utf-8'))
        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        win32print.ClosePrinter(printer_handle)
        print("Etiqueta ZPL enviada para a impressora.")
    except Exception as e:
        print(f"Erro ao enviar etiqueta ZPL: {e}")

# Função para imprimir PDF usando diferentes Métodos
def print_pdf_method_1(pdf_file_path):
    """Método 1: win32api.ShellExecute"""
    try:
        win32api.ShellExecute(0, "print", pdf_file_path, None, ".", 0)
        print(f"PDF enviado para impressão via ShellExecute: {pdf_file_path}")
        time.sleep(3)
        return True
    except Exception as e:
        print(f"Método 1 falhou: {e}")
        return False

def print_pdf_method_2(pdf_file_path):
    """Método 2: PowerShell"""
    try:
        cmd = f'Start-Process -FilePath "{pdf_file_path}" -Verb Print -WindowStyle Hidden'
        result = subprocess.run([
            'powershell', '-Command', cmd
        ], capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print(f"PDF impresso via PowerShell: {pdf_file_path}")
            time.sleep(3)
            return True
        else:
            print(f"PowerShell retornou erro: {result.stderr}")
            return False
    except Exception as e:
        print(f"Método 2 falhou: {e}")
        return False

def print_pdf_method_3(pdf_file_path):
    """Método 3: Automação (abrir e Ctrl+P)"""
    try:
        # Abre o PDF
        subprocess.Popen([pdf_file_path], shell=True)
        print(f"PDF aberto para impressão: {pdf_file_path}")
        
        # Aguarda o PDF abrir
        time.sleep(4)
        
        # Simula Ctrl+P para imprimir
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(2)
        
        # Pressiona Enter para confirmar impressão
        pyautogui.press('enter')
        print("Comando de impressão enviado via automação")
        
        # Fecha a janela somente se a opção estiver ativada
        if fechar_telas:
            time.sleep(2)
            pyautogui.hotkey('alt', 'f4')
            # Evita o último clique, a menos que explicitamente habilitado
            if clicar_apos_fechar:
                time.sleep(0.3)
                pyautogui.click()

        return True
    except Exception as e:
        print(f"Método 3 falhou: {e}")
        return False

def print_pdf_method_4(pdf_file_path):
    """Método 4: Usando SumatraPDF (se disponível)"""
    try:
        # Procura pelo SumatraPDF em locais comuns
        sumatra_paths = [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            r"C:\Users\{}\AppData\Local\SumatraPDF\SumatraPDF.exe".format(os.getenv('USERNAME'))
        ]
        
        sumatra_exe = None
        for path in sumatra_paths:
            if os.path.exists(path):
                sumatra_exe = path
                break
        
        if sumatra_exe:
            # Usa SumatraPDF para impressão silenciosa
            subprocess.run([sumatra_exe, '-print-to-default', pdf_file_path], 
                         timeout=15, check=True)
            print(f"PDF impresso via SumatraPDF: {pdf_file_path}")
            time.sleep(2)
            return True
        else:
            print("SumatraPDF não encontrado")
            return False
    except Exception as e:
        print(f"Método 4 falhou: {e}")
        return False


def print_pdf_method_4_silent(pdf_file_path):
    """SumatraPDF silencioso, direcionando para impressora selecionada (sem UI)."""
    try:
        user = os.getenv('USERNAME') or ''
        sumatra_paths = [
            r"C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
            r"C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe",
            rf"C:\\Users\\{user}\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe",
        ]
        # Adiciona diretório atual e do executável (PyInstaller)
        sumatra_paths.append(os.path.join(os.getcwd(), 'SumatraPDF.exe'))
        try:
            exe_dir = os.path.dirname(sys.executable)
            sumatra_paths.append(os.path.join(exe_dir, 'SumatraPDF.exe'))
        except Exception:
            pass

        sumatra_exe = None
        for path in sumatra_paths:
            if os.path.exists(path):
                sumatra_exe = path
                break

        if not sumatra_exe:
            print("SumatraPDF não encontrado")
            return False

        try:
            prn = selected_printer_name or win32print.GetDefaultPrinter()
        except Exception:
            prn = selected_printer_name

        if prn:
            cmd = [sumatra_exe, '-print-to', prn, '-exit-on-print', pdf_file_path]
        else:
            cmd = [sumatra_exe, '-print-to-default', '-exit-on-print', pdf_file_path]

        subprocess.run(cmd, timeout=30, check=True)
        print(f"PDF impresso via SumatraPDF (silent): {pdf_file_path}")
        time.sleep(1)
        return True
    except Exception as e:
        print(f"Método 4 (silent) falhou: {e}")
        return False

# Função principal para imprimir PDF
def print_pdf(pdf_file_path):
    global Método_impressão_pdf
    
    methods = {
        1: print_pdf_method_1,
        2: print_pdf_method_2,
        3: print_pdf_method_3,
        4: print_pdf_method_4_silent,
    }
    
    # Tenta o Método Selecionado primeiro
    if Método_impressão_pdf in methods:
        if methods[Método_impressão_pdf](pdf_file_path):
            return
    
    # Se o Método principal falhou, tenta os outros
    print(f"Método {Método_impressão_pdf} falhou. Tentando outros Métodos...")
    for method_num, method_func in methods.items():
        if method_num != Método_impressão_pdf:
            print(f"Tentando Método {method_num}...")
            if method_func(pdf_file_path):
                print(f"Sucesso com Método {method_num}")
                return
    
    print("Todos os Métodos de impressão falharam!")
    messagebox.showerror("Erro", f"Não foi possível imprimir o PDF: {os.path.basename(pdf_file_path)}")

# Descoberta de suporte PDF
def _find_sumatra_exe():
    try:
        user = os.getenv('USERNAME') or ''
        paths = [
            r"C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
            r"C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe",
            rf"C:\\Users\\{user}\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe",
            os.path.join(os.getcwd(), 'SumatraPDF.exe'),
        ]
        try:
            exe_dir = os.path.dirname(sys.executable)
            paths.append(os.path.join(exe_dir, 'SumatraPDF.exe'))
        except Exception:
            pass
        for p in paths:
            if os.path.exists(p):
                return p
    except Exception:
        pass
    return None

def _has_pdf_association():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r".pdf") as key:
            val, _ = winreg.QueryValueEx(key, None)
            return bool(val)
    except Exception:
        return False

# Função para processar o arquivo ZIP e enviar para a impressora
def process_zip(zip_file_path):
    global fechar_telas, clicar_apos_fechar
    try:
        extract_dir = os.path.join(os.path.dirname(zip_file_path), 'temp_extract')
        os.makedirs(extract_dir, exist_ok=True)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # Procura por arquivos TXT (ZPL) primeiro
        zpl_file = next(
            (os.path.join(root, file) for root, _, files in os.walk(extract_dir) for file in files if file.endswith('.txt')),
            None
        )

        # Se não encontrou TXT, procura por PDF
        pdf_file = next(
            (os.path.join(root, file) for root, _, files in os.walk(extract_dir) for file in files if file.endswith('.pdf')),
            None
        )

        if zpl_file:
            with open(zpl_file, 'r') as file:
                send_to_printer(file.read())
            os.remove(zpl_file)
        elif pdf_file:
            print_pdf(pdf_file)
            os.remove(pdf_file)

        # Limpa a pasta temporária
        for root, _, files in os.walk(extract_dir, topdown=False):
            for file in files:
                os.remove(os.path.join(root, file))
            os.rmdir(root)

        time.sleep(0.5)
        os.remove(zip_file_path)
        
        # Só fecha as telas se a opção estiver ativada
        if fechar_telas:
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
            if clicar_apos_fechar:
                pyautogui.click()

    except Exception as e:
        print(f"Erro ao processar ZIP: {e}")

# Função para processar arquivos .txt diretamente
def process_txt(txt_file_path):
    global fechar_telas, clicar_apos_fechar
    try:
        with open(txt_file_path, 'r') as file:
            send_to_printer(file.read())
        os.remove(txt_file_path)
        print(f"Arquivo TXT processado: {txt_file_path}")

        # Só fecha as telas se a opção estiver ativada
        if fechar_telas:
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
            if clicar_apos_fechar:
                pyautogui.click()

    except Exception as e:
        print(f"Erro ao processar TXT: {e}")

# Função para processar arquivos .pdf diretamente
def process_pdf(pdf_file_path):
    global fechar_telas, clicar_apos_fechar
    try:
        print_pdf(pdf_file_path)
        os.remove(pdf_file_path)
        print(f"Arquivo PDF processado: {pdf_file_path}")

        # Só fecha as telas se a opção estiver ativada
        if fechar_telas:
            time.sleep(3)  # Aguarda mais tempo para PDFs
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
            if clicar_apos_fechar:
                pyautogui.click()

    except Exception as e:
        print(f"Erro ao processar PDF: {e}")

# Função para alternar o estado do fechamento de telas
def toggle_fechar_telas(checkbox_var, status_fechar_label):
    global fechar_telas
    fechar_telas = checkbox_var.get()
    status_text = "Ativado" if fechar_telas else "Desativado"
    color = "green" if fechar_telas else "red"
    status_fechar_label.config(text=f"Fechamento de telas: {status_text}", fg=color)

def toggle_clicar_apos_fechar(checkbox_var):
    global clicar_apos_fechar
    clicar_apos_fechar = checkbox_var.get()


# Função para alterar Método de impressão PDF
def change_pdf_method(method_var, method_label):
    global Método_impressão_pdf
    Método_impressão_pdf = method_var.get()
    methods = {
        1: "ShellExecute",
        2: "PowerShell", 
        3: "Automação",
        4: "SumatraPDF",
        5: "Ghostscript",
    }
    method_label.config(text=f"Método PDF: {methods[Método_impressão_pdf]}")

# Função para monitorar pastas
def monitor_etiquetas_shopee(base_dir, status_label, log_text, select_button):
    global monitorando
    monitorando = True
    log_text.insert(tk.END, f"Monitorando a pasta: {base_dir}\n")
    log_text.insert(tk.END, f"Tipos de arquivo suportados: ZIP, TXT (ZPL), PDF\n")
    log_text.insert(tk.END, f"Método de impressão PDF: {Método_impressão_pdf}\n")
    if selected_printer_name:
        log_text.insert(tk.END, f"Impressora atual: {selected_printer_name}\n")
    log_text.yview(tk.END)

    try:
        select_button.pack_forget()  # Esconde o botão de Selecionar pasta
        while monitorando:
            status_label.config(text=f"Buscando arquivos na pasta: {base_dir}", fg="blue")
            log_text.yview(tk.END)

            for file_name in os.listdir(base_dir):
                file_path = os.path.join(base_dir, file_name)
                
                # Verifica se o arquivo é ZIP
                if file_name.endswith('.zip'):
                    status_label.config(text=f"Ativo - Processando ZIP: {file_name}", fg="green")
                    log_text.insert(tk.END, f"Processando ZIP: {file_name}\n")
                    log_text.yview(tk.END)
                    process_zip(file_path)
                    break
                
                # Verifica se o arquivo é TXT (ZPL)
                elif file_name.endswith('.txt'):
                    status_label.config(text=f"Ativo - Processando TXT: {file_name}", fg="green")
                    log_text.insert(tk.END, f"Processando TXT (ZPL): {file_name}\n")
                    log_text.yview(tk.END)
                    process_txt(file_path)
                    break
                
                # Verifica se o arquivo é PDF
                elif file_name.endswith('.pdf'):
                    status_label.config(text=f"Ativo - Processando PDF: {file_name}", fg="green")
                    log_text.insert(tk.END, f"Processando PDF: {file_name}\n")
                    log_text.yview(tk.END)
                    process_pdf(file_path)
                    break

            time.sleep(3)

        log_text.insert(tk.END, "Monitoramento finalizado...\n")
    except Exception as e:
        log_text.insert(tk.END, f"Erro: {e}\n")
        monitorando = False

# Função para parar o monitoramento
def stop_monitoramento(status_label, log_text, select_button):
    global monitorando
    monitorando = False
    status_label.config(text="Inativo", fg="red")
    log_text.insert(tk.END, "Monitoramento parado.\n")
    log_text.yview(tk.END)
    select_button.pack(pady=10)  # Mostra novamente o botão de Selecionar pasta

# Função para Selecionar pasta
def select_folder(status_label, log_text, select_button):
    folder_path = filedialog.askdirectory(title="Selecione a pasta a ser monitorada")
    if folder_path:
        status_label.config(text="Ativo", fg="green")
        stop_button.pack(pady=10)
        threading.Thread(
            target=monitor_etiquetas_shopee, 
            args=(folder_path, status_label, log_text, select_button), 
            daemon=True
        ).start()
    else:
        status_label.config(text="Inativo", fg="red")
        log_text.insert(tk.END, "Nenhuma pasta selecionada.\n")
        stop_button.pack_forget()

# Função para testar impressão PDF
def test_pdf_print():
    file_path = filedialog.askopenfilename(
        title="selecione um PDF para testar",
        filetypes=[("PDF files", "*.pdf")]
    )
    if file_path:
        print(f"Testando impressão do PDF: {file_path}")
        # Cria uma cópia temporária para não apagar o original
        temp_file = os.path.join(tempfile.gettempdir(), f"test_{os.path.basename(file_path)}")
        shutil.copy2(file_path, temp_file)
        print_pdf(temp_file)
        # Remove a cópia temporária
        if os.path.exists(temp_file):
            os.remove(temp_file)

# Função principal
def main():
    global stop_button, fechar_telas, Método_impressão_pdf
    
    root = tk.Tk()
    # Evita janela em branco enquanto exibe o login
    try:
        root.withdraw()
    except Exception:
        pass
    # Login obrigatório antes de continuar
    try:
        show_login_dialog(root)
    except Exception as e:
        print(f"Falha ao exibir login: {e}")
    if not _auth_session.get("email"):
        try:
            root.destroy()
        finally:
            return
    # Reexibe janela principal após login
    try:
        root.deiconify()
    except Exception:
        pass
    root.title("GAC - Monitor de Etiquetas (ZPL + PDF)")
    root.geometry("700x600")

    # Frame para controles superiores
    control_frame = tk.Frame(root)
    control_frame.pack(pady=10)

    # Primeira linha de controles
    first_row = tk.Frame(control_frame)
    first_row.pack(fill=tk.X, pady=5)

    # Checkbox para controlar o fechamento de telas
    checkbox_var = tk.BooleanVar(value=fechar_telas)
    checkbox = tk.Checkbutton(
        first_row,
        text="Fechar telas após impressão",
        variable=checkbox_var,
        font=("Arial", 10),
        command=lambda: toggle_fechar_telas(checkbox_var, status_fechar_label)
    )
    checkbox.pack(side=tk.LEFT, padx=10)

    # Label para mostrar o status do fechamento de telas
    status_fechar_label = tk.Label(
        first_row,
        text="Fechamento de telas: Ativado",
        font=("Arial", 10),
        fg="green"
    )
    status_fechar_label.pack(side=tk.LEFT, padx=10)

    # Checkbox para controlar o clique após fechar
    clicar_var = tk.BooleanVar(value=clicar_apos_fechar)
    clicar_checkbox = tk.Checkbutton(
        first_row,
        text="Clique após fechar",
        variable=clicar_var,
        font=("Arial", 10),
        command=lambda: toggle_clicar_apos_fechar(clicar_var)
    )
    clicar_checkbox.pack(side=tk.LEFT, padx=10)

    # Linha de status da assinatura
    subs_row = tk.Frame(control_frame)
    subs_row.pack(fill=tk.X, pady=2)
    subs_label_var = tk.StringVar(value="Assinatura: validando...")
    subs_label = tk.Label(subs_row, textvariable=subs_label_var, font=("Arial", 10), fg="green")
    subs_label.pack(side=tk.LEFT, padx=10)

    # Segunda linha de controles
    second_row = tk.Frame(control_frame)
    second_row.pack(fill=tk.X, pady=5)

    # Seletor de Método de impressão PDF
    tk.Label(second_row, text="Método impressão PDF:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

    # Define método padrão de forma inteligente:
    # - Se Sumatra portátil/existente for encontrado: usar Método 4 (silencioso)
    # - Senão, se não houver associação de PDF no Windows: usar Método 3 (automação)
    # - Caso contrário: Método 1 (ShellExecute) ou a configuração global
    try:
        if _find_sumatra_exe():
            default_method = 4
        elif not _has_pdf_association():
            default_method = 3
        else:
            default_method = Método_impressão_pdf
    except Exception:
        default_method = Método_impressão_pdf

    method_var = tk.IntVar(value=default_method)
    method_options = [
        ("ShellExecute", 1),
        ("PowerShell", 2),
        ("Automação", 3),
        ("SumatraPDF", 4),
    ]
    
    _method_names = {1: 'ShellExecute', 2: 'PowerShell', 3: 'Automação', 4: 'SumatraPDF'}
    method_label = tk.Label(second_row, text=f"Método PDF: {_method_names.get(method_var.get(), 'ShellExecute')}", font=("Arial", 10), fg="blue")
    
    for text, value in method_options:
        tk.Radiobutton(
            second_row,
            text=text,
            variable=method_var,
            value=value,
            font=("Arial", 9),
            command=lambda: change_pdf_method(method_var, method_label)
        ).pack(side=tk.LEFT, padx=5)
    
    method_label.pack(side=tk.LEFT, padx=10)
    # Sincroniza método inicial com o estado atual
    try:
        change_pdf_method(method_var, method_label)
    except Exception:
        pass
    
    # (Removido) Ghostscript como padrão — mantido apenas métodos 1-4
    # Exibe validade e agenda rechecagem da assinatura
    try:
        exp_at = _auth_session.get("expires_at")
        if exp_at:
            subs_label_var.set(f"Assinatura: válida até {_format_expire_date_brt(exp_at)}")
        else:
            subs_label_var.set("Assinatura: ativa")
        schedule_subscription_recheck(root, subs_label_var)
        periodic_recheck(root, subs_label_var, minutes=int(os.getenv('SUBS_RECHECK_MINUTES', '60')))
    except Exception:
        pass
    
    # Atualiza label de assinatura com data do login
    try:
        exp_at = _auth_session.get("expires_at")
        if exp_at:
            subs_label_var.set(f"Assinatura: válida até {_format_expire_date_brt(exp_at)}")
        else:
            subs_label_var.set("Assinatura: ativa")
    except Exception:
        pass

    # Terceira linha - seleção de impressora
    printer_row = tk.Frame(control_frame)
    printer_row.pack(fill=tk.X, pady=5)
    tk.Label(printer_row, text="Impressora:", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
    try:
        _current_default_prn = win32print.GetDefaultPrinter()
    except Exception:
        _current_default_prn = ""
    printer_label_var = tk.StringVar(value=(selected_printer_name or _current_default_prn))
    printer_label = tk.Label(printer_row, textvariable=printer_label_var, font=("Arial", 10), fg="purple")
    printer_label.pack(side=tk.LEFT, padx=8)

    def do_change_printer():
        choose_printer_dialog(root)
        try:
            printer_label_var.set(selected_printer_name or win32print.GetDefaultPrinter())
        except Exception:
            printer_label_var.set(selected_printer_name or "")

    tk.Button(printer_row, text="Trocar", command=do_change_printer).pack(side=tk.LEFT, padx=10)

    # Terceira linha - botão de teste
    third_row = tk.Frame(control_frame)
    third_row.pack(fill=tk.X, pady=5)

    test_button = tk.Button(
        third_row,
        text="Testar impressão PDF",
        font=("Arial", 10),
        bg="#2196F3",
        fg="white",
        command=test_pdf_print
    )
    test_button.pack(side=tk.LEFT, padx=10)

    # Separador visual
    separator = tk.Frame(root, height=2, bg="gray")
    separator.pack(fill=tk.X, padx=20, pady=10)

    # Label de informações sobre tipos de arquivo
    info_label = tk.Label(
        root, 
        text="Suporte a: ZIP, TXT (ZPL), PDF | Múltiplos Métodos de impressão PDF", 
        font=("Arial", 10), 
        fg="gray"
    )
    info_label.pack(pady=5)

    status_label = tk.Label(root, text="Selecione uma pasta para monitorar", font=("Arial", 14), width=50)
    status_label.pack(pady=20)

    log_text = tk.Text(root, height=12, width=80, font=("Arial", 9), wrap=tk.WORD)
    log_text.pack(pady=10)

    select_button = tk.Button(
        root, text="Selecionar Pasta", font=("Arial", 12), bg="#4CAF50", fg="white", 
        command=lambda: select_folder(status_label, log_text, select_button)
    )
    select_button.pack(pady=10)

    stop_button = tk.Button(
        root, text="Parar Monitoramento", font=("Arial", 12), bg="#f44336", fg="white", 
        command=lambda: stop_monitoramento(status_label, log_text, select_button)
    )

    # Solicitar seleção de impressora ao abrir
    try:
        choose_printer_dialog(root)
    except Exception as e:
        print(f"Falha ao abrir seleção de impressora: {e}")

    # Inicia automaticamente a seleção de pasta após 100ms
    root.after(100, lambda: select_folder(status_label, log_text, select_button))

    root.mainloop()

if __name__ == '__main__':
    main()














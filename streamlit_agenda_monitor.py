"""
Agenda Impettus - Monitor de E-mails Zimbra
============================================
App Streamlit que monitora a caixa de entrada via IMAP e encaminha
e-mails de agendamento para a Edge Function process-incoming-email.

Uso:
    pip install streamlit requests
    streamlit run streamlit_agenda_monitor.py
"""

import streamlit as st
import imaplib
import email
import requests
import json
import time
import threading
from datetime import datetime
from email.header import decode_header
from email.utils import parseaddr

# ─── Configurações ───────────────────────────────────────────────
IMAP_HOST = "smtp.emailzimbraonline.com"
IMAP_PORT = 993
IMAP_USER = "agenda@grupoimpettus.com.br"
IMAP_PASS = "upHh&@Cp6W"

ENDPOINT_URL = "https://rmkrsuwncqxsavykgrad.supabase.co/functions/v1/process-incoming-email"

# ─── Estado da sessão ────────────────────────────────────────────
if "logs" not in st.session_state:
    st.session_state.logs = []
if "is_running" not in st.session_state:
    st.session_state.is_running = True
if "total_processed" not in st.session_state:
    st.session_state.total_processed = 0
if "total_success" not in st.session_state:
    st.session_state.total_success = 0
if "total_conflict" not in st.session_state:
    st.session_state.total_conflict = 0
if "total_rejected" not in st.session_state:
    st.session_state.total_rejected = 0
if "total_error" not in st.session_state:
    st.session_state.total_error = 0
if "last_check" not in st.session_state:
    st.session_state.last_check = None
if "imap_status" not in st.session_state:
    st.session_state.imap_status = "verificando..."
if "initialized" not in st.session_state:
    st.session_state.initialized = False


def add_log(message: str, level: str = "info"):
    """Adiciona uma entrada ao log de atividade."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    icons = {"info": "ℹ️", "success": "✅", "warning": "⚠️", "error": "❌", "conflict": "🔴"}
    icon = icons.get(level, "ℹ️")
    st.session_state.logs.insert(0, {
        "time": timestamp,
        "message": message,
        "level": level,
        "icon": icon,
    })
    # Manter apenas os últimos 100 logs
    if len(st.session_state.logs) > 100:
        st.session_state.logs = st.session_state.logs[:100]


def decode_mime_header(header_value: str) -> str:
    """Decodifica cabeçalhos MIME (Subject, From, etc.)."""
    if not header_value:
        return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="ignore"))
        else:
            result.append(part)
    return " ".join(result)


def extract_email_body(msg) -> str:
    """Extrai o corpo do e-mail (prioriza text/plain, depois text/html)."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            if "attachment" in content_disposition:
                continue
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = payload.decode(charset, errors="ignore")
                    break
            elif content_type == "text/html" and not body:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    body = payload.decode(charset, errors="ignore")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="ignore")
    return body.strip()


def test_imap_connection() -> bool:
    """Testa a conexão IMAP e atualiza o status."""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASS)
        mail.select("INBOX", readonly=True)
        mail.logout()
        st.session_state.imap_status = "conectado"
        return True
    except Exception as e:
        st.session_state.imap_status = f"erro: {str(e)[:60]}"
        return False


def process_emails() -> int:
    """Conecta ao IMAP, processa e-mails não lidos e envia ao endpoint."""
    processed = 0
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_USER, IMAP_PASS)
        mail.select("INBOX")
        st.session_state.imap_status = "conectado"

        status, messages = mail.search(None, "UNSEEN")
        if status != "OK" or not messages[0]:
            add_log("Nenhum e-mail novo encontrado.", "info")
            st.session_state.last_check = datetime.now().strftime("%H:%M:%S")
            return 0

        msg_ids = messages[0].split()
        add_log(f"Encontrado(s) {len(msg_ids)} e-mail(s) não lido(s).", "info")

        for msg_id in msg_ids:
            try:
                status, data = mail.fetch(msg_id, "(RFC822)")
                if status != "OK":
                    add_log(f"Erro ao buscar e-mail ID {msg_id.decode()}", "error")
                    continue

                msg = email.message_from_bytes(data[0][1])

                # Extrair campos
                raw_from = msg.get("From", "")
                from_name, from_addr = parseaddr(raw_from)
                from_display = decode_mime_header(raw_from)
                subject = decode_mime_header(msg.get("Subject", ""))
                body = extract_email_body(msg)

                # Validação de domínio: apenas @grupoimpettus.com.br
                if not from_addr.lower().endswith("@grupoimpettus.com.br"):
                    add_log(f"Rejeitado (domínio externo): {from_addr}", "warning")
                    mail.store(msg_id, "+FLAGS", "\\Seen")
                    st.session_state.total_rejected += 1
                    st.session_state.total_processed += 1
                    processed += 1
                    continue

                add_log(f"Processando: \"{subject}\" de {from_addr}", "info")

                # Extrair CC e To para capturar participantes em cópia
                cc_raw = msg.get("Cc", "") or ""
                cc_addrs = [parseaddr(addr.strip())[1] for addr in cc_raw.split(",") if addr.strip() and parseaddr(addr.strip())[1]]

                to_raw = msg.get("To", "") or ""
                to_addrs = [parseaddr(addr.strip())[1] for addr in to_raw.split(",") if addr.strip() and parseaddr(addr.strip())[1]]

                if cc_addrs:
                    add_log(f"  CC detectados: {', '.join(cc_addrs)}", "info")
                if len(to_addrs) > 1:
                    add_log(f"  To detectados: {', '.join(to_addrs)}", "info")

                # Enviar ao endpoint
                payload = {
                    "from": from_addr or from_display,
                    "subject": subject,
                    "body": body,
                    "cc": cc_addrs,
                    "to": to_addrs,
                }

                response = requests.post(
                    ENDPOINT_URL,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                    timeout=60,
                )

                result = response.json() if response.headers.get("content-type", "").startswith("application/json") else {}

                if response.status_code == 200 and result.get("success"):
                    add_log(
                        f"Reunião criada: \"{result.get('title', subject)}\" em {result.get('date', '?')} - {result.get('location', '?')}",
                        "success",
                    )
                    st.session_state.total_success += 1
                elif result.get("reason") == "conflict":
                    add_log(
                        f"Conflito de horário: \"{subject}\" na {result.get('location', '?')} em {result.get('date', '?')}",
                        "conflict",
                    )
                    st.session_state.total_conflict += 1
                elif result.get("reason") in ("missing_date_time", "missing_room"):
                    add_log(f"Rejeitado ({result.get('reason')}): \"{subject}\"", "warning")
                    st.session_state.total_rejected += 1
                elif result.get("reason") == "unauthorized_domain":
                    add_log(f"Rejeitado (domínio não autorizado): \"{subject}\" de {from_addr}", "warning")
                    st.session_state.total_rejected += 1
                else:
                    error_msg = result.get("error", f"HTTP {response.status_code}")
                    add_log(f"Erro ao processar \"{subject}\": {error_msg}", "error")
                    st.session_state.total_error += 1

                # Marcar como lido
                mail.store(msg_id, "+FLAGS", "\\Seen")
                processed += 1
                st.session_state.total_processed += 1

            except Exception as e:
                add_log(f"Erro ao processar e-mail: {str(e)[:80]}", "error")
                st.session_state.total_error += 1

        st.session_state.last_check = datetime.now().strftime("%H:%M:%S")
        return processed

    except imaplib.IMAP4.error as e:
        st.session_state.imap_status = f"erro IMAP: {str(e)[:60]}"
        add_log(f"Erro IMAP: {str(e)[:80]}", "error")
        return 0
    except Exception as e:
        add_log(f"Erro geral: {str(e)[:80]}", "error")
        return 0
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass


# ─── Interface Streamlit ─────────────────────────────────────────

st.set_page_config(
    page_title="Agenda Impettus - Monitor",
    page_icon="📧",
    layout="wide",
)

# Auto-testar conexão IMAP ao carregar a página
if not st.session_state.initialized:
    st.session_state.initialized = True
    test_imap_connection()

st.title("📧 Agenda Impettus - Monitor de E-mails")
st.caption("Monitora a caixa de entrada e encaminha e-mails de agendamento para o sistema.")

# ─── Status e Métricas ───────────────────────────────────────────
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    status_color = "🟢" if st.session_state.imap_status == "conectado" else "🔴"
    st.metric("IMAP", f"{status_color} {st.session_state.imap_status[:15]}")
with col2:
    st.metric("Processados", st.session_state.total_processed)
with col3:
    st.metric("Sucesso", st.session_state.total_success)
with col4:
    st.metric("Conflitos", st.session_state.total_conflict)
with col5:
    st.metric("Rejeitados", st.session_state.total_rejected)

st.divider()

# ─── Controles ────────────────────────────────────────────────────
col_ctrl1, col_ctrl2, col_ctrl3 = st.columns([2, 2, 3])

with col_ctrl1:
    if st.button("🔍 Verificar Agora", use_container_width=True):
        with st.spinner("Verificando e-mails..."):
            count = process_emails()
        if count > 0:
            st.success(f"{count} e-mail(s) processado(s)!")
        else:
            st.info("Nenhum e-mail novo.")
        st.rerun()

with col_ctrl2:
    if st.button("🔌 Testar Conexão IMAP", use_container_width=True):
        with st.spinner("Testando conexão..."):
            ok = test_imap_connection()
        if ok:
            st.success("Conexão IMAP OK!")
        else:
            st.error("Falha na conexão IMAP.")
        st.rerun()

with col_ctrl3:
    interval = st.slider(
        "Intervalo de polling (segundos)",
        min_value=60, max_value=600, value=300, step=60,
        help="Intervalo entre verificações automáticas (padrão: 5 min)",
    )

# ─── Polling Automático ──────────────────────────────────────────
st.divider()

auto_col1, auto_col2 = st.columns([1, 3])

with auto_col1:
    auto_poll = st.toggle("🔄 Polling Automático", value=True)

with auto_col2:
    if auto_poll:
        st.info(f"✅ Monitoramento ativo — verificando a cada {interval // 60} min.")
        with st.spinner(f"Verificando e-mails..."):
            count = process_emails()
            if count > 0:
                st.toast(f"✅ {count} e-mail(s) processado(s)!")
        time.sleep(interval)
        st.rerun()
    else:
        if st.session_state.last_check:
            st.caption(f"Última verificação: {st.session_state.last_check}")

# ─── Log de Atividade ────────────────────────────────────────────
st.divider()
st.subheader("📋 Log de Atividade")

if not st.session_state.logs:
    st.info("Nenhuma atividade registrada. Clique em \"Verificar Agora\" para iniciar.")
else:
    # Filtro
    filter_level = st.selectbox(
        "Filtrar por tipo:",
        ["Todos", "Sucesso", "Conflito", "Rejeitado", "Erro", "Info"],
        index=0,
    )
    level_map = {
        "Sucesso": "success", "Conflito": "conflict",
        "Rejeitado": "warning", "Erro": "error", "Info": "info",
    }

    for log in st.session_state.logs:
        if filter_level != "Todos" and log["level"] != level_map.get(filter_level):
            continue
        color = {
            "success": "green", "conflict": "red",
            "warning": "orange", "error": "red", "info": "blue",
        }.get(log["level"], "gray")
        st.markdown(
            f"<span style='color:gray;font-size:12px;'>{log['time']}</span> "
            f"{log['icon']} <span style='color:{color};'>{log['message']}</span>",
            unsafe_allow_html=True,
        )

    if st.button("🗑️ Limpar Logs"):
        st.session_state.logs = []
        st.rerun()

# ─── Informações ──────────────────────────────────────────────────
with st.expander("ℹ️ Informações do Sistema"):
    st.markdown(f"""
| Configuração | Valor |
|---|---|
| **Servidor IMAP** | `{IMAP_HOST}:{IMAP_PORT}` |
| **Conta** | `{IMAP_USER}` |
| **Endpoint** | `{ENDPOINT_URL}` |
| **Método** | `POST (JSON)` |
""")
    st.markdown("""
### Como funciona
1. O monitor conecta via IMAP na caixa de entrada
2. Busca e-mails **não lidos**
3. Extrai remetente, assunto e corpo
4. Envia via HTTP POST para a Edge Function
5. A IA extrai dados da reunião e cria no sistema
6. O remetente recebe um e-mail de confirmação, conflito ou rejeição
7. O e-mail é marcado como lido
""")

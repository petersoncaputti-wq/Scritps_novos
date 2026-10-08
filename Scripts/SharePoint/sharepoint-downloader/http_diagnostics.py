"""Diagnostico limitado de erros HTTP, sem despejar corpos ou headers."""
import json
import re
import uuid


def sanitize(value, limit=1200):
    if not isinstance(value, (str, int, float)):
        return ''
    text = str(value)
    text = re.sub(r'https?://[^\s<>"\']+', '[url-removida]', text, flags=re.I)
    text = re.sub(r'\bBearer\s+\S+', 'Bearer [removido]', text, flags=re.I)
    text = re.sub(r'''(?i)\b(?:authorization|cookie|set-cookie|access_token|refresh_token|id_token|client_secret|password|senha|sig|token)\b["']?\s*[:=]\s*(?:"[^"]*"|'[^']*')''', '[credencial-removida]', text)
    text = re.sub(r'(?i)\b(?:authorization|cookie|set-cookie|access_token|refresh_token|id_token|client_secret|password|senha|sig|token)\b\s*[:=]\s*[^\s;,]+', '[credencial-removida]', text)
    text = re.sub(r'\beyJ[A-Za-z0-9_-]+(?:\.[A-Za-z0-9_-]+){1,2}\b', '[token-removido]', text)
    text = re.sub(r'[A-Za-z0-9_+/=-]{64,}', '[valor-opaco-removido]', text)
    text = re.sub(r'\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b', '[email-removido]', text)
    text = re.sub(r'[\x00-\x1f\x7f]', ' ', text)
    return ' '.join(text.split())[:limit]


def error_details(response):
    content_type = response.headers.get('Content-Type', '')
    content_type = content_type if isinstance(content_type, str) else ''
    media = content_type.split(';', 1)[0].lower().strip()
    details = {'formato': sanitize(media, 80), 'codigo': '', 'mensagem': '', 'corpo_truncado': False}
    for header in ('SPRequestGuid', 'request-id', 'client-request-id'):
        value = response.headers.get(header)
        try:
            details[header] = str(uuid.UUID(value)) if isinstance(value, str) else ''
        except ValueError:
            details[header] = ''
    # Read at most 16 KiB of decompressed response, never response.text/json().
    data = bytearray()
    try:
        for chunk in response.iter_content(chunk_size=4096):
            if not isinstance(chunk, bytes):
                break
            remaining = 16384 - len(data)
            data.extend(chunk[:remaining])
            if len(data) >= 16384:
                details['corpo_truncado'] = True
                break
    except Exception as error:
        details['leitura_erro'] = type(error).__name__
    text = bytes(data).decode('utf-8', errors='replace')
    details['bytes_corpo_lidos'] = len(data)
    try:
        body = json.loads(text)
        error = body.get('error', body.get('odata.error', {})) if isinstance(body, dict) else {}
        if isinstance(error, dict):
            details['codigo'] = sanitize(error.get('code'), 200)
            message = error.get('message', '')
            if isinstance(message, dict):
                message = message.get('value', '')
            details['mensagem'] = sanitize(message)
    except (ValueError, TypeError):
        if media == 'text/plain' or (not media and '<' not in text and '{' not in text):
            details['mensagem'] = sanitize(text)
        elif media == 'text/html':
            # Only page title, never login forms, scripts, or hidden fields.
            title = re.search(r'<title\b[^>]*>(.*?)</title>', text, re.I | re.S)
            if title:
                details['mensagem'] = sanitize(re.sub(r'<[^>]+>', '', title[1]))
    return details

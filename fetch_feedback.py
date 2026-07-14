#!/usr/bin/env python3
"""
Lê o canal 'Feedback Plataforma' do Teams (últimos N dias, incluindo respostas
de thread), classifica as mensagens e prepara a revisão:
 - feedbacks são inseridos no index.html (formato accordion), SEM commit automático
 - bugs vão para pending_review.json para triagem no Linear (não inseridos)

Uso: python3 fetch_feedback.py [dias]   (padrão: 7, ou env FETCH_SINCE_DAYS)
"""

import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import msal
import requests
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
CHANNEL_NAME = os.getenv("TEAMS_CHANNEL_NAME", "feedback plataforma")
TEAM_ID = os.getenv("TEAMS_TEAM_ID", "")
CHANNEL_ID = os.getenv("TEAMS_CHANNEL_ID", "")
DEFAULT_SINCE_DAYS = int(os.getenv("FETCH_SINCE_DAYS", "7"))

LLM_MODEL = os.getenv("LITELLM_MODEL", "gemini/gemini-3.1-flash-lite")
llm = OpenAI(
    api_key=os.environ["LITELLM_API_KEY"],
    base_url=os.getenv("LITELLM_BASE_URL", "https://llm.nodian.com.br"),
)

INDEX_HTML = Path(__file__).parent / "index.html"
PROCESSED_IDS_FILE = Path(__file__).parent / ".processed_message_ids.json"
PENDING_REVIEW_FILE = Path(__file__).parent / "pending_review.json"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Marcador de fechamento da seção de feedback no index.html (formato accordion).
# Cards novos são inseridos imediatamente antes deste bloco.
FEEDBACK_CLOSE = "\n</div>\n</details>\n</div> <!-- close tab-feedback -->"


def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Autenticação falhou: {result.get('error_description')}")
    return result["access_token"]


def graph_get(token, path, params=None):
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(f"{GRAPH_BASE}{path}", headers=headers, params=params)
    resp.raise_for_status()
    return resp.json()


def graph_get_url(token, url):
    """GET numa URL absoluta (usado para paginação via @odata.nextLink)."""
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.json()


def find_channel(token):
    if TEAM_ID and CHANNEL_ID:
        return TEAM_ID, CHANNEL_ID
    teams = graph_get(token, "/teams")["value"]
    for team in teams:
        channels = graph_get(token, f"/teams/{team['id']}/channels")["value"]
        for ch in channels:
            if CHANNEL_NAME.lower() in ch["displayName"].lower():
                return team["id"], ch["id"]
    raise RuntimeError(f"Canal '{CHANNEL_NAME}' não encontrado.")


def _created_at(msg):
    try:
        return datetime.fromisoformat(msg.get("createdDateTime", "").replace("Z", "+00:00"))
    except ValueError:
        return None


def get_new_messages(token, team_id, channel_id, processed_ids, since_days):
    """Coleta mensagens raiz E respostas de thread dos últimos `since_days` dias.

    A API de canais só retorna as mensagens-raiz em /messages; as respostas
    ficam em /messages/{id}/replies. Sem ler replies a maioria dos feedbacks
    (que ficam em threads) passa despercebida.
    """
    cutoff = datetime.now(timezone.utc) - timedelta(days=since_days)
    collected = []
    url = f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages?$top=50"
    pages = 0
    while url and pages < 20:
        data = graph_get_url(token, url)
        stop = False
        for m in data.get("value", []):
            created = _created_at(m)
            if created is not None and created < cutoff:
                stop = True  # mensagens vêm da mais nova p/ mais antiga
                continue
            collected.append(m)
            try:
                replies = graph_get(
                    token,
                    f"/teams/{team_id}/channels/{channel_id}/messages/{m['id']}/replies",
                    params={"$top": 50},
                )
                collected.extend(replies.get("value", []))
            except requests.HTTPError:
                pass  # sem replies acessíveis para esta mensagem
        if stop:
            break
        url = data.get("@odata.nextLink")
        pages += 1

    candidates = [
        m for m in collected
        if m["id"] not in processed_ids
        and m.get("messageType") == "message"
        and m.get("body", {}).get("content", "").strip()
        and m.get("from") is not None
    ]
    # deduplicar por conteúdo (edições e replies que ecoam a raiz)
    seen_content = set()
    unique = []
    for m in candidates:
        content = re.sub(r"<[^>]+>", "", m["body"]["content"]).strip()
        if content and content not in seen_content:
            seen_content.add(content)
            unique.append(m)
    return unique


def llm_call(system, user, max_tokens=1024):
    resp = llm.chat.completions.create(
        model=LLM_MODEL,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    return resp.choices[0].message.content.strip()


def classify_bug(content, author, date):
    result = llm_call(
        system="""Você recebe uma mensagem de usuário descrevendo um bug em uma plataforma de dados corporativa.
Retorne apenas JSON válido (sem markdown) com:
- title: título descritivo, máx 70 chars
- priority: "Alto", "Médio" ou "Baixo"
- user_story: frase no formato "Como [persona], ao [ação], [problema observado]."
- steps: lista de strings com os passos para reproduzir (máx 5)
- current_flow: lista de objetos {from, to} representando o fluxo atual com no máximo 4 nós. Use texto curto sem aspas especiais.
- expected_flow: lista de objetos {from, to} representando o fluxo esperado com no máximo 4 nós. Use texto curto sem aspas especiais.
- acceptance_criteria: lista de strings com os critérios de aceite (máx 4)
- consequence: frase descrevendo a consequência do bug para o usuário""",
        user=f"Usuário {author} ({date}):\n\n{content}",
    )
    return json.loads(result)


def classify_feedback(content, author, date):
    result = llm_call(
        system="""Você recebe uma mensagem de usuário com um feedback sobre uma plataforma de dados corporativa.
Retorne apenas JSON válido (sem markdown) com:
- title: título descritivo, máx 70 chars
- category: emoji + rótulo curto, ex: "🕳️ Lacuna de funcionalidade", "🧩 Usabilidade", "💡 Sugestão", "⚡ Performance", "🎨 Design"
- summary: descrição expandida em 2-3 frases, terceira pessoa
- priority: "Alto", "Médio" ou "Baixo" """,
        user=f"Usuário {author} ({date}):\n\n{content}",
        max_tokens=512,
    )
    return json.loads(result)


def classify_type(content):
    result = llm_call(
        system="""Classifique a mensagem de um usuário de plataforma.
Responda apenas com uma palavra:
- "bug" — relato de comportamento incorreto, erro ou falha
- "feedback" — sugestão, opinião ou dificuldade de uso com conteúdo acionável
- "ignorar" — mensagem sem conteúdo acionável (saudações, testes, mensagens vazias, avisos administrativos)""",
        user=content,
        max_tokens=10,
    )
    return result.lower()


def feedback_card(item):
    """Card de feedback no formato accordion atual do dashboard."""
    original = re.sub(r"\s+", " ", item["original"]).strip()
    return f"""<details class="issue-row">
<summary>
<div class="issue-row-id">FEEDBACK</div>
<div class="issue-row-title">{item['title']}</div>
<span class="priority-tag priority-medium">{item.get('category', '💬 Feedback')}</span>
<svg class="issue-row-chevron" fill="none" height="14" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" viewbox="0 0 24 24" width="14"><polyline points="9 18 15 12 9 6"></polyline></svg>
</summary>
<div class="issue-row-body">
<div class="feedback-meta">{item['date']} · {item['author']}</div>
<p class="issue-summary"><em>"{original}"</em></p>
<p class="issue-summary">{item.get('summary', '')}</p>
</div>
</details>"""


def update_html(feedbacks):
    """Insere os cards de feedback antes do fechamento da seção e ajusta o contador."""
    if not feedbacks:
        return
    html = INDEX_HTML.read_text(encoding="utf-8")
    if FEEDBACK_CLOSE not in html:
        raise RuntimeError(
            "Marcador de fechamento da seção de feedback não encontrado no index.html. "
            "A estrutura mudou — insira os cards manualmente ou atualize FEEDBACK_CLOSE."
        )
    cards = "\n".join(feedback_card(f) for f in feedbacks)
    html = html.replace(FEEDBACK_CLOSE, f"\n{cards}{FEEDBACK_CLOSE}", 1)

    def _bump(match):
        return f'Feedbacks <span class="issue-section-count">{int(match.group(1)) + len(feedbacks)}</span>'

    html, n = re.subn(r'Feedbacks <span class="issue-section-count">(\d+)</span>', _bump, html, count=1)
    if n == 0:
        print("Aviso: contador da seção de feedback não encontrado — ajuste manualmente.")
    INDEX_HTML.write_text(html, encoding="utf-8")


def save_pending_review(bugs, feedbacks):
    existing = []
    if PENDING_REVIEW_FILE.exists():
        try:
            existing = json.loads(PENDING_REVIEW_FILE.read_text())
        except Exception:
            existing = []
    all_items = existing + [{"type": "bug", **b} for b in bugs] + [{"type": "feedback", **f} for f in feedbacks]
    PENDING_REVIEW_FILE.write_text(json.dumps(all_items, ensure_ascii=False, indent=2))


def load_processed():
    if PROCESSED_IDS_FILE.exists():
        return set(json.loads(PROCESSED_IDS_FILE.read_text()))
    return set()


def save_processed(ids):
    PROCESSED_IDS_FILE.write_text(json.dumps(list(ids)))


def main():
    try:
        since_days = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SINCE_DAYS
    except ValueError:
        print(f"Argumento inválido '{sys.argv[1]}' — informe o número de dias. Ex: python3 fetch_feedback.py 14")
        return

    processed = load_processed()

    print("Autenticando com Microsoft Graph...")
    token = get_token()

    print(f"Localizando canal '{CHANNEL_NAME}'...")
    team_id, channel_id = find_channel(token)

    print(f"Buscando mensagens dos últimos {since_days} dia(s), incluindo respostas de thread...")
    messages = get_new_messages(token, team_id, channel_id, processed, since_days)

    if not messages:
        print("Nenhuma mensagem nova.")
        return

    print(f"{len(messages)} mensagem(ns) nova(s). Classificando com LLM...")
    bugs, feedbacks = [], []

    for msg in messages:
        content = re.sub(r"<[^>]+>", "", msg["body"]["content"]).strip()
        if not content:
            continue
        author = (msg.get("from") or {}).get("user", {}).get("displayName", "Usuário")
        date = msg["createdDateTime"][:10]

        try:
            msg_type = classify_type(content)
            if msg_type == "bug":
                data = classify_bug(content, author, date)
                data.update({"author": author, "date": date, "original": content, "msg_id": msg["id"]})
                bugs.append(data)
                print(f"  🐛 Bug: {data.get('title', content[:50])}")
            elif msg_type == "feedback":
                data = classify_feedback(content, author, date)
                data.update({"author": author, "date": date, "original": content, "msg_id": msg["id"]})
                feedbacks.append(data)
                print(f"  💬 Feedback: {data.get('title', content[:50])}")
            else:
                print(f"  ⏭️  Ignorado: {content[:60]}")
        except Exception as e:
            print(f"  Aviso: erro ao processar mensagem de {author}: {e}")

    if feedbacks:
        print(f"Inserindo {len(feedbacks)} feedback(s) no index.html...")
        update_html(feedbacks)

    if bugs or feedbacks:
        save_pending_review(bugs, feedbacks)

    # Marca TODAS as mensagens vistas como processadas (evita reprocessar em runs futuros).
    save_processed(processed | {m["id"] for m in messages})

    print(f"\n✓ {len(feedbacks)} feedback(s) inseridos no index.html (NÃO commitados).")
    if bugs:
        print(f"⚠ {len(bugs)} bug(s) precisam de triagem no Linear (issue NOD-xxx) — NÃO inseridos.")
        print("   Detalhes em pending_review.json.")
    print("\nPróximos passos:")
    print("  1. Revise:  python3 review.py")
    print("  2. No Claude Code: peça para revisar/ajustar os cards, criar issues dos bugs e então commitar + push.")


if __name__ == "__main__":
    main()

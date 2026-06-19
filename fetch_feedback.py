#!/usr/bin/env python3
"""
Lê o canal 'Feedback Plataforma' do Teams, classifica as mensagens,
adiciona cards no index.html, faz git commit e salva pending_review.json.
"""

import json
import os
import re
import subprocess
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

LLM_MODEL = os.getenv("LITELLM_MODEL", "gemini/gemini-3.1-flash-lite")
llm = OpenAI(
    api_key=os.environ["LITELLM_API_KEY"],
    base_url=os.getenv("LITELLM_BASE_URL", "https://llm.nodian.com.br"),
)

INDEX_HTML = Path(__file__).parent / "index.html"
PROCESSED_IDS_FILE = Path(__file__).parent / ".processed_message_ids.json"
PENDING_REVIEW_FILE = Path(__file__).parent / "pending_review.json"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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


def get_new_messages(token, team_id, channel_id, processed_ids):
    since = (datetime.now(timezone.utc) - timedelta(hours=24)).strftime("%Y-%m-%dT%H:%M:%SZ")
    data = graph_get(token, f"/teams/{team_id}/channels/{channel_id}/messages", params={"$top": 50})
    messages = data.get("value", [])
    candidates = [
        m for m in messages
        if m["id"] not in processed_ids
        and m.get("messageType") == "message"
        and m.get("body", {}).get("content", "").strip()
        and m.get("from") is not None
        and m.get("createdDateTime", "") >= since
    ]
    # deduplicar por conteúdo (evita processar edições como mensagem nova)
    seen_content = set()
    unique = []
    for m in candidates:
        content = re.sub(r"<[^>]+>", "", m["body"]["content"]).strip()
        if content not in seen_content:
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


def build_mermaid_flow(nodes):
    lines = ["flowchart TD"]
    for i, edge in enumerate(nodes):
        src = f'N{i}["{edge["from"]}"]'
        dst = f'N{i+1}["{edge["to"]}"]'
        lines.append(f"    {src} --> {dst}")
    return "\n".join(lines)


def bug_card(item):
    priority_class = {"Alto": "priority-high", "Médio": "priority-medium", "Baixo": "priority-low"}.get(
        item.get("priority", "Médio"), "priority-medium"
    )
    steps_html = "\n".join(f"      <li>{s}</li>" for s in item.get("steps", []))
    criteria_html = "\n".join(f"      <li>{c}</li>" for c in item.get("acceptance_criteria", []))
    current_flow = build_mermaid_flow(item.get("current_flow", []))
    expected_flow = build_mermaid_flow(item.get("expected_flow", []))

    return f"""
  <div class="issue-card">
    <div class="issue-top">
      <div>
        <div class="issue-id">BUG · {item['date']} · {item['author']}</div>
        <div class="issue-title">{item['title']}</div>
      </div>
      <span class="priority-tag {priority_class}">{item.get('priority', 'Médio')}</span>
    </div>
    <p class="issue-summary"><em>"{item['original']}"</em></p>
    <p class="issue-summary">{item.get('user_story', '')}</p>

    <div class="flow-label">Passos para reproduzir</div>
    <ol style="margin: 8px 0 20px 20px; font-size: 14px; color: var(--stone-600); line-height: 1.9;">
{steps_html}
    </ol>

    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px;">
      <div>
        <div class="flow-label">Fluxo atual</div>
        <pre class="mermaid">{current_flow}</pre>
      </div>
      <div>
        <div class="flow-label">Fluxo esperado</div>
        <pre class="mermaid">{expected_flow}</pre>
      </div>
    </div>

    <div class="flow-label">Critérios de aceite</div>
    <ul style="margin: 8px 0 20px 20px; font-size: 14px; color: var(--stone-600); line-height: 1.9;">
{criteria_html}
    </ul>

    <div style="background: var(--red-100); border-radius: 8px; padding: 14px 18px; font-size: 13px; color: var(--stone-700); line-height: 1.6;">
      <strong>Consequência:</strong> {item.get('consequence', '')}
    </div>
  </div>"""


def feedback_card(item):
    return f"""
  <div class="issue-card">
    <div class="issue-top">
      <div>
        <div class="issue-id">FEEDBACK · {item['date']} · {item['author']}</div>
        <div class="issue-title">{item['title']}</div>
      </div>
      <span class="priority-tag priority-medium">{item.get('category', '💬 Feedback')}</span>
    </div>
    <p class="issue-summary"><em>"{item['original']}"</em></p>
    <p class="issue-summary">{item.get('summary', '')}</p>
  </div>"""


def update_html(bugs, feedbacks):
    html = INDEX_HTML.read_text(encoding="utf-8")

    if bugs:
        cards_html = "\n".join(bug_card(b) for b in bugs)
        html = html.replace(
            "  </div> <!-- close tab-bugs -->",
            f"{cards_html}\n  </div> <!-- close tab-bugs -->",
        )

    if feedbacks:
        cards_html = "\n".join(feedback_card(f) for f in feedbacks)
        html = html.replace(
            "  </div> <!-- close tab-feedback -->",
            f"{cards_html}\n  </div> <!-- close tab-feedback -->",
        )

    INDEX_HTML.write_text(html, encoding="utf-8")


def git_commit(n_bugs, n_feedbacks):
    try:
        subprocess.run(["git", "add", "index.html"], cwd=INDEX_HTML.parent, check=True)
        date_str = datetime.now().strftime("%d/%m/%Y")
        msg = f"feedback Teams {date_str}: {n_bugs} bug(s), {n_feedbacks} feedback(s)"
        subprocess.run(["git", "commit", "-m", msg], cwd=INDEX_HTML.parent, check=True)
        print(f"✓ Commit feito: {msg}")
    except subprocess.CalledProcessError as e:
        print(f"Aviso: git commit falhou — {e}")


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
    processed = load_processed()

    print("Autenticando com Microsoft Graph...")
    token = get_token()

    print(f"Localizando canal '{CHANNEL_NAME}'...")
    team_id, channel_id = find_channel(token)

    print("Buscando mensagens das últimas 24h...")
    messages = get_new_messages(token, team_id, channel_id, processed)

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

    if bugs or feedbacks:
        print("Atualizando index.html...")
        update_html(bugs, feedbacks)
        git_commit(len(bugs), len(feedbacks))
        save_pending_review(bugs, feedbacks)
        print(f"\n✓ {len(bugs)} bug(s) e {len(feedbacks)} feedback(s) adicionados.")
        print("  Abra o Claude Code e rode: python3 review.py")
        print("  Para revisar e publicar no Vercel.")

    save_processed(processed | {m["id"] for m in messages})


if __name__ == "__main__":
    main()

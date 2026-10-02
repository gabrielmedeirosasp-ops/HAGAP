"""
HAGAP - gerador.py
Compatibilidade segura para o programa ENVIAR_PROGRAMADOS_DRIVE.

Este arquivo NÃO contém credenciais embutidas.

Autenticação aceita:
1. Variável GOOGLE_APPLICATION_CREDENTIALS apontando para um JSON de conta de serviço; ou
2. credenciais_google.json na mesma pasta; ou
3. service_account.json na mesma pasta.

A conta usada precisa ter acesso à pasta do Google Drive do HAGAP.
"""

from __future__ import annotations

import os
from pathlib import Path


def _get_drive_service_local():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as exc:
        print(
            "[Drive] Bibliotecas Google ausentes. Instale com:\n"
            "python -m pip install google-api-python-client google-auth"
        )
        print(f"[Drive] Detalhe: {exc}")
        return None

    base = Path(__file__).resolve().parent
    candidatos = []

    env = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if env:
        candidatos.append(Path(env))

    candidatos.extend([
        base / "credenciais_google.json",
        base / "service_account.json",
        base / "google_service_account.json",
    ])

    credencial = next((p for p in candidatos if p.exists() and p.is_file()), None)

    if not credencial:
        print("[Drive] Nenhuma credencial Google localizada.")
        print("[Drive] Coloque o JSON da conta de serviço na mesma pasta com o nome:")
        print("        credenciais_google.json")
        print("[Drive] ou configure GOOGLE_APPLICATION_CREDENTIALS no Windows.")
        return None

    try:
        creds = service_account.Credentials.from_service_account_file(
            str(credencial),
            scopes=["https://www.googleapis.com/auth/drive"],
        )

        service = build("drive", "v3", credentials=creds, cache_discovery=False)

        about = service.about().get(fields="user(emailAddress)").execute()
        email = (about.get("user") or {}).get("emailAddress") or "(conta de serviço)"
        print(f"[Drive] ✅ Conectado como {email}")
        print(f"[Drive] Credencial: {credencial}")

        return service

    except Exception as exc:
        print(f"[Drive] ❌ Falha ao autenticar no Google Drive: {exc}")
        return None


if __name__ == "__main__":
    svc = _get_drive_service_local()
    if svc:
        print("[Drive] Integração pronta.")
    else:
        print("[Drive] Integração não configurada.")

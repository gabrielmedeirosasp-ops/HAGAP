"""
gerar_base64.py — Execute no SEU PC para gerar a string de credenciais
Copie o resultado e cole na variável GOOGLE_SERVICE_ACCOUNT_JSON no Render
"""
import base64, json, os

# Caminho para o arquivo JSON da Service Account (baixado no Passo 2)
# Edite esta linha com o caminho correto no seu PC
CAMINHO_JSON = r"C:\Users\SEU_USUARIO\Downloads\service_account.json"

def main():
    if not os.path.exists(CAMINHO_JSON):
        print(f"❌ Arquivo não encontrado: {CAMINHO_JSON}")
        print("\n✅ Edite o arquivo e coloque o caminho correto do seu JSON.")
        print("   Exemplo: CAMINHO_JSON = r'C:\\Users\\João\\Downloads\\hagap-drive-xxxxx.json'")
        return

    with open(CAMINHO_JSON, "rb") as f:
        dados = f.read()

    encoded = base64.b64encode(dados).decode("ascii")

    print("\n" + "="*60)
    print("  ✅ BASE64 GERADO!")
    print("="*60)
    print("\n📋 COPIE O TEXTO ABAIXO (uma linha gigante):")
    print("-" * 60)
    print(encoded)
    print("-" * 60)
    print("\n📌 No Render, crie a variável de ambiente:")
    print("   Name:  GOOGLE_SERVICE_ACCOUNT_JSON")
    print("   Value: cole todo o texto acima")
    print("\n💡 Para facilitar, o texto também foi salvo em base64_output.txt")
    
    with open("base64_output.txt", "w", encoding="utf-8") as f:
        f.write(encoded)
    print("\n✅ Salvo em: base64_output.txt")

if __name__ == "__main__":
    main()
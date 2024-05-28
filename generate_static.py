from app import app
import os

# Defina o diretório onde os arquivos estáticos serão salvos
output_dir = 'static_site'

# Crie o diretório se não existir
os.makedirs(output_dir, exist_ok=True)

def save_static_html(url, filename):
    with app.test_client() as client:
        response = client.get(url)
        if response.status_code == 200:
            filepath = os.path.join(output_dir, filename)
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(response.get_data(as_text=True))
        else:
            print(f"Erro ao acessar {url}: {response.status_code}")

# Liste suas rotas aqui e o nome dos arquivos HTML
routes = {
    '/': 'index.html',
    '/about': 'about.html',
    # Adicione mais rotas conforme necessário
}

with app.app_context():
    for url, filename in routes.items():
        save_static_html(url, filename)

print("Site estático gerado com sucesso!")

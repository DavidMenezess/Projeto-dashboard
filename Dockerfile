# Imagem base leve com Python já instalado
FROM python:3.12-slim

# Evita arquivos .pyc e deixa os logs aparecerem em tempo real
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instala as dependências primeiro (camada separada do código):
# assim, se só o código mudar, o Docker reaproveita essa camada e o
# build fica bem mais rápido.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Agora copia o resto do projeto
COPY . .

# Pasta onde ficam o banco SQLite e as planilhas — precisa existir e
# ser gravável dentro do container.
RUN mkdir -p /app/data

EXPOSE 8000

# Sem --reload: isso é só para desenvolvimento local, não para produção.
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]

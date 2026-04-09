# Renumerador de Recibos - Vercel

## Rodar localmente
pip install -r requirements.txt
python app.py

Acesse:
http://127.0.0.1:5000

## Deploy na Vercel
1. Envie esta pasta para um repositório no GitHub.
2. Na Vercel, clique em Add New Project.
3. Importe o repositório.
4. Mantenha o Framework Preset como Other ou deixe a detecção automática.
5. Deploy.

A aplicação expõe um app Flask WSGI no arquivo `app.py`, formato aceito pelo Python Runtime da Vercel.

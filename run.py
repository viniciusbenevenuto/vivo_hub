"""run.py — Entrypoint para desenvolvimento local e Docker.

Uso local:
    python run.py

Uso com Gunicorn (produção/Docker):
    gunicorn -w 2 -b 0.0.0.0:5000 run:application
"""
import os
from app import create_app

application = create_app()   # nome compatível com WSGI (Gunicorn/uWSGI)

if __name__ == "__main__":
    # HOST 0.0.0.0 obrigatório para funcionar em Docker/WSL/VMs
    port = int(os.environ.get("PORT", 5000))
    application.run(
        host="0.0.0.0",
        port=port,
        debug=False,
        use_reloader=False,
    )


"""routes/__init__.py — Registra todos os Blueprints na aplicação Flask."""
from flask import Flask


def register_blueprints(app: Flask) -> None:
    from routes.auth import bp as auth_bp
    from routes.central import bp as central_bp
    from routes.atacado import bp as atacado_bp
    from routes.engenharia import bp as engenharia_bp
    from routes.api import bp as api_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(central_bp)
    app.register_blueprint(atacado_bp)
    app.register_blueprint(engenharia_bp)
    app.register_blueprint(api_bp)

"""cli.py — Comandos CLI do Flask (flask create-user, flask sbc-check, etc.)."""
import logging
import sqlite3

import click
from flask import Flask
from werkzeug.security import generate_password_hash

from config import ANEXO5_XLSX_PATH
from db import apply_cn_metadata, get_db, seed_cns
from operadoras import import_anexo5
from sbc.analyzer import SBCAnalyzer

logger = logging.getLogger(__name__)


def register_cli_commands(app: Flask) -> None:

    # -------------------------------------------------------------------------
    @app.cli.command("create-user")
    @click.argument("email")
    @click.argument("password")
    @click.argument("role")
    def create_user_cmd(email: str, password: str, role: str) -> None:
        """Cria um novo usuário. ROLE deve ser 'engenharia' ou 'atacado'."""
        role = role.strip().lower()
        if role not in ("engenharia", "atacado"):
            click.echo("Role inválida. Use: engenharia ou atacado.")
            return
        db = get_db()
        try:
            db.execute(
                "INSERT INTO users (email, password_hash, role) VALUES (?, ?, ?)",
                (email.strip().lower(), generate_password_hash(password), role),
            )
            db.commit()
            click.echo(f"Usuário '{email}' criado com sucesso.")
        except sqlite3.IntegrityError:
            click.echo(f"E-mail '{email}' já cadastrado.")

    # -------------------------------------------------------------------------
    @app.cli.command("seed-cns")
    def seed_cns_cmd() -> None:
        """Aplica/atualiza o seed de CNs e metadados no banco."""
        db = get_db()
        seed_cns(db)
        apply_cn_metadata(db)
        click.echo("Seed de CNs aplicado/atualizado.")

    # -------------------------------------------------------------------------
    @app.cli.command("import-anexo5")
    @click.argument("xlsx_path", required=False)
    def import_anexo5_cmd(xlsx_path: str | None) -> None:
        """Reimporta o Anexo 5 (ABR Telecom) para a tabela de operadoras."""
        path = xlsx_path or ANEXO5_XLSX_PATH
        try:
            total = import_anexo5(get_db(), path)
            click.echo(f"{total} operadoras importadas de: {path}")
        except (OSError, ValueError) as exc:
            click.echo(f"Falha na importação: {exc}")

    # -------------------------------------------------------------------------
    @app.cli.command("sbc-check")
    def sbc_check_cmd() -> None:
        """Verifica o status do diretório e arquivo XLSX de SBCs."""
        analyzer: SBCAnalyzer = app.config["SBC_ANALYZER"]
        health = analyzer.health_check()
        status_label = "OK" if health["data_dir_exists"] else "NÃO ENCONTRADO"
        click.echo(f"Diretório: {health['data_dir']} ({status_label})")
        click.echo(f"XLSX mais recente: {health.get('latest_file', 'Nenhum')}")
        if health["file_info"].get("filename"):
            info = health["file_info"]
            click.echo(f"  Medições:         {info['total_measurements']}")
            click.echo(f"  Linhas agregadas: {info['total_aggregated']}")
            click.echo(f"  Modificado:       {info['modified_at']}")

    # -------------------------------------------------------------------------
    @app.cli.command("sbc-suggest")
    @click.argument("cn")
    def sbc_suggest_cmd(cn: str) -> None:
        """Sugere SBCs para o CN informado."""
        analyzer: SBCAnalyzer = app.config["SBC_ANALYZER"]
        result = analyzer.suggest_for_cn(cn)
        sep = "=" * 60
        click.echo(f"\n{sep}")
        click.echo(f"  SBC Suggestion para CN {result.cn} — {result.cidade} ({result.uf})")
        click.echo(f"  Regional: {result.regional}")
        click.echo(f"  Fonte:    {result.source_file} ({result.source_modificado_em})")
        click.echo(f"  Medição:  {result.data_medicao}")
        if result.fallback_usado:
            click.echo(f"  ⚠️  Fallback: {result.fallback_origem}")
        click.echo(sep)
        click.echo(f"  {result.mensagem}")
        click.echo(f"{sep}\n")

        if not result.sbcs:
            click.echo("  Nenhum SBC encontrado.")
            return

        for i, sbc in enumerate(result.sbcs):
            saude = sbc.get("saude", "")
            if sbc.get("recomendado"):
                icon = "✅"
            elif saude in ("moderado", "critico"):
                icon = "⚠️"
            else:
                icon = "  "

            click.echo(
                f"  {icon} #{i+1} {sbc['nome']} — "
                f"{sbc.get('cidade', '')} ({sbc['uf']})"
            )
            click.echo(
                f"     Modelo: {sbc['modelo']} | "
                f"Serviços: {', '.join(sbc.get('servicos', []))}"
            )
            click.echo(
                f"     CAPS avg: {sbc['caps_avg']} | "
                f"máx {sbc['caps_max']} | mín {sbc['caps_min']} | "
                f"último {sbc['caps_ultimo']}"
            )
            click.echo(
                f"     Tendência: {sbc['caps_tendencia']} | "
                f"Medições: {sbc['total_medicoes']} | "
                f"Status: {sbc['status_fonte']} | Saúde: {saude}"
            )
            rec_label = " 🏆 RECOMENDADO" if sbc.get("recomendado") else ""
            click.echo(f"     Score: {sbc['score']}/100{rec_label}")
            if sbc.get("responsavel"):
                click.echo(f"     Responsável: {sbc['responsavel']}")
            if sbc.get("prazo"):
                click.echo(f"     Prazo: {sbc['prazo']}")
            click.echo(f"     Motivo: {sbc['motivo']}")
            click.echo()

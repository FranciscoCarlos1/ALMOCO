from pathlib import Path

from flask import Blueprint, render_template, request, redirect, url_for, jsonify, abort, send_file
from datetime import date
from db import DB_DIR, get_conn

bp_main = Blueprint('main', __name__)

TURMAS = [
    "TIN I", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]
INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]
CARDAPIO_DIR = DB_DIR / "cardapio_imagens"

@bp_main.route("/")
def index():
    sucesso = request.args.get("sucesso") == "1"
    erro = request.args.get("erro")
    hoje = date.today().isoformat()
    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT descricao, imagem_path
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (hoje,),
        ).fetchone()
    return render_template(
        "index.html",
        turmas=TURMAS,
        intencoes=INTENCOES,
        sucesso=sucesso,
        erro=erro,
        hoje=hoje,
        cardapio_hoje=cardapio["descricao"] if cardapio else None,
        cardapio_imagem=cardapio["imagem_path"] if cardapio else None,
    )


@bp_main.route("/cardapio/imagens/<path:nome_arquivo>")
def cardapio_imagem(nome_arquivo: str):
    caminho = (CARDAPIO_DIR / nome_arquivo).resolve()
    diretorio = CARDAPIO_DIR.resolve()
    if diretorio not in caminho.parents or not caminho.is_file():
        abort(404)
    return send_file(Path(caminho))

@bp_main.route("/aluno")
def buscar_aluno():
    matricula = request.args.get("matricula", "").strip()
    if not matricula:
        return jsonify({"ok": False, "erro": "Matrícula não informada."}), 400
    with get_conn() as conn:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma
            FROM alunos
            WHERE matricula = ?
            """,
            (matricula,),
        ).fetchone()
    if not aluno:
        return jsonify({"ok": False, "erro": "Matrícula não encontrada."}), 404
    return jsonify({
        "ok": True,
        "nome": aluno["nome"],
        "matricula": aluno["matricula"],
        "turma": aluno["turma"],
    })

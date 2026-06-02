from io import BytesIO

from flask import Blueprint, render_template, request, redirect, url_for, jsonify, abort, send_file
from datetime import date
from db import get_conn

bp_main = Blueprint('main', __name__)

TURMAS = [
    "TIN I", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]
INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]

@bp_main.route("/")
def index():
    sucesso = request.args.get("sucesso") == "1"
    erro = request.args.get("erro")
    hoje = date.today().isoformat()
    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT descricao, imagem_blob
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
        cardapio_imagem=hoje if cardapio and cardapio["imagem_blob"] else None,
    )


@bp_main.route("/cardapio/imagens/<path:nome_arquivo>")
def cardapio_imagem(nome_arquivo: str):
    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT imagem_blob, imagem_mime
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (nome_arquivo,),
        ).fetchone()

    if not cardapio or not cardapio["imagem_blob"]:
        abort(404)

    imagem_blob = cardapio["imagem_blob"]
    if isinstance(imagem_blob, memoryview):
        imagem_blob = imagem_blob.tobytes()

    return send_file(
        BytesIO(imagem_blob),
        mimetype=cardapio["imagem_mime"] or "application/octet-stream",
    )

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

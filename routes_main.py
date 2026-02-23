from flask import Blueprint, render_template, request, redirect, url_for, jsonify, abort
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
    return render_template(
        "index.html",
        turmas=TURMAS,
        intencoes=INTENCOES,
        sucesso=sucesso,
        erro=erro,
        hoje=hoje,
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

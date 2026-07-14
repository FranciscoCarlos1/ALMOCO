from io import BytesIO
from pathlib import Path
import os

from flask import Blueprint, render_template, request, redirect, url_for, jsonify, abort, send_file
from datetime import date, datetime, timedelta
from db import get_conn

bp_main = Blueprint('main', __name__)

TURMAS = [
    "TIN I A", "TIN I B", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]
INTENCOES = ["SIM", "NAO"]
DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex"]
EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".webp", ".gif"}


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def week_start(given_date: date) -> date:
    return given_date - timedelta(days=given_date.weekday())


def obter_biometria_token() -> str:
    return os.getenv("ALMOCO_BIOMETRIA_TOKEN") or os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")


def limpar_cpf(valor: str) -> str:
    return "".join(caractere for caractere in valor if caractere.isdigit())


def buscar_aluno_por_identificador(conn, matricula: str, cpf: str, identificador_biometrico: str):
    if identificador_biometrico:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma, cpf, identificador_biometrico
            FROM alunos
            WHERE identificador_biometrico = ?
            """,
            (identificador_biometrico,),
        ).fetchone()
        if aluno:
            return aluno

    if matricula:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma, cpf, identificador_biometrico
            FROM alunos
            WHERE matricula = ?
            """,
            (matricula,),
        ).fetchone()
        if aluno:
            return aluno

    if cpf:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma, cpf, identificador_biometrico
            FROM alunos
            WHERE cpf = ?
            """,
            (cpf,),
        ).fetchone()
        if aluno:
            return aluno

    return None


def listar_galeria_imagens() -> list[str]:
    pasta_galeria = Path(__file__).resolve().parent / "static" / "galeria"
    if not pasta_galeria.exists():
        return []

    imagens = []
    for arquivo in sorted(pasta_galeria.iterdir()):
        if arquivo.is_file() and arquivo.suffix.lower() in EXTENSOES_IMAGEM:
            imagens.append(f"galeria/{arquivo.name}")
    return imagens

@bp_main.route("/")
def index():
    sucesso = request.args.get("sucesso") == "1"
    erro = request.args.get("erro")
    hoje_data = date.today()
    hoje = hoje_data.isoformat()
    dia_atual = DIAS_SEMANA[hoje_data.weekday()] if hoje_data.weekday() < len(DIAS_SEMANA) else None
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
        dia_atual=dia_atual,
        cardapio_hoje=cardapio["descricao"] if cardapio else None,
        cardapio_imagem=hoje if cardapio and cardapio["imagem_blob"] else None,
        galeria_imagens=listar_galeria_imagens(),
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
            SELECT nome, matricula, turma, cpf
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
        "cpf": aluno["cpf"],
    })


@bp_main.post("/biometria/registrar")
def registrar_biometria():
    payload = request.get_json(silent=True) or request.form
    token = (payload.get("token") or request.args.get("token") or "").strip()
    if token != obter_biometria_token():
        return jsonify({"ok": False, "erro": "Token biométrico inválido."}), 403

    identificador_biometrico = (payload.get("identificador_biometrico") or "").strip()
    matricula = (payload.get("matricula") or "").strip()
    cpf = limpar_cpf((payload.get("cpf") or "").strip())
    data_almoco = (payload.get("data_almoco") or date.today().isoformat()).strip()

    if not any([identificador_biometrico, matricula, cpf]):
        return jsonify({
            "ok": False,
            "erro": "Informe identificador_biometrico, matricula ou cpf para localizar o aluno.",
        }), 400

    try:
        data_registro = parse_iso_date(data_almoco)
    except ValueError:
        return jsonify({"ok": False, "erro": "Data inválida."}), 400

    with get_conn() as conn:
        aluno = buscar_aluno_por_identificador(conn, matricula, cpf, identificador_biometrico)
        if not aluno:
            return jsonify({"ok": False, "erro": "Aluno não encontrado para os dados informados."}), 404

        conn.execute(
            """
            INSERT INTO respostas (nome, matricula, turma, data_almoco, intencao)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(matricula, data_almoco)
            DO UPDATE SET
                nome = excluded.nome,
                turma = excluded.turma,
                intencao = 'SIM',
                criado_em = CURRENT_TIMESTAMP
            """,
            (aluno["nome"], aluno["matricula"], aluno["turma"], data_registro.isoformat(), "SIM"),
        )
        conn.commit()

    return jsonify({
        "ok": True,
        "mensagem": "Almoço registrado com sucesso.",
        "data_almoco": data_registro.isoformat(),
        "aluno": {
            "nome": aluno["nome"],
            "turma": aluno["turma"],
            "matricula": aluno["matricula"],
            "cpf": aluno["cpf"],
            "identificador_biometrico": aluno["identificador_biometrico"],
        },
    })


@bp_main.post("/enviar")
def enviar():
    nome = request.form.get("nome", "").strip()
    matricula = request.form.get("matricula", "").strip()
    turma = request.form.get("turma", "").strip()
    data_referencia = request.form.get("data_almoco", "").strip()
    dias_raw = request.form.getlist("dias")

    dias_marcados: list[str] = []
    for raw in dias_raw:
        normalizado = raw.replace(";", ",").replace(" ", ",")
        partes = [item.strip().lower() for item in normalizado.split(",") if item.strip()]
        dias_marcados.extend(partes)
    dias_marcados = list(dict.fromkeys(dias_marcados))

    if not dias_marcados:
        dia_semana_atual = date.today().weekday()
        if dia_semana_atual < len(DIAS_SEMANA):
            dias_marcados = [DIAS_SEMANA[dia_semana_atual]]

    if not nome:
        return redirect(url_for("main.index", erro="Informe seu nome."))
    if turma not in TURMAS:
        return redirect(url_for("main.index", erro="Selecione uma turma válida."))
    if not dias_marcados:
        return redirect(url_for("main.index", erro="Hoje não é um dia letivo para envio automático."))
    if any(item not in DIAS_SEMANA for item in dias_marcados):
        return redirect(url_for("main.index", erro="Seleção de dias inválida."))

    if data_referencia:
        try:
            data_ref = parse_iso_date(data_referencia)
        except ValueError:
            return redirect(url_for("main.index", erro="Informe uma data válida."))
    else:
        data_ref = date.today()

    if not matricula:
        matricula = f"AUTO::{turma}::{nome}".upper()
    segunda = week_start(data_ref)
    datas_semana = {
        "seg": segunda,
        "ter": segunda + timedelta(days=1),
        "qua": segunda + timedelta(days=2),
        "qui": segunda + timedelta(days=3),
        "sex": segunda + timedelta(days=4),
    }

    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO alunos (matricula, nome, turma)
            VALUES (?, ?, ?)
            ON CONFLICT(matricula)
            DO UPDATE SET
                nome = excluded.nome,
                turma = excluded.turma,
                atualizado_em = CURRENT_TIMESTAMP
            """,
            (matricula, nome, turma),
        )

        for dia, data_almoco in datas_semana.items():
            intencao = "SIM" if dia in dias_marcados else "NAO"
            conn.execute(
                """
                INSERT INTO respostas (nome, matricula, turma, data_almoco, intencao)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(matricula, data_almoco)
                DO UPDATE SET
                    nome = excluded.nome,
                    turma = excluded.turma,
                    intencao = excluded.intencao,
                    criado_em = CURRENT_TIMESTAMP
                """,
                (nome, matricula, turma, data_almoco.isoformat(), intencao),
            )
        conn.commit()

    return redirect(url_for("main.index", sucesso=1))


@bp_main.get("/enviar")
def enviar_redirect():
    return redirect(url_for("main.index"))

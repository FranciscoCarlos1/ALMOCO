"""
Rotas para integração com porteiro e biometria.
Permite capturar biometria de alunos e registrar presença automática.
"""
import os
from datetime import date
from flask import Blueprint, request, jsonify
from db import get_conn

bp_biometria = Blueprint('biometria', __name__)

PORTEIRO_TOKEN = os.getenv("ALMOCO_PORTEIRO_TOKEN", "porteiro-sbs")


def normalize_rfid_uid(raw_uid: str) -> str:
    """Normaliza UID de RFID removendo caracteres especiais e espaços."""
    return "".join(c for c in raw_uid.upper() if c.isalnum())


@bp_biometria.post("/porteiro/consulta")
def porteiro_consulta():
    """
    Consulta aluno por UID de RFID ou código de biometria.
    Retorna dados do aluno para confirmação visual no porteiro.
    
    Payload esperado:
    {
        "token": "porteiro-sbs",
        "uid": "DF8C5A12",  # UID do cartão RFID
        "or" pode enviar: "biometria_id", "matricula", "cpf"
    }
    """
    payload = request.get_json(silent=True) or {}
    token = (payload.get("token") or "").strip()
    
    if token != PORTEIRO_TOKEN:
        return jsonify({"ok": False, "erro": "Token inválido"}), 403
    
    uid = (payload.get("uid") or payload.get("rfid_uid") or "").strip()
    matricula = (payload.get("matricula") or "").strip()
    cpf = (payload.get("cpf") or "").strip()
    identificador_biometrico = (payload.get("biometria_id") or payload.get("identificador_biometrico") or "").strip()
    
    if not any([uid, matricula, cpf, identificador_biometrico]):
        return jsonify({
            "ok": False,
            "erro": "Informe uid, matricula, cpf ou identificador_biometrico"
        }), 400
    
    # Normalizar RFID se fornecido
    if uid:
        uid = normalize_rfid_uid(uid)
    
    with get_conn() as conn:
        aluno = None
        
        # Buscar por RFID primeiro
        if uid:
            aluno = conn.execute(
                """
                SELECT nome, matricula, turma, cpf, identificador_biometrico
                FROM alunos
                WHERE identificador_biometrico = ?
                """,
                (uid,),
            ).fetchone()
        
        # Depois por identificador_biometrico
        if not aluno and identificador_biometrico:
            aluno = conn.execute(
                """
                SELECT nome, matricula, turma, cpf, identificador_biometrico
                FROM alunos
                WHERE identificador_biometrico = ?
                """,
                (identificador_biometrico,),
            ).fetchone()
        
        # Depois por matrícula
        if not aluno and matricula:
            aluno = conn.execute(
                """
                SELECT nome, matricula, turma, cpf, identificador_biometrico
                FROM alunos
                WHERE matricula = ?
                """,
                (matricula,),
            ).fetchone()
        
        # Finalmente por CPF
        if not aluno and cpf:
            aluno = conn.execute(
                """
                SELECT nome, matricula, turma, cpf, identificador_biometrico
                FROM alunos
                WHERE cpf = ?
                """,
                (cpf,),
            ).fetchone()
    
    if not aluno:
        return jsonify({
            "ok": False,
            "erro": "Aluno não encontrado",
            "uid": uid
        }), 404
    
    return jsonify({
        "ok": True,
        "aluno": {
            "nome": aluno["nome"],
            "matricula": aluno["matricula"],
            "turma": aluno["turma"],
            "cpf": aluno["cpf"] or "",
            "identificador_biometrico": aluno["identificador_biometrico"] or ""
        }
    })


@bp_biometria.post("/porteiro/registrar")
def porteiro_registrar():
    """
    Registra presença/almoço para o aluno identificado por RFID/biometria.
    Cria ou atualiza a entrada em 'respostas' com intencao='SIM'.
    
    Payload esperado:
    {
        "token": "porteiro-sbs",
        "uid": "DF8C5A12",
        "data_almoco": "2026-01-20" (opcional, hoje se omitido)
    }
    """
    payload = request.get_json(silent=True) or {}
    token = (payload.get("token") or "").strip()
    
    if token != PORTEIRO_TOKEN:
        return jsonify({"ok": False, "erro": "Token inválido"}), 403
    
    uid = (payload.get("uid") or payload.get("rfid_uid") or "").strip()
    data_almoco = (payload.get("data_almoco") or date.today().isoformat()).strip()
    
    if not uid:
        return jsonify({"ok": False, "erro": "UID do cartão RFID é obrigatório"}), 400
    
    # Normalizar RFID
    uid = normalize_rfid_uid(uid)
    
    # Buscar aluno
    with get_conn() as conn:
        aluno = conn.execute(
            """
            SELECT nome, matricula, turma
            FROM alunos
            WHERE identificador_biometrico = ?
            """,
            (uid,),
        ).fetchone()
    
    if not aluno:
        return jsonify({
            "ok": False,
            "erro": "Aluno não encontrado para este cartão RFID"
        }), 404
    
    # Registrar presença
    with get_conn() as conn:
        conn.execute(
            """
            INSERT INTO respostas (nome, matricula, turma, data_almoco, intencao)
            VALUES (?, ?, ?, ?, 'SIM')
            ON CONFLICT(matricula, data_almoco)
            DO UPDATE SET
                nome = excluded.nome,
                turma = excluded.turma,
                intencao = 'SIM',
                criado_em = CURRENT_TIMESTAMP
            """,
            (aluno["nome"], aluno["matricula"], aluno["turma"], data_almoco),
        )
        conn.commit()
    
    return jsonify({
        "ok": True,
        "mensagem": "Almoço registrado com sucesso",
        "aluno": {
            "nome": aluno["nome"],
            "matricula": aluno["matricula"],
            "turma": aluno["turma"]
        },
        "data_almoco": data_almoco
    })


@bp_biometria.post("/admin/cadastrar_biometria")
def cadastrar_biometria():
    """
    Cadastra UID de cartão RFID para um aluno.
    
    Payload esperado:
    {
        "token": "admin-token",
        "matricula": "2026001",
        "uid": "DF8C5A12",
        or
        "uid": "DF8C5A12",
        "cpf": "12345678900"
    }
    """
    payload = request.get_json(silent=True) or {}
    token = (payload.get("token") or "").strip()
    admin_token = os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")
    
    if token != admin_token:
        return jsonify({"ok": False, "erro": "Token de admin inválido"}), 403
    
    uid = (payload.get("uid") or payload.get("rfid_uid") or "").strip()
    matricula = (payload.get("matricula") or "").strip()
    cpf = (payload.get("cpf") or "").strip()
    
    if not uid:
        return jsonify({"ok": False, "erro": "UID do cartão é obrigatório"}), 400
    
    if not any([matricula, cpf]):
        return jsonify({"ok": False, "erro": "Informe matrícula ou CPF do aluno"}), 400
    
    # Normalizar RFID
    uid = normalize_rfid_uid(uid)
    
    with get_conn() as conn:
        # Buscar aluno
        aluno = None
        if matricula:
            aluno = conn.execute(
                """SELECT matricula, nome, turma FROM alunos WHERE matricula = ?""",
                (matricula,)
            ).fetchone()
        elif cpf:
            # Limpar CPF
            cpf_clean = "".join(c for c in cpf if c.isdigit())
            aluno = conn.execute(
                """SELECT matricula, nome, turma FROM alunos WHERE cpf = ?""",
                (cpf_clean,)
            ).fetchone()
        
        if not aluno:
            return jsonify({
                "ok": False,
                "erro": "Aluno não encontrado"
            }), 404
        
        # Atualizar com UID de RFID
        conn.execute(
            """
            UPDATE alunos
            SET identificador_biometrico = ?
            WHERE matricula = ?
            """,
            (uid, aluno["matricula"])
        )
        conn.commit()
    
    return jsonify({
        "ok": True,
        "mensagem": "Cartão RFID cadastrado com sucesso",
        "aluno": {
            "nome": aluno["nome"],
            "matricula": aluno["matricula"],
            "turma": aluno["turma"],
            "uid_rfid": uid
        }
    })


@bp_biometria.get("/admin/listar_biometrias")
def listar_biometrias():
    """
    Lista todos os alunos com seus UIDs de RFID cadastrados.
    Útil para visualizar o estado dos cadastros.
    """
    token = request.args.get("token", "")
    admin_token = os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")
    
    if token != admin_token:
        return jsonify({"ok": False, "erro": "Token inválido"}), 403
    
    with get_conn() as conn:
        alunos = conn.execute(
            """
            SELECT matricula, nome, turma, cpf, identificador_biometrico
            FROM alunos
            ORDER BY turma, nome
            """
        ).fetchall()
    
    alunos_list = [
        {
            "matricula": a["matricula"],
            "nome": a["nome"],
            "turma": a["turma"],
            "cpf": a["cpf"] or "",
            "uid_rfid": a["identificador_biometrico"] or "",
            "biometria_cadastrada": bool(a["identificador_biometrico"])
        }
        for a in alunos
    ]
    
    total = len(alunos_list)
    com_biometria = sum(1 for a in alunos_list if a["biometria_cadastrada"])
    
    return jsonify({
        "ok": True,
        "total_alunos": total,
        "com_biometria": com_biometria,
        "sem_biometria": total - com_biometria,
        "alunos": alunos_list
    })

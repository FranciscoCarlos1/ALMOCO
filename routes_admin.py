import os
from pathlib import Path
from uuid import uuid4

from flask import Blueprint, render_template, request, abort, redirect, url_for
from datetime import date
from werkzeug.utils import secure_filename

from db import DB_DIR, get_conn

bp_admin = Blueprint('admin', __name__)

ADMIN_TOKEN = os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")
CARDAPIO_DIR = DB_DIR / "cardapio_imagens"
ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif"}

TURMAS = [
    "TIN I", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]

@bp_admin.route("/admin")
def admin():
    token = request.args.get("token", "")
    if token != ADMIN_TOKEN:
        abort(403, "Acesso negado. Informe um token válido na URL.")
    data_filtro = request.args.get("data") or date.today().isoformat()
    cardapio_salvo = request.args.get("cardapio_salvo") == "1"

    with get_conn() as conn:
        cardapio = conn.execute(
            """
            SELECT descricao, imagem_path
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (data_filtro,),
        ).fetchone()

    return render_template(
        "admin.html",
        resumo={},
        token=token,
        data_filtro=data_filtro,
        periodo=request.args.get("periodo", "semana"),
        importado=False,
        import_error=None,
        importado_quadro=False,
        import_quadro_error=None,
        backup_restaurado=False,
        backup_restore_file=None,
        backup_restore_error=None,
        backup_manual=False,
        backup_manual_error=None,
        total_sim=0,
        total_nao=0,
        total_geral=0,
        periodo_label="Semana",
        periodo_inicio=data_filtro,
        periodo_fim=data_filtro,
        total_semana_periodo=0,
        total_mes_periodo=0,
        total_ano_periodo=0,
        total_periodo_sim=0,
        total_periodo_nao=0,
        semana_inicio=data_filtro,
        semana_fim=data_filtro,
        quadro_rows=[],
        semana_sim={"seg": 0, "ter": 0, "qua": 0, "qui": 0, "sex": 0},
        total_semana_geral=0,
        respostas=[],
        cardapio_texto=cardapio["descricao"] if cardapio else "",
        cardapio_salvo=cardapio_salvo,
        cardapio_url=url_for("admin.painel_cardapio", token=token, data=data_filtro),
    )


def _validar_token(token: str) -> None:
    if token != ADMIN_TOKEN:
        abort(403, "Acesso negado. Informe um token válido na URL.")


def _obter_cardapio(data_filtro: str):
    with get_conn() as conn:
        return conn.execute(
            """
            SELECT descricao, imagem_path
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (data_filtro,),
        ).fetchone()


def _salvar_imagem_cardapio(arquivo, data_filtro: str) -> str:
    nome_seguro = secure_filename(arquivo.filename or "")
    extensao = Path(nome_seguro).suffix.lower()
    if extensao not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError("Envie uma imagem PNG, JPG, JPEG, WEBP ou GIF.")

    CARDAPIO_DIR.mkdir(parents=True, exist_ok=True)
    nome_arquivo = f"{data_filtro}-{uuid4().hex}{extensao}"
    destino = CARDAPIO_DIR / nome_arquivo
    arquivo.save(destino)
    return nome_arquivo


def _remover_imagem(nome_arquivo: str | None) -> None:
    if not nome_arquivo:
        return
    caminho = CARDAPIO_DIR / nome_arquivo
    if caminho.exists():
        caminho.unlink()


@bp_admin.route("/admin/cardapio", methods=["GET", "POST"])
def painel_cardapio():
    token = request.values.get("token", "")
    _validar_token(token)

    data_filtro = request.values.get("data", "").strip() or date.today().isoformat()

    if request.method == "POST":
        descricao = request.form.get("descricao", "").strip()
        remover_imagem = request.form.get("remover_imagem") == "1"
        arquivo = request.files.get("imagem")
        cardapio_atual = _obter_cardapio(data_filtro)
        imagem_path = cardapio_atual["imagem_path"] if cardapio_atual else None

        try:
            if arquivo and arquivo.filename:
                nova_imagem = _salvar_imagem_cardapio(arquivo, data_filtro)
                _remover_imagem(imagem_path)
                imagem_path = nova_imagem
            elif remover_imagem:
                _remover_imagem(imagem_path)
                imagem_path = None
        except ValueError as exc:
            return render_template(
                "cardapio_admin.html",
                token=token,
                data_filtro=data_filtro,
                cardapio_texto=descricao,
                cardapio_imagem=imagem_path,
                cardapio_salvo=False,
                erro_cardapio=str(exc),
            )

        with get_conn() as conn:
            if descricao or imagem_path:
                conn.execute(
                    """
                    INSERT INTO cardapios (data_almoco, descricao, imagem_path, atualizado_em)
                    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(data_almoco) DO UPDATE SET
                        descricao = excluded.descricao,
                        imagem_path = excluded.imagem_path,
                        atualizado_em = CURRENT_TIMESTAMP
                    """,
                    (data_filtro, descricao, imagem_path),
                )
            else:
                conn.execute(
                    "DELETE FROM cardapios WHERE data_almoco = ?",
                    (data_filtro,),
                )
            conn.commit()

        return redirect(url_for("admin.painel_cardapio", token=token, data=data_filtro, salvo=1))

    cardapio = _obter_cardapio(data_filtro)
    return render_template(
        "cardapio_admin.html",
        token=token,
        data_filtro=data_filtro,
        cardapio_texto=cardapio["descricao"] if cardapio else "",
        cardapio_imagem=cardapio["imagem_path"] if cardapio else None,
        cardapio_salvo=request.args.get("salvo") == "1",
        erro_cardapio=None,
    )

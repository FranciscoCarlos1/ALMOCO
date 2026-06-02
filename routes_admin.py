import os

from flask import Blueprint, render_template, request, abort, redirect, url_for
from datetime import date

from db import get_conn

bp_admin = Blueprint('admin', __name__)

ADMIN_TOKEN = os.getenv("ALMOCO_ADMIN_TOKEN", "ifc-sbs")
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
            SELECT descricao, imagem_blob
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
        cardapio_imagem=data_filtro if cardapio and cardapio["imagem_blob"] else None,
        cardapio_url=url_for("admin.painel_cardapio", token=token, data=data_filtro),
    )


def _validar_token(token: str) -> None:
    if token != ADMIN_TOKEN:
        abort(403, "Acesso negado. Informe um token válido na URL.")


def _obter_cardapio(data_filtro: str):
    with get_conn() as conn:
        return conn.execute(
            """
            SELECT descricao, imagem_blob, imagem_mime
            FROM cardapios
            WHERE data_almoco = ?
            """,
            (data_filtro,),
        ).fetchone()


def _salvar_imagem_cardapio(arquivo) -> tuple[bytes, str]:
    nome_arquivo = arquivo.filename or ""
    extensao = os.path.splitext(nome_arquivo)[1].lower()
    if extensao not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError("Envie uma imagem PNG, JPG, JPEG, WEBP ou GIF.")

    mime_type = (arquivo.mimetype or "").strip() or {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif",
    }[extensao]
    return arquivo.read(), mime_type


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
        imagem_blob = cardapio_atual["imagem_blob"] if cardapio_atual else None
        imagem_mime = cardapio_atual["imagem_mime"] if cardapio_atual else None

        try:
            if arquivo and arquivo.filename:
                imagem_blob, imagem_mime = _salvar_imagem_cardapio(arquivo)
            elif remover_imagem:
                imagem_blob = None
                imagem_mime = None
        except ValueError as exc:
            return render_template(
                "cardapio_admin.html",
                token=token,
                data_filtro=data_filtro,
                cardapio_texto=descricao,
                cardapio_imagem=data_filtro if imagem_blob else None,
                cardapio_salvo=False,
                erro_cardapio=str(exc),
            )

        with get_conn() as conn:
            if descricao or imagem_blob:
                conn.execute(
                    """
                    INSERT INTO cardapios (data_almoco, descricao, imagem_blob, imagem_mime, atualizado_em)
                    VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(data_almoco) DO UPDATE SET
                        descricao = excluded.descricao,
                        imagem_blob = excluded.imagem_blob,
                        imagem_mime = excluded.imagem_mime,
                        atualizado_em = CURRENT_TIMESTAMP
                    """,
                    (data_filtro, descricao, imagem_blob, imagem_mime),
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
        cardapio_imagem=data_filtro if cardapio and cardapio["imagem_blob"] else None,
        cardapio_salvo=request.args.get("salvo") == "1",
        erro_cardapio=None,
    )

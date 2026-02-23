from flask import Blueprint, render_template, request, abort
from datetime import date
from db import get_conn

bp_admin = Blueprint('admin', __name__)

TURMAS = [
    "TIN I", "TIN II", "TIN III",
    "TAI I", "TAI II", "TAI III",
    "TST I", "TST II", "TST III", "SERVIDORES"
]

@bp_admin.route("/admin")
def admin():
    # Exemplo simplificado, adapte conforme sua lógica
    if request.args.get("token") != "ifc-sbs":
        abort(403, "Acesso negado. Informe um token válido na URL.")
    data_filtro = request.args.get("data") or date.today().isoformat()
    # ... restante da lógica ...
    return render_template("admin.html", resumo={}, data_filtro=data_filtro)

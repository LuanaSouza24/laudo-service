import os
import re
import base64
import tempfile
import shutil
import importlib
import uuid
import time
import threading
from io import BytesIO
from typing import Optional, List
from urllib.parse import quote

import requests
from PIL import Image
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse
from pydantic import BaseModel

app = FastAPI()

# ============================================================
# ARMAZENAMENTO DE JOBS (em memória)
# Cada job tem: status, result, error, criado_em
# ============================================================
_jobs: dict = {}
_jobs_lock = threading.Lock()

JOB_TTL_SEGUNDOS = 3600  # remove jobs com mais de 1 hora automaticamente


def _limpar_jobs_antigos():
    agora = time.time()
    with _jobs_lock:
        expirados = [jid for jid, j in _jobs.items()
                     if agora - j["criado_em"] > JOB_TTL_SEGUNDOS]
        for jid in expirados:
            del _jobs[jid]
    if expirados:
        print(f"[INFO] {len(expirados)} job(s) expirado(s) removido(s).")


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None
    image_paths: Optional[List[str]] = None


def normalizar_rel_path(path: str) -> str:
    """
    Normaliza o caminho relativo vindo do Power Automate / Excel script.
    Remove barras invertidas, barras duplicadas e espacos laterais.
    """
    if not path:
        return ""
    rel = str(path).strip().replace("\\", "/")
    rel = re.sub(r"/+", "/", rel)
    rel = rel.lstrip("/")
    return rel


def montar_url_sharepoint(rel_path: str) -> str:
    base_url = os.getenv("SHAREPOINT_BASE_URL", "").strip().rstrip("/")
    if not base_url:
        raise Exception("Variavel de ambiente SHAREPOINT_BASE_URL nao configurada.")
    rel = normalizar_rel_path(rel_path)
    if not rel:
        raise Exception("Caminho da imagem vazio ao montar URL do SharePoint.")
    return f"{base_url}/{quote(rel, safe='/:()_- ')}"


def obter_headers_sharepoint() -> dict:
    token = os.getenv("SHAREPOINT_BEARER_TOKEN", "").strip()
    if not token:
        raise Exception("Variavel de ambiente SHAREPOINT_BEARER_TOKEN nao configurada.")
    return {"Authorization": f"Bearer {token}"}


def baixar_arquivo_sharepoint(rel_path: str, destino_abs: str):
    url = montar_url_sharepoint(rel_path)
    headers = obter_headers_sharepoint()
    resp = requests.get(url, headers=headers, stream=True, timeout=180)
    if resp.status_code != 200:
        raise Exception(
            f"Erro ao baixar arquivo do SharePoint. "
            f"Status={resp.status_code}, path={rel_path}, url={url}"
        )
    os.makedirs(os.path.dirname(destino_abs), exist_ok=True)
    with open(destino_abs, "wb") as f:
        for chunk in resp.iter_content(chunk_size=1024 * 1024):
            if chunk:
                f.write(chunk)


def eh_imagem(caminho: str) -> bool:
    ext = os.path.splitext(caminho)[1].lower()
    return ext in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


def comprimir_imagem_no_mesmo_arquivo(caminho: str, max_lado: int = 1600, qualidade: int = 75):
    if not os.path.exists(caminho):
        return
    if not eh_imagem(caminho):
        return
    try:
        with Image.open(caminho) as img:
            img.thumbnail((max_lado, max_lado))
            if img.mode in ("RGBA", "LA", "P"):
                fundo = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                fundo.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
                img = fundo
            elif img.mode != "RGB":
                img = img.convert("RGB")
            ext = os.path.splitext(caminho)[1].lower()
            if ext in {".jpg", ".jpeg"}:
                img.save(caminho, format="JPEG", quality=qualidade, optimize=True)
                return
            novo_caminho = os.path.splitext(caminho)[0] + ".jpg"
            img.save(novo_caminho, format="JPEG", quality=qualidade, optimize=True)
        if novo_caminho != caminho and os.path.exists(caminho):
            os.remove(caminho)
    except Exception as e:
        print(f"[AVISO] Falha ao comprimir imagem '{caminho}': {e}")


def preparar_excel(work_dir: str, excel_base64: str) -> str:
    excel_path = os.path.join(work_dir, "Cautelar.xlsx")
    with open(excel_path, "wb") as f:
        f.write(base64.b64decode(excel_base64))
    return excel_path


def preparar_template(work_dir: str, template_base64: Optional[str]) -> str:
    dst_template = os.path.join(work_dir, "tamplete.docx")
    if template_base64:
        with open(dst_template, "wb") as f:
            f.write(base64.b64decode(template_base64))
    else:
        template_src = os.path.join(os.path.dirname(__file__), "tamplete.docx")
        if not os.path.exists(template_src):
            raise Exception("Template nao enviado e arquivo tamplete.docx nao encontrado no repositorio.")
        shutil.copy(template_src, dst_template)
    return dst_template


def preparar_imagens(work_dir: str, image_paths: Optional[List[str]]):
    if not image_paths:
        print("[INFO] Nenhuma imagem recebida em image_paths.")
        return
    total = len(image_paths)
    print(f"[INFO] Iniciando download de {total} imagem(ns).")
    for i, original_path in enumerate(image_paths, start=1):
        rel = normalizar_rel_path(original_path)
        if not rel:
            print(f"[AVISO] Caminho vazio ignorado no indice {i}.")
            continue
        destino_abs = os.path.join(work_dir, rel)
        print(f"[INFO] [{i}/{total}] Baixando: {rel}")
        baixar_arquivo_sharepoint(rel, destino_abs)
        if eh_imagem(destino_abs):
            comprimir_imagem_no_mesmo_arquivo(destino_abs)
    print("[INFO] Download/preparo de imagens concluido.")


def localizar_docx_gerado(work_dir: str) -> str:
    out_dir = os.path.join(work_dir, "saida")
    if not os.path.isdir(out_dir):
        raise Exception(f"Pasta de saida nao encontrada: {out_dir}")
    docx_files = [
        os.path.join(out_dir, nome)
        for nome in os.listdir(out_dir)
        if nome.lower().endswith(".docx")
    ]
    if not docx_files:
        raise Exception("Laudo nao foi gerado (nenhum .docx na pasta de saida).")
    docx_files.sort(key=os.path.getmtime, reverse=True)
    return docx_files[0]


def gerar_laudo_no_modulo(id_vistoria: str):
    import gerar_laudo as gl
    importlib.reload(gl)
    gl.gerar_laudo(id_vistoria)


# ============================================================
# PROCESSAMENTO EM BACKGROUND
# Roda em thread separada para nao bloquear o Power Automate
# ============================================================

def _processar_job(job_id: str, p: Payload):
    """Executa todo o processamento do laudo em background."""
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        with _jobs_lock:
            _jobs[job_id]["status"] = "running"

        print(f"========== JOB {job_id} INICIADO ==========")
        print(f"[INFO] id_vistoria={p.id_vistoria}")
        print(f"[INFO] image_paths={len(p.image_paths or [])}")

        os.environ["LAUDO_BASE_DIR"] = work

        preparar_excel(work, p.excel_base64)
        preparar_template(work, p.template_base64)
        preparar_imagens(work, p.image_paths)
        gerar_laudo_no_modulo(p.id_vistoria)

        out_path = localizar_docx_gerado(work)
        filename = os.path.basename(out_path)

        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        with _jobs_lock:
            _jobs[job_id]["status"] = "done"
            _jobs[job_id]["result"] = {"filename": filename, "docx_base64": docx_b64}

        print(f"[INFO] JOB {job_id} CONCLUIDO: {filename}")

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(f"=== ERRO NO JOB {job_id} ===")
        print(tb)
        with _jobs_lock:
            _jobs[job_id]["status"] = "error"
            _jobs[job_id]["error"] = str(e)

    finally:
        shutil.rmtree(work, ignore_errors=True)
        _limpar_jobs_antigos()


# ============================================================
# ENDPOINTS
# ============================================================

@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate")
def generate(p: Payload, background_tasks: BackgroundTasks):
    """
    Recebe os dados, dispara o processamento em background
    e retorna imediatamente o job_id.
    O Power Automate consulta /status/{job_id} ate receber "done".
    """
    job_id = str(uuid.uuid4())

    with _jobs_lock:
        _jobs[job_id] = {
            "status": "pending",
            "result": None,
            "error": None,
            "criado_em": time.time(),
        }

    background_tasks.add_task(_processar_job, job_id, p)

    print(f"[INFO] Job criado: {job_id} para vistoria {p.id_vistoria}")
    return JSONResponse({"job_id": job_id}, status_code=202)


@app.get("/status/{job_id}")
def status(job_id: str):
    """
    Retorna o status atual do job.
    Valores possiveis: pending | running | done | error
    """
    with _jobs_lock:
        job = _jobs.get(job_id)

    if job is None:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    resposta = {"status": job["status"]}
    if job["status"] == "error":
        resposta["error"] = job["error"]

    return JSONResponse(resposta)


@app.get("/result/{job_id}")
def result(job_id: str):
    """
    Retorna o laudo gerado (filename + docx_base64).
    So disponivel quando status = done.
    Remove o job da memoria apos a entrega.
    """
    with _jobs_lock:
        job = _jobs.get(job_id)

    if job is None:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    if job["status"] != "done":
        raise HTTPException(
            status_code=400,
            detail=f"Job ainda nao concluido. Status atual: {job['status']}"
        )

    result_data = job["result"]

    with _jobs_lock:
        _jobs.pop(job_id, None)

    return JSONResponse(result_data)
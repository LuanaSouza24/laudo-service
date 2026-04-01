import os
import re
import json
import base64
import tempfile
import shutil
import importlib
import uuid
import time
import threading
from typing import Optional, List
from urllib.parse import quote

import requests
from PIL import Image
from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse
from pydantic import BaseModel

app = FastAPI()

JOBS_FILE = "/tmp/laudo_jobs.json"
_jobs_lock = threading.Lock()
JOB_TTL_SEGUNDOS = 3600


# ─────────────────────── Jobs ───────────────────────

def _ler_jobs() -> dict:
    if not os.path.exists(JOBS_FILE):
        return {}
    try:
        with open(JOBS_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def _salvar_jobs(jobs: dict):
    try:
        with open(JOBS_FILE, "w") as f:
            json.dump(jobs, f)
    except Exception as e:
        print(f"[AVISO] Nao foi possivel salvar jobs: {e}")


def _get_job(job_id: str) -> dict:
    with _jobs_lock:
        return _ler_jobs().get(job_id)


def _set_job(job_id: str, dados: dict):
    with _jobs_lock:
        jobs = _ler_jobs()
        jobs[job_id] = dados
        _salvar_jobs(jobs)


def _delete_job(job_id: str):
    with _jobs_lock:
        jobs = _ler_jobs()
        jobs.pop(job_id, None)
        _salvar_jobs(jobs)


def _limpar_jobs_antigos():
    agora = time.time()
    with _jobs_lock:
        jobs = _ler_jobs()
        expirados = [jid for jid, j in jobs.items()
                     if agora - j.get("criado_em", 0) > JOB_TTL_SEGUNDOS]
        for jid in expirados:
            del jobs[jid]
        if expirados:
            _salvar_jobs(jobs)
            print(f"[INFO] {len(expirados)} job(s) expirado(s) removido(s).")


# ─────────────────────── Utilitários ───────────────────────

def normalizar_rel_path(path: str) -> str:
    if not path:
        return ""
    rel = str(path).strip().replace("\\", "/")
    rel = re.sub(r"/+", "/", rel)
    rel = rel.lstrip("/")
    return rel


def eh_imagem(caminho: str) -> bool:
    ext = os.path.splitext(caminho)[1].lower()
    return ext in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}


def comprimir_imagem_no_mesmo_arquivo(caminho: str, max_lado: int = 1600, qualidade: int = 75):
    if not os.path.exists(caminho) or not eh_imagem(caminho):
        return
    try:
        novo_caminho = caminho
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
            raise Exception("Template nao enviado e tamplete.docx nao encontrado no repositorio.")
        shutil.copy(template_src, dst_template)
    return dst_template


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


# ─────────────────────── Processamento ───────────────────────

def _processar_job_v2(job_id: str):
    job = _get_job(job_id)
    if not job:
        print(f"[ERRO] Job {job_id} nao encontrado para processar.")
        return

    work = job.get("work_dir", "")
    try:
        _set_job(job_id, {**job, "status": "running"})

        print(f"========== JOB {job_id} GERANDO ==========")
        print(f"[INFO] id_vistoria={job['id_vistoria']}")

        os.environ["LAUDO_BASE_DIR"] = work

        preparar_excel(work, job["excel_base64"])
        preparar_template(work, job.get("template_base64") or None)

        gerar_laudo_no_modulo(job["id_vistoria"])

        out_path = localizar_docx_gerado(work)
        filename = os.path.basename(out_path)

        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        _set_job(job_id, {
            "status": "done",
            "result": {"filename": filename, "docx_base64": docx_b64},
            "error": None,
            "criado_em": job.get("criado_em", time.time()),
            "work_dir": "",
            "id_vistoria": job["id_vistoria"],
            "excel_base64": "",
            "template_base64": ""
        })
        print(f"[INFO] JOB {job_id} CONCLUIDO: {filename}")

    except Exception as e:
        import traceback
        print(f"=== ERRO NO JOB {job_id} ===")
        print(traceback.format_exc())
        _set_job(job_id, {**job, "status": "error", "error": str(e)})

    finally:
        if work and os.path.isdir(work):
            shutil.rmtree(work, ignore_errors=True)
        _limpar_jobs_antigos()


# ─────────────────────── Models ───────────────────────

class PayloadIniciar(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None


class PayloadFoto(BaseModel):
    path: str
    b64: str


# ─────────────────────── Endpoints ───────────────────────

@app.get("/health")
def health():
    return {"ok": True}


@app.post("/iniciar")
async def iniciar(p: PayloadIniciar):
    """Cria o job e reserva diretório de trabalho. Retorna job_id."""
    job_id = str(uuid.uuid4())
    work = tempfile.mkdtemp(prefix="laudo_")

    _set_job(job_id, {
        "status": "aguardando_fotos",
        "work_dir": work,
        "id_vistoria": p.id_vistoria,
        "excel_base64": p.excel_base64,
        "template_base64": p.template_base64 or "",
        "result": None,
        "error": None,
        "criado_em": time.time()
    })

    print(f"[INFO] Job iniciado: {job_id} | vistoria={p.id_vistoria}")
    return JSONResponse({"job_id": job_id}, status_code=202)


@app.post("/foto/{job_id}")
async def receber_foto(job_id: str, p: PayloadFoto):
    """Recebe uma foto por vez (base64) e salva no diretório do job."""
    job = _get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    work = job.get("work_dir", "")
    if not work or not os.path.isdir(work):
        raise HTTPException(status_code=400, detail="Diretorio de trabalho invalido ou expirado.")

    rel = normalizar_rel_path(p.path)
    if not rel:
        raise HTTPException(status_code=400, detail="Campo 'path' invalido ou vazio.")
    if not p.b64:
        raise HTTPException(status_code=400, detail="Campo 'b64' invalido ou vazio.")

    destino = os.path.join(work, rel)
    os.makedirs(os.path.dirname(destino), exist_ok=True)

    with open(destino, "wb") as f:
        f.write(base64.b64decode(p.b64))

    if eh_imagem(destino):
        comprimir_imagem_no_mesmo_arquivo(destino)

    print(f"[INFO] Foto salva: {rel}")
    return JSONResponse({"ok": True, "path": rel})


@app.post("/gerar/{job_id}")
async def gerar(job_id: str, background_tasks: BackgroundTasks):
    """Dispara a geração do laudo após todas as fotos terem sido enviadas."""
    job = _get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    if job.get("status") not in ("aguardando_fotos", "pending"):
        raise HTTPException(
            status_code=400,
            detail=f"Job em status inesperado: {job.get('status')}. Esperado: aguardando_fotos."
        )

    background_tasks.add_task(_processar_job_v2, job_id)
    print(f"[INFO] Geracao disparada para job {job_id}")
    return JSONResponse({"ok": True, "job_id": job_id}, status_code=202)


@app.get("/status/{job_id}")
def status(job_id: str):
    """Retorna o status atual do job."""
    job = _get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")
    resposta = {"status": job["status"]}
    if job["status"] == "error":
        resposta["error"] = job["error"]
    return JSONResponse(resposta)


@app.get("/result/{job_id}")
def result(job_id: str):
    """Retorna o laudo gerado em base64 e remove o job."""
    job = _get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job nao encontrado.")
    if job["status"] != "done":
        raise HTTPException(
            status_code=400,
            detail=f"Job ainda nao concluido. Status atual: {job['status']}"
        )
    result_data = job["result"]
    _delete_job(job_id)
    return JSONResponse(result_data)

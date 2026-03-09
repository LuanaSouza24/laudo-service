import os
import base64
import tempfile
import shutil
import importlib
import requests

from urllib.parse import quote
from typing import Optional, List

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel

app = FastAPI()


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None
    image_paths: Optional[List[str]] = None


def baixar_arquivo_sharepoint(url_ou_path: str, destino: str):
    sharepoint_base = os.getenv("SHAREPOINT_BASE_URL", "").rstrip("/")
    sharepoint_token = os.getenv("SHAREPOINT_BEARER_TOKEN", "")

    if not sharepoint_base:
        raise Exception("SHAREPOINT_BASE_URL não configurado.")
    if not sharepoint_token:
        raise Exception("SHAREPOINT_BEARER_TOKEN não configurado.")

    rel = (url_ou_path or "").strip()

    if rel.lower().startswith("http://") or rel.lower().startswith("https://"):
        url = rel
    else:
        if not rel.startswith("/"):
            rel = "/" + rel
        url = f"{sharepoint_base}{quote(rel, safe='/:()_- ')}"

    headers = {
        "Authorization": f"Bearer {sharepoint_token}"
    }

    r = requests.get(url, headers=headers, stream=True, timeout=120)
    if r.status_code != 200:
        raise Exception(f"Erro ao baixar imagem do SharePoint: {r.status_code} - {url}")

    os.makedirs(os.path.dirname(destino), exist_ok=True)
    with open(destino, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)


@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        os.environ["LAUDO_BASE_DIR"] = work

        excel_path = os.path.join(work, "Cautelar.xlsx")
        with open(excel_path, "wb") as f:
            f.write(base64.b64decode(p.excel_base64))

        dst_template = os.path.join(work, "tamplete.docx")
        if p.template_base64:
            with open(dst_template, "wb") as f:
                f.write(base64.b64decode(p.template_base64))
        else:
            template_src = os.path.join(os.path.dirname(__file__), "tamplete.docx")
            shutil.copy(template_src, dst_template)

        if p.image_paths:
            for path in p.image_paths:
                rel = (path or "").lstrip("/").replace("\\", "/")
                abs_path = os.path.join(work, rel)
                baixar_arquivo_sharepoint(path, abs_path)

        import gerar_laudo as gl
        importlib.reload(gl)

        gl.gerar_laudo(p.id_vistoria)

        out_dir = os.path.join(work, "saida")
        if not os.path.isdir(out_dir):
            raise HTTPException(500, f"Pasta de saída não encontrada: {out_dir}")

        docx_files = [x for x in os.listdir(out_dir) if x.lower().endswith(".docx")]
        if not docx_files:
            raise HTTPException(500, "Laudo não foi gerado (nenhum .docx na pasta de saída).")

        docx_files.sort(key=lambda n: os.path.getmtime(os.path.join(out_dir, n)), reverse=True)
        filename = docx_files[0]
        out_path = os.path.join(out_dir, filename)

        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        return JSONResponse({"filename": filename, "docx_base64": docx_b64})

    except Exception as e:
        import traceback
        print("=== ERRO NO /generate ===")
        print(traceback.format_exc())
        raise HTTPException(500, f"Erro gerando laudo: {e}")

    finally:
        shutil.rmtree(work, ignore_errors=True)
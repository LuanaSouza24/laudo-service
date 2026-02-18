import os
import base64
import tempfile
import shutil
import traceback
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel

import gerar_laudo as gl

app = FastAPI()


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None


@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        # Diretório de trabalho (Render) onde Excel/Template/saida existirão
        os.environ["LAUDO_BASE_DIR"] = work

        # Salva o Excel enviado pelo Power Automate
        excel_path = os.path.join(work, "Cautelar.xlsx")
        with open(excel_path, "wb") as f:
            f.write(base64.b64decode(p.excel_base64))

        # Template: se vier no payload, usa; senão copia o do repositório
        template_dst = os.path.join(work, "tamplete.docx")
        if p.template_base64:
            with open(template_dst, "wb") as f:
                f.write(base64.b64decode(p.template_base64))
        else:
            template_src = os.path.join(os.path.dirname(__file__), "tamplete.docx")
            shutil.copy(template_src, template_dst)

        # Garante que o módulo use os caminhos atualizados (LAUDO_BASE_DIR)
        if hasattr(gl, "refresh_paths"):
            gl.refresh_paths()

        # Gera o laudo
        gl.gerar_laudo(p.id_vistoria)

        out_dir = os.path.join(work, "saida")
        if not os.path.isdir(out_dir):
            raise HTTPException(500, "Pasta de saída não encontrada (saida).")

        files = [x for x in os.listdir(out_dir) if x.lower().endswith(".docx")]
        if not files:
            raise HTTPException(500, "Laudo não foi gerado (nenhum .docx em saida).")

        filename = files[0]
        out_path = os.path.join(out_dir, filename)

        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        return JSONResponse({"filename": filename, "docx_base64": docx_b64})

    except Exception as e:
        print("=== ERRO NO /generate ===")
        print(traceback.format_exc())
        raise HTTPException(500, f"Erro gerando laudo: {e}")

    finally:
        shutil.rmtree(work, ignore_errors=True)

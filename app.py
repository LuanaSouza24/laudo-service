import os, base64, tempfile, shutil
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.responses import JSONResponse

import gerar_laudo as gl

app = FastAPI()

class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str  # conteúdo do Cautelar.xlsx vindo do Power Automate
    # opcional:
    # template_base64: str

@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        os.environ["LAUDO_BASE_DIR"] = work

        # salva o Excel
        excel_path = os.path.join(work, "Cautelar.xlsx")
        with open(excel_path, "wb") as f:
            f.write(base64.b64decode(p.excel_base64))

        # template: usar o do repo (mais simples)
        template_src = os.path.join(os.path.dirname(__file__), "tamplete.docx")
        shutil.copy(template_src, os.path.join(work, "tamplete.docx"))

        # roda o laudo
        gl.gerar_laudo(p.id_vistoria)

        out_dir = os.path.join(work, "saida")
        files = [x for x in os.listdir(out_dir) if x.lower().endswith(".docx")]
        if not files:
            raise HTTPException(500, "Laudo não foi gerado.")

        filename = files[0]
        out_path = os.path.join(out_dir, filename)
        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        return JSONResponse({"filename": filename, "docx_base64": docx_b64})

    except Exception as e:
        raise HTTPException(500, f"Erro gerando laudo: {e}")

    finally:
        shutil.rmtree(work, ignore_errors=True)

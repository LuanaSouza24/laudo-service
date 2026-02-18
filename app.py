import os
import base64
import tempfile
import shutil
import importlib
from typing import Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.responses import JSONResponse

app = FastAPI()


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None


@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        # Diretório temporário onde vamos montar Cautelar.xlsx, tamplete.docx e a saída
        os.environ["LAUDO_BASE_DIR"] = work

        # 1) Salvar o Excel
        excel_path = os.path.join(work, "Cautelar.xlsx")
        with open(excel_path, "wb") as f:
            f.write(base64.b64decode(p.excel_base64))

        # 2) Salvar o template (preferir o enviado; senão usar o que está no repo)
        dst_template = os.path.join(work, "tamplete.docx")
        if p.template_base64:
            with open(dst_template, "wb") as f:
                f.write(base64.b64decode(p.template_base64))
        else:
            template_src = os.path.join(os.path.dirname(__file__), "tamplete.docx")
            shutil.copy(template_src, dst_template)

        # 3) Importar o gerar_laudo APÓS setar LAUDO_BASE_DIR
        import gerar_laudo as gl
        importlib.reload(gl)  # garante que os paths globais respeitem o LAUDO_BASE_DIR

        # 4) Gerar o laudo
        gl.gerar_laudo(p.id_vistoria)

        # 5) Ler o DOCX gerado
        out_dir = os.path.join(work, "saida")
        if not os.path.isdir(out_dir):
            raise HTTPException(500, f"Pasta de saída não encontrada: {out_dir}")

        docx_files = [x for x in os.listdir(out_dir) if x.lower().endswith(".docx")]
        if not docx_files:
            raise HTTPException(500, "Laudo não foi gerado (nenhum .docx na pasta de saída).")

        # pegar o mais recente
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

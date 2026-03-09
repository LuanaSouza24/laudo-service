import os
import re
import base64
import tempfile
import shutil
import importlib
from io import BytesIO
from typing import Optional, List
from urllib.parse import quote

import requests
from PIL import Image
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel

app = FastAPI()


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None
    image_paths: Optional[List[str]] = None


def normalizar_rel_path(path: str) -> str:
    """
    Normaliza o caminho relativo vindo do Power Automate / Excel script.
    Remove barras invertidas, barras duplicadas e espaços laterais.
    """
    if not path:
        return ""
    rel = str(path).strip().replace("\\", "/")
    rel = re.sub(r"/+", "/", rel)
    rel = rel.lstrip("/")
    return rel


def montar_url_sharepoint(rel_path: str) -> str:
    """
    Monta a URL final do SharePoint a partir do caminho relativo.
    Exige a variável SHAREPOINT_BASE_URL.
    Ex.:
      SHAREPOINT_BASE_URL=https://dominio.sharepoint.com
      rel_path=Documentos Compartilhados/pasta/arquivo.jpg
    """
    base_url = os.getenv("SHAREPOINT_BASE_URL", "").strip().rstrip("/")
    if not base_url:
        raise Exception("Variável de ambiente SHAREPOINT_BASE_URL não configurada.")

    rel = normalizar_rel_path(rel_path)
    if not rel:
        raise Exception("Caminho da imagem vazio ao montar URL do SharePoint.")

    return f"{base_url}/{quote(rel, safe='/:()_- ')}"


def obter_headers_sharepoint() -> dict:
    """
    Usa bearer token já pronto, salvo em SHAREPOINT_BEARER_TOKEN.
    """
    token = os.getenv("SHAREPOINT_BEARER_TOKEN", "").strip()
    if not token:
        raise Exception("Variável de ambiente SHAREPOINT_BEARER_TOKEN não configurada.")

    return {
        "Authorization": f"Bearer {token}"
    }


def baixar_arquivo_sharepoint(rel_path: str, destino_abs: str):
    """
    Baixa um arquivo do SharePoint e salva no destino absoluto informado.
    """
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
    """
    Comprime/redimensiona a imagem no mesmo caminho.
    Converte para JPEG quando necessário.
    """
    if not os.path.exists(caminho):
        return

    if not eh_imagem(caminho):
        return

    try:
        with Image.open(caminho) as img:
            img.thumbnail((max_lado, max_lado))

            # Remove transparência se houver
            if img.mode in ("RGBA", "LA", "P"):
                fundo = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                fundo.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
                img = fundo
            elif img.mode != "RGB":
                img = img.convert("RGB")

            ext = os.path.splitext(caminho)[1].lower()

            # Se já for jpg/jpeg, sobrescreve
            if ext in {".jpg", ".jpeg"}:
                img.save(caminho, format="JPEG", quality=qualidade, optimize=True)
                return

            # Para png/bmp/etc, substitui por .jpg
            novo_caminho = os.path.splitext(caminho)[0] + ".jpg"
            img.save(novo_caminho, format="JPEG", quality=qualidade, optimize=True)

        # remove o original se mudou extensão
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
            raise Exception("Template não enviado e arquivo tamplete.docx não encontrado no repositório.")
        shutil.copy(template_src, dst_template)

    return dst_template


def preparar_imagens(work_dir: str, image_paths: Optional[List[str]]):
    """
    Baixa as imagens do SharePoint e recria a estrutura local esperada pelo gerar_laudo.py.
    """
    if not image_paths:
        print("[INFO] Nenhuma imagem recebida em image_paths.")
        return

    total = len(image_paths)
    print(f"[INFO] Iniciando download de {total} imagem(ns).")

    for i, original_path in enumerate(image_paths, start=1):
        rel = normalizar_rel_path(original_path)
        if not rel:
            print(f"[AVISO] Caminho vazio ignorado no índice {i}.")
            continue

        destino_abs = os.path.join(work_dir, rel)
        print(f"[INFO] [{i}/{total}] Baixando: {rel}")

        baixar_arquivo_sharepoint(rel, destino_abs)

        if eh_imagem(destino_abs):
            comprimir_imagem_no_mesmo_arquivo(destino_abs)

    print("[INFO] Download/preparo de imagens concluído.")


def localizar_docx_gerado(work_dir: str) -> str:
    out_dir = os.path.join(work_dir, "saida")
    if not os.path.isdir(out_dir):
        raise Exception(f"Pasta de saída não encontrada: {out_dir}")

    docx_files = [
        os.path.join(out_dir, nome)
        for nome in os.listdir(out_dir)
        if nome.lower().endswith(".docx")
    ]

    if not docx_files:
        raise Exception("Laudo não foi gerado (nenhum .docx na pasta de saída).")

    docx_files.sort(key=os.path.getmtime, reverse=True)
    return docx_files[0]


def gerar_laudo_no_modulo(id_vistoria: str):
    """
    Importa/recarrega o módulo após definir LAUDO_BASE_DIR.
    """
    import gerar_laudo as gl
    importlib.reload(gl)
    gl.gerar_laudo(id_vistoria)


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")

    try:
        print("========== /generate ==========")
        print(f"[INFO] id_vistoria={p.id_vistoria}")
        print(f"[INFO] image_paths={len(p.image_paths or [])}")

        # Diretório base consumido pelo gerar_laudo.py
        os.environ["LAUDO_BASE_DIR"] = work

        # 1) Excel
        preparar_excel(work, p.excel_base64)

        # 2) Template
        preparar_template(work, p.template_base64)

        # 3) Imagens
        preparar_imagens(work, p.image_paths)

        # 4) Gerar laudo
        gerar_laudo_no_modulo(p.id_vistoria)

        # 5) Encontrar docx gerado
        out_path = localizar_docx_gerado(work)
        filename = os.path.basename(out_path)

        # 6) Retornar em base64 para o Power Automate
        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        print(f"[INFO] Laudo gerado com sucesso: {filename}")
        return JSONResponse({
            "filename": filename,
            "docx_base64": docx_b64
        })

    except Exception as e:
        import traceback
        print("=== ERRO NO /generate ===")
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Erro gerando laudo: {e}")

    finally:
        shutil.rmtree(work, ignore_errors=True)
import os
from pathlib import Path

from flask import Flask, redirect, render_template_string, request, url_for
from werkzeug.utils import secure_filename


DATA_DIR = Path(os.getenv("CDI_DATA_DIR", "/var/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED_EXTENSIONS = {".xls", ".xlsx"}

app = Flask(__name__)


PAGE_TEMPLATE = """
<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>CDI Upload</title>
    <style>
      :root {
        color-scheme: dark;
        --bg: #0b1220;
        --panel: #121b2d;
        --muted: #9fb0cf;
        --text: #e8eefc;
        --accent: #63a4ff;
        --accent-2: #35d49a;
        --border: #22314f;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background:
          radial-gradient(circle at top left, rgba(99, 164, 255, 0.16), transparent 28%),
          radial-gradient(circle at top right, rgba(53, 212, 154, 0.14), transparent 22%),
          var(--bg);
        color: var(--text);
      }
      .wrap {
        max-width: 980px;
        margin: 0 auto;
        padding: 32px 20px 60px;
      }
      .hero, .panel {
        background: rgba(18, 27, 45, 0.92);
        border: 1px solid var(--border);
        border-radius: 18px;
        box-shadow: 0 20px 60px rgba(0, 0, 0, 0.28);
      }
      .hero {
        padding: 28px;
        margin-bottom: 22px;
      }
      h1 {
        margin: 0 0 10px;
        font-size: 34px;
      }
      p {
        color: var(--muted);
        line-height: 1.5;
      }
      .actions {
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        margin-top: 18px;
      }
      .btn, button {
        appearance: none;
        border: 0;
        border-radius: 12px;
        padding: 12px 18px;
        font-size: 15px;
        font-weight: 600;
        cursor: pointer;
        text-decoration: none;
      }
      .btn-primary, button {
        background: linear-gradient(135deg, var(--accent), #3d7cff);
        color: white;
      }
      .btn-secondary {
        background: rgba(255, 255, 255, 0.04);
        border: 1px solid var(--border);
        color: var(--text);
      }
      .grid {
        display: grid;
        grid-template-columns: 1.1fr 0.9fr;
        gap: 22px;
      }
      .panel {
        padding: 24px;
      }
      input[type=file] {
        width: 100%;
        padding: 14px;
        border: 1px dashed var(--border);
        border-radius: 14px;
        background: rgba(255, 255, 255, 0.03);
        color: var(--muted);
        margin: 10px 0 16px;
      }
      .flash {
        margin-bottom: 16px;
        padding: 12px 14px;
        border-radius: 12px;
        border: 1px solid var(--border);
        background: rgba(99, 164, 255, 0.08);
      }
      .file {
        display: flex;
        justify-content: space-between;
        gap: 16px;
        padding: 14px 0;
        border-bottom: 1px solid rgba(255, 255, 255, 0.06);
      }
      .file:last-child {
        border-bottom: 0;
      }
      .meta {
        color: var(--muted);
        font-size: 14px;
      }
      code {
        background: rgba(255, 255, 255, 0.06);
        padding: 2px 8px;
        border-radius: 8px;
      }
      @media (max-width: 780px) {
        .grid { grid-template-columns: 1fr; }
      }
    </style>
  </head>
  <body>
    <div class="wrap">
      <section class="hero">
        <h1>CDI - Upload de Planilhas</h1>
        <p>Envie o arquivo <code>.xls</code> ou <code>.xlsx</code> do seu computador. Depois abra o app e selecione o arquivo em <code>/var/data</code>.</p>
        <div class="actions">
          <a class="btn btn-primary" href="/open-app">Abrir app</a>
          <a class="btn btn-secondary" href="/healthz">Health check</a>
        </div>
      </section>

      <div class="grid">
        <section class="panel">
          <h2>Enviar arquivo</h2>
          {% if message %}
            <div class="flash">{{ message }}</div>
          {% endif %}
          <form method="post" action="/upload" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx" required>
            <button type="submit">Enviar para /var/data</button>
          </form>
          <p>O app remoto foi ajustado para abrir o seletor de arquivos diretamente em <code>{{ data_dir }}</code>.</p>
        </section>

        <section class="panel">
          <h2>Arquivos no servidor</h2>
          {% if files %}
            {% for file in files %}
              <div class="file">
                <div>
                  <div>{{ file.name }}</div>
                  <div class="meta">{{ file.size }} MB</div>
                </div>
                <a class="btn btn-secondary" href="/open-app">Usar no app</a>
              </div>
            {% endfor %}
          {% else %}
            <p>Nenhum arquivo enviado ainda.</p>
          {% endif %}
        </section>
      </div>
    </div>
  </body>
</html>
"""


def list_uploaded_files():
    files = []
    for path in sorted(DATA_DIR.glob("*")):
        if path.is_file():
            files.append(
                {
                    "name": path.name,
                    "size": f"{path.stat().st_size / (1024 * 1024):.2f}",
                }
            )
    return files


@app.get("/")
def index():
    return render_template_string(
        PAGE_TEMPLATE,
        files=list_uploaded_files(),
        data_dir=str(DATA_DIR),
        message=request.args.get("message", ""),
    )


@app.get("/open-app")
def open_app():
    return redirect("/novnc/vnc.html?autoconnect=1&resize=scale&quality=9&path=novnc/websockify")


@app.get("/healthz")
def healthz():
    return {"status": "ok", "data_dir": str(DATA_DIR)}


@app.post("/upload")
def upload():
    uploaded = request.files.get("file")
    if uploaded is None or not uploaded.filename:
        return redirect(url_for("index", message="Selecione um arquivo antes de enviar."))

    filename = secure_filename(uploaded.filename)
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return redirect(url_for("index", message="Envie apenas arquivos .xls ou .xlsx."))

    target = DATA_DIR / filename
    uploaded.save(target)
    return redirect(url_for("index", message=f"Arquivo enviado: {filename}"))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8001, debug=False)

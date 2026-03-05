# Deploy do `analise_tempo_atendimento_cdi.py` no Render

Este pacote publica somente o app `analise_tempo_atendimento_cdi.py`.

## Como funciona

O app original e um desktop Qt (`PySide6`). Para manter a mesma interface no Render, ele roda dentro de um container Linux com display virtual e e exposto no navegador via noVNC.

Isso preserva a UI do app.

## Estrutura

```text
deploy_render_analise_tempo_atendimento_cdi/
├── app/
│   └── analise_tempo_atendimento_cdi.py
├── Dockerfile
├── README.md
├── render.yaml
├── requirements.txt
├── start.sh
└── web/
    └── index.html
```

## Publicar

1. Crie um repositorio novo no GitHub.
2. Suba somente o conteudo desta pasta.
3. No Render, crie `New +` -> `Blueprint`.
4. Selecione o repositorio.

O `render.yaml` cria um Web Service Docker com disco persistente em `/var/data`.

## Acesso

Depois do deploy, abra a URL raiz do servico no Render. A raiz `/` agora redireciona automaticamente para o cliente VNC com autoconnect.

## Arquivos Excel

No pacote de deploy, a copia do app foi ajustada para usar `/var/data` como diretorio padrao de abrir/salvar arquivos.

Limite importante:

- O app continua sendo desktop remoto.
- O seletor de arquivos ve os arquivos do servidor, nao os arquivos locais do seu computador.

Ou seja: a interface fica igual, mas a experiencia de carregar planilhas no Render nao e identica a um app local. Isso e uma limitacao do Render para apps desktop nativos.

## O que foi alterado na copia de deploy

- Remocao de caminhos fixos do macOS.
- Diretorio padrao de abrir/salvar arquivos movido para `CDI_DATA_DIR` com fallback em `/var/data`.

O arquivo original do seu projeto nao foi alterado.

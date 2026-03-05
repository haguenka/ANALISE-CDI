# Deploy do `analise_tempo_atendimento_cdi.py` no Render

Este pacote publica somente o app `analise_tempo_atendimento_cdi.py`.

## Como funciona

O app original e um desktop Qt (`PySide6`). Para manter a mesma interface no Render, ele roda dentro de um container Linux com display virtual e e exposto no navegador via noVNC.

Isso preserva a UI do app.

A raiz do servico agora e uma pagina web de upload. O desktop remoto do app fica em `/novnc/`.

## Estrutura

```text
deploy_render_analise_tempo_atendimento_cdi/
├── app/
│   └── analise_tempo_atendimento_cdi.py
├── Dockerfile
├── README.md
├── nginx.conf.template
├── render.yaml
├── requirements.txt
├── start.sh
└── upload_server.py
```

## Publicar

1. Crie um repositorio novo no GitHub.
2. Suba somente o conteudo desta pasta.
3. No Render, crie `New +` -> `Blueprint`.
4. Selecione o repositorio.

O `render.yaml` cria um Web Service Docker com disco persistente em `/var/data`.

## Acesso

Depois do deploy, abra a URL raiz do servico no Render.

Fluxo:

1. Faça upload do `.xls` ou `.xlsx` pela pagina inicial.
2. Clique em abrir o app.
3. No desktop remoto, o seletor de arquivos ja abre em `/var/data`.

## Arquivos Excel

No pacote de deploy, a copia do app foi ajustada para usar `/var/data` como diretorio padrao de abrir/salvar arquivos.

Limite importante:

- O app continua sendo desktop remoto.
- O arquivo sobe primeiro pela interface web e depois fica disponivel no servidor para o app Qt.

Ou seja: a interface fica igual, mas a experiencia de carregar planilhas no Render nao e identica a um app local. Isso e uma limitacao do Render para apps desktop nativos.

## O que foi alterado na copia de deploy

- Remocao de caminhos fixos do macOS.
- Diretorio padrao de abrir/salvar arquivos movido para `CDI_DATA_DIR` com fallback em `/var/data`.
- Adicao de pagina web para upload de planilhas.
- Proxy reverso para servir upload web e noVNC no mesmo dominio.

O arquivo original do seu projeto nao foi alterado.

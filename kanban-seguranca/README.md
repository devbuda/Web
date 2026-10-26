# Plano de Ação – Implantação de Segurança da Informação (Kanban)

Painel visual estilo Kanban para acompanhamento em tempo real das demandas do **Plano de Ação – Implantação de Segurança da Informação** (Julho de 2026 a Novembro de 2026).

## Funcionalidades
- **Quadro Kanban Interativo (Drag & Drop):** Arraste os cards entre as colunas **Pendente**, **Em Andamento** e **Concluído**.
- **Banco de Dados SQLite + Tempo Real:** Todas as alterações são salvas automaticamente no banco SQLite (`kanban.db`) no servidor e sincronizadas em tempo real para todos os usuários da rede interna.
- **Indicadores de Progresso:** Visão geral de demandas e acompanhamento detalhado por Solução.

## Como executar no Servidor (Ubuntu)

```bash
PORT=8080 python3 server.py
```
O servidor iniciará na porta específica `8080` (`http://192.168.0.53:8080`), deixando a porta `80` livre para outros projetos web.

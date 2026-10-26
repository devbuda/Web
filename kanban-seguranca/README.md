# Plano de Ação – Implantação de Segurança da Informação (Kanban)

Painel visual estilo Kanban para acompanhamento em tempo real das demandas do **Plano de Ação – Implantação de Segurança da Informação** (Julho de 2026 a Novembro de 2026).

## Funcionalidades
- **Quadro Kanban Interativo (Drag & Drop):** Arraste os cards entre as colunas **Pendente**, **Em Andamento** e **Concluído**.
- **Salvamento Automático:** Todas as alterações são salvas automaticamente no servidor (`tasks_data.json`) e sincronizadas para todos os usuários da rede interna.
- **Indicadores de Progresso:** Visão geral de demandas e acompanhamento detalhado por Solução.

## Como executar no Servidor (Ubuntu)

```bash
sudo python3 server.py
```
O servidor iniciará por padrão na porta `80` (`http://192.168.0.53`).

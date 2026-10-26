#!/bin/bash
set -e

echo "=== Instalando Kanban Web Server na porta 8080 ==="
sudo mkdir -p /opt/kanban
sudo cp /tmp/index.html /opt/kanban/index.html
sudo cp /tmp/server.py /opt/kanban/server.py
sudo cp /tmp/initial_tasks.json /opt/kanban/initial_tasks.json
sudo chmod +x /opt/kanban/server.py

sudo tee /etc/systemd/system/kanban.service > /dev/null << 'EOF'
[Unit]
Description=Kanban Seguranca da Informacao Web Server
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/opt/kanban
Environment="PORT=8080"
ExecStart=/usr/bin/python3 /opt/kanban/server.py
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
sudo systemctl enable kanban.service
sudo systemctl restart kanban.service

if command -v ufw >/dev/null 2>&1; then
    sudo ufw allow 8080/tcp || true
fi

echo "=== Pronto! Servidor ativo em http://192.168.0.53:8080 ==="

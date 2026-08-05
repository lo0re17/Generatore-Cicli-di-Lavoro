#!/bin/sh
# Generatore Cicli - Deploy
# Eseguire da: /share/Container/fn-cicli/fn-cicli

echo "════════════════════════════════════════"
echo "  Generatore Cicli - Deploy"
echo "════════════════════════════════════════"

echo "→ Fermando container..."
docker compose down 2>/dev/null

mkdir -p templates anagrafica web/instance

echo "→ Building..."
docker compose build

echo "→ Avviando..."
docker compose up -d

echo "→ Attesa avvio (5s)..."
sleep 5

echo ""
docker compose ps
echo ""
echo "✅ DEPLOY COMPLETATO!"
echo "📱 URL: http://$(hostname -I | awk '{print $1}'):8071"
echo "════════════════════════════════════════"

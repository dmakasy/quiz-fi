#!/bin/bash
cd "$(dirname "$0")"
if ! command -v node >/dev/null 2>&1; then
  echo "Node.js n'est pas installe. Installe-le depuis https://nodejs.org (version LTS) puis relance ce fichier."
  read -p "Appuie sur Entree pour fermer."
  exit 1
fi
if [ ! -d node_modules ]; then
  echo "Premiere fois : installation des composants (necessite Internet)..."
  npm install
fi
echo ""
echo "Demarrage du quiz... Laisse cette fenetre ouverte pendant la partie."
node server.js

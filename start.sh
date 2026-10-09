#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
if [ ! -d node_modules ]; then
  echo "Installing dependencies (internet connection required the first time)..."
  npm install
fi
npm run dev

#!/bin/sh
# Gera download-page/index.html (página única, com o código embutido)
cd "$(dirname "$0")/.." || exit 1
npx --yes esbuild pagina/ponte.ts --bundle --minify --format=iife --target=es2019 --outfile=pagina/ponte.js >/dev/null || exit 1
python3 - <<'PY'
t=open('pagina/template.html').read(); b=open('pagina/ponte.js').read().replace('</script','<\\/script')
open('download-page/index.html','w').write(t.replace('/*BUNDLE*/',b))
PY

name: Update TrackLink Santa Marta Report

on:
  schedule:
    - cron: '0 12 * * 1'  # Cada lunes a las 12:00 UTC (08:00 hora Chile)
  workflow_dispatch:        # Botón manual en la pestaña Actions

jobs:
  download-report:
    runs-on: ubuntu-latest
    permissions:
      contents: write       # Necesario para hacer push del xlsx actualizado

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Setup Node.js 20
        uses: actions/setup-node@v4
        with:
          node-version: '20'

      - name: Install dependencies
        run: npm install puppeteer adm-zip

      - name: Download report from TrackGTS
        env:
          TL_USER: ${{ secrets.TL_USER }}
          TL_PASSWORD: ${{ secrets.TL_PASSWORD }}
          TL_DOMAIN: ${{ secrets.TL_DOMAIN }}
        run: node download.js

      - name: Commit and push report
        run: |
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git config user.name "github-actions[bot]"
          git add "INFORME EXCESOS DE VELOCIDAD.xlsx"
          if git diff --staged --quiet; then
            echo "Sin cambios — el reporte ya estaba actualizado."
          else
            git commit -m "chore: actualiza reporte semanal $(date -u +%Y-%m-%d)"
            git push
          fi

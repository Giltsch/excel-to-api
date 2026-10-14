# Deployment Guide — Excel to API

## Voraussetzungen

- **Server:** Hetzner CX22 (2 vCPU, 4 GB RAM, ~€4/Monat) mit Ubuntu 24.04
- **Domain:** z.B. `excel2api.deinedomain.de` — DNS A-Record auf die Server-IP zeigen lassen

---

## 1. Server einrichten

```bash
# Per SSH verbinden
ssh root@DEINE_SERVER_IP

# Updates & Pakete
apt update && apt upgrade -y
apt install -y docker.io docker-compose-plugin caddy git

# Docker ohne Neustart aktivieren
systemctl enable --now docker
```

## 2. Projekt deployen

```bash
# Repo klonen
git clone https://github.com/Giltsch/excel-to-api.git /opt/excel-to-api
cd /opt/excel-to-api

# Environment konfigurieren
cp .env.example .env
nano .env
```

In der `.env` setzen:

```env
PUBLIC_MODE=true
OPENROUTER_API_KEY=          # leer lassen! User bringen ihren eigenen Key mit
OPENROUTER_MODEL=anthropic/claude-sonnet-4
```

## 3. App starten

```bash
docker compose up -d
```

Prüfen ob es läuft:

```bash
docker compose logs -f
# Strg+C zum Beenden der Logs
```

Die App läuft jetzt auf `http://DEINE_SERVER_IP:8501` (ohne HTTPS).

## 4. Domain + HTTPS einrichten

### DNS konfigurieren

Bei deinem Domain-Provider einen **A-Record** erstellen:

| Typ | Name | Wert |
|-----|------|------|
| A   | excel2api | DEINE_SERVER_IP |

### Caddy konfigurieren

```bash
# Caddyfile anpassen
nano /opt/excel-to-api/Caddyfile
```

`YOUR_DOMAIN` durch deine echte Domain ersetzen (z.B. `excel2api.deinedomain.de`):

```
excel2api.deinedomain.de {
    reverse_proxy localhost:8501

    header {
        X-Content-Type-Options nosniff
        X-Frame-Options SAMEORIGIN
        Referrer-Policy strict-origin-when-cross-origin
        Permissions-Policy "camera=(), microphone=(), geolocation=()"
        -Server
    }

    request_body {
        max_size 12MB
    }
}
```

Dann aktivieren:

```bash
cp /opt/excel-to-api/Caddyfile /etc/caddy/Caddyfile
systemctl reload caddy
```

Caddy holt sich automatisch ein Let's Encrypt-Zertifikat. Deine App ist nach wenigen Sekunden unter `https://excel2api.deinedomain.de` erreichbar.

## 5. Firewall einrichten

```bash
# Nur SSH, HTTP und HTTPS erlauben
ufw allow 22/tcp
ufw allow 80/tcp
ufw allow 443/tcp
ufw --force enable
```

> **Wichtig:** Port 8501 ist damit von außen nicht mehr direkt erreichbar — nur über Caddy (HTTPS).

---

## Updates deployen

```bash
cd /opt/excel-to-api
git pull
docker compose up -d --build
```

## Logs & Monitoring

```bash
# App-Logs
docker compose logs -f

# Caddy-Logs
journalctl -u caddy -f

# Ressourcen prüfen
docker stats
```

## Troubleshooting

| Problem | Lösung |
|---------|--------|
| HTTPS funktioniert nicht | DNS prüfen: `dig excel2api.deinedomain.de` — muss deine Server-IP zeigen |
| App startet nicht | `docker compose logs` prüfen |
| "Rate limit reached" | Warten oder `RATE_LIMIT_UPLOADS` in `.env` erhöhen, dann `docker compose up -d` |
| Upload schlägt fehl | Dateigröße prüfen (max 10 MB), nur `.xlsx` erlaubt |

---

## Sicherheits-Checkliste

- [x] Public Mode aktiv (`PUBLIC_MODE=true`)
- [x] Kein API-Key auf dem Server (Bring Your Own Key)
- [x] Excel-Validierung (Größe, Format, Struktur)
- [x] Rate-Limiting (10 Uploads/Stunde pro Session)
- [x] Non-root Docker User
- [x] Read-only Container-Filesystem
- [x] HTTPS via Caddy + Let's Encrypt
- [x] Firewall (nur 22, 80, 443)
- [x] Security Headers (nosniff, SAMEORIGIN, etc.)
- [x] Live-Preview deaktiviert (kein Subprocess-Risiko)

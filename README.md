# Miramar Gastos

App web simple para cargar, consultar y exportar gastos mensuales de Miramar.

## Deploy con Docker Compose

Subir estos archivos a GitHub:

- `server.py`
- `Dockerfile`
- `docker-compose.yml`
- `.gitignore`
- `README.md`

En Portainer o similar, elegir este repositorio y usar `docker-compose.yml` como compose file.

La app queda disponible en:

```text
http://IP_DEL_SERVIDOR:8765
```

## Datos

La base SQLite se guarda en el volumen Docker `miramar_data`, por eso no se pierde al reiniciar el contenedor.

No subir a GitHub:

- `miramar.db`
- `exports/`
- archivos `.log`

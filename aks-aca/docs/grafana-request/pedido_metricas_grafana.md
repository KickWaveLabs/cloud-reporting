# Pedido de métricas por "micro"-servicio (Grafana)

## Para qué

Queremos calcular, para cada "micro"-servicio de producción, cuántas réplicas necesitaría según su tráfico real.
Hoy lo medimos por API en APIM, pero APIM solo ve lo que entra al cluster. No ve las llamadas internas entre
"micro"-servicios, y tampoco a qué "micro"-servicio va cada llamada que entra por el ingress.

Por eso necesitamos las solicitudes medidas **en el propio "micro"-servicio** (del lado del servidor), no en el
ingress: así se cuentan también las llamadas internas.

## Alcance

- **Clusters:** todos los AKS de producción.
- **Qué se mide:** cada "micro"-servicio (deployment) de cada namespace de aplicación. Sin namespaces de plataforma
  (kube-system, monitoreo, ingress).
- **Periodo:** octubre 2026 completo (del 1 al 31), que es el mismo mes del costo. Si no guardan tanto, los últimos 30
  días que tengan, indicando las fechas.
- **Hora:** en hora de Lima (UTC-5), o en UTC indicándolo.

## Qué necesitamos (dos archivos)

### Archivo 1: una fila por "micro"-servicio, día y hora

Son unas 2,500 "micro"-servicios × 31 días × 24 horas: alrededor de 1.8 millones de filas. Mejor en CSV.

| Columna | Qué es |
|---|---|
| `cluster` | Nombre del cluster AKS |
| `namespace` | Namespace |
| `deployment` | Nombre del deployment (o del service, si así lo tienen) |
| `date` | Fecha (AAAA-MM-DD) |
| `hour` | Hora del día, de 0 a 23 |
| `requests` | Solicitudes recibidas en esa hora (total de la hora) |
| `requests_max_minute` | Solicitudes del minuto con más solicitudes dentro de esa hora |
| `requests_min_minute` | Solicitudes del minuto con menos solicitudes dentro de esa hora |
| `zero_minutes` | Cuántos minutos de esa hora tuvieron 0 solicitudes |
| `duration_sum_seconds` | Suma de la duración de todas las solicitudes de la hora (para sacar el promedio) |
| `duration_max_seconds` | La solicitud más lenta de la hora |
| `errors_5xx` | Solicitudes con error 5xx en la hora (opcional) |
| `replicas_avg` | Réplicas (pods) disponibles en promedio en la hora (opcional) |
| `replicas_max` | Máximo de réplicas en la hora (opcional) |

No necesitamos percentiles (p95, p99). Con la suma y el conteo sacamos el promedio, y con el máximo, el pico.

### Archivo 2: la hora más cargada de cada "micro"-servicio, minuto a minuto

Para cada "micro"-servicio: su hora con más solicitudes del mes (una sola hora), con los 60 minutos.

| Columna | Qué es |
|---|---|
| `cluster`, `namespace`, `deployment` | Igual que arriba |
| `date`, `hour` | La hora más cargada del mes de ese "micro"-servicio |
| `minute` | Minuto, de 0 a 59 |
| `requests` | Solicitudes en ese minuto |

Son unos 2,500 × 60 = 150 mil filas.

## Métricas de origen (referencia)

Depende de cómo lo midan. Unos ejemplos en PromQL, por minuto:

- **Si las aplicaciones exponen métricas Micrometer** (Spring Boot o Quarkus):
  - Solicitudes: `sum by (namespace, deployment) (increase(http_server_requests_seconds_count[1m]))`
  - Duración (suma): `sum by (namespace, deployment) (increase(http_server_requests_seconds_sum[1m]))`
  - Duración máxima: `max by (namespace, deployment) (max_over_time(http_server_requests_seconds_max[1m]))`
  - Errores: lo mismo que solicitudes, filtrando `status=~"5.."`.
- **Réplicas** (kube-state-metrics): `kube_deployment_status_replicas_available`.
- **Si solo tienen métricas del ingress** (por ejemplo `nginx_ingress_controller_requests`), igual nos sirven, pero
  avísennos: ahí no aparecen las llamadas internas.

Si los nombres de las métricas son otros, lo que importa es lo de las tablas de arriba.

## Cómo lo vamos a usar

Es la misma lógica que ya aplicamos con APIM, pero por "micro"-servicio y con el minuto en vez de 5 minutos:

1. **TPS de cada hora** = `requests / 3600`. Con eso armamos el perfil de un día hábil y de un fin de semana:
   para cada hora del día, el promedio de todos los días hábiles (o de fin de semana) del mes.
2. **Tres formas de dimensionar**, que se eligen en el workbook:
   - **Pico:** en cada hora, el minuto más cargado (`requests_max_minute / 60`).
   - **Hora más cargada:** se repite minuto a minuto la hora más cargada del mes (archivo 2).
   - **Promedio:** el promedio de la hora.
3. **Réplicas** = 1 por cada 10 TPS (la regla de escalado de Container Apps). Mínimo 1 en horario de oficina y 0
   fuera de él: los minutos sin solicitudes (`zero_minutes`) son los que se pueden apagar.
4. **Duración promedio** = `duration_sum_seconds / requests`. Sirve para ver cuánto rinde una réplica, y para
   revisar la latencia.

## Contacto

Cualquier duda sobre columnas o formato, me escriben.

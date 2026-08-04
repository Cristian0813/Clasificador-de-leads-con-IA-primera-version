# Clasificador Automático de Leads por Correo (IA)

Lee los correos no leídos de una bandeja, los clasifica automáticamente por
**categoría** (Venta, Consulta, Queja, Soporte, Spam) y **urgencia**
(Alta, Media, Baja) usando IA, y entrega un Excel listo con colores para
que el dueño del negocio sepa a quién responder primero, sin leer correo
por correo.

## Para quién es esto (tu pitch de venta)

Cualquier negocio que reciba más de 15-20 correos al día y no dé abasto
priorizando: inmobiliarias, clínicas/consultorios, despachos de abogados,
agencias de viajes, e-commerce con atención al cliente por correo.

**Problema que resuelve:** "Se me pierden los correos importantes entre
el spam y las consultas triviales."
**Resultado que vendes:** un Excel automático, generado cada mañana (o
cada hora si lo programan), con los leads urgentes resaltados en rojo
arriba de todo.

## Instalación

```bash
pip install -r requirements.txt
cp .env.example .env
# Edita .env con las credenciales reales del cliente
python main.py
```

### Cómo obtener la contraseña de aplicación de Gmail
1. El cliente activa verificación en 2 pasos en su cuenta de Google.
2. Va a myaccount.google.com → Seguridad → Contraseñas de aplicaciones.
3. Genera una y la pega en `EMAIL_APP_PASSWORD`.

### Cómo obtener la API key de Anthropic
1. console.anthropic.com → API Keys → Create Key.
2. Cada clasificación cuesta fracciones de centavo; para 20 correos/día
   el costo mensual es de un par de dólares.

## Automatización (opcional, para venderlo como servicio recurrente)
En Windows: usa el Programador de Tareas (Task Scheduler) para correrlo
cada hora — la misma técnica que ya dominas de tu proyecto anterior.
En Linux/servidor: un cron job `0 * * * * python main.py`.

## Cómo lo vendes
1. Corres esto una vez, gratis, con el correo real de un negocio (con su
   permiso) y le mandas el Excel resultante como demostración.
2. Le cobras un setup inicial (ej. $150.000–300.000 COP) + una mensualidad
   pequeña ($50.000–100.000 COP) si quiere que lo automatices y lo tengas
   corriendo solo.
3. Con el tiempo, empaquetas esto como plantilla en Gumroad para que
   otros lo compren y lo configuren ellos mismos.

## Próximos pasos técnicos (para subir el precio del servicio)
- Enviar notificación a WhatsApp cuando llega un correo "Urgencia: Alta"
  (reutilizas tu experiencia con bots de WhatsApp).
- Guardar el histórico en una base de datos en vez de sobrescribir el Excel.
- Panel web simple (Streamlit) en vez de abrir el Excel manualmente.

"""Avisa de la alerta: abre un aviso en GitHub (te llega por correo) y, si hay credenciales de correo
configuradas (MAIL_USER y MAIL_PASS), manda un correo a Davidsm004@gmail.com con el respaldo adjunto."""
import os, glob, subprocess, smtplib, ssl, datetime
from email.message import EmailMessage
alertas = open('alerta.txt').read().strip()
run = f"{os.environ['GITHUB_SERVER_URL']}/{os.environ['GITHUB_REPOSITORY']}/actions/runs/{os.environ['GITHUB_RUN_ID']}"
hora = (datetime.datetime.utcnow() - datetime.timedelta(hours=6)).strftime('%d/%m/%Y %H:%M')
cuerpo = (f"Se detectó una posible amenaza en el sitio de la Teletón 2026 ({hora}, hora de Costa Rica):\n\n"
          + '\n'.join('• ' + a for a in alertas.splitlines())
          + f"\n\nRespaldo de la última versión buena del sitio y de los datos (descargable 90 días): {run}\n"
          + "Para volver a levantar el sitio, seguí RESTAURAR.md dentro del respaldo.\n"
          + "Si el cambio lo hiciste vos, podés ignorar este aviso.")
subprocess.run(['gh', 'issue', 'create', '--title', f'⚠ Alerta de seguridad {hora}', '--body', cuerpo], check=False)
u, p = os.environ.get('MAIL_USER'), os.environ.get('MAIL_PASS')
if u and p:
    m = EmailMessage(); m['Subject'] = f'⚠ Teletón 2026: alerta de seguridad ({hora})'; m['From'] = u; m['To'] = os.environ.get('MAIL_TO') or 'Davidsm004@gmail.com'
    m.set_content(cuerpo)
    for z in sorted(glob.glob('respaldo/*.zip'))[-1:]:
        m.add_attachment(open(z, 'rb').read(), maintype='application', subtype='zip', filename=os.path.basename(z))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context()) as s:
        s.login(u, p); s.send_message(m)
    print('Correo enviado')
else:
    print('Sin credenciales de correo: el aviso llega por las notificaciones de GitHub.')

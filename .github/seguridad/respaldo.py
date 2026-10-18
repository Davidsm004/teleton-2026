"""Arma un respaldo listo para volver a levantar el sitio: páginas (rama main) y datos cifrados (rama datos).
Si la alerta vino de un cambio, se respalda la versión ANTERIOR a ese cambio (la última buena)."""
import os, subprocess, zipfile, io, datetime
sh = lambda *a: subprocess.run(a, capture_output=True).stdout
ev, ref, antes = os.environ.get('EVENT', ''), os.environ.get('REF', ''), os.environ.get('BEFORE', '')
alerta = os.path.exists('alerta.txt') and open('alerta.txt').read().strip()
def punto(rama):
    if alerta and ev == 'push' and ref == f'refs/heads/{rama}' and antes and set(antes) - {'0'}: return antes
    return f'origin/{rama}'
os.makedirs('respaldo', exist_ok=True)
hoy = datetime.datetime.utcnow() - datetime.timedelta(hours=6)
nombre = f"respaldo/teleton-2026-respaldo-{hoy:%Y-%m-%d-%H%M}.zip"
with zipfile.ZipFile(nombre, 'w', zipfile.ZIP_DEFLATED) as z:
    for rama, carpeta in (('main', 'sitio'), ('datos', 'datos')):
        p = punto(rama)
        tar = sh('git', 'archive', '--format=zip', p)
        if not tar: continue
        with zipfile.ZipFile(io.BytesIO(tar)) as src:
            for n in src.namelist():
                if n.endswith('/') or n.startswith('.github/'): continue
                z.writestr(f'{carpeta}/{n}', src.read(n))
        z.writestr(f'{carpeta}/VERSION.txt', f'Rama {rama} en {p}\n' + sh('git', 'log', '-1', '--format=%H %cd', p).decode())
    z.writestr('RESTAURAR.md', sh('git', 'show', 'origin/main:RESTAURAR.md') or b'Ver RESTAURAR.md en el repositorio.')
print(nombre)

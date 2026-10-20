"""Revisión de seguridad del sitio Teletón 2026.
Busca señales de un ataque o de cambios no autorizados y, si encuentra algo, deja la alerta en alerta.txt
(el paso siguiente abre un aviso, manda el correo y adjunta el respaldo)."""
import os, re, json, hashlib, subprocess, urllib.request, datetime

def sh(*a):
    return subprocess.run(a, capture_output=True, text=True).stdout.strip()

ev = os.environ.get('EVENT', '')
ref = os.environ.get('REF', '')
alertas = []
def alerta(t): alertas.append(t); print('ALERTA:', t)

sh('git', 'fetch', '--quiet', 'origin', '+refs/heads/*:refs/remotes/origin/*')
ramas = [r.replace('origin/', '') for r in sh('git', 'branch', '-r', '--format=%(refname:short)').split() if r not in ('origin', 'origin/HEAD')]
ESPERADAS = {'main', 'datos'}

# 1. Ramas: solo deben existir main y datos
for r in ramas:
    if r not in ESPERADAS: alerta(f'Apareció una rama que no es del sitio: «{r}».')
for r in ESPERADAS:
    if r not in ramas: alerta(f'Falta la rama «{r}» (¿la borraron?).')
if ev == 'delete': alerta(f"Se borró la {os.environ.get('REF_TYPE', 'referencia')} «{os.environ.get('DREF', '')}».")

# 2. Historial reescrito (push forzado)
if ev == 'push' and os.environ.get('FORCED') == 'true':
    alerta(f'Se reescribió el historial de «{ref}» (push forzado). Puede borrar o esconder cambios.')

# 3. Archivos permitidos en cada rama
PERM_MAIN = re.compile(r'^(index\.html|(captura|finanzas|accesos|produccion)/index\.html|informes/[^/]+\.(js|md)|informes/fuentes/[^/]+\.ttf|lib/[^/]+\.js|README\.md|RESTAURAR\.md|\.nojekyll|\.github/.*)$')
PERM_DATOS = re.compile(r'^(datos\.json|usuarios\.json|sesiones\.json|chat\.json|README\.md|adjuntos/[0-9a-f]{24}\.bin)$')
for rama, perm in (('main', PERM_MAIN), ('datos', PERM_DATOS)):
    if rama not in ramas: continue
    for f in sh('git', 'ls-tree', '-r', '--name-only', f'origin/{rama}').splitlines():
        if not perm.match(f): alerta(f'Archivo inesperado en «{rama}»: {f}')

# 4. Cambios en el propio monitoreo
if ev == 'push' and ref == 'refs/heads/main':
    antes, despues = os.environ.get('BEFORE', ''), os.environ.get('AFTER', '')
    if antes and not set(antes) <= {'0'}:
        cambiados = sh('git', 'diff', '--name-only', antes, despues).splitlines()
        if any(f.startswith('.github/') for f in cambiados):
            alerta('Se modificó el monitoreo de seguridad (.github). Si no lo pediste vos, alguien intentó apagar las alertas.')

# 5. Integridad de las páginas: sin scripts ni conexiones a sitios externos, y con la política de seguridad activa
HOSTS_OK = {'api.github.com', 'raw.githubusercontent.com', 'formspree.io'}
for pag in ('index.html', 'finanzas/index.html', 'captura/index.html', 'accesos/index.html', 'produccion/index.html'):
    try: html = sh('git', 'show', f'origin/main:{pag}')
    except Exception: html = ''
    if not html: alerta(f'Falta la página {pag}.'); continue
    if "default-src 'none'" not in html or 'Content-Security-Policy' not in html:
        alerta(f'La página {pag} perdió la política de seguridad (Content-Security-Policy).')
    if re.search(r'<script[^>]+src=["\']https?://', html, re.I):
        alerta(f'La página {pag} carga un script desde otro sitio.')
    for h in sorted(set(re.findall(r'https?://([a-zA-Z0-9.-]+)', html))):
        if h not in HOSTS_OK: alerta(f'La página {pag} se conecta a un sitio no autorizado: {h}')
    if 'noindex' not in html: alerta(f'La página {pag} ya no pide a los buscadores que no la indexen.')

# 6. Huellas de las librerías y del generador de informes
hp = '.github/seguridad/huellas.json'
if os.path.exists(hp):
    esperadas = json.load(open(hp))
    for f, h in esperadas.items():
        data = subprocess.run(['git', 'show', f'origin/main:{f}'], capture_output=True).stdout
        if not data: alerta(f'Falta el archivo {f}.'); continue
        if hashlib.sha256(data).hexdigest() != h: alerta(f'El archivo {f} fue modificado (no coincide con la versión revisada).')

# 7. Datos: formato correcto, sin borrado masivo, sin cambios fuera de la aplicación
if 'datos' in ramas:
    for f in ('datos.json', 'usuarios.json'):
        raw = sh('git', 'show', f'origin/datos:{f}')
        if not raw: alerta(f'Falta {f} en la rama de datos.'); continue
        try: j = json.loads(raw)
        except Exception: alerta(f'{f} no tiene un formato válido (posible daño).'); continue
        if f == 'datos.json' and not (j.get('iv') and j.get('ct')): alerta('datos.json ya no está cifrado como corresponde.')
        if f == 'usuarios.json' and not (j.get('entries') and j.get('dir')): alerta('usuarios.json está vacío o dañado: nadie podría entrar.')
    hist = sh('git', 'log', '-2', '--format=%H', 'origin/datos', '--', 'datos.json').split()
    if len(hist) == 2:
        a = len(sh('git', 'show', f'{hist[1]}:datos.json')); b = len(sh('git', 'show', f'{hist[0]}:datos.json'))
        if a > 20000 and b < a * 0.4: alerta(f'El archivo de datos se redujo de {a:,} a {b:,} caracteres en el último cambio (posible borrado masivo).')
    desde = os.environ.get('DESDE', '')
    rng = f'{desde}..origin/datos' if desde and not set(desde) <= {'0'} else 'origin/datos~5..origin/datos'
    for linea in sh('git', 'log', '--format=%h|%s', rng).splitlines():
        h, msg = (linea.split('|', 1) + [''])[:2]
        if not re.match(r'^(Datos|Usuarios) \d', msg) and not msg.startswith('Usuarios: '):
            alerta(f'Cambio en los datos hecho fuera de la aplicación (commit {h}: «{msg[:60]}»).')

# 8. El sitio publicado coincide con el repositorio
if ev in ('schedule', 'workflow_dispatch'):
    ult = int(sh('git', 'log', '-1', '--format=%ct', 'origin/main') or 0)
    if ult and datetime.datetime.now().timestamp() - ult > 1800:
        for pag in ('index.html', 'finanzas/index.html', 'captura/index.html', 'accesos/index.html', 'produccion/index.html'):
            url = 'https://davidsm004.github.io/teleton-2026/' + pag.replace('index.html', '')
            try:
                vivo = urllib.request.urlopen(url, timeout=30).read()
                repo = subprocess.run(['git', 'show', f'origin/main:{pag}'], capture_output=True).stdout
                if hashlib.sha256(vivo).digest() != hashlib.sha256(repo).digest(): alerta(f'La página publicada ({url}) no coincide con la del repositorio.')
            except Exception as e:
                print('No se pudo revisar', url, e)

# 9. Recordatorio: renovar el token de GitHub del sitio («Captura Teletón», vence el 18/10/2027). Avisa 2 días antes.
VENCE = datetime.date(2027, 10, 18)
falta = (VENCE - datetime.date.today()).days
if falta <= 2:
    titulo = 'Recordatorio: renovar el token de GitHub «Captura Teletón»'
    abiertos = sh('gh', 'issue', 'list', '--state', 'open', '--search', titulo, '--json', 'title', '--jq', '.[].title')
    if titulo not in abiertos:
        subprocess.run(['gh', 'issue', 'create', '--title', titulo, '--body', f'El token vence el 18/10/2027 (faltan {max(falta, 0)} días). Regeneralo en https://github.com/settings/personal-access-tokens y pasáselo a quien administra el sitio para publicarlo. Si vence, el sitio se puede ver pero no guardar.'], check=False)

open('alerta.txt', 'w').write('\n'.join(alertas))
with open(os.environ['GITHUB_OUTPUT'], 'a') as o: o.write(f"alerta={'1' if alertas else '0'}\n")
print('Revisión terminada:', len(alertas), 'alertas')

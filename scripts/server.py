#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Servidor HTTP que envuelve a tekla_to_allied.py.
N8N lo invoca con un POST a http://python-runner:5000/run

Endpoints:
    GET  /health  - chequeo de vida (devuelve {"status": "ok"})
    POST /run     - ejecuta el script y devuelve su JSON de resultado
    GET  /run     - igual que POST, util para pruebas desde el navegador
"""

from flask import Flask, jsonify
import subprocess
import json
import os
import sys

app = Flask(__name__)
SCRIPT_PATH = '/scripts/tekla_to_allied.py'


@app.route('/health', methods=['GET'])
def health():
    """Endpoint que N8N (o un docker healthcheck) puede usar para
    verificar que el servicio esta vivo."""
    return jsonify({'status': 'ok'})


@app.route('/run', methods=['POST', 'GET'])
def run_script():
    """Ejecuta el script Tekla -> Allied y devuelve su JSON de resultado.

    El script imprime el JSON final en stdout y los logs en stderr.
    Aca leemos ambos, devolvemos el JSON al cliente y adjuntamos los logs
    en el campo 'logs' por si quieren mostrarse en el email.
    """
    try:
        result = subprocess.run(
            ['python3', SCRIPT_PATH],
            capture_output=True,
            text=True,
            timeout=300  # 5 minutos como tope
        )

        stdout = (result.stdout or '').strip()
        stderr = result.stderr or ''

        # El script puede imprimir varios mensajes en stdout, pero el JSON
        # final siempre es la ultima linea no vacia.
        try:
            lines = [l for l in stdout.split('\n') if l.strip()]
            data = json.loads(lines[-1]) if lines else {}
        except (json.JSONDecodeError, IndexError):
            return jsonify({
                'status': 'parse_error',
                'message': 'No se pudo parsear el output del script',
                'stdout': stdout,
                'stderr': stderr,
                'returncode': result.returncode
            }), 500

        # Adjuntar logs para visibilidad en N8N
        data['logs'] = stderr
        return jsonify(data)

    except subprocess.TimeoutExpired:
        return jsonify({
            'status': 'timeout',
            'message': 'El script tardo mas de 5 minutos en terminar'
        }), 500

    except Exception as e:
        return jsonify({
            'status': 'fatal_error',
            'message': str(e)
        }), 500


if __name__ == '__main__':
    print('Servidor Python escuchando en :5000', file=sys.stderr, flush=True)
    # host=0.0.0.0 para que sea accesible desde otros contenedores en la red docker
    app.run(host='0.0.0.0', port=5000, debug=False)

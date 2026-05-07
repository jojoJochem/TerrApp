import logging
logging.basicConfig(level=logging.INFO)

from waitress import serve
from app import app

serve(app, host="0.0.0.0", port=8000, threads=8)
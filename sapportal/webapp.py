"""The Flask app object, shared by every module that has routes."""

from flask import Flask
from flask_cors import CORS

app = Flask('server')
CORS(app)

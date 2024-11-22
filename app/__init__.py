from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from app.routes import main

db = SQLAlchemy()
bcrypt = Bcrypt()

def create_app():
    app = Flask(__name__)
    app.config.from_pyfile('config.py')

    db.init_app(app)
    bcrypt.init_app(app)

    app.register_blueprint(main)

    return app
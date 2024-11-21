import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'dev_secret'
    SQLALCHEMY_DATABASE_URI = 'sqlite:///passwords.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
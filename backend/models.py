from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Usuario(db.Model):
    id            = db.Column(db.Integer, primary_key=True)
    email         = db.Column(db.String(150), unique=True, nullable=False)
    contraseña    = db.Column(db.String(200), nullable=False)
    facturas_usadas = db.Column(db.Integer, default=0)
    es_premium    = db.Column(db.Boolean, default=False)

class InvitadoIP(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ip = db.Column(db.String(50), unique=True, nullable=False)
    facturas_usadas = db.Column(db.Integer, default=0)
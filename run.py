from flask import Flask
from flask_restful import Api
from flask_sqlalchemy import SQLAlchemy
from resources import *
from config import Config
import pymysql
pymysql.install_as_MySQLdb()
db = SQLAlchemy()
app = Flask(__name__)
app.config.from_object(Config)
api = Api(app)


api.add_resource(Excel, "/excel")

if __name__ == "__main__":
    db.init_app(app)
    app.run(port=5000, debug=True)

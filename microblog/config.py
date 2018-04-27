import os

class Config(object):
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'you-will-never-guess'
    
    #Database config
    db_host = 'mow03-rdps13'
    db_port = '1433'
    db_name = 'flask'
    db_user = 'flask'
    db_pass = 'flask'
    
    SQLALCHEMY_DATABASE_URI = 'mssql+pyodbc://{}:{}@{}:{}/{}?driver=SQL+Server+Native+Client+11.0'.format(db_user, db_pass, db_host, db_port, db_name)
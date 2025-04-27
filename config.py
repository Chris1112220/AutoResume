# config.py
DB_USER = 'postgres'
DB_PASSWORD = 'Topher1212'
DB_NAME = 'autoresume_v2'
DB_HOST = 'localhost'
DB_PORT = '5432'

SQLALCHEMY_DATABASE_URI = (
    f'postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
)

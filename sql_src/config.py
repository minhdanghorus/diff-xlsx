# Connection profiles for PostgreSQL databases.
# Each key is a profile name shown to the user when selecting a connection.
# Fill in your actual connection details below.

CONNECTIONS = {
    "local_docker": {
        "host": "localhost",
        "port": 5432,
        "database": "your_database",
        "user": "postgres",
        "password": "postgres",
    },
    # Add more profiles as needed:
    # "staging": {
    #     "host": "staging.example.com",
    #     "port": 5432,
    #     "database": "mydb",
    #     "user": "app_user",
    #     "password": "secret",
    # },
}

from app import create_app
import os

app = create_app()

if __name__ == '__main__':
    debug = os.getenv('FLASK_DEBUG', '').lower() in {'1', 'true', 'yes', 'on'}
    app.run(debug=debug)
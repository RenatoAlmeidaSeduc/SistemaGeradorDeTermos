from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/')
def index():
    return jsonify({'message': 'Hello, world!'})

def lambda_handler(event, context):
    return app(event, context)

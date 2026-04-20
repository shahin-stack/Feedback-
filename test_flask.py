from flask import Flask, jsonify
app = Flask(__name__)

@app.route("/")
def index():
    return {"status": "error", "message": "hello"}, 500

if __name__ == "__main__":
    app.run(port=9099)

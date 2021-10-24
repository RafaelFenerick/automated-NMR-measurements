from __future__ import division

from flask import Flask, request

app = Flask(__name__)

@app.route('/')
def hello_world():
	return 'Hello World'

@app.route('/receiver', methods=['POST'])
def worker():
    # read json + reply
    data = request.get_data()

    result = data
    print(data)

    return result

app.run(host='0.0.0.0', port="7500")
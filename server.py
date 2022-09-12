from flask import Flask, render_template
from services.getFunctions import GetFunctions

app = Flask(__name__)
get_functions = GetFunctions()

@app.route("/")
def index():
    funs = get_functions.get_data("./POB.xlsm")
    return render_template("index.html", funs=funs)



if __name__ == "__main__":
    app.run(port=8080, debug=True)
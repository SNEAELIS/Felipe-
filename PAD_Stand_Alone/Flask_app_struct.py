from flask import Flask, render_template, request
import PAD_Exec

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run', methods=['POST'])
def run_scritp():
    # Get user inputs from the form
    user_input = request.form.get('input_field')

    # Run your script logic
    result = PAD_Exec.main()

    return f'Result: {result}'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
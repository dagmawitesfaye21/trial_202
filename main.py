from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    df = pd.read_excel(file)
    column_name = request.form['column_name']  # Get the column name from the request
    column_data = df[column_name]
    item_counts = column_data.value_counts()
    result_df = pd.DataFrame({'Element': item_counts.index, 'Frequency': item_counts.values})

    # Calculate the sum of the frequency
    frequency_sum = result_df['Frequency'].sum()

    result_directory = 'D:\Final Results'  # Specify the desired directory path
    os.makedirs(result_directory, exist_ok=True)
    result_file = os.path.join(result_directory, 'result.xlsx')
    result_df.to_excel(result_file, index=False)

    # Sort the result dataframe by 'Element' column in alphabetical order
    result_df = result_df.sort_values(by='Element')

    return render_template('result.html', result=result_df.to_dict('records'), frequency_sum=frequency_sum)

if __name__ == '__main__':
    app.run()
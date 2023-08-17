import PySimpleGUI as sg
import mysql.connector
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv()

sg.theme('SandyBeach')

layout = [
    [sg.Text('Please enter your SQL query:')],
    [sg.Multiline(size=(50, 5), key='textbox')],
    [sg.Submit(), sg.Cancel()]
]

window = sg.Window('Report Generator', layout)

while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, 'Cancel'):
        break
    elif event == 'Submit':
        query = values['textbox']

        #db connection
        try:
            connect = mysql.connector.connect(
            host=os.getenv('DB_HOST'),
            user=os.getenv('DB_USER'),
            passwd=os.getenv('DB_PASSWORD'),
            database=os.getenv('DB_NAME')
    )

            df = pd.read_sql(query, connect)
            connect.close()  #Close the connection when done

            #Save to Excel
            excel_file = 'myReport.xlsx'
            df.to_excel(excel_file, index=False)

            sg.popup_ok(f'Data exported to {excel_file}', title='Export Successful')
            
            # Open the Excel file
            os.startfile(excel_file)
        except mysql.connector.Error as err:
            sg.popup_error(f"MySQL Error: {err}", title='Error')
        except Exception as e:
            sg.popup_error(f"An error occurred: {e}", title='Error')

window.close()

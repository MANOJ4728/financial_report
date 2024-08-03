from flask import Flask, request, redirect, render_template, flash, url_for
from werkzeug.utils import secure_filename
import os
import pandas as pd
import json
import datetime

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'supersecretkey'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'excel' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['excel']
    
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        flash('File successfully uploaded')

        main_keys = ["executive summary", "statistics", "income statement", "cash_flow_statement", "balance_sheet", "financial ratios & key insights"]
        sub_keys = ["overview", "key financial metrics", "Rooms available", "Rooms sold", "Occupancy%", "ADR", "RevPAR", "revenue_variance", "expense_variance", "expense_departments", "net_profit/loss_variance", "operating_activities", "investing_activities", "financing_activities", "cash_flow_variance", "Assets", "Liabilities", "Equity", "liquidity ratios", "profitability ratios", "solvency ratios", "Efficiency ratios"]

        file_path = 'final_bookkeepers_template.xlsx'  
        df = pd.read_excel(file_path)

        json_structure = {}
        current_main_key = None
        current_subkey = None

        def convert_value(value):
            if isinstance(value, datetime.datetime): 
                return value.strftime('%Y-%m-%d')
            elif pd.isna(value):
                return ""
            return value

        for index, row in df.iterrows():
            key = row['keys']
            value = row['values']
            
            if key in main_keys:
                current_main_key = key
                json_structure[current_main_key] = {}
                current_subkey = None
            elif key in sub_keys:
                current_subkey = key
                json_structure[current_main_key][current_subkey] = {}
            elif current_main_key and current_subkey:
                json_structure[current_main_key][current_subkey][key] = convert_value(value)
            elif current_main_key:
                json_structure[current_main_key][key] = convert_value(value)

        json_output = json.dumps(json_structure, indent=2)
        print(json_output)

        # Save the JSON to a file
        output_file_path = 'op.json'
        with open(output_file_path, 'w') as f:
            json.dump(json_structure, f, indent=2)

        def read_json_file(file_path):
            with open(file_path, 'r') as file:
                return json.load(file)
        
        def write_json_to_file(json_data, file_path):
            with open(file_path, 'w') as file:
                json.dump(json_data, file, indent=4)

        # Function to copy values into template
        def copy_values_from_input_to_template(input_json, template_json):
            def update_dict(template_dict, input_dict):
                for key, value in input_dict.items():
                    if isinstance(value, dict):
                        if key not in template_dict or not isinstance(template_dict[key], dict):
                            template_dict[key] = {}
                        update_dict(template_dict[key], value)
                    else:
                        template_dict[key] = value

            update_dict(template_json, input_json)
            print("Template JSON after copying:")
            print(json.dumps(template_json, indent=4))
            return template_json

        def format_percentage(value):
            return round(value * 100, 2)

        def format_value(value):
            return round(value, 2)

        def calculate_derived_values(template):
            def format_percentage(value):
                return round(value * 100, 2)  

            def format_value(value):
                return round(value, 2) 

            template["executive summary"]["key financial metrics"]["net profit/loss"] = format_value(
                template["executive summary"]["key financial metrics"]["total revenue"] - 
                template["executive summary"]["key financial metrics"]["total expenses"]
            )
            
            # Rooms available statistics
            template["statistics"]["Rooms available"]["variance amount"] = format_value(
                template["statistics"]["Rooms available"]["rooms available"] - 
                template["statistics"]["Rooms available"]["previous_period"]
            )
            template["statistics"]["Rooms available"]["variance percentage"] = format_percentage(
                template["statistics"]["Rooms available"]["variance amount"] / 
                template["statistics"]["Rooms available"]["previous_period"]
            )
            template["statistics"]["Rooms available"]["budget variance amount "] = format_value(
                template["statistics"]["Rooms available"]["rooms available"] - 
                template["statistics"]["Rooms available"]["budget"]
            )
            template["statistics"]["Rooms available"]["budget variance percentage"] = format_percentage(
                template["statistics"]["Rooms available"]["budget variance amount "] / 
                template["statistics"]["Rooms available"]["budget"]
            )
            
            # Rooms sold statistics
            template["statistics"]["Rooms sold"]["variance amount"] = format_value(
                template["statistics"]["Rooms sold"]["rooms sold"] - 
                template["statistics"]["Rooms sold"]["previous period"]
            )
            template["statistics"]["Rooms sold"]["variance percentage"] = format_percentage(
                template["statistics"]["Rooms sold"]["variance amount"] / 
                template["statistics"]["Rooms sold"]["previous period"]
            )
            template["statistics"]["Rooms sold"]["budget variance amount "] = format_value(
                template["statistics"]["Rooms sold"]["rooms sold"] - 
                template["statistics"]["Rooms sold"]["budget"]
            )
            template["statistics"]["Rooms sold"]["budget variance percentage"] = format_percentage(
                template["statistics"]["Rooms sold"]["budget variance amount "] / 
                template["statistics"]["Rooms sold"]["budget"]
            )
            
            # Occupancy statistics
            template["statistics"]["occupancy"]["occupancy %"] = format_percentage(
                template["statistics"]["Rooms sold"]["rooms sold"] / 
                template["statistics"]["Rooms available"]["rooms available"]
            )
            template["statistics"]["occupancy"]["previous period"] = format_percentage(
                template["statistics"]["Rooms sold"]["previous period"] / 
                template["statistics"]["Rooms available"]["previous_period"]
            )
            template["statistics"]["occupancy"]["variance amount"] = format_value(
                template["statistics"]["occupancy"]["occupancy %"] - 
                template["statistics"]["occupancy"]["previous period"]
            )
            template["statistics"]["occupancy"]["variance percentage"] = format_percentage(
                template["statistics"]["occupancy"]["variance amount"] / 
                template["statistics"]["occupancy"]["previous period"]
            )
            template["statistics"]["occupancy"]["budget"] = format_percentage(
                template["statistics"]["Rooms sold"]["budget"] / 
                template["statistics"]["Rooms available"]["budget"]
            )
            template["statistics"]["occupancy"]["budget variance amount"] = format_value(
                template["statistics"]["occupancy"]["occupancy %"] - 
                template["statistics"]["occupancy"]["budget"]
            )
            template["statistics"]["occupancy"]["budget variance percentage"] = format_percentage(
                template["statistics"]["occupancy"]["budget variance amount"] / 
                template["statistics"]["occupancy"]["budget"]
            )

            # ADR statistics
            template["statistics"]["ADR"]["variance amount"] = format_value(
                template["statistics"]["ADR"]["adr"] - 
                template["statistics"]["ADR"]["previous period"]
            )
            template["statistics"]["ADR"]["variance percentage"] = format_percentage(
                template["statistics"]["ADR"]["variance amount"] / 
                template["statistics"]["ADR"]["previous period"]
            )
            template["statistics"]["ADR"]["budget variance amount "] = format_value(
                template["statistics"]["ADR"]["adr"] - 
                template["statistics"]["ADR"]["budget"]
            )
            template["statistics"]["ADR"]["budget variance percentage"] = format_percentage(
                template["statistics"]["ADR"]["budget variance amount "] / 
                template["statistics"]["ADR"]["budget"]
            )
            
            # RevPAR statistics
            template["statistics"]["RevPAR"]["variance amount"] = format_value(
                template["statistics"]["RevPAR"]["revpar"] - 
                template["statistics"]["RevPAR"]["previous period"]
            )
            template["statistics"]["RevPAR"]["variance percentage"] = format_percentage(
                template["statistics"]["RevPAR"]["variance amount"] / 
                template["statistics"]["RevPAR"]["previous period"]
            )
            template["statistics"]["RevPAR"]["budget variance amount "] = format_value(
                template["statistics"]["RevPAR"]["revpar"] - 
                template["statistics"]["RevPAR"]["budget"]
            )
            template["statistics"]["RevPAR"]["budget variance percentage"] = format_percentage(
                template["statistics"]["RevPAR"]["budget variance amount "] / 
                template["statistics"]["RevPAR"]["budget"]
            )
            
            # Income statement calculations
            template["income statement"]["Revenue"]["Total revenue"] = template["executive summary"]["key financial metrics"]["total revenue"]
            template["income statement"]["revenue_variance"]["actual revenue"] = template["executive summary"]["key financial metrics"]["total revenue"]
            template["income statement"]["revenue_variance"]["variance amount"] = format_value(
                template["income statement"]["revenue_variance"]["actual revenue"] - 
                template["income statement"]["revenue_variance"]["previous period revenue"]
            )
            template["income statement"]["revenue_variance"]["variance percentage"] = format_percentage(
                template["income statement"]["revenue_variance"]["variance amount"] / 
                template["income statement"]["revenue_variance"]["previous period revenue"]
            )
            template["income statement"]["revenue_variance"]["budget variance amount "] = format_value(
                template["income statement"]["revenue_variance"]["actual revenue"] - 
                template["income statement"]["revenue_variance"]["budgeted revenue"]
            )
            template["income statement"]["revenue_variance"]["budget variance percentage"] = format_percentage(
                template["income statement"]["revenue_variance"]["budget variance amount "] / 
                template["income statement"]["revenue_variance"]["budgeted revenue"]
            )
            
            # Expense Variance
            template["income statement"]["expense_variance"]["actual expense"] = template["executive summary"]["key financial metrics"]["total expenses"]
            template["income statement"]["expense_variance"]["variance amount"] = format_value(
                template["income statement"]["expense_variance"]["actual expense"] - 
                template["income statement"]["expense_variance"]["previous period expense"]
            )
            template["income statement"]["expense_variance"]["variance percentage"] = format_percentage(
                template["income statement"]["expense_variance"]["variance amount"] / 
                template["income statement"]["expense_variance"]["previous period expense"]
            )
            template["income statement"]["expense_variance"]["budget variance amount "] = format_value(
                template["income statement"]["expense_variance"]["actual expense"] - 
                template["income statement"]["expense_variance"]["budgeted expense"]
            )
            template["income statement"]["expense_variance"]["budget variance percentage"] = format_percentage(
                template["income statement"]["expense_variance"]["budget variance amount "] / 
                template["income statement"]["expense_variance"]["budgeted expense"]
            )
            #expense departments-total expenses
            template["income statement"]["expense_departments"]["total expenses"] = template["executive summary"]["key financial metrics"]["total expenses"]
            

            # Net Profit/Loss Variance
            template["income statement"]["Net Profit/Loss"]["net profit/loss"] = template["executive summary"]["key financial metrics"]["net profit/loss"]
            template["income statement"]["net_profit/loss_variance"]["actual net profit/loss"] = template["executive summary"]["key financial metrics"]["net profit/loss"]
            template["income statement"]["net_profit/loss_variance"]["variance amount"] = format_value(
                template["income statement"]["net_profit/loss_variance"]["actual net profit/loss"] -
                template["income statement"]["net_profit/loss_variance"]["previous period net profit/loss"]
            )
            template["income statement"]["net_profit/loss_variance"]["variance percentage"] = format_percentage(
                template["income statement"]["net_profit/loss_variance"]["variance amount"] / 
                template["income statement"]["net_profit/loss_variance"]["previous period net profit/loss"]
            )
            template["income statement"]["net_profit/loss_variance"]["budget variance amount "] = format_value(
                template["income statement"]["net_profit/loss_variance"]["actual net profit/loss"] - 
                template["income statement"]["net_profit/loss_variance"]["budgeted net profit/loss"]
            )
            template["income statement"]["net_profit/loss_variance"]["budget variance percentage"] = format_percentage(
                template["income statement"]["net_profit/loss_variance"]["budget variance amount "] / 
                template["income statement"]["net_profit/loss_variance"]["budgeted net profit/loss"]
            )
            
            # Operating Activities
            template["cash_flow_statement"]["operating_activities"]["net cash from operating activities"] = format_value(
                template["cash_flow_statement"]["operating_activities"]["cash inflows"] - 
                template["cash_flow_statement"]["operating_activities"]["cash outflows"]
            )
            
            template["cash_flow_statement"]["investing_activities"]["net cash from investing activities"] = format_value(
                template["cash_flow_statement"]["investing_activities"]["cash inflows"] - 
                template["cash_flow_statement"]["investing_activities"]["cash outflows"]
            )
            
            template["cash_flow_statement"]["financing_activities"]["net cash from financing activities"] = format_value(
                template["cash_flow_statement"]["financing_activities"]["cash inflows"] - 
                template["cash_flow_statement"]["financing_activities"]["cash outflows"]
            )
            
            template["cash_flow_statement"]["financing_activities"]["net increase/decrease in cash"] = format_value(
                template["cash_flow_statement"]["operating_activities"]["net cash from operating activities"] + 
                template["cash_flow_statement"]["investing_activities"]["net cash from investing activities"] + 
                template["cash_flow_statement"]["financing_activities"]["net cash from financing activities"]
            )

            # Cash Flow Variance
            template["cash_flow_statement"]["cash_flow_variance"]["actual cash flow"] = format_value(
                template["cash_flow_statement"]["financing_activities"]["net increase/decrease in cash"]
            )

            template["cash_flow_statement"]["cash_flow_variance"]["variance amount"] = format_value(
                template["cash_flow_statement"]["cash_flow_variance"]["actual cash flow"] - 
                template["cash_flow_statement"]["cash_flow_variance"]["previous period cash flow"]
            )
            template["cash_flow_statement"]["cash_flow_variance"]["variance percentage"] = format_percentage(
                template["cash_flow_statement"]["cash_flow_variance"]["variance amount"] / 
                template["cash_flow_statement"]["cash_flow_variance"]["previous period cash flow"]
            )

            # Financial ratios
            
            # Liquidity ratios
            template["financial ratios & key insights"]["liquidity ratios"]["current ratio"] = format_value(
                template["balance_sheet"]["Assets"]["current assets"] / 
                template["balance_sheet"]["Liabilities"]["current liabilities"]
            )
            
            template["financial ratios & key insights"]["liquidity ratios"]["quick ratio"] = format_value(
                (template["balance_sheet"]["Assets"]["current assets"] - 
                template["financial ratios & key insights"]["liquidity ratios"]["inventory"] - 
                template["financial ratios & key insights"]["liquidity ratios"]["prepaid_expenses"]) /
                template["balance_sheet"]["Liabilities"]["current liabilities"]
            )
            
            # Profitability ratios
            template["financial ratios & key insights"]["profitability ratios"]["gross profit margin"] = format_percentage(
                template["financial ratios & key insights"]["profitability ratios"]["gross_operating_profit"] / 
                template["income statement"]["Revenue"]["Total revenue"]
            )
            
            template["financial ratios & key insights"]["profitability ratios"]["net profit margin"] = format_percentage(
                template["executive summary"]["key financial metrics"]["net profit/loss"] / 
                template["income statement"]["Revenue"]["Total revenue"]
            )
            
            template["financial ratios & key insights"]["profitability ratios"]["return on equity"] = format_percentage(
                template["executive summary"]["key financial metrics"]["net profit/loss"] / 
                template["balance_sheet"]["Equity"]["total equity"]
            )
            
            # Solvency ratios
            template["financial ratios & key insights"]["solvency ratios"]["debt to equity ratio"] = format_value(
                template["balance_sheet"]["Liabilities"]["total liabilities"] / 
                template["balance_sheet"]["Equity"]["total equity"]
            )
            
            template["financial ratios & key insights"]["solvency ratios"]["interest coverage ratio"] = format_value(
                template["financial ratios & key insights"]["solvency ratios"]["ebit"] / 
                template["income statement"]["expense_departments"]["interest expense"]
            )
            
            # Efficiency ratios
            template["financial ratios & key insights"]["Efficiency ratios"]["asset turnover ratio"] = format_value(
                template["executive summary"]["key financial metrics"]["net profit/loss"] / 
                (template["financial ratios & key insights"]["Efficiency ratios"]["previous_period_opening_total_assets"] + 
                template["balance_sheet"]["Assets"]["total assets"])
            )

            return template

        def create_final_json(input_json_path, template_json_path, output_folder):
            input_json = read_json_file(input_json_path)
            template_json = read_json_file(template_json_path)
            template_json = copy_values_from_input_to_template(input_json, template_json)
            extended_json = calculate_derived_values(template_json)
            
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            output_json_path = os.path.join(output_folder, 'output.json')
            write_json_to_file(extended_json, output_json_path)
            return extended_json

        input_json_path = 'op.json'
        template_json_path = 'template.json'
        output_folder_path = 'output'

        create_final_json(input_json_path, template_json_path, output_folder_path)
        
        return redirect(url_for('report'))

    else:
        flash('Allowed file types are xls, xlsx')
        return redirect(request.url)

@app.route('/report')
def report():
    # Read JSON data
    with open('output/output.json') as json_file:
        data = json.load(json_file)
    
    return render_template('report.html', data=data)


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)

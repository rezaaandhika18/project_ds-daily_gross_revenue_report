# Create daily gross revenue report each product and each Gender use OOP
# Dataset - https://www.kaggle.com/datasets/aungpyaeap/supermarket-sales
import json
from product import *

# call class
pr = product()

# Opening config json file
f = open('configs/config.json')
data = json.load(f)

pr.input_file = data['input_file']
pr.output_file = data['output_file']
pr.date = '2019-02-05'
pr.sheet_name = 'Report Daily Gross Product'
pr.with_chart = True
pr.with_total = True
pr.colomns = 'Product line'
pr.values = 'gross income'
pr.index = 'Gender'
pr.title = 'Sales Report '+pr.date

if __name__ == "__main__":
    # process ...
    pr.process_daily()
import openpyxl
import psycopg2


def connect():
    conn = None
    try:
        conn = psycopg2.connect(database='pharmacystock', user='django', password='django', host='localhost')
        print("Connected!")

        return conn

    except(Exception) as e:
        print(e)

#Function to write values to the database
def write_data(drugs_and_prices, drugs_and_purchaserates,
               drugs_and_quantitities, drugs_and_categories, drugs_list):
    conn = connect()
    cur = conn.cursor()

    try:
        for each_drug in drugs_list:
            price_perunit = drugs_and_prices[each_drug]
            purchase_rate = drugs_and_purchaserates[each_drug]
            quantity = drugs_and_quantitities[each_drug]
            category = drugs_and_categories[each_drug]

            cur.execute("""INSERT INTO mainapp_drug (product_name, purchase_rate, sales_rate,
                        category)
                        VALUES (%s, %s, %s, %s)"""
                        , (each_drug, purchase_rate, price_perunit, category,))

            conn.commit()
        print("Data saved Successfully")

    except(Exception) as e:
        print(e, "Fail to Insert Data")



#Function to format data into dictionaries of drugs and prices and drugs and purchase rates
def format_data(drugs_list, price_list, purchaserate_list, quantity_list, category_list):
    drugs_and_prices = {}
    drugs_and_purchaserates = {}
    drugs_and_quantities = {}
    drugs_and_categories = {}

    for each_drug_index in range(0, len(drugs_list)):
        drugs_and_prices[drugs_list[each_drug_index]] = price_list[each_drug_index]
    print(drugs_and_prices)

    for each_drug_index in range(0, len(drugs_list)):
        drugs_and_purchaserates[drugs_list[each_drug_index]] = purchaserate_list[each_drug_index]

    for each_drug_index in range(0, len(drugs_list)):
        drugs_and_quantities[drugs_list[each_drug_index]] = quantity_list[each_drug_index]

    for each_drug_index in range(0, len(drugs_list)):
        drugs_and_categories[drugs_list[each_drug_index]] = category_list[each_drug_index]

    print(drugs_and_purchaserates)

    write_data(drugs_and_prices, drugs_and_purchaserates,
               drugs_and_quantities, drugs_and_categories, drugs_list)


#Loading the Excel file to read from
#Using openpyxl, file's extension must be .xlsx
products_file = openpyxl.load_workbook("newproducts.xlsx")

#Choosing the sheet of the file to read from
sheet = products_file["A"]

#Function to copy data from the excel file (copies just a single column)
def copyData(startRow, endRow, col,  sheet):
    data_selected = []

    for i in range(startRow, endRow+1):
        data = sheet.cell(row=i, column=col).value
        #print(data)

        #Removing white spaces on the right of each drug's name
        if type(data) is not int:
            data_selected.append(data.rstrip())
        else:
            data_selected.append(data)

    #print(data_selected)

    return data_selected


if __name__ == '__main__':
    drugs_list = copyData(2, 4219, 2, sheet)
    price_list = copyData(2, 4219, 7, sheet)
    purchaserate_list = copyData(2, 4219, 6, sheet)
    category_list = copyData(2, 4219, 5, sheet)
    quantity_list = copyData(2, 4219, 4, sheet)
    print(len(drugs_list), len(price_list), len(purchaserate_list), len(category_list), len(quantity_list))
    format_data(drugs_list, price_list, purchaserate_list, quantity_list, category_list)




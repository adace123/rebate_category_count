import pygsheets
import random
import psycopg2

conn = None
rebate_sheet = None

def connect_to_db(host, db, user, password):
    global conn
    try:
        conn = psycopg2.connect(host=host, database=db, user=user, password=password)
    except:
        raise Exception("Could not connect to database")
    
def get_sheet():
    global rebate_sheet
    gc = pygsheets.authorize(outh_file='client_secret.json')
    rebate_sheet = gc.open('rebate_count_test').worksheets()[0]
    return rebate_sheet.range("A2:B21")

# for local testing only
def fill_sheet_rebates():
    states = ['AL', 'AK','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD']
    states.extend(('MA','MI','MN','MS','MO','MN','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA'))
    states.extend(('RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY'))
    categories = ['Gas Water Heater', 'Electric Water Heater', 'Dryer', 'Thermostat', 'Freezer', 'Refrigerator']
    for i in range(2, 22):
        rebate_sheet.update_cell('A{}'.format(i), random.choice(states))
        rebate_sheet.update_cell('B{}'.format(i), 'utility_{}'.format(i - 1))

def fill_category_count(index, value):
    global rebate_sheet
    #if(value > 0):
    rebate_sheet.update_cell('C{}'.format(index), value)

def get_category_count(state, utility):
    global conn
    db = conn.cursor()
    db.execute("""
    SELECT COUNT(DISTINCT category) FROM rebates WHERE 
    state = '{}' AND utility = '{}'
    """.format(state, utility))
    result = db.fetchone()
    db.close()
    return result[0]

connect_to_db("localhost", "test", "Aaron", "") 

for state, utility in get_sheet():
    # print(state.value, utility.value)
    count = get_category_count(state.value, utility.value)
    fill_category_count(state.row, count)
# columns = list(filter(lambda x: x, rebate_sheet[0]))

conn.close()
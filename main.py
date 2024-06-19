import pyodbc

#driver = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
#print(f'MS-Access Drivers: {driver}')

#Checking connection
try:
    con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\charl\Desktop\VSCode2\Access Project1\pythonDB.accdb;'
    conn = pyodbc.connect(con_string)
    print("Connected To Database")

except pyodbc.Error as e:
    print("Error in Connection", e)

#Practice
try:
    cursor = conn.cursor()

    # =============== Add data ===============
    newUser = (
        (12, 'Water', 'Water@gmail.com', 'dihab'),
        (13, 'python', 'python@gmail.com', 'MotDePasse'),
        (14, 'java', 'java@gmail.com', 'CanIGiveYouADeliciousFresca?'),
    )
    cursor.executemany('INSERT INTO users VALUES (?,?,?,?)', newUser)
    conn.commit()
    print('Data Inserted')


    # =============== Remove data =============== 
    # Deleting user ID 2
    user_id = 2
    cursor.execute('DELETE FROM users WHERE id = ?', (user_id))
    conn.commit()
    print("Data Deleted ")

    # =============== Iterating =============== 
    # Part 1: Printing everything
    cursor.execute('SELECT * FROM Users')
    for row in cursor.fetchall():
        print(row)

    # Part 2: Printing Specific Columns
    cursor.execute('SELECT Name FROM Users')
    for names in cursor.fetchall():
        print(names)
    

    # =============== Searching ===============
    # Manually
    cursor.execute('SELECT * FROM Users')
    for col in cursor.fetchall():
        name = col[1]  # Access the first element of the tuple
        if name == 'Walter':
            print('Found user named Walter:',col)
    # With implemented commands
    cursor.execute('SELECT * FROM Users WHERE Name = ?', ('Walter',))
    print('Found user named Walter:',cursor.fetchall()[0])


except pyodbc.Error as e:
    print("Error in connection", e)

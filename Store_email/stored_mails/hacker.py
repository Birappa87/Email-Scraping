'''Importing the required libraries'''
import os
import pandas as pd
from bs4 import BeautifulSoup
import pyodbc

def getdata(file):
    '''Get the html content'''
    HTMLFile = open(file, "r")
  
    # Reading the file
    index = HTMLFile.read()

    #parsing into beautifulsoup
    soup = BeautifulSoup(index, 'html.parser')
    return soup

def GetProfileName(soup):
    '''Profile Name'''
    try:
        profile = soup.find('div', {'class' : 'text profile-name'}).text
        return profile.strip()
    except:
        pass
    try:
        profile = soup.find('div', {'class' : 'text'}).text
        return profile.strip()
    except:
        pass

def gettable_data(soup):
    '''Table Data which contains the detailed information'''
    dict_ = {}
    try:
        parent_tag = soup.find('td', {'class': 'divider-top detail-list-padding'})
        table_data = parent_tag.find_all('tr', {'class' : 'detail-row'})
        for i in table_data:
            label = i.find('div', {'class' : 'label'}).text.strip()
            value = i.find('div', {'class' : 'value'}).text.strip()
            if label == 'Amount':
                dict_[label] = value
            elif label == 'Destination':
                dict_[label] = value
            elif label == 'Identifier':
                dict_[label] = value
            elif label == 'To':
                dict_[label] = value
            elif label == 'From':
                dict_[label] = value
            else:
                pass
    except:
        # If key not is present, will add to dictionary
        try:
            keynote = soup.find('div', {'class' : 'text note'}).text.strip()
            dict_['key note'] = keynote
        except:
            pass
        
        #if only description is given
        try:
            keynote = soup.find('div', {'class' : 'subtitle text'}).text.strip()
            dict_['key_note'] = keynote
        except:
            pass

        #if only description is given
        try:
            keynote = soup.find('td', {'class' : 'mobBodyStandardFontSize mobBodyStandardLineHeight'})
            txt = keynote.find('span').text.strip()
            dict_['key_note'] = txt
        except:
            pass

    return dict_

def writing_to_database(df):
    #Database connection string
    conn = pyodbc.connect(
        Trusted_Connected = 'Yes',
        Driver = {'Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)'},
        Server = 'BIRAPPA_G',
        Database = 'employee'
    )
    #Database cursor
    cursor = conn.cursor()

    #creating table
    cursor.execute('''
        CREATE TABLE mail_info 
        (   'path varchar(500)',
            'name varchar(50)',
            'From varchar(50)', 
            'To varchar(50)', 
            'Amount float',
            'Identifier varchar(25)', 
            'Destination varchar(1000)',
            'key_note varchar(5000)'
        )
        ''')
    
    #inserting records 
    for row in df.itertuples():
        cursor.execute(
            '''INSERT INTO employee.dbo.mail_info VALUES (?, ?, ?, ?,?,?,?,?)''',
            row.path,
            row.name,
            row.From,
            row.To,
            row.Amount,
            row.Identifier,
            row.Destination,
            row.key_note
        )
    cursor.commit()

def clean_data(df):
    '''func to extract ID from path using pandas'''
    def extract_id(path):
        id = path.split('-')[1].replace('.html','')
        return id


    df['Destination'] = 'CashApp'
    df.dropna(subset=['Amount'], inplace = True)  #using subset we can remove null values from specific columns
    df['Id'] = df.apply(lambda x : extract_id(x['path']), axis=1) #axis = 1 implies filter by row wise
    df.to_csv('email_stored.csv')


def main():
    '''Reading and writing the files'''
    folderpath = r"C:\Users\birap\OneDrive\Desktop\Conda Analytics\Store_email\stored_mails\stored_mails"
    filepaths  = [os.path.join(folderpath, name) for name in os.listdir(folderpath)]

    details = []
    for path in filepaths:
        #process only the html files
        if path.endswith('html'):
            soup = getdata(path)
            name = GetProfileName(soup)
            table_data = gettable_data(soup)
            table_data['name'] = name
            table_data['path'] = path
            if table_data['name'] == 'Uber':
                table_data['To'] = 'Uber'
                table_data['From'] = 'CHRISTOPHER C MAGHAS'
            details.append(table_data)
            

    #Converting dictionary to dataframe using Pandas {key : columns , values : rows}
    df = pd.DataFrame(details)
    clean_data(df)

if __name__ == "__main__":
    main()
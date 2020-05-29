import pyodbc
import sys
import csv
mypassword = sys.argv[1]
myserver='aix7.prod.bcidaho.loc'
myport=10060
mydatabase='bcidaho_db'
myuserid ='bhar2001'
mydriver ='Adaptive Server Enterprise'
host = 'aix7.prod.bcidaho.loc:10060'

myserver2='aix5.prod.bcidaho.loc'
myport2=10032
mydb2='bcidaho_db'
host2 = 'aix5.prod.bcidaho.loc:10032'

cnxsrc = pyodbc.connect(driver='{Adaptive Server Enterprise}', server= myserver, port=myport,trusted_connection='yes', user=myuserid, password=mypassword,mydb=mydb2)
connsrc=cnxsrc.cursor()
rows = connsrc.execute('''
    select 
    SBSB_ID,
    subscriber_check_amt,
    zip,
    MEME_CK,
    CSPI_CARD_STOCK,
    CLCL_PAID_DT,
    CLCL_PAY_PR_IND,
    rpt_sort,
    rpt_segment,
    rpt_name,
    rpt_array
    From bcidaho_db.dbo.bci_cl_crspd_eob_rpt 
''')
# rows = connsrc.fetchall()


connsrc.commit()

# cnxtgt = pyodbc.connect(driver='{Adaptive Server Enterprise}', server= myserver2, port=myport2,trusted_connection='yes', user=myuserid, password=mypassword, mydb = 'extract_db')
# conntgt=cnxtgt.cursor()
# print(rows)
with open(r'C:\Workspace\DentalEobRunBook\DentalEob.csv', 'w+', encoding='utf-8' ,newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow([x[0] for x in connsrc.description])  # column headers
    for row in rows:
        writer.writerow(row)

print('File has been loaded successfully')



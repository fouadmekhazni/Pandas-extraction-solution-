import win32com.client as win32

def create_connection(server_name, database_name):
    def _connection_string(database_name):
        """
        uid --> username
        pwd --> password
        """
        connection_string = f"""
            Provider=MSDASQL.1;
            driver={{SQL Server}};
            server={server_name};
            database={database_name};
        """
        return connection_string
    conn_string = _connection_string(database_name)
    try:
        conn = win32.Dispatch('ADODB.Connection')
        conn.Open(conn_string)
        return conn
    except Exception as e:
        raise Exception(e)

def run_query(conn, sql_query):
    rst = win32.Dispatch('ADODB.Recordset')
    try:
        rst.Open(sql_query, conn)
        return rst
    except Exception as e:
        raise Exception(e)
    
def copy_to_excel(recordset):
    xlApp = win32.Dispatch('Excel.Application')
    xlApp.Visible = True

    wb = xlApp.Workbooks.Add()
    
    ws = wb.Sheets(1)

    if not recordset.EOF:
        ws.Range('B1').CopyFromRecordset(recordset)

        for i in range(recordset.Fields.count):
            ws.Cells(1, i+2).Value = recordset.Fields(i).Name


    wb.SaveAs('NEW PDR 1.xlsx')
    wb.Close()
    

    xlApp.Quit()

SERVER_NAME = '192.168.1.5'  # 
DATABASE_NAME = 'Database1'  
conn = create_connection(SERVER_NAME, DATABASE_NAME)

sql_query = """
SELECT  [No_]      ,[No_ 2]      ,[Description]      ,[Search Description]      ,[Description 2]      ,[Bill of Materials]      ,[Base Unit of Measure]      ,[Price Unit Conversion]      ,[Inventory Posting Group]      ,[Shelf No_]      ,[Item Disc_ Group]      ,[Allow Invoice Disc_]      ,[Statistics Group]      ,[Commission Group]      ,[Unit Price]      ,[Price_Profit Calculation]      ,[Profit %]      ,[Costing Method]      ,[Unit Cost]      ,[Standard Cost]      ,[Last Direct Cost]      ,[Indirect Cost %]      ,[Cost is Adjusted]      ,[Allow Online Adjustment]      ,[Vendor No_]      ,[Vendor Item No_]      ,[Lead Time Calculation]      ,[Reorder Point]      ,[Maximum Inventory]      ,[Reorder Quantity]      ,[Alternative Item No_]      ,[Unit List Price]      ,[Duty Due %]      ,[Duty Code]      ,[Gross Weight]      ,[Net Weight]      ,[Units per Parcel]      ,[Unit Volume]      ,[Durability]      ,[Freight Type]      ,[Tariff No_]      ,[Duty Unit Conversion]      ,[Country_Region Purchased Code]      ,[Budget Quantity]      ,[Budgeted Amount]      ,[Budget Profit]      ,[Blocked]      ,[Last Date Modified]      ,[Price Includes VAT]      ,[VAT Bus_ Posting Gr_ (Price)]      ,[Gen_ Prod_ Posting Group]      ,[Country_Region of Origin Code]      ,[Automatic Ext_ Texts]      ,[No_ Series]      ,[Tax Group Code]      ,[VAT Prod_ Posting Group]      ,[Reserve]      ,[Global Dimension 1 Code]      ,[Global Dimension 2 Code]      ,[Low-Level Code]      ,[Lot Size]      ,[Serial Nos_]      ,[Last Unit Cost Calc_ Date]      ,[Rolled-up Material Cost]      ,[Rolled-up Capacity Cost]      ,[Scrap %]      ,[Inventory Value Zero]      ,[Discrete Order Quantity]      ,[Minimum Order Quantity]      ,[Maximum Order Quantity]      ,[Safety Stock Quantity]      ,[Order Multiple]      ,[Safety Lead Time]      ,[Flushing Method]      ,[Replenishment System]      ,[Rounding Precision]      ,[Sales Unit of Measure]      ,[Purch_ Unit of Measure]      ,[Reorder Cycle]      ,[Reordering Policy]      ,[Include Inventory]      ,[Manufacturing Policy]      ,[Manufacturer Code]      ,[Item Category Code]      ,[Created From Nonstock Item]      ,[Product Group Code]      ,[Service Item Group]      ,[Item Tracking Code]      ,[Lot Nos_]      ,[Expiration Calculation]      ,[Special Equipment Code]      ,[Put-away Template Code]      ,[Put-away Unit of Measure Code]      ,[Phys Invt Counting Period Code]      ,[Last Counting Period Update]      ,[Next Counting Period]      ,[Use Cross-Docking]      ,[Kit BOM No_]      ,[Kit Disassembly BOM No_]      ,[Components on Sales Orders]      ,[Components on Shipments]      ,[Components on Invoices]      ,[Components on Pick Tickets]      ,[Roll-up Kit Pricing]      ,[Automatic Build Kit BOM]      ,[Prorated VAT]      ,[Pro_ VAT Prod_ Posting Group]      ,[Receip Unit]      ,[Routing No_]      ,[Production BOM No_]      ,[Single-Level Material Cost]      ,[Single-Level Capacity Cost]      ,[Single-Level Subcontrd_ Cost]      ,[Single-Level Cap_ Ovhd Cost]      ,[Single-Level Mfg_ Ovhd Cost]      ,[Overhead Rate]      ,[Rolled-up Subcontracted Cost]      ,[Rolled-up Mfg_ Ovhd Cost]      ,[Rolled-up Cap_ Overhead Cost]      ,[Order Tracking Policy]      ,[Critical]      ,[Common Item No_]      ,[Shelf No_ BARAKI]      ,[Shelf No_ SETIF]      ,[Shelf No_ AKBOU]  

FROM "Tchin Lait WMS$Item"
WHERE [No_]  LIKE '%PDR%' 
"""
recordset = run_query(conn, sql_query)


if not recordset.EOF:

 
    copy_to_excel(recordset)

    '''
    # Afficher les noms des colonnes
    for i in range(recordset.Fields.Count):
        print(recordset.Fields(i).Name, end='\t')
    print()  # Nouvelle ligne pour les données

    # Afficher les données
    while not recordset.EOF:
        for i in range(recordset.Fields.Count):
            print(recordset.Fields(i).Value, end='\t')
        print()  # Nouvelle ligne pour les données
        recordset.MoveNext()
    '''


conn.Close()



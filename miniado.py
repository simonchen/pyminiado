# A mini library used with connecting MS Access DB using ADO from Win32 COM.
import types
#import win32api
import codecs
from win32com import client
import pythoncom
from threading import Lock

def reform_sql(sql, args=[]):
    if args is None:
        return sql
    
    #print sql, args
    newargs = []
    for arg in args:
        if isinstance(arg, types.StringTypes):
            newargs.append("'" + arg.replace("'", "''") + "'")
        else:
            if arg is None:
                arg = 'NULL'
            else:
                arg = str(arg)
            newargs.append(arg)
    sql = sql %tuple(newargs)
    
    return sql

class Errors(Exception):
    CONNECTION_FAILED = 1
    OPEN_RECORDSET_FAILED = 2
    EXECUTE_SQL_FAILED = 3
    
    def __init__(self, value, sql, errcode, msg, exc=None):
        self.value = value
        self.sql = sql
        self.errcode = errcode
        self.msg = msg
        self.exc = exc
    
    def __str__(self):
        str = 'Unknown error!'
        
        if self.value == Errors.CONNECTION_FAILED:
            str = 'Connection is failed, please check datasource , user name, password is or not correct!'
        elif self.value == Errors.OPEN_RECORDSET_FAILED:
            str = 'Open recordset is failed, sql = %s' %self.sql
        elif self.value == Errors.EXECUTE_SQL_FAILED:
            str = 'Executing sql statements is failed, sql = %s' %self.sql
        
        #str += u"\nThe ADO call failed with code %d: %s" % (self.errcode, self.msg.encode('utf16'))
        if self.exc is None:
            str += "\nThere is no extended error information"
        else:
            wcode, source, text, helpFile, helpId, scode = self.exc
            str += "\nThe source of the error is %s" %source.encode('mbcs')
            str += "The error message is %s" %text.encode('mbcs')
            str += "More info can be found in %s (id=%d)" % (helpFile, helpId)    
        
        #print str
        return str

class AdoDB:
    def __init__(self, filepath="", username="", passwd="", sync=False):
        self.conn = client.Dispatch("ADODB.Connection")
        self.sync = sync
        self.lock = Lock()
        try:
            self.conn.Open('PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=%s;Persist Security Info=False;'%filepath, username, passwd)
            self.rs = client.Dispatch("ADODB.Recordset")
        except pythoncom.com_error, (hr, msg, exc, arg):
            raise Errors(Errors.CONNECTION_FAILED, "", hr, msg, exc)
        
    def __del__(self):
        if self.conn:
            self.conn.Close()
            
    def get_tables(self):
        '''Get list of tables'''
        tables = []
        cat = client.Dispatch("ADOX.CataLog")
        cat.ActiveConnection = self.conn
        for x in cat.Tables:
            if x.Type == 'TABLE':
                tables.append(x.Name.encode('utf-8'))
                
        return tables
            
    def execute(self, sql, args=[], last_insert_id=False):
        '''Used for update, insert sql statements. returning (rows, description)'''
        sql = reform_sql(sql, args)
        result = None
        
        try:
            if self.sync:
                #print 'waiting lock..'
                self.lock.acquire()
                
            self.conn.Execute(sql)
            if last_insert_id:
                result = self.execute_sel('SELECT @@IDENTITY')[0][0][0]
        except pythoncom.com_error, (hr, msg, exc, arg):
            raise Errors(Errors.EXECUTE_SQL_FAILED, sql, hr, msg, exc)
        finally:
            if self.sync:
                try:
                    #print 'unlocking...'
                    self.lock.release()
                except:
                    pass
        
        return result
        
    def execute_sel(self, sql, args=[]):
        '''Used for select sql statements. returning (rows, description)'''
        sql = reform_sql(sql, args)

        try:
            self.rs.Open(sql, self.conn, 1, 3) # 1, 3 means adOpenKeyset and adLockOptimistic 
        except pythoncom.com_error, (hr, msg, exc, arg):
            raise Errors(Errors.OPEN_RECORDSET_FAILED, sql, hr, msg, exc)
        
        #description with fields
        desc = []
        for i in range(self.rs.Fields.Count):
            desc.append((self.rs.Fields.Item(i).Name.encode('utf-8'), self.rs.Fields.Item(i).Type, self.rs.Fields.Item(i).DefinedSize))
        
        #rows
        rows = []
        while 1:
            if self.rs.EOF:
                break
            
            row = []
            for field in desc:
                v = self.rs.Fields.Item(field[0]).Value
                if v is not None and isinstance(v, types.StringTypes):
                    row.append(v.encode('utf-8'))
                else:
                    row.append(v)
                
            rows.append(row)
            
            self.rs.MoveNext()
        
        self.rs.Close()
            
        return (rows, desc)
    
    def convertToDictList(self, rows, desc):
        fields = [d[0] for d in desc]
        for row in rows:
            yield dict(zip(fields, row))

        
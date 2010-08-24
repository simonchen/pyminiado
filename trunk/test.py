# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
#
# The test.py aims to test the module to connect MS Access DB.
# You can read more details via the source code and comments below.
#
# Author: Xinyi Chen
# Created at: 2010/8/10
# Last modified: 2010/8/24
#
# Email: simonchen@likefreelancer.com
#
# @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
import imp, os, sys, time, string, random, traceback
from threading import Thread

import miniado # import the miniado module

_debug = False

def main_is_frozen():
    ''' main_is_frozen() returns True when running the exe, and False when running from a script. '''
    return (hasattr(sys, "frozen") or # new py2exe
           hasattr(sys, "importers") # old py2exe
           or imp.is_frozen("__main__")) # tools/freeze

def get_main_dir():
   if main_is_frozen():
       return os.path.dirname(sys.executable)
   return os.path.dirname(sys.argv[0])

# Hard code with db path, you can change it according to you needs.
db_path = os.path.join(get_main_dir(), 'test.mdb')

# The TableObject is an abstract class to describe the structure of the table.
# You shouldn't use it as an instance.
class TableObject:
    table = '' # The defination of the table name
    fields = [] # The defination of the table fields as an array
    
    # A static variable - db is for all instances of class who's derived from
    # TableObject. the variable aims to construct a database object of ADO,
    # then you can utilize it to query / execute Sql statements as soon.
    # Note: the sync parameter indicates that we now use
    db = miniado.AdoDB(db_path, username='', passwd='', sync=True)
    
    def table_exists(self, table=None):
        tables = self.db.get_tables()
        if not table:
            table = self.table
            
        if table not in tables:
            return False
        return True
    
    def insert(self, args={}, want_last_id=False):
        sql = 'insert into %s(%s) values(%s)'
        keys, values, params = [], [], []
        for key in args:
            keys.append(key)
            values.append('%s')
            params.append(args[key])

        sql = sql %(self.table, ','.join(keys), ','.join(values))
        try:
            result = self.db.execute(sql, params, True)
            if want_last_id:
                print 'Last insert id: %d' %result
        except miniado.Errors, e:
            print e
    
# The Table_Of_Test is derived from TableObject class,
# it will be instantiated for testing purpose. (see below)
class Table_Of_Test(TableObject):
    table = 'Table1'
    
    fields = [
        'id', # Auto number
        'field1', # Integer
        'field2', # Text
        'field3', # Double
        'field4', # Date time
    ]
    
    def create_table(self):
        sql = 'create table %s(%s)'
        toadd = []

        for field in self.fields:
            if field == 'id':
                toadd.append('id counter not null primary key')
            elif field == 'field1':
                toadd.append('%s int' %field)
            elif field == 'field2':
                toadd.append('%s text' %field)
            elif field == 'field3':
                toadd.append('%s double' %field)
            elif field == 'field4':
                toadd.append('%s datetime' %field)

        sql = sql %(self.table, ','.join(toadd))
        try:
            self.db.execute(sql)
        except miniado.Errors, e:
            print e
            
    def write_table(self):
        for i in range(0, 10):
            item = {
                'field1': random.randint(0, 100),
                'field2': ''.join(random.sample(string.ascii_letters, 10)),
                'field3': random.randint(0, 100) * 0.5,
                'field4': time.strftime('%Y/%m/%d %H:%M:%S', time.localtime())
            }
            # calling the insert method derived from parent class.
            self.insert(item, want_last_id=True) 
            
            print 'Sleep 1 sec.'
            time.sleep(1)
        
    def read_table(self):
        for field in self.fields:
            sql = 'select * from %s order by %s asc' %(self.table, field)
            print 'SQL = %s' %sql
            try:
                # calling db.execute_sel method to fetch records
                rows, desc = self.db.execute_sel(sql)
                # the returned rows & descriptions of fields wouldn't be friendly,
                # we can convert them to an array of records represtented by dictionaries
                result = self.db.convertToDictList(rows, desc)

                print 'The result is:'
                for r in result:
                    print repr(r);
            except miniado.Errors, e:
                print e
                
            print 'Sleep 2 secs.'
            time.sleep(2)
        
    def test_simple(self):
        print 'Try to create table..'
        self.create_table()
        print 'Try to write few records..'
        self.write_table()
        print 'Try to read few records..'
        self.read_table()
        
        
if __name__ == '__main__':
    test = Table_Of_Test()
    test.test_simple()
    

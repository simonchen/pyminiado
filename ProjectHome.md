A lightweight python module to connect Microsoft Access Database. the module is to use ADO engine through Win32 COM. the module need another Win32 API module, you can find it at the below url:

Pyminiado是一个轻量级的Python访问Access数据库的接口，单连接(connection)，插入／更新可选同步模式，支持UTF-8，多线程。需要下面的Win32 COM模块支持:

http://python.net/crew/mhammond/win32/Downloads.html


<br /><br />
The key features to the module:<br />
1) Supports multi-threading to insert records as exclusively.<br /><br />
2) Supports UTF8/Unicode string as input data in insert/update/select, The returned string is as UTF8 encoding when using 'select' statement.<br /><br />
3) Provides exception handling, very simple to use.<br /><br />
4) Provides sample code guiding you how to use the module.<br /><br />
5) Use of the module is very familiar with MySQLdb module.<br /><br />
<br /><br />
Usage:<br />
<br />
Refers to sample program - test.py in archive zip.
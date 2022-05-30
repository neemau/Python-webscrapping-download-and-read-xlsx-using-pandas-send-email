<?php
$retval = 'C:\\Users\\mappadmin\\AppData\Local\\Programs\\Python\\Python38-32\\python.exe D:\\wamp64\\www\\projects\\taghash\\python\\posintegration\\soultree\\integration.py 2>&1';

$last_line = system ($retval);
$new_last_line = json_decode($last_line);
print($new_last_line);

?>
<form method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">

<?php include 'conn.php';
// Query the DB to see what fields are available and setup what field to serach in
echo "Search in: <select name='searchin'>";
      $querycolumns = "SELECT `COLUMN_NAME` FROM `INFORMATION_SCHEMA`.`COLUMNS` WHERE `TABLE_SCHEMA`='$dbname' AND `TABLE_NAME`='$tablename2';";
              $columnvalue = mysql_query ($querycolumns);
              while($column = mysql_fetch_array($columnvalue)){
              echo "<option value=$column[COLUMN_NAME]>$column[COLUMN_NAME]</option>";
              }
              mysql_close();
echo "</select>";
// Ask the User what value to look for in the above selected 
echo "for values that begin with: <input type='text' name='searchfor'>";
echo "<input type='submit' name='submit' value='Submit'>";
?>
</form>
<hr>

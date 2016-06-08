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

<?php
// Take values from entry form and pass into search query
if(isset($_POST['submit']))
{ 
    $searchelement = $_POST['searchin'];
    $searchvalue = $_POST['searchfor'];
echo " Results of '$searchelement' beginging with '$searchvalue'<hr>";
}
?>

<table border = 1 width = 1280>
      <tr>
            <?php include 'conn.php';
            // Setup the header row for the search results by querying the DB for the field names and add an additional column for the link based on PDFs with Access Numbers as file names
            $querycolumns = "SELECT `COLUMN_NAME` FROM `INFORMATION_SCHEMA`.`COLUMNS` WHERE `TABLE_SCHEMA`='$dbname' AND `TABLE_NAME`='$tablename';";
                  $columnvalue = mysql_query ($querycolumns);
                  while($column = mysql_fetch_assoc($columnvalue)){
                  echo "<th>$column[COLUMN_NAME]</th>";
                  $colcount[] = $column['COLUMN_NAME'];
                  }
             
             mysql_close();
            ?>
      <th>Link</th>
      </tr>




</table>

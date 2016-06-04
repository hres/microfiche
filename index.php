<form method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">

<?php include 'conn.php';
// Setup what column to look in
echo "Search in: <select name='searchin'>";
      $querycolumns = "SELECT `COLUMN_NAME` FROM `INFORMATION_SCHEMA`.`COLUMNS` WHERE `TABLE_SCHEMA`='$dbname' AND `TABLE_NAME`='$tablename';";
              $columnvalue = mysql_query ($querycolumns);
              while($column = mysql_fetch_array($columnvalue)){
              echo "<option value=$column[COLUMN_NAME]>$column[COLUMN_NAME]</option>";
              }
              mysql_close();
echo "</select>";
// Setup what value to look for
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
            // Setup the header row for the search results and add an additional column for the link based on PDFs with Access Numbers as file names
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

<?php include 'conn.php';
$result = mysql_query("SELECT * FROM $tablename WHERE $searchelement LIKE '$searchfor%'");
while ($row = mysql_fetch_array($result)) {
       echo "<tr>";
       for($i = 0, $j = count($colcount); $i < $j ; $i++) {
                        echo "<td>";
                        echo $row[$colcount[$i]];
                        echo "</td>";
           }
       echo "<td> <a href=";
       echo $pdfpath;
       echo $row[$pdfname];
       echo ".pdf>Link</a> </td>";
       echo "</tr>";
       }
mysql_close();
?>
</table>

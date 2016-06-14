<form method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">
<?php 
//replacement for include
$config = parse_ini_file('./db.ini');
//connection to the Server
$dbhandle = mysqli_connect($config['hostname'], $config['username'], $config['password'],$config['dbname']) or die("Unable to connect to Server");

// Query the DB to see what fields are available and setup what field to serach in
echo "Search in: <select name='searchin'>";
	  $querycolumns_prepare = mysqli_stmt_init($dbhandle);
	  mysqli_stmt_prepare($querycolumns_prepare, "SELECT `COLUMN_NAME` FROM `INFORMATION_SCHEMA`.`COLUMNS` WHERE `TABLE_SCHEMA`=? AND `TABLE_NAME`=?;");
 	  mysqli_stmt_bind_param($querycolumns_prepare, "ss", $config['dbname'], $config['tablename']);
	  mysqli_stmt_execute($querycolumns_prepare);
	  mysqli_stmt_bind_result($querycolumns_prepare,$col);
		
              while(mysqli_stmt_fetch($querycolumns_prepare)){
              	echo "<option value=$col>$col</option>";
              }
		//mysqli_stmt_close($querycolumns_prepare); //dont close until we reused down below
		//mysqli_close($dbhandle);//close at end of page, saves you from having to keep opening
echo "</select>";
// Ask the User what value to look for in the above selected 
echo "for values that begin with: <input type='text' name='searchfor' value='1'>";
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
            <?php 
	        // Setup the header row for the search results by querying the DB for the field names and add an additional column for the link based on PDFs with Access Numbers as file names
          //re use the above prepare just run again with different output
	  mysqli_stmt_execute($querycolumns_prepare);
          mysqli_stmt_bind_result($querycolumns_prepare,$col);

              while(mysqli_stmt_fetch($querycolumns_prepare)){
                  echo "<th>$col</th>";
                  $colcount[] = $col;
                  }
	        mysqli_stmt_close($querycolumns_prepare);
                //mysqli_close($dbhandle);//close at end of page, saves you from having to keep opening 
            ?>
      <th>Link</th>
      </tr>

<?php
	 $query = mysqli_stmt_init($dbhandle);
          mysqli_stmt_prepare($query, "SELECT * FROM ".$config['tablename']." WHERE ".$searchelement." LIKE ?;");
	  $searchvalue = "$searchvalue%";
	  $searchelement = "$searchelement";
          mysqli_stmt_bind_param($query,'s',$searchvalue);
          mysqli_stmt_execute($query);
	  //echo mysqli_stmt_num_rows($query);
	  //define which columns will be displayed, more columns needed? add them below eg: $column['newcolumn']
          mysqli_stmt_bind_result($query,$column['AccessNum'],$column['MFRCode'],$column['ClassNum'],$column['NotificationDate'],$column['ProductName'],$column['DIN'],$column['Form'],$column['Route']);
	//beacuse we are not using the mysqlnd native driver we cannot use fetch_result
	while(mysqli_stmt_fetch($query)){
		echo "<tr>";
		foreach($column as $col){
                        echo "<td>";
                        echo $col;
                        echo "</td>";

		}
	       // Output a link using the AccessNum as the file name
	       echo "<td> <a href=";
	       echo $config['pdfpath'] . rawurlencode($column[$config['pdfname']]);
	       echo ".pdf>Link</a> </td>";
	       echo "</tr>";
	}
mysqli_close($dbhandle); //close the connection
?>
</table>

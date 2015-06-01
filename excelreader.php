<html>
  <head>
    <style type="text/css">
    table {
    	border-collapse: collapse;
    }        
    td {
    	border: 1px solid black;
    	padding: 0 0.5em;
    }        
    </style>
  </head>
  <body>
    <table>
    <?php
    // include class file
    include 'Excel/reader.php';
    
    // initialize reader object
    $excel = new Spreadsheet_Excel_Reader();
    
    // read spreadsheet data
    $excel->read('book1.xls');    
    
    // iterate over spreadsheet cells and print as HTML table
    $x=1;
    while($x<=$excel->sheets[0]['numRows']) {
      echo "\t<tr>\n";
      $y=1;
      while($y<=$excel->sheets[0]['numCols']) {
        $cell = isset($excel->sheets[0]['cells'][$x][$y]) ? $excel->sheets[0]['cells'][$x][$y] : '';
        echo "\t\t<td>$cell</td>\n";  
        $y++;
      }  
      echo "\t</tr>\n";
      $x++;
    }
    ?>    
    </table>
  </body>
</html>